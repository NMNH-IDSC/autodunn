"""Sends dunning emails based on a report from enmnhtransactions"""

import logging
from pathlib import Path
from pprint import pprint

import pandas as pd

from nmnh_ms_tools.records.transactions import Transaction, create_transaction
from nmnh_ms_tools.utils import prompt
from xmu import EMuReader, EMuRecord, write_group

from config.dunns import Dunn, prep_loans, save_preflight


# Set up logger to provide detailed info
logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=getattr(logging, "INFO"),
    handlers=[
        logging.FileHandler("autodunn.log", "a", encoding="utf-8"),
    ],
)

if __name__ == "__main__":

    logging.info("===================")
    logging.info("Running autodunn.py")
    logging.info(f"Config: {Dunn.trn_config}")

    # Set paths for groups
    groups = Path("groups")
    if Dunn.trn_config["debug"]:
        grp_dunned = groups / "dunn_succeeded_debug.xml"
        grp_skipped = groups / "dunn_failed_debug.xml"
    else:
        grp_dunned = groups / "dunn_succeeded.xml"
        grp_skipped = groups / "dunn_failed.xml"

    # Ensure that required directories exist
    groups.mkdir(parents=True, exist_ok=True)
    Path("letters").mkdir(parents=True, exist_ok=True)

    # Check transactions that have already been handled since the last export.
    # If group files are older than the EMu export, get rid of them.
    dunned = []
    skipped = []
    for path, lst in [(grp_dunned, dunned), (grp_skipped, skipped)]:
        try:
            reader = EMuReader(path, rec_class=EMuRecord)
        except FileNotFoundError:
            pass
        else:
            if Path("xmldata.xml").stat().st_mtime > path.stat().st_mtime:
                path.unlink()
            else:
                lst.extend(
                    [
                        EMuRecord({"irn": r["Keys_tab"][0]}, module="enmnhtransactions")
                        for r in reader
                    ]
                )
    irns = {r["irn"] for r in dunned + skipped}

    # Iterate through the export file to get item data for each transaction
    reader = EMuReader("xmldata.xml")
    transactions = {}
    for rec in reader:
        transactions[int(rec["TraNumber"])] = create_transaction(rec)
        if (
            Transaction.trn_config["debug_num"]
            and rec["TraNumber"] != Transaction.trn_config["debug_num"]
        ):
            continue
        if Transaction.trn_config["debug_num"]:
            pprint(rec)
        reader.report_progress()

    # Remove closed transactions and all associated metadata from preflight
    if Transaction.trn_config["remove_closed_transactions"]:
        try:
            df = pd.read_excel("preflight.xlsx")
        except FileNotFoundError:
            pass
        else:
            df["DueDate"] = df["DueDate"].dt.date
            df["LastInteraction"] = df["LastInteraction"].dt.date
            active = df[df["TransactionNumber"].isin(transactions)]
            active.fillna("").to_excel(
                "preflight.xlsx", sheet_name="Loans", index=False, freeze_panes=(1, 0)
            )

    # Prepare loans
    loans = prep_loans(transactions)

    # Warn user when preparing to send emails
    if not Dunn.trn_config["debug"]:
        resp = prompt(
            "***The script will send out actual dunning emails to actual"
            " people! Are you sure you want to continue?***",
            {"y": True, "n": False},
        )
        if resp and not Dunn.trn_config["safe_send"]:
            resp = prompt(
                "***You have disabled the safe send option! This is your"
                " last chance to bail before sending a dunning letter to"
                " everyone with an overdue loan. Are you sure you want to"
                " continue?***",
                {"y": True, "n": False},
            )
        if not resp:
            raise RuntimeError("User chose not to proceed")

    # Filter and sort loans
    loans = [t for t in loans if t.is_open() and t.contact]
    loans = sorted(loans, key=lambda t: t.contact.name)

    # Dunn loans and find errors
    try:
        for loan in loans:
            tranum = loan["TraNumber"]
            # Check for debug number
            if Dunn.trn_config["debug_num"] and tranum != Dunn.trn_config["debug_num"]:
                continue
            # Filter out loans that have been dunned since the last export
            if loan["irn"] in irns:
                msg = f"{loan['TraNumber']}: Dunn already processed"
                logging.info(msg)
                print(msg)
                continue
            # Dunn overdue loans
            if loan.is_overdue() or loan.is_almost_due():
                rec = EMuRecord({"irn": loan["irn"]}, module="enmnhtransactions")
                try:
                    if not loan.dunn(send=not Dunn.trn_config["debug"]):
                        raise Exception("Dunn failed")
                except Exception as e:
                    msg = f"{loan['TraNumber']}: Dunn failed"
                    logging.exception(msg)
                    print(msg)
                    skipped.append(rec)
                else:
                    msg = f"{loan['TraNumber']}: Dunn succeeded"
                    logging.info(msg)
                    print(msg)
                    dunned.append(rec)
        print("Done!")
    except:
        raise
    finally:
        # Write imports into egroups for successful and failed dunns. Note that
        # this does not take into account whether the dunning email went through.
        if dunned:
            write_group(dunned, grp_dunned, name="DMS_DunnSucceeded")
        if skipped:
            write_group(skipped, grp_skipped, name="DMS_DunnFailed")
        save_preflight(loan.preflight, "preflight.xlsx", False)
