import logging
import os
import re
import time
import warnings
import webbrowser as wb
from datetime import datetime, timedelta
from pathlib import Path

import inflect
import pandas as pd
import yaml

try:
    import win32com.client as win32
except ModuleNotFoundError:
    print("Cannot send dunning letters (win32 module not installed)")

from nmnh_ms_tools.records.transactions import LoanOutgoing, Transaction


AUTODUNN_CODES = [
    "[AUTODUNN] Collection excluded",
    "[AUTODUNN] Contains errors",
    "[AUTODUNN] Not due yet",
    "[AUTODUNN] Not in export",
    "[AUTODUNN] Recent interaction",
]
CONFIG_DIR = Path(__file__).parent


class Dunn(LoanOutgoing):
    """Container for transactions to dunn"""

    with open(CONFIG_DIR / "template.htm", "r", encoding="utf-8") as f:
        template = f.read()

    with open(CONFIG_DIR / "components.yml", "r", encoding="utf-8") as f:
        components = yaml.safe_load(f)

    trn_config = Transaction.trn_config
    outlook = None
    preflight = None
    _supervisors = {}

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

    def dunn(self, send=False):
        """Verifies the loan is dunnable and sends the dunning letter"""

        try:
            preflight = self.preflight[
                self.preflight["TransactionNumber"] == self["TraNumber"]
            ].iloc[0]
        except IndexError:
            logging.warning("{}: Not found in preflight".format(self["TraNumber"]))
            return False

        if not is_empty(preflight["DoNotDunn"]):
            logging.warning(
                "{}: Do not dunn ({})".format(self["TraNumber"], preflight["DoNotDunn"])
            )
            return False

        # Check if dunnable
        errors = self.find_errors()
        if errors:
            logging.warning("\n".join(errors))
            return False

        return_date = datetime.now() + timedelta(days=30)
        mailing_address = (
            self.trn_config["mailing_address"]
            .format(**self.coll_contact)
            .replace("\n", "<br>")
        )
        shipping_address = (
            self.trn_config["shipping_address"]
            .format(**self.coll_contact)
            .replace("\n", "<br>")
        )

        # Compile basic info about this transaction
        eng = inflect.engine()
        dunn_info = {
            "tranum": self["TraNumber"],
            "greeting": _greeting(self.contact),
            "name": self.contact.name,
            "org": self.org.name if self.org else "",
            "recipient": self.recipient,
            "due_date": self.due_date.strftime("%d %b %Y"),
            "return_date": return_date.strftime("%d %b %Y"),
            "num_dunns": self.num_dunns,  # only used to decide about showing a warning
            "nth": eng.number_to_words(eng.ordinal(self.num_dunns + 1)),
            "orig_contact": self.orig_contact.name,
            "org_change": "",
            "kind": "Recalled" if self.level == "recall" else "Overdue",
            "coll_mailing_address": mailing_address,
            "coll_shipping_address": shipping_address,
        }

        # Note a change of affiliation by the original contact if the original
        # loan was made to an organization
        if dunn_info["org"] and dunn_info["org"] != self.contact.affiliation:
            dunn_info["org_change"] = " " + self.get_component(
                "org_change", **dunn_info
            )

        # Add collections contact info
        dunn_info.update({f"coll_{k}": v for k, v in self.coll_contact.items()})

        # Escalate if previous dunning letters have been ignored
        supervisor = None
        if self.escalate():
            supervisor = self.get_supervisor(preflight, dunn_info)

        # Customize intro based on whether this is a reminder
        intro_key = "intro_due" if self.is_overdue() else "intro_reminder"

        # Update the mask with the components defined in the config file
        components = {
            "greeting": self.get_component("greeting", **dunn_info),
            "intro": self.get_component(intro_key, **dunn_info),
            "summary": self.summarize(),
            "escalation": self.get_component("escalate", **dunn_info),
            "action": self.get_component("action", **dunn_info),
            "data_return": self.get_component("data_return", **dunn_info),
            "sender": self.trn_config["dunner"]["name"],
        }
        components.update(
            {f"coll_{k}": v for k, v in self.trn_config["dunner"].items()}
        )

        # Get the recipient"s email. Adjust wording if not the original contact.
        if self.contact.is_deceased() or (
            self.orig_contact and self.orig_contact.is_deceased()
        ):
            components["intro"] = self.get_component(
                intro_key, "deceased_contact", **dunn_info
            )
        elif self.contact["NamLast"] != self.orig_contact["NamLast"]:
            components["intro"] = self.get_component(
                intro_key, "new_contact", **dunn_info
            )

        # Construct the email from the components
        body = self.template.format(**components).strip()
        body = re.sub(r"<br />", "<br>", body)
        body = re.sub(r"((?:</(?:blockquote|h\d|li|p|table|ul)>\s*)+)", r"\1<br>", body)

        if self.level == "recall":
            body = body.replace("must be renewed or returned", "has been recalled")

        subject = "{kind} loan from the Smithsonian: {tranum}".format(**dunn_info)
        if self.trn_config["debug"]:
            subject += " [DEBUG]"

        fp = os.path.join("letters", f"{self['TraNumber']}_{self.level}.htm")
        with open(fp, "w", encoding="utf-8", newline="") as f:
            # Add subject and recipients to the HTML preview
            metadata = [
                "<span class='metadata'>Subject:</span> " + subject,
                "<span class='metadata'>To:</span> " + self.contact.email,
            ]
            if supervisor:
                cc = "; ".join((self.contact.email, dunn_info["coll_email"]))
                # Flip the cc/to emails if escalating
                metadata[1] = "<span class='metadata'>To:</span> " + supervisor
                metadata.append("<span class='metadata'>Cc:</span> " + cc)
            else:
                cc = dunn_info["coll_email"]
                metadata.append("<span class='metadata'>Cc:</span> " + cc)
            recipients = "<body>\n<p>" + "<br>".join(metadata) + "</p><hr />"
            f.write(body.replace("<body>", recipients))

        sent = False
        if send or self.trn_config["send_to_me"]:
            if self.trn_config["safe_send"]:
                # Preview the dunning email if using safe send
                wb.open(fp)
            elif self.trn_config["debug"] and not self.trn_config["send_to_me"]:
                # Last chance to trap errors before you actually send an email
                raise Exception("Trying to send email while in debug mode")

            self.send(
                subject,
                body,
                self.contact.email,
                dunn_info["coll_email"],
                supervisor,
            )
            if not self.trn_config["safe_send"]:
                time.sleep(1)

            sent = True

        # Group XML files used to track recent dunns instead
        # if sent:
        #    cond = self.preflight["TransactionNumber"] == self["TraNumber"]
        #    self.preflight.loc[cond, "LastInteraction"] = datetime.now()
        #    save_preflight(self.preflight, "preflight.xlsx", False)

        return sent if (send or self.trn_config["send_to_me"]) else True

    def to_preflight(self):
        """Maps basic metadata to the fields used in the preflight file"""
        return {
            "TransactionNumber": self["TraNumber"],
            "Catalog": self.catalog,
            "DueDate": self.due_date.value if self.due_date else "",
            "Level": self.level.title(),
            "Contact": str(self.contact),
            "Organization": str(self.org) if self.org else "",
            "SupervisorEmail": "",
            "DunnCount": self.num_dunns,
            "LastInteraction": self.last_interaction.value,
            "DoNotDunn": "",
            "Errors": "; ".join(self.find_errors()),
            "Notes": "",
        }

    def get_component(self, key, level=None, **kwargs):
        """Returns the formatted string for part of the dunning email"""

        if level is None:
            level = self.level

        # HACK: Reminder is a useful label but the only change is managed through
        # the label for the into. Process as default.
        if level == "reminder":
            level = "default"

        # HACK: Show the too-many-dunns warning even when recalling the loan
        if level == "recall" and key == "escalate":
            if self.escalate():
                level = "escalate"
            elif self.warn():
                level = "warn"

        try:
            return self.components[level][key].format(**kwargs)
        except KeyError as exc:
            if str(exc) != repr(key):
                raise
            try:
                return self.components["default"][key].format(**kwargs)
            except KeyError:
                # Only escalation can go undefined
                if key == "escalate":
                    return ""
                raise

    def get_supervisor(self, preflight, dunn_info):
        """Determines supervisor of contact for a loan with too many dunns"""
        supervisor = preflight["SupervisorEmail"]
        key = "{name} ({org})".format(**dunn_info)
        if is_empty(supervisor):
            try:
                supervisor = self._supervisors[key]
            except KeyError:
                print(
                    "This is the {nth} dunning letter for"
                    " {tranum}!".format(**dunn_info)
                )
                while True:
                    supervisor = input(
                        "Escalate contact for {name}" " ({org}): ".format(**dunn_info)
                    )
                    if supervisor:
                        break
        self._supervisors[key] = supervisor
        return supervisor

    def find_errors(self):
        """Verifies loan has enough info to autodunn"""
        errors = []
        tranum = self["TraNumber"]
        contact = self.contact

        if not contact:
            errors.append("{}: No contact provided".format(tranum))
        elif self.contact.is_deceased():
            errors.append("{}: Contact is deceased".format(tranum))
            return errors
        elif not self.contact.email:
            errors.append("{}: No email address".format(tranum))
        elif self.contact.email.count("@") != 1 or " " in self.contact.email:
            errors.append("{}: Bad email address".format(tranum))
        elif (
            self.contact.is_person()
            and not self.contact.get("NamTitle")
            and not self.contact.get("NamFirst")
        ):
            errors.append("{}: No title or first name".format(tranum))

        if not self.due_date:
            errors.append("{}: No due date".format(tranum))

        if not self.open_date:
            errors.append("{}: No open date".format(tranum))

        if not [i for i in self.tr_items if i.is_outstanding()]:
            errors.append("{}: No outstanding items".format(tranum))

        return errors

    def summarize(self):
        """Creates a high-level summary of the transaction"""
        tranum = self["TraNumber"]
        name = self.orig_contact.name
        org = self.org.name if self.org else ""
        try:
            start_date = " on " + self.open_date.strftime("%d %b %Y")
        except ValueError:
            start_date = ""
        n = len(list(self.tr_items))
        s = "s" if n != 1 else ""
        summary = ["<h2>Transaction {}</h2>".format(tranum)]
        summary.append(
            "<p>The National Museum of Natural History loaned {}"
            " item{} to {} on behalf of {}{}. The following"
            " objects are overdue:</p>".format(n, s, org, name, start_date)
        )
        summary[-1] = summary[-1].replace("to  on behalf of", "to")
        summary.append(self.item_table())
        return "".join(summary)

    def item_table(self):
        """Constructs a table with data about transaction items"""
        table = [
            "<table>\n"
            "<tr>"
            "<th>Catalog number</th>"
            "<th>Object</th>"
            "<th>Type</th>"
            "<th>Description</th>"
            "<th># outstanding</th>"
            "</tr>\n"
        ]
        items = list(self.tr_items)
        items.sort(key=lambda d: (d["ItmObjectName"], d["ItmCatalogueNumber"]))
        for item in items:
            if item["ItmObjectCountOutstanding"]:
                table.append(
                    "<tr>"
                    "<td>{ItmCatalogueNumber}</td>"
                    "<td>{ItmObjectName}</td>"
                    "<td>{ItmPreparation}</td>"
                    "<td>{ItmDescription}</td>"
                    "<td>{ItmObjectCountOutstanding}/{ItmObjectCount}</td>"
                    "</tr>\n".format(**item)
                )
        table.append("</table>")
        return "".join(table)

    def send(self, subject, body, recipient, coll_email, supervisor=None):
        """Sends the dunning email using the current user"s Outlook"""
        # Initiate Outlook if either not in debug mode or send to me is enabled.
        # Send criteria should really be settled by now, so this is likely redundant.
        if self.outlook is None and (
            not self.trn_config["debug"] or self.trn_config["send_to_me"]
        ):
            self.__class__.outlook = win32.Dispatch("Outlook.Application")
        # Create mail item
        mail = self.outlook.CreateItem(0)
        # Set to send from the specified Outlook account
        sender = self.trn_config["dunner"]["email"]
        account = None
        for account in self.outlook.Session.Accounts:
            if account.SmtpAddress.lower() == sender.lower():
                # From https://stackoverflow.com/questions/35908212
                mail._oleobj_.Invoke(*(64209, 0, 8, 0, account))
                break
        else:
            raise ValueError("Cannot send from {}".format(sender))
        # Set the to and cc fields
        if self.trn_config["send_to_me"]:
            mail.To = self.trn_config["dunner"]["email"]
            cc = [self.trn_config["dunner"]["email"]]
            if supervisor:
                cc.append(self.trn_config["dunner"]["email"])
            mail.CC = "; ".join(cc)
        elif supervisor:
            mail.To = supervisor
            mail.CC = "; ".join([recipient, coll_email])
        else:
            mail.To = recipient
            mail.CC = coll_email
        mail.Subject = subject
        mail.HTMLBody = body
        if self.trn_config["safe_send"]:
            msg = "Send dunning email to " + mail.To
            if cc:
                msg += " (cc: " + mail.CC + ")"
            msg += "? Press ENTER to send or CTRL+C to quit."
            input(msg)
        mail.Send()


def prep_loans(transactions, fp="preflight.csv"):
    """Reads transction metadata from the preflight file"""

    loans = [Dunn(t) for t in transactions.values() if isinstance(t, LoanOutgoing)]

    # Get basic metadata from the list of loans
    rows = []
    for loan in loans:

        if loan.is_open():
            row = loan.to_preflight()

            # Set DoNotDunn code
            if row["Errors"]:
                row["DoNotDunn"] = "[AUTODUNN] Contains errors"
            elif loan.catalog in loan.trn_config["exclude_codes"]:
                row["DoNotDunn"] = "[AUTODUNN] Collection excluded"
            elif not loan.is_overdue() and not loan.is_almost_due():
                row["DoNotDunn"] = "[AUTODUNN] Not due yet"

            rows.append(row)

    if not rows:
        raise ValueError("No loans found!")

    # Read the preflight file if it already exists
    preflight_new = pd.DataFrame(rows)
    preflight_new["DueDate"] = pd.to_datetime(preflight_new["DueDate"])
    preflight_new["LastInteraction"] = pd.to_datetime(preflight_new["LastInteraction"])
    preflight_new = preflight_new.fillna("")

    try:
        preflight_old = pd.read_excel("preflight.xlsx")
    except FileNotFoundError:
        save_preflight(preflight_new, "preflight.xlsx")
    else:
        cols = preflight_new.columns
        preflight = pd.merge(
            preflight_new,
            preflight_old,
            how="outer",
            on="TransactionNumber",
            suffixes=("", "_old"),
        )

        preflight["DunnCount"] = preflight["DunnCount"].fillna(0).astype(int)

        # Migrate supervisor email
        preflight["SupervisorEmail"] = preflight["SupervisorEmail_old"]

        # Migrate more recent interactions
        cond = preflight["LastInteraction_old"] > preflight["LastInteraction"]
        preflight.loc[cond, "LastInteraction"] = preflight.loc[
            cond, "LastInteraction_old"
        ]

        # Migrate DoNotDunn
        cond = ~(
            pd.isna(preflight["DoNotDunn_old"])
            | (preflight["DoNotDunn_old"].isin(AUTODUNN_CODES))
        )
        preflight.loc[cond, "DoNotDunn"] = preflight.loc[cond, "DoNotDunn_old"]

        # Note recent interactions
        cond = pd.isna(preflight["DoNotDunn"]) & (
            (datetime.now() - preflight["LastInteraction"])
            < timedelta(days=loans[0].trn_config["num_days"])
        )
        preflight.loc[cond, "DoNotDunn"] = "[AUTODUNN] Recent interaction"

        # Migrate and flag rows that do not appear in current export
        cond = pd.isna(preflight["Catalog"])
        for col in cols:
            warnings.filterwarnings("error")
            if col != "TransactionNumber":
                try:
                    preflight.loc[cond, col] = preflight.loc[cond, f"{col}_old"]
                except Exception as exc:
                    raise ValueError(col)

        cond = cond & (pd.isna(preflight["DoNotDunn"]))
        preflight.loc[cond, "DoNotDunn"] = "[AUTODUNN] Not in export"

        # Remove old columns and sort by transaction number
        preflight = (
            preflight[cols]
            .sort_values("TransactionNumber", ascending=False)
            .reset_index(drop=True)
        )
        new = preflight.fillna("").to_dict("records")
        old = preflight_old.fillna("").to_dict("records")
        if new != old:

            print("Found differences! Updating preflight file...")

            new = {r["TransactionNumber"]: r for r in new}
            old = {r["TransactionNumber"]: r for r in old}

            new_only = sorted(set(new) - set(old))
            if new_only:
                print(f"Records added to preflight: {sorted(new_only)}")

            old_only = sorted(set(old) - set(new))
            if old_only:
                print(f"Record removed from preflight: {sorted(old_only)}")

            for tranum in set(new) & set(old):
                new_rec = new[tranum]
                old_rec = old[tranum]
                for key in sorted(set(new_rec) & set(old_rec)):
                    new_val = new_rec.get(key)
                    old_val = old_rec.get(key)
                    if new_val != old_val:
                        print(
                            f"{tranum}: {key} changed from {repr(old_val)} to {repr(new_val)}"
                        )

            preflight = preflight.fillna("").replace(r"^None$", "")
            save_preflight(preflight, "preflight.xlsx")

    Dunn.preflight = preflight
    return loans


def save_preflight(df, path, exit_on_change=True):
    df["DueDate"] = df["DueDate"].dt.date
    df["LastInteraction"] = df["LastInteraction"].dt.date
    df = df.sort_values("TransactionNumber", ascending=False)
    while True:
        try:
            df.fillna("").to_excel(
                path, sheet_name="Loans", index=False, freeze_panes=(1, 0)
            )
            break
        except PermissionError:
            input(
                f"Could not save {path}! Please close the file and hit ENTER to try again"
            )
    if exit_on_change:
        print(
            "Updated preflight file! Review preflight.xlsx and re-run this notebook to send dunns."
        )
        import sys

        sys.exit()


def _greeting(contact) -> str:
    """Determines the proper greeting for a dunning letter"""
    if contact.is_person():
        title = (
            contact["NamTitle"] if contact["NamTitle"] else contact["NamFirst"][0]
        ).rstrip(".")
        return f"Dear {title}. {contact['NamLast']}:"
    return "To whom it may concern:"


def is_empty(val):
    return not val or pd.isna(val)
