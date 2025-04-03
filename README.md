# autodunn

## Installation

1. Install mambaforge
2. Install a text editor
3. Install the environment

If you will be using a different network accounts to send dunns, you will also need to set up a profile for that account in Outlook.

## Usage

Record the path to the autodunn directory, which should also be the directory that contains this file. The instructions below use `path/to/directory` to refer to this directory.

In EMu:

1. Open Transactions
2. Search EMu for Transaction Type = LOAN OUTGOING and Transaction Status = OPEN
3. Report the matching records using the DMS_Autodunn report
4. Save the report to `path/to/directory`

Now that you have the loan data, you can run the autodunn script as follows:

1. Open the Miniforge Prompt
2. Navigate to the autodunn directory: `cd /path/to/autodunn`
3. Activate the autodunn environment: `conda activate autodunn`
4. Run the autodunn script: `python autodunn.py`

The first time the autodunn script is run with a new export file, it will try to update the preflight file with updated data from EMu. If changes are found, the user will be prompted to review the file and run the script again. If no changes are found, it will proceed with generating and sending dunns.

As the script runs, it produces three outputs:

- **autodunn.log** logs information about the script
- **groups** contains imports for the EMu Groups module. If imported into EMu, they can be used to view records that were processed by the autodunn script. Successful and failed dunns are recorded in separate files.
- **letters** contains HTML files with the letters produced for each transaction. When the script is run in debug mode, these can be used to review the emails before they go out.

Dunns must be recorded in EMu manually. The most accurate way to do this is to look through the sent mail on the account that was used to send the dunns. This allows you to verify that each email went out as expected and to catch bouncebacks. 

If a run is interrupted, it is generally best to update the completed dunns in EMu and re-export. The script attempts to catch these transactions based on the XML files in the groups folder, but updating EMu is the safest way to avoid accidentally sending duplicate dunning emails.

### Configuration

The configuration/config.yml file allows you to change the behavior of the script. For example, you can specify the number of dunns that are required to trigger a warning email or the date after which a loan must be recalled. 

To enable debug mode, you can set the debug key to True.

### Components

The template.htm and components.yml files in the configuration directory allow you to modify the next of the dunning emails. The contents of these files are tightly linked to the script itself, and generally these files should not be modified.

### Preflight file

The preflight file is an Excel workbook that contains basic metadata about each loan that can be used to review dunns before they are sent. Loan metadata is pulled directly from EMu, and changes to most columns will be overwritten the next time the autodunn script is run. However, two columns can be overwritten manually. Changes to these fields will be retained the next time the script it run, except for the special case noted in the description of DoNotDunn.

- **SupervisorEmail** allows you to specify the email address for a recipient's supervisor to be used when escalating. Data in this field should only be populated when needed and will not be used unless the dunn is escalated.
- **DoNotDunn** allows you to mark a loan that should not be dunned, for example, because a staff member is aware that it is already being prepped for return. The script will populate this field if there is an error with the loan record, for example, a missing email address. When the script populates this field, it uses the prefix \[AUTODUNN\]. Entries with this prefix will be overwritten by the autodunn script the next time it is run. All other entries are retained.
