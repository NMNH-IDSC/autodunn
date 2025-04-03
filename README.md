# autodunn

## Installation

First install the following software:

- [mambaforge](https://github.com/conda-forge/miniforge)
- [git](https://git-scm.com/downloads)
- [VS Code](https://code.visualstudio.com/download)

We'll use git to download autodunn. Open the Miniforge Prompt. By default, Miniforge Prompt opens in your home directory. If you want to download the script files to a different locations, use the `cd` command to change the directory.

```
cd /path/to/directory
git clone https://github.com/NMNH-IDSC/autodunn
```

Next we'll use mamba to set up the environment:

```
cd /path/to/directory
mamba create -f environment.yml
```

If you will be using a different network account to send dunns, you will also need to set up a profile for that account in Outlook.

## Usage

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
- **letters** contains HTML files with the letters produced for each transaction. These are generated when the script is run in debug mode and can be used to review the emails before they go out.

**Dunns must be recorded in EMu manually.** The most accurate way to do this is to look through the sent mail on the account that was used to send the dunns. This allows you to verify that each email went out as expected and to catch bouncebacks. 

If a run is interrupted, it is safest to update the completed dunns in EMu and re-export before sending additional dunns. The script attempts to catch these transactions based on the XML files in the groups folder, but updating EMu is the best way to avoid accidentally sending duplicate dunning emails.

### Configuration

The configuration/config.yml file allows you to change the behavior of the script. For example, you can specify the number of dunns that are required to trigger a warning email or the date after which a loan must be recalled. 

To enable debug mode, you can set the debug key to True.

### Components

The template.htm and components.yml files in the configuration directory allow you to modify the next of the dunning emails. The contents of these files are tightly linked to the script itself, and generally these files should not be modified.

### Preflight file

The preflight file is an Excel workbook that contains basic metadata about each loan that can be used to review dunns before they are sent. Loan metadata is pulled directly from EMu, and changes to most columns will be overwritten the next time the autodunn script is run. However, two columns can be overwritten manually:

- **SupervisorEmail** allows you to specify the email address for a recipient's supervisor to be used when escalating. Data in this field should only be populated when needed and will not be used unless the dunn is escalated.
- **DoNotDunn** allows you to mark a loan that should not be dunned, for example, because a staff member is aware that it is already being prepped for return. The script will populate this field if there is an error with the loan record, for example, a missing email address. When the script populates this field, it uses the prefix \[AUTODUNN\]. Entries with this prefix will be overwritten by the autodunn script the next time it is run. All other entries are retained.

Changes to these fields will be retained the next time the script it run, except for the special case noted for DoNotDunn above.
