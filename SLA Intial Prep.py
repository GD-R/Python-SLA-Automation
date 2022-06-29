import pandas as pd
import re
from openpyxl import load_workbook

slaReport = pd.read_excel(r"AZDOA_SLA.xlsx", sheet_name='Sheet1', engine='openpyxl', header=None)

slaReport.columns = ["Issue Key", "Project", "Subsystem", "Topic - Subtopic", "Summary", "SLA Comments",
                     "Actions Taken", "Created",
                     "Closure Date", "Status", "Severity", "Priority", "Team", "Assignee",
                     "External Ticket ID", "Reporter", "Patch Set", "Defect ID", "Patch Request ID",
                     "Contract Number"]

slaReport['PMO Notes'] = ""

slaReport = slaReport.fillna("")

slaReport.loc[(slaReport.Severity == 'Severity 1'), 'Severity'] = 'Critical'
slaReport.loc[(slaReport.Severity == 'Severity 2'), 'Severity'] = 'Serious'
slaReport.loc[(slaReport.Severity == 'Severity 3'), 'Severity'] = 'Moderate'
slaReport.loc[(slaReport.Severity == 'Severity 4'), 'Severity'] = 'Minor'

slaReport.loc[(slaReport.Priority == 'Critical'), 'Priority'] = 'Urgent'
slaReport.loc[(slaReport.Priority == 'Medium'), 'Priority'] = 'Normal'

searchFor = ['completed', 'complete', 'resolved', 'resolve', 'closed', 'closing', 'delivered', 'attached', 'implemented', 'packaged',
             'duplicate', 'successful', 'cancel', 'uploaded', 'executed', 'finished', 'close', 'processed']
slaKey = ['downtime', 'outage']

statusKeywords = ['closed', 'resolved', 'patch', 'Delivered', 'Defect', 'Shipped', 'Cloud']


def substring_after(str1, delim):
    s1 = int(str1.partition(delim)[2][3:])
    if s1 <= 21:
        return True
    else:
        return False


for i in range(len(slaReport.index)):
    if "Maintenance Window" in slaReport.at[i, 'Summary'] and "Placeholder" in slaReport.at[i, 'Actions Taken']:
        slaReport.at[i, 'PMO Notes'] = "Please Update MW minutes"
        break
    else:
        slaReport.at[i, 'PMO Notes'] = ""


for i in range(len(slaReport.index)):
    if (slaReport.at[i, 'Actions Taken'] == "" or slaReport.at[i, 'Status'] == "In Review" or (slaReport.at[i, 'Status'] == "Work In Progress" and "MS" in slaReport.at[i, 'Project'])) and (slaReport.at[i, 'Closure Date'] == ""):
        slaReport.at[i, 'PMO Notes'] += "\n Ticket needs review, update Text 2 and/or Activity Stage as per recent Status"
    else:
        slaReport.at[i, 'PMO Notes'] = ""



for i in range(len(slaReport.index)):
    if (slaReport.at[i, 'Status'] == "Closed") and (slaReport.at[i, 'Closure Date'] == ""):
        slaReport.at[i, 'PMO Notes'] = slaReport.at[
                                           i, 'PMO Notes'] + "\n Activity Stage is closed, please close the Ticket in JIRA as well."
    else:
        slaReport.at[i, 'PMO Notes'] = slaReport.at[i, 'PMO Notes'] + ""


for i in range(len(slaReport.index)):
    if (slaReport.at[i, 'Status'] == "Resolved - Pending Client Confirmation") and \
            (slaReport.at[i, 'Closure Date'] == "") and (substring_after(str(slaReport.at[i, 'Created']), "-")):
        slaReport.at[i, 'PMO Notes'] = slaReport.at[i, 'PMO Notes'] + "\n Please Update JIRA status."
    else:
        slaReport.at[i, 'PMO Notes'] = slaReport.at[i, 'PMO Notes'] + ""




for i in range(len(slaReport.index)):
    if (not (slaReport.at[i, 'Closure Date'] == "")) and not (re.compile('|'.join(statusKeywords), re.IGNORECASE).search(slaReport.at[i, 'Status'])):
        slaReport.at[i, 'PMO Notes'] = slaReport.at[
                                           i, 'PMO Notes'] + "\n Ticket is Closed/Resolved in JIRA, " \
                                                             "please update Activity Stage to Closed/RPCC and make sure closure comment is added in Text 2"
    else:
        slaReport.at[i, 'PMO Notes'] = slaReport.at[i, 'PMO Notes'] + ""


for i in range(len(slaReport.index)):
    if (slaReport.at[i, 'Status'] == "Closed" or slaReport.at[i, 'Status'] == "Completed") and not \
            (re.compile('|'.join(searchFor), re.IGNORECASE).search(slaReport.at[i, 'Actions Taken'])) \
            and not (slaReport.at[i, 'Actions Taken'] == "") and not ("Closed/Resolved" in slaReport.at[i, 'PMO Notes']):
        slaReport.at[i, 'PMO Notes'] = slaReport.at[
                                           i, 'PMO Notes'] + "\n Please add a comment that justifies the closed status in Text 2."
    else:
        slaReport.at[i, 'PMO Notes'] = slaReport.at[i, 'PMO Notes'] + ""


for i in range(len(slaReport.index)):
    if (re.compile('|'.join(slaKey), re.IGNORECASE).search(slaReport.at[i, 'Summary'])) and not \
            (re.compile('|'.join(slaKey), re.IGNORECASE).search(slaReport.at[i, 'SLA Comments'])):
        slaReport.at[i, 'PMO Notes'] = slaReport.at[
                                           i, 'PMO Notes'] + "\n Please enter the application outage/downtime minutes in Text 1."
    else:
        slaReport.at[i, 'PMO Notes'] = slaReport.at[i, 'PMO Notes'] + ""

# *****************************************************************************************************************

for i in range(len(slaReport.index)):
    if (slaReport.at[i, 'Subsystem'] == "") or (slaReport.at[i, 'Topic - Subtopic'] == "" and "MS" in slaReport.at[i, 'Project']) or \
            ((not re.search(r'-', slaReport.at[i, 'Topic - Subtopic'])) and "MS" in slaReport.at[i, 'Project']):
        slaReport.at[i, 'PMO Notes'] = slaReport.at[
                                           i, 'PMO Notes'] + "\n Please Update Subsystem/Topic - Subtopic"
    else:
        slaReport.at[i, 'PMO Notes'] = slaReport.at[i, 'PMO Notes'] + ""

for i in range(len(slaReport.index)):
    if ("Managed Services" in slaReport.at[i, 'Topic - Subtopic']) and not ("Topic" in slaReport.at[i, 'PMO Notes']):
        slaReport.at[i, 'PMO Notes'] = slaReport.at[
                                           i, 'PMO Notes'] + "\n MS ticket Please Update Select List 2 accordingly."
    else:
        slaReport.at[i, 'PMO Notes'] = slaReport.at[i, 'PMO Notes'] + ""

for i in range(len(slaReport.index)):
    if slaReport.at[i, 'Severity'] == "":
        slaReport.at[i, 'PMO Notes'] = slaReport.at[i, 'PMO Notes'] + "\n Please Update Severity"
    else:
        slaReport.at[i, 'PMO Notes'] = slaReport.at[i, 'PMO Notes'] + ""

for i in range(len(slaReport.index)):
    if ((slaReport.at[i, 'Severity'] == "Serious") or (slaReport.at[i, 'Severity'] == "Critical")) and \
            (slaReport.at[i, 'SLA Comments'] == ""):
        slaReport.at[i, 'PMO Notes'] = slaReport.at[i, 'PMO Notes'] + "\n Serious/Critical Issues requires SLA Comment"
    else:
        slaReport.at[i, 'PMO Notes'] = slaReport.at[i, 'PMO Notes'] + ""

for i in range(len(slaReport.index)):
    if slaReport.at[i, 'Assignee'] == "":
        slaReport.at[i, 'PMO Notes'] = slaReport.at[i, 'PMO Notes'] + "\n Please Update Assignee"
    else:
        slaReport.at[i, 'PMO Notes'] = slaReport.at[i, 'PMO Notes'] + ""

for i in range(len(slaReport.index)):
    if slaReport.at[i, 'Contract Number'] == "":
        slaReport.at[i, 'PMO Notes'] = slaReport.at[i, 'PMO Notes'] + "\n Please Update Contract Number"
    else:
        slaReport.at[i, 'PMO Notes'] = slaReport.at[i, 'PMO Notes'] + ""

for i in range(len(slaReport.index)):
    if ("MS" in slaReport.at[i, 'Project']) and not ("MS" in slaReport.at[i, 'Contract Number']) \
            and not (slaReport.at[i, 'Contract Number'] == ""):
        slaReport.at[i, 'PMO Notes'] = slaReport.at[
                                           i, 'PMO Notes'] + "\n Please Check Contract Number for MS Ticket."
    else:
        slaReport.at[i, 'PMO Notes'] = slaReport.at[i, 'PMO Notes'] + ""

for i in range(len(slaReport.index)):
    if not re.search(r'/', slaReport.at[i, 'Actions Taken']) and not (slaReport.at[i, 'Actions Taken'] == ""):
        slaReport.at[i, 'PMO Notes'] = slaReport.at[i, 'PMO Notes'] + "\n Please add the Date in Text 2"
    else:
        slaReport.at[i, 'PMO Notes'] = slaReport.at[i, 'PMO Notes'] + ""


slaReport['PMO Notes'] = slaReport['PMO Notes'].str.lstrip()

for i in range(len(slaReport.index)):
    if "\n" in slaReport.at[i, 'PMO Notes']:
        lines = slaReport.at[i, 'PMO Notes'].split("\n")
        for j in range(len(lines)):
            lines[j] = str(j + 1) + ") " + lines[j]
        slaReport.at[i, 'PMO Notes'] = "\n".join(lines)
    else:
        slaReport.at[i, 'PMO Notes'] = slaReport.at[i, 'PMO Notes']


with pd.ExcelWriter(r"AZDOA_SLA.xlsx", engine="openpyxl", mode="a",
                    if_sheet_exists="replace") as writer:
    slaReport.to_excel(writer, 'SLA', index=False)


