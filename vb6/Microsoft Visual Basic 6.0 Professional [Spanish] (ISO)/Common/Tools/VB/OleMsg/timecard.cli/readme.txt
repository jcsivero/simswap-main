Time Card Sample

This sample application illustrates using MAPI OLE Messaging in VB
applications.

OLE Messaging features illustrated:
+  creating a MAPI session object and logging on
+  creating new messages
+  using standard message properties: subject, body, recipients,
   attachments
+  sending and receiving both IPM and IPC messages
+  using custom fields and multi-valued fields
+  browsing through messages in a folder and through folders in
   personal/public folders
+  using Session.AddressBook dialog
+  saving address book entries between sessions

Installation:
In order to use the sample, first compile the client executable 
(tmcli.exe) in the client directory.
Next, in file server\server.bas set the value of constant
ClientExePath to the full path to the client executable (tmcli.exe).
Finally, recompile the tmserv.exe.

Terminology:
server app (server) - a program that keeps a list of all the users,
sends out time report requests to each user on the list, and produces
a summary report for all the users for a given pay period.
client app (client) - a program that scans the inbox for a request
from the server. If a request is found, it displays and prompts the
user to fill out  a time report form.
request message (request)- an IPM message sent from the server app to
the users. This message contains client.exe as an attachment and
report categories, number of report categories, pay period saved in
the message's named properties.
report message (report) - an IPC message sent from a client app to the
server. This message contains all the information that request does
plus number of hours user worked divided by  report categories and
days of the week saved in the message's named properties.
pay period - pay period for this sample is a week (Sunday through
Saturday)  that is identified by the date of its Friday.

Description:
The sample performs collecting of time information for hourly
employees for payroll purposes. It consists of two standalone parts:
a server (tmserv.exe) and a client (tmcli.exe).

Server:
The main window of the server is divided into two parts: a list of
users and a list of report categories. Each part has "Add" and
"Remove" buttons to manipulate the contents of the corresponding list.
The File menu Save command saves both user and category lists to files
so that next time you start up the server they are there.
The Report menu "Send Requests" command first asks list owner for a
pay period and then sends requests to all the users on the list.
The Report menu "Generate Report" command first asks user for a pay
period and then generates a report for the given pay period.
The Report menu "Clean Up" command deletes all processed messages of
the report message class from the topmost folder of the default
message store.

Client:
When launched, the client app searches its inbox for a request
message. Then it displays a form based on the data in the message.
When user is done filling out the form, the information is sent to the
server in a report message.

How it works:
The request message contains three named properties:
+  "NumReportCategories" - number of report categories (integer)
+  "ReportCategories" - report categories (multi-valued string)
+  "PayPeriod"  - pay period (date)
The report message contains all of the request message named
properties plus one multi-valued property (array of doubles) for every
report category. Names of these properties are "ReportedTimek", k = 1,
2, ... NumberOfCategories. For example, the name of the data property
corresponding to the first report category, which name is the first
element of the array stored in the report categories property, is
"ReportedTime1".


Issues to consider:
Following is a list of some of questions that everybody who writes an
application similar to this one has to face. Some of them are solved
in this sample, for others a possible solution is proposed.

1. If there is more than one request message in user's inbox, the server
app can display a listbox showing pay-period and time received for
each message and let user choose the one that he wants to use.
2. If a user unintentionally sends more than one message for the same pay
period, the client app can remove the request message after it is used,
but then users will not be able to resubmit report in case of an
error.
3. Suppose a user made a mistake in his report and wants to resubmit it.
The server can use the latest report message from the user for a given
report period.. Whoever runs the server app has to make sure that  he
waits until all the report messages are received or regenerates the
report after new submissions. This would also be a solution for 2. On
the client side: after reading a request message, client can scan the
SentMail folder and if it finds a message for the requested pay period,
it intializes the form with the data from the message.
4. If some report messages are lost either in transmission or due to
corruption of the store on the server side, "Remind" button on the
Report dialog can be used to send out new requests to the users whose
reports were lost. If the solution for 3. is implemented then the only
thing the user will have to do after launching the form is to hit the
"Send" button.
5. If a report from somebody not on the user list is received, the server
can either ignore it or give the list owner an option to add the user
to the list.
6. Validating pay period. In this sample the validation is performed in
frmCalendar.ValidatePayPeriod function.  To change the definition of
the pay period modify the logic of this function.

