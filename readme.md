##How-To##

* This is a really rough script to export rules from outlook and create filters in gmail. Right now you have to manually customise the export script with the name of the mailbox you want to extract rules from before you run it, I will update the script shortly to allow a user select the desired mailbox.

* Tested under Win 7 & Outlook 2010
* Add Developer tab to Outlook ribbon
** File -> Options -> Customize Ribbon Check the "Developer" checkbox
* Enable macros in Outlook 2010
** On Developer tab click "Macro Security" and select "Notifications for all Macros"
* Customise the export script
** Replace "INSERT_ACCOUNT_NAME_HERE" with the name of the mailbox containing the rules you want to export
* Create & Run export macro. Note that not all types of rules will be exported
* Pull the `rules.csv` on the root level of your hard drive into a text editor and format the rule folder redirects into gmail-like labels
* Run `convert.rb > ~/Desktop/gmail_rules.xml`
* Enable gmail filters import in gmail
** Setting -> Labs enable "Filter import/export" and click save changes
* Add the gmail labs rules 
** Settings -> Filters click "Import Filters" and import `gmail_rules.xml`

##Resources##

* <http://www.outlookcode.com/codedetail.aspx?id=1266>
* <http://easyvsto.wordpress.com/category/vsto/office12/>
* <http://www.add-in-express.com/forum/read.php?FID=5&TID=6942>
* <http://zo-d.com/blog/archives/programming/vba-writing-to-a-text-file-ms-project-excel.html>
* <http://www.dimastr.com/outspy/home.htm>
* <http://www.outlooksharp.de/blog/programming-outlook-sender-or-receiver-smtp-address>
* <http://technicalreflections.wordpress.com/2011/08/03/common-outlook-tasks-via-vsto/>

###Note
* Updated origional scripts to work on Win 7 with Outlook 2010
* There is a lot of stuff in the export script thats not executed to do with differnt types of rules, as it is it only exports simple move to folder type rules
* Importing existing filters into gmail is handled gracefully, no duplicate filters are created