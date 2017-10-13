This is an utility that helps to parse the mails from the outlook and fetch the attachments in a specified folder & extracts the sender, subject, mail content (body), received time and a flag that speficies whether the mail has attachments or not.

The configuration files contains the location and mailbox details which needs to be updated before the utility is run.

    <add key ="SharedMailBox" value="Mailbox / Inbox address"/> - This needs to be defined as something@something.com
    
    <add key ="LogLocation" value="Environment.SpecialFolder.MyDocuments"/> - Logging will be introduced in the next release
    
    <add key ="DownloadLocation" value="The folder in which the attachments needs to be extracted"/> - Please note tha the folder must be created prior to the run of the utility.
    
    <add key="ExcelPath" value="The Excel path where the mail content to be written"/> - Please note that the excel has to be in .xls format and needs to be created prior to the run of the utility.
    
    <add key="IsSharedMailbox" value=""/> - Place 'Y' when you are parsing the shared mailbox and 'N' if that's a INBOX.
