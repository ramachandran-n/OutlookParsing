PREREQUISITE:

This utility was built using the .NET Framework 4.5.2. Hence the client running the tool must have .NET Framework 4.5.2 or greater installed.

This is helps to parse the mails from the outlook and fetch the attachments in a specified folder & extracts the sender, subject, mail content (body), received time and a flag that speficies whether the mail has attachments or not. (This flag just states whether the mail contains attachments and NOT whether multiple attachments are present. Also, if the mail has pictures / any media as part of the mail body the same is considered as an attachment and extracted)

The utility uses the Interop for Outlook and Excel applications and DOESN'T INCLUDE any 3rd party DLLs thus ensuring the safety and adhere the compliance and security policies of any organization. The Interop DLLs are installed as part of the .NET and the Office installations and doesn't need any manual installation unless specifically required (business cases may vary from user to user)

The configuration files contains the location and mailbox details which needs to be updated before the utility is run. Please note that the download location folder and the excel files must be created in prior to the execution of the tool and the path provided must match the path of the files/ folders.

    <add key ="SharedMailBox" value="Mailbox / Inbox address"/> - This needs to be defined as something@something.com
    
    <add key ="LogLocation" value="Environment.SpecialFolder.MyDocuments"/> - Logging will be introduced in the next release
    
    <add key ="DownloadLocation" value="The folder in which the attachments needs to be extracted"/> - Please note tha the folder must be created prior to the run of the utility.
    
    <add key="ExcelPath" value="The Excel path where the mail content to be written"/> - Please note that the excel has to be in .xls format and needs to be created prior to the run of the utility.
    
    <add key="IsSharedMailbox" value=""/> - Place 'Y' when you are parsing the shared mailbox and 'N' if that's a INBOX.
