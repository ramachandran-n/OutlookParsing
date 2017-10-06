using System.Collections.Generic;
using System.Configuration;
using Microsoft.Office.Interop.Outlook;
using System.IO;
using System.Data;
using Excel = Microsoft.Office.Interop.Excel;
using Logging;


namespace OutlookEmailParsing
{
    /// <summary>
    /// Created by Ramachandran Narayanan - 04 October 2017
    /// This class is used to parse through the unread mails from the outlook application and write the contents to an excel and extract attachments to a folder
    /// </summary>
    public class ExchangeRepository : IMailParsing
    {
        public static DataTable tempTable = new DataTable();
        string textToWrite = string.Empty;
        Microsoft.Office.Interop.Outlook.Application app = null;
        Microsoft.Office.Interop.Outlook._NameSpace ns = null;
        Microsoft.Office.Interop.Outlook.MAPIFolder inboxFolder = null;
        List<MailItem> mails = null;

        string downloadPath = ConfigurationManager.AppSettings["DownloadLocation"].ToString();

        public void AddToTable(List<MailItem> mails)
        {
            foreach (MailItem m in mails)
            {
                string mailSubject = m.Subject;
                string mailSender = m.Sender.Name.ToString();
                string mailBody = m.Body;
                string receivedTime = m.ReceivedTime.ToShortDateString();
                string attachment_count = m.Attachments.Count > 0 ? "Y" : "N";
                tempTable.Rows.Add(mailSubject, mailSender, mailBody, receivedTime, attachment_count);
            }
        }


        /// <summary>
        /// Created by Ramachandran Narayanan - 04 October 2017
        /// Export the data to excel
        /// </summary>
        /// <param name="mails"></param>
        public void ExportToExcel()
        {
            Excel.Application excelApp = new Excel.Application();

            //Create an Excel workbook instance and open it from the predefined location
            Excel.Workbook excelWorkBook = excelApp.Workbooks.Open(ConfigurationManager.AppSettings["ExcelPath"].ToString());

            //Add a new worksheet to workbook with the Datatable name
            Excel.Worksheet excelWorkSheet = excelWorkBook.Sheets.Add();
            
            for (int i = 1; i < tempTable.Columns.Count + 1; i++)
            {
                excelWorkSheet.Cells[1, i] = tempTable.Columns[i - 1].ColumnName;
            }

            for (int j = 0; j < tempTable.Rows.Count; j++)
            {
                for (int k = 0; k < tempTable.Columns.Count; k++)
                {
                    excelWorkSheet.Cells[j + 2, k + 1] = tempTable.Rows[j].ItemArray[k].ToString();
                }
            }
            excelWorkBook.Save();
            excelWorkBook.Close();
            excelApp.Quit();
        }


        /// <summary>
        /// Created by Ramachandran Narayanan - 04 October 2017
        /// Initialize the data table with columns
        /// The table contains 5 columns - Sender, Subject, Body of the mail, Received Time and a flag that says 'Y' when the mail contains attachment
        /// </summary>
        public void InitializeDataTable()
        {
            tempTable.Columns.Add("Sender");
            tempTable.Columns.Add("Subject");
            tempTable.Columns.Add("Body");
            tempTable.Columns.Add("ReceivedTime");
            tempTable.Columns.Add("Contains_Attachments");
        }

        /// <summary>
        /// Created by Ramachandran Narayanan - 04 October 2017
        /// Parse the outlook and pick unread mails
        /// 
        /// </summary>
        /// <returns> List<MailItem></returns>
        public List<MailItem> ParseOutlookApp()
        {
            try
            {
                app = new Microsoft.Office.Interop.Outlook.Application();
                ns = app.GetNamespace("MAPI");
                ns.Logon(null, null, false, false);
                var recipient = ns.CreateRecipient(ConfigurationManager.AppSettings["SharedMailBox"].ToString());
                recipient.Resolve();
                switch(ConfigurationManager.AppSettings["IsSharedMailbox"].ToString())
                {
                    case "Y":
                        inboxFolder = ns.GetSharedDefaultFolder(recipient, Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderInbox);
                        break;
                    case "N":
                        inboxFolder = ns.GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderInbox);
                        break;
                }
                
                //string downloadPath = ConfigurationManager.AppSettings["DownloadLocation"].ToString();
                string pathToSave = string.Empty;
                mails = new List<MailItem>();
                foreach (Microsoft.Office.Interop.Outlook.MailItem mail in inboxFolder.Items)
                {
                    if (mail is Microsoft.Office.Interop.Outlook.MailItem && mail.UnRead)
                    {
                        mails.Add(mail as MailItem);
                        mail.UnRead = false;
                        if(mail.Attachments.Count > 0)
                        {
                            string sub = mail.Subject.ToString().Replace("FW:","").Replace("RE:","").Trim();
                                
                            if (!Directory.Exists(Path.Combine(downloadPath + sub)))
                            {
                                Directory.SetCurrentDirectory(downloadPath);
                                Directory.CreateDirectory(sub);
                            }
                            pathToSave = downloadPath + sub + "\\";
                            for(int i=1 ; i<= mail.Attachments.Count ; i++)
                            {
                                mail.Attachments[i].SaveAsFile(pathToSave + mail.Attachments[i].FileName);
                            }
                        }
                    }   
                }
                AddToTable(mails);
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                throw ex;
            }
            catch(System.Exception e)
            {
                throw e;
            }
            finally
            {
                ns = null;
                app = null;
                inboxFolder = null;
            }
            return mails;
        }
    }
}
