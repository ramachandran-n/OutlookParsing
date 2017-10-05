using System.Collections.Generic;
using Microsoft.Office.Interop.Outlook;

namespace OutlookEmailParsing
{
    /// <summary>
    /// Created by Ramachandran Narayanan - 04 October 2017
    /// Interface file for the creation of the class
    /// </summary>
    public interface IMailParsing
    {
        List<MailItem> ParseOutlookApp();
        void ExportToExcel();
        void InitializeDataTable();
    }
}
