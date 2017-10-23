using Logging;
using System;

namespace OutlookEmailParsing
{
    /// <summary>
    /// Created by Ramachandran Narayanan - 04 October 2017
    /// The startup class for the execution of the exe. This program accepts no input params
    /// </summary>
    class Startup
    {
        public static void Main(String[] argv)
        {
            if (DotNetVersionFinder.Get45PlusFromRegistry())
            {
                try
                {
                    Logger.LogMessage("Starting Utility", true);
                    IMailParsing ex = new ExchangeRepository();
                    ex.InitializeDataTable();
                    ex.ParseOutlookApp();
                    ex.ExportToExcel();
                    Logger.LogMessage("Utility Completed Successfully", false);
                }
                catch (Exception ex)
                {
                    Logger.LogException(ex);
                }
            }
            else
            {
                Environment.Exit(0);
            }
        }
    }
}
