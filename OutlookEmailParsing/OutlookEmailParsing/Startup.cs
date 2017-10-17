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
            if (GetDotNetVersion.Get45PlusFromRegistry())
            {
                try
                {
                    IMailParsing ex = new ExchangeRepository();
                    ex.InitializeDataTable();
                    ex.ParseOutlookApp();
                    ex.ExportToExcel();
                }
                catch (Exception ex)
                {
                    //Do nothing
                }
            }
            else
            {
                Environment.Exit(0);
            }
        }
    }
}
