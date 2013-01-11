using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System.Diagnostics;
using System.Text.RegularExpressions; 
using System.Xml.Linq;

namespace InventoryDataMerge2013
{
    static class Start
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        
        internal const string mbCaption = "Tax-Aide Inventory Data Merge 2013";

        [STAThread]
        static void Main()
        {
            DateTime endDate = new DateTime(2013, 1, 31);
            if (endDate < DateTime.Now)
            {
                MessageBox.Show("This program version is intended for use in the 2013 Tax-Aide Inventory reporting activity.\r\nTherefore this version stopped working in Jan 31 2013.\r\n\r\nQuestions? Please contact your TCS or TaxAideTech", "AARP Foundation Tax-Aide");
                Environment.Exit(0);
            }
            Regex myPatt = new Regex(@"\((.*)\)"); //extract process friendly name from full process
            Process myProc = Process.GetCurrentProcess();
            Match myMatch = myPatt.Match(myProc.ToString());
            String myProcFriendly = myMatch.Value.Substring(1, myMatch.Length - 2);//get rid of parentheses
            Process[] myProcArray = Process.GetProcessesByName(myProcFriendly);
            if (myProcArray.GetLength(0) > 1) 
            {
                MessageBox.Show("This program is already running", "IDC Data Merge", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                Environment.Exit(1);
            }
            InventoryWorkBookClass xcel = new InventoryWorkBookClass();   // opens spreadsheet brings spreadsheet data across into a List(rowData) 
            xcel.SetupRange();
            DialogResult dR = DialogResult.Abort;
            while (dR != DialogResult.No)
            {
                InvXMLFile xmlData = new InvXMLFile();
                if (!xmlData.GetIDCXmlData())
                    xcel.DisposeX();
                foreach (XElement el in xmlData.systems)
                {
                    xcel.IDCSysProcess(el);
                }
                dR = MessageBox.Show(String.Format("The IDC data merge 2013 is complete with the following results\r\nIDC records processed : {0}\r\nIDC records added to the spreadsheet: {3}\r\nIDC records merged with existing spreadsheet rows (Identical data in search columns) : {1}\r\nIDC records merged with existing spreadsheet rows (one or more search column matches) : {2}\r\nIDC records merged with existing spreadsheet rows (with user help) : {4}\r\n\r\nProcess another TaxAideInv2013.xml file?", xcel.rowsXMLrecsImported, xcel.rowsIdentical, xcel.rowsMerged, xcel.rowsAdded, xcel.rowsMergedHelp), mbCaption,MessageBoxButtons.YesNo,MessageBoxIcon.Question);
                xcel.rowsIdentical = 0;
                xcel.rowsMerged = 0;
                xcel.rowsXMLrecsImported = 0;
                xcel.rowsMergedHelp = 0;
                xcel.rowsAdded = 0;
            }
        }
    }
}
