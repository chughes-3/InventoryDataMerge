using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;

namespace InventoryDataMerge2013
{
    class InventoryWorkBookClass
    {
        const int colMfgSer = 8;    //MUST BE CHANGED IF MFG_SERIAL_NUM Column is changedChanged for 13
        const string colMfgSerH = "H";
        const int colAssTag = 3;    //MUST BE CHANGED IF Asset_Tag column changes!!Changed for 13
        const string colAssTagC = "C";
        const string colMRedSerT = "T";
        const string colMRedSerHdr = "MR_Serial_Number"; //spec'd here to make any col name change obvious
        const string colMfgSerHdr = "Mfg_Serial_Number";
        const string colAssTagHdr = "Asset_Tag";
        const int colIDCEquality = 10;  //used in proc that checks existing IDC data against new idc data. Will need to change if change spreadsheet
        char[] alpha = { 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z' }; //4 convert A1 to R1C1
        List<string> colHdrs = new List<string>() { "district_number", "state", "asset_tag", "category", "provider", "mfg_id", "mfg_model", "mfg_serial_number", "mfg_date", "status", "processor_speed", "memory", "hard_drive_size", "notes", "custodial_vol_id", "custodian", "computer_name", "mr_manufacturer", "mr_model", "mr_serial_number", "os_name", "os_version", "os_width", "os_product_key_type", "os_partial_product_key", "os_product_key", "lac_mac", "lac_name" };

        Excel.Application xlApp;
        //Excel.Workbooks xlWBooks;
        Excel.Workbook xlWBook = null;
        Excel.Sheets xlWSheets;
        Excel.Worksheet xlWsheet = null;
        List<RowData> rowList = new List<RowData>() { new RowData() { lAssTag = "", lMfgSerNum = "", lMRedSerNum = "" } };  //initial entry to make indexing = excel indexing
        class RowData
        {
            public string lAssTag;
            public string lMfgSerNum;
            public string lMRedSerNum;
        }
        int rowStart = 0;   //first line of real data in spreadsheet
        int rowEnd;         //last row of real data in spreadsheet
        int rowEndImport;   //last row of imported XML data
        internal int rowsIdentical;
        internal int rowsMerged;
        internal int rowsXMLrecsImported;
        internal int rowsAdded;
        internal int rowsMergedHelp;

        internal InventoryWorkBookClass()
        {
            #region Initialization, open spreadsheet, select worksheet setup ranges
            try
            {
                xlApp = System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application") as Excel.Application;
            }
            catch (Exception)
            {
                xlApp = new Excel.Application();
            }
            Microsoft.Office.Core.FileDialog fd = this.xlApp.get_FileDialog(Microsoft.Office.Core.MsoFileDialogType.msoFileDialogOpen);
            fd.AllowMultiSelect = false;
            fd.Filters.Clear();
            fd.Filters.Add("Excel Files", "*.xls;*.xlsx");
            fd.Filters.Add("All Files", "*.*");
            fd.Title = "Open Tax-Aide Equipment Inventory Spreadsheet 2013";
            fd.InitialFileName = Properties.Settings.Default.wBookMRU;
            if (fd.Show() == -1)
                try
                {
                    fd.Execute();
                }
                catch (Exception)
                {
                    MessageBox.Show("There is a problem opening the excel file.\r\nPlease close any open Excel applications and start the program again.", Start.mbCaption);
                    Environment.Exit(1);
                }
            else
                Environment.Exit(1);
            Properties.Settings.Default.wBookMRU = xlApp.ActiveWorkbook.Name;
            Properties.Settings.Default.Save();
            xlApp.Visible = true;
            xlWBook = xlApp.ActiveWorkbook;
            xlWSheets = xlWBook.Sheets;
            foreach (Excel.Worksheet wsht in xlWSheets)
            {
                if (wsht.Name == "State Inventory")
                    xlWsheet = wsht;
            }
            if (xlWsheet == null)
            {
                MessageBox.Show("Cannot find the State Inventory Worksheet.\r\nDo you have the correct Tax-Aide Equipment Inventory Report Spreadsheet", Start.mbCaption);
                Environment.Exit(1);
            }

        }

        internal void SetupRange()
        {
            //xlWBook.Activate();
            //xlWsheet.Activate();
            Excel.Range stateSearchRng = xlWsheet.Range["A1:A40"];  //40 rows should be enough to find the State
            object[,] stateSearchObj = new object[40, 1];
            stateSearchObj = stateSearchRng.Value2;
            for (int i = 1; i < 40; i++)    //40 rows should be enough to find the State
            {
                if (stateSearchObj[i, 1] != null && stateSearchObj[i, 1].ToString() == "District_Number")
                //string cellValue = (xlWsheet.Cells[i, 1].Value != null) ? xlWsheet.Cells[i, 1].Value.ToString() : "";
                //if (cellValue == "State")
                {
                    rowStart = i + 1;
                    break;
                }
            }
            if (rowStart == 0)
            {
                MessageBox.Show("Unable to find \"District_Number\" column label in the first column\r\rExiting", "IDC Merge");
                DisposeX();
            }
            //Let us check column headers
            Excel.Range colHdrsRng = xlWsheet.Range["A" + (rowStart - 1).ToString() + ":AB" + (rowStart - 1).ToString()];
            object[,] colHdrsObj = new object[28, 1];
            colHdrsObj = colHdrsRng.Value2;
            var colHdrsStr = colHdrsObj.Cast<string>().Select(i => i == null ? "" : i.ToLower());
            //Debug.WriteLine(string.Join(Environment.NewLine + "\t",colHdrsStr.Select(x => x.ToString())));
            var result = colHdrsStr.SequenceEqual(colHdrs);
            if (!result)
            {
                MessageBox.Show("The Column Headers in this spreadsheet do not conform to the EIR 2013 specification.\r\nIs the correct spreadsheet open?\r\n\r\nExiting", Start.mbCaption);
                DisposeX();
            }
            //First make sure that working range is one area and has no blank rows at the end.
            if (xlWsheet.UsedRange.Areas.Count != 1)
            {
                MessageBox.Show("The used range on this spreadsheet is not contiguous, something is wrong.\r\nIs the correct spreadsheet open?\r\n\r\nExiting", Start.mbCaption);
                DisposeX();
            }
            string colHeadAss = "";
            string colHeadSer = "";
            try
            {//This is first place access a cell and likely will give issues if spreadsheet in edit mode.
                xlApp.CutCopyMode = (Excel.XlCutCopyMode)0;     //Here to resolve issue if user has left copy or cut selected. Prgram takes control
                colHeadAss = (xlWsheet.Cells[rowStart - 1, colAssTag].Value != null) ? xlWsheet.Cells[rowStart - 1, colAssTag].Value.ToString() : "";
                colHeadSer = (xlWsheet.Cells[rowStart - 1, colMfgSer].Value != null) ? xlWsheet.Cells[rowStart - 1, colMfgSer].Value.ToString() : "";
            }
            catch (Exception)
            {
                MessageBox.Show("The spreadsheet is not accepting programmatic input.\rThe simplest way to fix this is to start the program using a freshly opened spreadsheet in which no editing has been done.\r\r   Exiting!", "IDC Data Merge");
                DisposeX();
            }
            if (colHeadAss != "Asset_Tag" || colHeadSer != "Mfg_Serial_Number")
            {
                MessageBox.Show("The Asset Tag and/or Mfg Serial Number column headings are not in the expected places.\rIs the program pointed at a correctly formatted spreadsheet?", "IDC Data Merge");
                DisposeX();
            }
            rowEnd = xlWsheet.Cells.Find("*", SearchOrder: Excel.XlSearchOrder.xlByRows, SearchDirection: Excel.XlSearchDirection.xlPrevious).Row;
            rowEndImport = rowEnd;
            Excel.Range rowsAssTag = xlWsheet.Range[colAssTagC + "1:" + colAssTagC + rowEnd.ToString()];  //start at 1 to keep indexing same as spreadsheet
            object[,] rowsAssTagObj = new object[rowsAssTag.Count, 1];
            Excel.Range rowsMfgSer = xlWsheet.Range[colMfgSerH + "1:" + colMfgSerH + rowEnd.ToString()];  //start 
            object[,] rowsMfgSerObj = new object[rowsMfgSer.Count, 1];
            Excel.Range rowsMred = xlWsheet.Range[colMRedSerT + "1:" + colMRedSerT + rowEnd.ToString()];  //start at 1 to keep indexing same as spreadsheet
            object[,] rowsMredObj = new object[rowsMred.Count, 1];
            rowsAssTagObj = rowsAssTag.Value2;
            rowsMfgSerObj = rowsMfgSer.Value2;
            rowsMredObj = rowsMred.Value2;
            for (int i = 1; i < rowsAssTag.Count + 1; i++)
            {
                rowList.Add(new RowData() { lAssTag = (rowsAssTagObj[i, 1] != null) ? rowsAssTagObj[i, 1].ToString().Trim() : "" });
                rowList[i].lMfgSerNum = ((rowsMfgSerObj[i, 1] != null) ? rowsMfgSerObj[i, 1].ToString().Trim() : "");
                rowList[i].lMRedSerNum = ((rowsMredObj[i, 1] != null) ? rowsMredObj[i, 1].ToString().Trim() : "");
            }
            xlWsheet.Rows[rowStart].Select();
        }
            #endregion

        internal void DisposeX()
        {
            xlApp = null;
            Environment.Exit(1);
        }
        internal void Dispose()
        {
            xlApp = null;
        }

        internal void IDCSysProcess(System.Xml.Linq.XElement el)
        {
            ImportXmlRow(el);   //get IDC element to bottom of spreadsheet
            var qryScope = from row in rowList
                           where (rowList.IndexOf(row) >= rowStart && rowList.IndexOf(row) <= rowEnd)
                           select row;
            RowData mrSerPresent = qryScope.FirstOrDefault(rw => rw.lMRedSerNum == el.Element("mr_serial_number").Value.Trim());
            if (mrSerPresent == null)
            {//nr not in spread sheet check for asset tag or mfg serial
                RowData assTagPres = qryScope.Where(rw => rw.lAssTag != "").FirstOrDefault(rw => rw.lAssTag == (el.Element("asset_tag") == null ? string.Empty : el.Element("asset_tag").Value.Trim()));
                if (assTagPres == null)
                {
                    RowData hrSerPres = qryScope.Where(rw => rw.lMfgSerNum != "").FirstOrDefault(rw => rw.lMfgSerNum == (el.Element("mfg_serial_number") == null ? string.Empty : el.Element("mfg_serial_number").Value.Trim()));
                    if (hrSerPres == null)
                    {//We have no mr match, no atag match and no hr serial match
                        //Check if mrSer == hrSer if we have blank AT and MfgSer
                        if ((el.Element("asset_tag") == null || el.Element("asset_tag").Value.Trim() == "") && (el.Element("mfg_serial_number") == null || el.Element("mfg_serial_number").Value.Trim() == ""))
                        {
                            RowData mrSerEqHrSer = qryScope.FirstOrDefault(rw => rw.lMfgSerNum == el.Element("mr_serial_number").Value.Trim());
                            int mrSerEqHrSerIndx = rowList.IndexOf(mrSerEqHrSer);
                            if (mrSerEqHrSer == null)
                            {//no mr match, no atag match, no hrserial match,  NO mrser=hrserial match
                                rowsAdded++;
                                return; //it is already at bottom of spreadsheet do not need return here done for code reading clarity
                            }
                            else
                            {//we have mrser=hrserial match we need to import XML IDC data to existing spreadsheet record 
                                MergeIDC2RowExist(mrSerEqHrSerIndx);
                                rowList[mrSerEqHrSerIndx].lAssTag = xlWsheet.Cells[mrSerEqHrSerIndx, colAssTag].Value != null ? xlWsheet.Cells[mrSerEqHrSerIndx, colAssTag].Value.ToString().Trim() : string.Empty;    //since we know that HR is different now than in xml
                                rowList[mrSerEqHrSerIndx].lMfgSerNum = xlWsheet.Cells[mrSerEqHrSerIndx, colMfgSer].Value != null ? xlWsheet.Cells[mrSerEqHrSerIndx, colMfgSer].Value.ToString().Trim() : string.Empty;
                            }
                        }
                        else
                        {//No Matches for mrser,hrser,Atag and blank IDC hr data. We have a full xml record to import to spreadsheet including HR data
                            rowsAdded++;
                            return; //already at bottom of spreadsheet do not need return here done for code reading clarity
                        }
                    }
                    else
                    {//We have a match for hrser data, no match for mrser and Atag.
                        int hrSerPresIndx = rowList.IndexOf(hrSerPres);
                        if (hrSerPres.lAssTag == string.Empty || (el.Element("asset_tag") == null || el.Element("asset_tag").Value.Trim() == string.Empty) || hrSerPres.lAssTag == el.Element("asset_tag").Value.Trim())
                        {//no asset tags so just merge idc row into hrSer matched row
                            MergeIDC2RowExist(hrSerPresIndx);
                        }
                        else
                        {//we have asset tags and they do not match must ask user
                            FldMatchErrAskUser(hrSerPresIndx, "There is a record match for the Mfg_Serial_Number column between the spreadsheet data in row {0} and the IDC data in row {1}, but the Asset_Tag fields are different(1).\r\nPlease make any changes in row {1} which has the purple text. This is the row that will be kept. After making all changes click OK in this dialog box.\r\nThe Row {0} will then be overwritten.");
                        }
                    }
                }
                else
                {//We have a match for asset tag, no match for mrSer - need to check if hrser has a match or not
                    int assTagPresIndx = rowList.IndexOf(assTagPres);
                    if (assTagPres.lMfgSerNum == string.Empty || el.Element("mfg_serial_number") == null || el.Element("mfg_serial_number").Value.Trim() == string.Empty || assTagPres.lMfgSerNum == el.Element("mfg_serial_number").Value.Trim())
                    {
                        MergeIDC2RowExist(assTagPresIndx);
                    }
                    else
                    {//we have matched asset tags but mismatched hrSernum must ask user
                        FldMatchErrAskUser(assTagPresIndx, "There is a record match for the Asset_Tag column between the spreadsheet data in row {0} and the IDC data in row {1}, but the Mfg_Serial_Number fields are different(2).\r\nPlease make any changes in row {1} which has the purple text. This is the row that will be kept. After making all changes click OK in this dialog box.\r\nThe Row {0} will then be overwritten.");
                    }
                }
            }
            else
            {//MrSerial found in spreadsheet NEED TO CHECK ASSTAG?HRSERIAL the same before copy paste
                int mrSerMatchIndex = rowList.IndexOf(mrSerPresent);
                if (el.Element("asset_tag") != null && el.Element("asset_tag").Value.Trim() == mrSerPresent.lAssTag && el.Element("mfg_serial_number") != null && el.Element("mfg_serial_number").Value.Trim() == mrSerPresent.lMfgSerNum)
                {//clean system import from XML to existing spreadsheet data row replacing identical existing data.
                    MergeIDC2RowExist(mrSerMatchIndex);
                    rowsIdentical++;
                }
                else
                {// MRSerno match but not a match in either HRSer or ATag, put problem rows together issue user message
                    FldMatchErrAskUser(mrSerMatchIndex, "There is an error in the Asset Tag or Mfg_Serial_Number column in either the original spreadsheet row {0} or in the imported IDC data row {1}.\r\nPlease make any changes in row {1} which has the purple text. This is the row that will be kept. After making all changes click OK in this dialog box.\r\nThe Row {0} will then be overwritten.");

                }

            }
        }

        private void MergeIDC2RowExist(int rowMergeIndx)
        {
            xlWsheet.Rows[rowEndImport].Copy();
            xlWsheet.Rows[rowMergeIndx].PasteSpecial(SkipBlanks: true);
            xlWsheet.Rows[rowEndImport].Delete();
            rowList.RemoveAt(rowList.Count - 1);
            rowEndImport--;
            rowsMerged++;
        }

        private void FldMatchErrAskUser(int rowMatchIndex, string messUserTxt)
        {
            xlWsheet.Rows[rowMatchIndex + 1].Insert();
            xlWsheet.Rows[rowEndImport + 1].Cut(xlWsheet.Rows[rowMatchIndex + 1]);
            xlWsheet.Rows[rowEndImport + 1].Delete();
            rowList.Insert(rowMatchIndex + 1, rowList[rowList.Count - 1]);
            rowList.RemoveAt(rowList.Count - 1);
            Excel.Range rngErr = xlWsheet.Rows.EntireRow[rowMatchIndex.ToString()];
            //var rowCol = rngErr.Interior.Color;
            rngErr.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightYellow);
            //var row1Col = rngErr.Offset[1,0].Interior.Color;
            rngErr.Offset[1, 0].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightYellow);
            //var rowTxtCol = rngErr.Font.Color;
            rngErr.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkGreen);
            //var rowTxt1Col = rngErr.Offset[1, 0].Font.Color;
            rngErr.Offset[1, 0].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Purple);
            DialogResult dR = MessageBox.Show(string.Format(messUserTxt, rowMatchIndex, rowMatchIndex + 1), Start.mbCaption, MessageBoxButtons.OKCancel, MessageBoxIcon.Error);
            if (dR == DialogResult.Cancel)
                DisposeX();
            //rngErr.Interior.Color = rowCol;
            //rngErr.Offset[1, 0].Interior.Color = row1Col;
            //rngErr.Font.Color = rowTxtCol;
            //rngErr.Offset[1, 0].Font.Color = rowTxt1Col;
            try
            {
                xlWsheet.Rows[rowMatchIndex + 1].Cut(xlWsheet.Rows[rowMatchIndex]);
            }
            catch
            {
                MessageBox.Show("The Excel program is not accepting programmatic input.\r\nMake sure that you have closed out all editing in the excel spreadsheet.\r\nMake sure that a blank cell is selected.\r\nThen click OK in this Dialog", Start.mbCaption, MessageBoxButtons.OK, MessageBoxIcon.Error);
                xlWsheet.Rows[rowMatchIndex + 1].Cut(xlWsheet.Rows[rowMatchIndex]);
            }
            xlWsheet.Rows[rowMatchIndex + 1].Delete();
            xlWsheet.Rows[rowMatchIndex].ClearFormats();
            rowEndImport--; //deleted a spreadsheet row so import row goes back
            //update ListRows to capture any At or hr ser changes
            rowList.RemoveAt(rowMatchIndex + 1);
            rowList[rowMatchIndex].lAssTag = xlWsheet.Cells[rowMatchIndex, colAssTag].Value != null ? xlWsheet.Cells[rowMatchIndex, colAssTag].Value.ToString().Trim() : string.Empty;
            rowList[rowMatchIndex].lMfgSerNum = xlWsheet.Cells[rowMatchIndex, colMfgSer].Value != null ? xlWsheet.Cells[rowMatchIndex, colMfgSer].Value.ToString().Trim() : string.Empty;
            rowsMergedHelp++;
        }
        void ImportXmlRow(System.Xml.Linq.XElement el)
        {
            string colEnd = "";
            if (colHdrs.Count < 26)
                colEnd = alpha[colHdrs.Count - 1].ToString();
            else if (colHdrs.Count < 52)
                colEnd = "A" + alpha[colHdrs.Count - 27].ToString();     //if more than 52 cols will throw an error
            object[,] objData = new object[1, colHdrs.Count];
            rowEndImport++;   //we are extending worksheet by one row
            Excel.Range rngIDC = xlWsheet.Range["A" + rowEndImport.ToString() + ":" + colEnd + rowEndImport.ToString()];
            for (int i = 0; i < colHdrs.Count; i++)
            {
                objData[0, i] = el.Element(colHdrs[i]) == null ? string.Empty : el.Element(colHdrs[i]).Value.Trim();
            }
            rngIDC.ClearFormats();
            rngIDC.Value2 = objData;
            //Next update Row list not really needed except need to keep it synchronized in case of error conditions so do it anyway
            rowList.Add(new RowData { lAssTag = el.Element(colAssTagHdr.ToLower()).Value.Trim(), lMfgSerNum = el.Element(colMfgSerHdr.ToLower()).Value.Trim(), lMRedSerNum = el.Element(colMRedSerHdr.ToLower()).Value.Trim() });
            rowsXMLrecsImported++;

        }
    }
}
