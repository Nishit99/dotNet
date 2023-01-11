using System;
using System.Drawing;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;
using CreateUserReadinessReport.EntityModel;

namespace ExcelAutomationToolLibrary
{
    public class ToolHelperClass
    {
        Excel.Application xlApp = null;
        Excel.Workbook xlWorkbook = null;
        Excel.Sheets sheets = null;
        Excel._Worksheet xlWorksheet = null;
        Excel.Range rngHeaderRow = null;
        ColumnConfig columnConfig = null;
        string _HeaderAddress = "";
        int columnIndex = 0;
        string srtTableArrayURL = "";
        Logger logger = null;

        /// <summary>
        /// Constructor used to initialize objects.
        /// </summary>
        private ToolHelperClass()
        {
            xlApp = new Excel.Application()
            {
                DisplayAlerts = false,
                Visible = false,
                AskToUpdateLinks = false,
            };

            logger = Logger.Instance;

            //Read configured columns from JOSN file.
            columnConfig = JsonConvert.DeserializeObject<ColumnConfig>(System.IO.File.ReadAllText(@"ColumnConfig\ColumnConfig.json"));
        }

        /// <summary>
        /// Static property to get instance of ToolHelperClass.
        /// </summary>
        public static ToolHelperClass GetWorkBookInstance { get => new ToolHelperClass(); }

        /// <summary>
        /// Method the run the ExcelAutomationTool.
        /// 1. Initialize ExcelAutomationTool objects.
        /// 2. Insert new column in source file and insert fomula in define range of cells.
        /// 3. Log the exception into logger file.
        /// 4. One all the action is completed release all the excel objects.
        /// </summary>
        /// <param name="SourceFileName">String type, Path of the source file.</param>
        /// <param name="ReferenceFilePath_URL">String type path of Reference file.</param>
        /// <param name="lstBoxSelectedWorksheets">String type, name of the selected worksheet..</param>
        public void RunExcelAutomationTool(string SourceFileName, string ReferenceFilePath_URL, string lstBoxSelectedWorksheets)
        {
            try
            {
                //Initialize ExcelAutomationTool objects.
                InitializeExcelAutomationTool(SourceFileName, ReferenceFilePath_URL, lstBoxSelectedWorksheets);

                //Insert new column in source file and insert fomula in define range of cells.
                //Otimize the code the make it more genric.
                #region Insert first column, do the style and fill the value.

                _HeaderAddress = GetColumnAddress(rngHeaderRow, columnConfig.SearchColumn.NewApplicationID);
                InsertColumn(xlWorksheet, _HeaderAddress, columnConfig.DynamicColumn.BaselineCoumn.ADMID);
                //InsertColumn(xlWorksheet, "", columnConfig.DynamicColumn.BaselineCoumn.ADMID);
                //CellStyle(xlWorksheet.get_Range(HeaderAddress), "#8EA9DB", 43);
                //
                //Fill the value in Inserted column.
                //
                //columnIndex = GetColumnIndex(rngHeaderRow, columnConfig.SearchColumn.NewApplicationID) - 1;
                columnIndex = GetColumnIndex(columnConfig.DynamicColumn.BaselineCoumn.ADMID) - 1;
                InsertXLOOKUP_Formula(srtTableArrayURL, "B", "B");
                #endregion

                #region Insert second column, do the style and fill the value.

                ////HeaderAddress = GetColumnAddress(rngHeaderRow, columnConfig.SearchColumn.NewApplicationID);
                ////InsertColumn(xlWorksheet, HeaderAddress, columnConfig.DynamicColumn.BaselineCoumn.FinalStatus);
                ////CellStyle(xlWorksheet.get_Range(HeaderAddress), "#8EA9DB", 43);
                //////
                //////Fill the value in Inserted column.
                //////
                ////columnIndex = GetColumnIndex(rngHeaderRow, columnConfig.SearchColumn.Category) - 1;
                ////InsertXLOOKUP_Formula(srtTableArrayURL, "B", "D");
                #endregion

                #region Insert Thired column, do the style and fill the value.

                ////HeaderAddress = GetColumnAddress(rngHeaderRow, columnConfig.SearchColumn.Category);
                ////InsertColumn(xlWorksheet, HeaderAddress, columnConfig.DynamicColumn.NewColumn.NewReadiness);
                ////CellStyle(xlWorksheet.get_Range(HeaderAddress), "#8EA9DB", 21);
                //////
                //////Fill the value in Inserted column.
                //////
                ////string srtLogicHeaderaddress = GetColumnAddress(rngHeaderRow, columnConfig.DynamicColumn.BaselineCoumn.FinalStatus);
                ////InsertIF_Formula(srtLogicHeaderaddress, "Available to Production Users", "Ready", "Not Ready");
                #endregion
            }
            catch (Exception ex)
            {
                //Log the exception in logger file.
                logger.LogMessage(ex.Message);
            }

            finally
            {
                //One all the action is completed release all the excel objects.
                xlWorkbook.Save();
                xlWorkbook.Close();
                xlApp.DisplayAlerts = true;
                xlApp.Visible = true;
                xlApp.Quit();
                Marshal.ReleaseComObject(xlWorkbook);
                Marshal.ReleaseComObject(xlApp);
            }
        }

        /// <summary>
        /// Public method to get all the workssheet from the selected reference excel file.
        /// 1. Get all the Worksheet from the reference file.
        /// 2. Add the worksheet name to a list object.
        /// 3. Log error into log file.
        /// </summary>
        /// <param name="ReferenceFilePath_URL">String type path of Reference file.</param>
        /// <returns>Return a collection object, A list of workseets.</returns>
        public List<string> GetAllWorksheets(string ReferenceFilePath_URL)
        {
            List<string> worksheetItem = new List<string>();
            try
            {
                if (!string.IsNullOrEmpty(ReferenceFilePath_URL) && worksheetItem !=null)
                {
                    //Open reference excel file.
                    Excel.Workbook tempXlWorkbook = xlApp.Workbooks.Open(FormatURL(ReferenceFilePath_URL, false), 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                    
                    //Loop through all the worksheet in a workbook.
                    foreach (Excel.Worksheet tempSheet in tempXlWorkbook.Worksheets)
                    {
                        //Add worksheet name to a list object.
                        worksheetItem.Add(tempSheet.Name);
                    }
                }
            }
            catch (Exception ex)
            {
                //Log error into log file.
                logger.LogMessage(ex.Message);
            }

            return worksheetItem;
        }

        #region Private Methods

        /// <summary>
        /// 
        /// </summary>
        /// <param name="SourceFileName"></param>
        /// <param name="ReferenceFilePath_URL"></param>
        /// <param name="lstBoxSelectedWorksheets"></param>
        private void InitializeExcelAutomationTool(string SourceFileName, string ReferenceFilePath_URL, string lstBoxSelectedWorksheets)
        {
            try
            {
                //Formate URL for excel formula.
                srtTableArrayURL = FormatURL(ReferenceFilePath_URL, true, lstBoxSelectedWorksheets);

                xlWorkbook = xlApp.Workbooks.Open(SourceFileName);
                sheets = xlWorkbook.Worksheets;
                xlWorksheet = (Excel.Worksheet)sheets.get_Item(1);
                Excel.Range xlRange = xlWorksheet.UsedRange;
                rngHeaderRow = xlRange.Rows[1];
            }
            catch (Exception ex)
            {
                logger.LogMessage(ex.Message);
            }
        }

        private string GetExcelColumnName(int columnNumber)
        {
            string columnName = "";

            while (columnNumber > 0)
            {
                int modulo = (columnNumber - 1) % 26;
                columnName = Convert.ToChar('A' + modulo) + columnName;
                columnNumber = (columnNumber - modulo) / 26;
            }

            return columnName;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="headerRange"></param>
        /// <param name="columnName"></param>
        /// <returns></returns>
        private string GetColumnAddress(Excel.Range headerRange, string columnName)
        {
            string ColumnAddress = "";
            try
            {
                Excel.Range rngResult = headerRange.Find(What: columnName);
                ColumnAddress = Regex.Replace(rngResult.Address[Excel.XlReferenceStyle.xlA1], @"[^\w\.@-]", "");
            }
            catch (ArgumentException ex)
            {
                logger.LogMessage(ex.Message);
            }
            return ColumnAddress;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="headerRange"></param>
        /// <param name="columnName"></param>
        /// <returns></returns>
        private int GetColumnIndex(string columnName)
        {
            int columnIndex = 0;
            try
            {
                Excel.Range headerRange = xlWorksheet.UsedRange;
                columnIndex = headerRange.Find(What: columnName) != null ? headerRange.Find(What: columnName).Column : -1;
            }
            catch (ArgumentException ex)
            {
                logger.LogMessage(ex.Message);
            }
            return columnIndex;
        }

        /// <summary>
        /// Void method, used for insert new column.
        /// </summary>
        /// <param name="xlTempWorksheet"> Excel worksheet object.</param>
        /// <param name="HeaderAddress">String Type. Optional paramater. Contains excel column header address (e.g: "A1","B1","C1" ...). 
        /// IF value is given new column will be inserted before the given header address else it will be inseted after last used column.</param>
        /// <param name="NewColumnHeaderName">String type contain new column header name.</param>
        private void InsertColumn(Excel._Worksheet xlTempWorksheet, [Optional]string HeaderAddress, string NewColumnHeaderName)
        {
            try
            {
                Excel.Range oRng = null;

                if (!string.IsNullOrEmpty(HeaderAddress))
                {
                    //Insert in between column.
                    oRng = xlTempWorksheet.Range[HeaderAddress];
                }
                else
                {
                    //Insert column afer last use column.
                    int intColumnIndex = xlTempWorksheet.UsedRange.Columns.Count;
                    HeaderAddress = GetExcelColumnName(intColumnIndex + 1) + 1;
                    _HeaderAddress = HeaderAddress;
                    oRng = xlTempWorksheet.Range[HeaderAddress];
                }

                if (oRng == null) return;

                oRng.EntireColumn.Insert();
                oRng = xlTempWorksheet.Range[HeaderAddress];
                oRng.Value2 = NewColumnHeaderName;
            }
            catch (Exception ex)
            {
                logger.LogMessage(ex.Message);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="formatRange"></param>
        /// <param name="BGColor"></param>
        /// <param name="columnwWidth"></param>
        private void CellStyle(Excel.Range formatRange, string BGColor, int columnwWidth)
        {
            try
            {
                formatRange.Interior.Color = ColorTranslator.FromHtml(BGColor);
                formatRange.ColumnWidth = columnwWidth;
            }
            catch (FormatException ex)
            {
                logger.LogMessage(ex.Message);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="ReferenceFilePath_URL"></param>
        /// <param name="isForFormula"></param>
        /// <param name="SelectedWorksheets"></param>
        /// <returns></returns>
        private string FormatURL(string ReferenceFilePath_URL, bool isForFormula, [Optional] string SelectedWorksheets)
        {
            string strTempformatedURL = "";
            try
            {
                //Uri uriWorkbookfilepath = new Uri(txt_ReferenceFilePath_URL.Text);
                Uri uriWorkbookfilepath = new Uri(ReferenceFilePath_URL);
                string filename = System.IO.Path.GetFileName(uriWorkbookfilepath.LocalPath);

                if (uriWorkbookfilepath.IsFile)
                {
                    if (uriWorkbookfilepath.Host != "")
                    {
                        strTempformatedURL = uriWorkbookfilepath.AbsoluteUri;
                    }
                    else
                    {
                        strTempformatedURL = $"{uriWorkbookfilepath.Host}{uriWorkbookfilepath.LocalPath}";
                    }
                }
                else
                {
                    strTempformatedURL = $"http://{uriWorkbookfilepath.Host}{uriWorkbookfilepath.LocalPath}";
                }

                if (isForFormula)
                {
                    //strTempformatedURL = $"'{strTempformatedURL.Replace(filename, "")}[{filename}]{lstBox_Worksheets.SelectedItem.ToString()}'";
                    strTempformatedURL = $"'{strTempformatedURL.Replace(filename, "")}[{filename}]{SelectedWorksheets}'";
                }
            }
            catch (Exception ex)
            {
                logger.LogMessage(ex.Message);
            }
            return strTempformatedURL;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="srtTableArrayURL"></param>
        /// <param name="lookupColumn"></param>
        /// <param name="returnColumn"></param>
        private void InsertXLOOKUP_Formula(string srtTableArrayURL, string lookupColumn, string returnColumn)
        {
            try
            {
                int countRows = xlWorksheet.UsedRange.Rows.Count;
                string strHeaderAddress = (_HeaderAddress.First<char>()).ToString();
                string strXLFormula = $"=XLOOKUP(A2,{srtTableArrayURL}!${lookupColumn}$2:${lookupColumn}$800,{srtTableArrayURL}!${returnColumn}$2:${returnColumn}$800,\"Not Found!\",0,1)";
                xlWorksheet.Range[strHeaderAddress + "2:" + strHeaderAddress + countRows].Formula = strXLFormula;
            }
            catch (Exception ex)
            {
                logger.LogMessage(ex.Message);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="srtLogicHeaderaddress"></param>
        /// <param name="srtToCheck"></param>
        /// <param name="srtTrueValue"></param>
        /// <param name="srtFalseValue"></param>
        private void InsertIF_Formula(string srtLogicHeaderaddress, string srtToCheck, string srtTrueValue, string srtFalseValue)
        {
            try
            {
                int countRows = xlWorksheet.UsedRange.Rows.Count;
                string strHeaderAddress = (_HeaderAddress.First<char>()).ToString();
                srtLogicHeaderaddress = (srtLogicHeaderaddress.First<char>()).ToString() + "2";
                string strXL_IF_Formula = $"=IF({srtLogicHeaderaddress}=0,{"0"},IF({srtLogicHeaderaddress}=\"{srtToCheck}\",\"{srtTrueValue}\",\"{srtFalseValue}\"))";

                xlWorksheet.Range[strHeaderAddress + "2:" + strHeaderAddress + countRows].Formula = strXL_IF_Formula;
            }
            catch (Exception ex)
            {
                logger.LogMessage(ex.Message);
            }
        }
        #endregion
    }
}
