using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Web;
using GU = ExcelUtil.GenUtils;

namespace ExcelUtil
{
    /// <summary>
    /// Microsoft Office Excel Library through OpenXML SDK
    /// </summary>
    public partial class ExcelLib
    {
        // debug information
        public static string timeCheck1 = "";
        public static string timeCheck2 = "";
        public static string timeCheck3 = "";
        public static string timeCheck4 = "";
        private const int maxExcelCellLength = 32767;

        #region "Excel Export Functions"

        // ***************************************************************************************************************************************************
        //   USAGE:
        //       1. Include "using ExcelUtil;" for C# or "import ExcelUtil" for VB
        //       2. (Optional) create a DataSet and assign TableName to each DataTable
        //       3. call one of the 4 overload routines
        //   4 overload routines, parameter options:
        //       DataTable / strSQL / DataSet - allow multiple DataTables from the last two types
        //       Optional allowing pre-formatted custom template Excel file
        //   Control Export through ExportParameter class object
        // ***************************************************************************************************************************************************
        //   FEATURES:
        //   The default templateFile is coded as a class
        //   Allow us to apply a custom pre-formatted templateFile by supplying templateFile parameter (need to upload template file to the folder)
        //       for example, "app_code/template_custom.xlsx"
        //   Each DataTable will be loaded into one worksheet
        //       The DataTable will use its name (TableName) to match the worksheet name, if the matching cannot be found, use the first available worksheet;
        //       if there is no more worksheet available in the workbook, create a new one.
        //   Header row - the cell style of the first cell "A1" will be copied to all header row cells that were not provided in the template
        //   Lookup Table - need to assign datatable to a name started with "LOOKUP" (e.g. ds.Tables(1).TableName = "LOOKUP1")
        //       - all Lookup worksheets will be hidden (by default)
        //       - the first cell "A1" provides the lookup table column name; the first column A2:Ax provides data source
        //       - the 2nd column (if exists) provides ID lookup (e.g. 2 columns: "PRODUCT" | "PRODUCT_ID", the "PRODUCT_ID" in non-lookup sheets will be looked up by the selection of "PRODUCT")
        //       - all non-lookup tables (non-LOOKUP-prefixed) worksheets' columns will be matched up with lookup column names
        //   Column Width - if the column is not existed in the original worksheet; the column width will be adjusted automatically,
        //       - if the string type column max length is more than 15, the column width will be set to 20px
        //       - if the string type column max length is more than 30, the column width will be set to 40px
        //	Hide Tables (Non-lookup) - if the table name ends with an "*" from the datatable, the "*" will be stripped and the table will be hidden (better using exportParameters.hideTables)
        //	Hide Columns - if the column name ends with an "*" from the datatable, the "*" will be stripped and the column will be hidden

        /// ***************************************************************************************************************************************************
        /// <summary>(overload) export to Excel (xlsx) format from a DataTable </summary>
        /// <param name="myDataTable">DataTable object instance</param>
        /// <param name="downloadFileName">(optional) return file name, e.g. myExcelExport.xlsx; if not provided, default to Excel_Export.xlsx </param>
        /// <param name="templateFile">(optional) Excel file used as a template, default to an internal template </param>
        /// ***************************************************************************************************************************************************
        public static void exportExcel(DataTable myDataTable, string downloadFileName = null, string templateFile = null)
        {
            DataSet ds = new DataSet();
            ds.Tables.Add(myDataTable.Copy());

            exportExcel(ds, downloadFileName, templateFile);
        }

        /// ***************************************************************************************************************************************************
        /// <summary>(overload) export to Excel (xlsx) format from a DataSet - a worksheet will be generated for each DataTable in the DataSet </summary>
        /// <param name="myDataSet">DataSet object instance</param>
        /// <param name="downloadFileName">(optional) return file name, e.g. myExcelExport.xlsx; if not provided, default to Excel_Export.xlsx</param>
        /// <param name="templateFile">(optional) Excel file used as a template, default to an internal template </param>
        /// ***************************************************************************************************************************************************
        public static void exportExcel(DataSet myDataSet, string downloadFileName = null, string templateFile = null)
        {
            ExportParameters exportParam = new ExportParameters();
            exportParam.SqlDataSet = myDataSet;
            if (downloadFileName != null)
                exportParam.DownloadFileName = downloadFileName;
            if (templateFile != null)
                exportParam.TemplateFile = templateFile;

            exportExcel(exportParam);
        }

        /// ***************************************************************************************************************************************************
        /// <summary>(overload) export to Excel (xlsx) format from a SQL query - a worksheet will be generated for each Select statement </summary>
        /// <param name="strSQL">a SQL Query - one of more select statements</param>
        /// <param name="downloadFileName">(optional) return file name, e.g. myExcelExport.xlsx; if not provided, default to Excel_Export.xlsx</param>
        /// <param name="templateFile">(optional) Excel file used as a template, default to an internal template </param>
        /// ***************************************************************************************************************************************************
        public static void exportExcel(string strSQL, string downloadFileName = null, string templateFile = null)
        {
            ExportParameters exportParam = new ExportParameters();
            exportParam.SqlQuery = strSQL;
            if (downloadFileName != null)
                exportParam.DownloadFileName = downloadFileName;
            if (templateFile != null)
                exportParam.TemplateFile = templateFile;

            exportExcel(exportParam);
        }

        /// ***************************************************************************************************************************************************
        /// <summary>(overload) export to Excel file (xlsx format) from a generic ExportParameters object instace</summary>
        /// <param name="exportParam">an ExportParameters object instance (required parameters: SqlDataSet (or SqlQuery))</param>
        /// <example><code>
        ///     ExportParameters exportParam = new ExportParameters() { SqlDataSet = myDS };    // myDS is a preset DataSet object
        ///     ExcelLib.exportExcel(exportParam);
        /// </code></example>
        /// ***************************************************************************************************************************************************
        public static void exportExcel(ExportParameters exportParam)
        {
            if (exportParam == null || (String.IsNullOrWhiteSpace(exportParam.SqlQuery) && exportParam.SqlDataSet == null))
                return;

            DateTime start = DateTime.Now;

            // log export
            GU.writeLog("EXPORT", System.Reflection.MethodBase.GetCurrentMethod().Name, exportParam.DownloadFileName);

            if (!String.IsNullOrWhiteSpace(exportParam.SqlQuery))
            {
                DataSet ds = new DataSet();
                using (SqlConnection sqlConn = new SqlConnection(exportParam.ConnectionString))
                {
                    SqlDataAdapter adapter = new SqlDataAdapter();
                    adapter.SelectCommand = new SqlCommand(exportParam.SqlQuery, sqlConn);
                    adapter.SelectCommand.CommandTimeout = 3600;
                    // if the query string does not contain spaces and tabs, assume it is a stored procedure
                    if (!exportParam.SqlQuery.Trim().Contains(" ") && !exportParam.SqlQuery.Trim().Contains("\t"))
                        adapter.SelectCommand.CommandType = CommandType.StoredProcedure;
                    adapter.Fill(ds);
                }
                exportParam.SqlDataSet = ds;
            }

            // time check
            timeCheck1 = (DateTime.Now - start).ToString();
            DateTime start2 = DateTime.Now;

            // populate TableParameter
            exp_populateTableParameter(exportParam);

            // check if returnedFile follows the file name restriction (also, it must be a .xlsx file extension)
            exportParam.DownloadFileName = GU.cleanupFileName(exportParam.DownloadFileName, "_");
            // fix file extension if it is not .xlsx
            if (Path.GetExtension(exportParam.DownloadFileName).ToLower() != ".xlsx")
                exportParam.DownloadFileName += ".xlsx";

            // if delay Excel download, create a zip file (named from a GUID) folder (in the default _exportDirectory)
            if (exportParam.excelFileforZip)
            {
                // add GUID (kept in the session variable) at the end of the "general" folder
                exportParam._ExportDirectory = GU.combinePath(exportParam._ExportDirectory, GU.cleanupFileName(String.IsNullOrWhiteSpace(exportParam.ZipDirectory) ? GU.getZipGUID() : exportParam.ZipDirectory));
                exportParam.DownloadFileName = GU.getUnusedFileName(exportParam.DownloadFileName);
                exportParam._serverDownloadFilePath = GU.getPhysicalPath(GU.combinePath(exportParam._ExportDirectory, exportParam.DownloadFileName));
            }
            else    // export only one Excel file
            {
                exportParam._serverDownloadFilePath = GU.getPhysicalPath(GU.combinePath(exportParam._ExportDirectory, Guid.NewGuid().ToString() + ".xlsx"));
            }

            // verify if the default _exportDirectory has already existed
            string expDir = GU.getPhysicalPath(exportParam._ExportDirectory);
            if (!Directory.Exists(expDir))
                Directory.CreateDirectory(expDir);

            // either take the custom templateFile or the default one
            if (!String.IsNullOrEmpty(exportParam.TemplateFile))
                File.Copy(HttpContext.Current.Server.MapPath(exportParam.TemplateFile), exportParam._serverDownloadFilePath, true);
            else
                DefaultExcelTemplate.createFile(exportParam._serverDownloadFilePath);

            // add worksheets to excel
            exp_addWorkSheets(exportParam._serverDownloadFilePath, exportParam);

            // time check
            timeCheck2 = (DateTime.Now - start2).ToString();
            DateTime start3 = DateTime.Now;

            // if delay is requested, wait until all files are ready, user has to export them (probably in a zip file) manually
            if (exportParam.downloadExcelFile)
                downloadExcelFile(exportParam._serverDownloadFilePath, exportParam.DownloadFileName);

            // time check
            timeCheck3 = (DateTime.Now - start3).ToString();
            //DateTime start4 = DateTime.Now;
            //timeCheck4 = (DateTime.Now - start4).ToString();
        }

        #endregion "Excel Export Functions"

        #region "Excel Import Functions"

        // $$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
        // $$$$$$$$$$$$    Import To a Dataset    $$$$$$$$$$$$
        // $$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$

        /// ***************************************************************************************************************************************************
        /// <summary>(overload) import from an Excel file (xlsx &amp; xls format) to a DataSet</summary>
        /// <param name="uploadFile">HttpPostedFile object from the return of a upload control</param>
        /// <param name="oDataSet">output DataSet object</param>
        /// <param name="bImportOnlyFirstWorksheet">(optional) import only the first worksheet (default to false)</param>
        /// <param name="bImportHiddenWorksheets">(optional) import hidden worksheets (default to true)</param>
        /// <param name="bHasHeader">(optional) indicate the first row of a worksheet or definedName is a header row (default to true)</param>
        /// <example><code>ExcelLib.importExcel(FileUpload1.PostedFile); </code></example>
        /// ***************************************************************************************************************************************************
        public static void importExcel(HttpPostedFile uploadFile, out DataSet oDataSet, bool bImportOnlyFirstWorksheet = false, bool bImportHiddenWorksheets = true, bool bHasHeader = true)
        {
            oDataSet = null;
            if (uploadFile == null)
                return;

            ImportParameters importParam = new ImportParameters()
            {
                UploadFile = uploadFile,
                importOnlyFirstWorksheet = bImportOnlyFirstWorksheet,
                importHiddenWorksheets = bImportHiddenWorksheets,
                hasHeader = bHasHeader
            };
            importExcel(importParam, out oDataSet);
        }

        /// ***************************************************************************************************************************************************
        /// <summary>(overload) import from an Excel file (xlsx &amp; xls format) to a DataSet
        ///     <para>code example:</para>
        ///     <para>- DataSet myDS = null;</para>
        ///     <para>- ImportParameters importParam = new ImportParameter() {UploadFile = FileUpload1.PostedFile};</para>
        ///     <para>- ExcelLib.importExcel(importParam, myDS);</para>
        /// </summary>
        /// <param name="importParam">an ImportParameters object instance (required parameters: UploadFile (or ImportFile))</param>
        /// <param name="oDataSet">output DataSet object</param>
        /// <example><code>
        ///     DataSet myDS = null;
        ///     ImportParameters importParam = new ImportParameter() {UploadFile = FileUpload1.PostedFile};
        ///     ExcelLib.importExcel(importParam, myDS);
        /// </code></example>
        /// <remarks>use excelToDataSet if required the DataSet object as a return parameter </remarks>
        /// ***************************************************************************************************************************************************
        public static void importExcel(ImportParameters importParam, out DataSet oDataSet)
        {
            oDataSet = excelToDataSet(importParam);
        }

        /// ***************************************************************************************************************************************************
        /// <summary>import from an Excel file (xlsx &amp; xls format) to a DataSet
        ///         <para>this method is provided for an alternative to get a returned DataSet</para>
        /// </summary>
        /// <param name="importParam">an ImportParameters object instance (required parameters: UploadFile (or ImportFile))</param>
        /// <returns>return a DataSet object with data retrieved from the Excel file</returns>
        /// <remarks>this method is provided for an alternative to get a returned DataSet </remarks>
        /// ***************************************************************************************************************************************************
        public static DataSet excelToDataSet(ImportParameters importParam)
        {
            // no importFile, no dataset!
            if (String.IsNullOrWhiteSpace(importParam.ImportFile) && importParam.UploadFile == null)
                return null;

            //bool bDeleteImportFile = true;

            // generate importFile from uploadFile
            if (importParam.UploadFile != null)
                importParam.ImportFile = uploadFile(importParam.UploadFile);
            //else
            //    // if no upload file, do not delete the importFile (designated from the server)
            //    bDeleteImportFile = false;

            DataSet ds = new DataSet();

            // log import
            GU.writeLog("IMPORT", System.Reflection.MethodBase.GetCurrentMethod().Name, Path.GetFileName(importParam.ImportFile));

            if (Path.GetExtension(importParam.ImportFile).ToLower() == ".xlsx")
                ds = imp_XLSX2Dataset(importParam);
            else    // extension == ".xls"
                ds = imp_XLS2Dataset(importParam);

            // remove the importFile if it was coming from the upload
            if (importParam.deleteImportFile)
            {
                string importPath = GU.getPhysicalPath(importParam.ImportFile);
                try
                {
                    if (File.Exists(importPath))
                        File.Delete(importPath);
                }
                catch (Exception ex)
                {
                    GU.writeLog("ERROR", System.Reflection.MethodBase.GetCurrentMethod().Name, ex.Message);
                }
            }

            return ds;
        }

        // $$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
        // $$$$$$$$$$    Import To a SQL Server    $$$$$$$$$$
        // $$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$

        /// *********************************************************************************************************************************************************************************
        /// <summary>(overload) import from an Excel file (xlsx &amp; xls format) to a SQL Server Table
        ///         <para>* Assume (require) the Excel Worksheet names are the same as the SQL Server "DestinationTableNames", also the columns will be matched </para>
        /// </summary>
        /// <param name="uploadFile">HttpPostedFile object from the return of a upload control</param>
        /// <param name="bImportOnlyFirstWorksheet">(optional) import only the first worksheet (default to false)</param>
        /// <param name="bImportHiddenWorksheets">(optional) import hidden worksheets (default to true)</param>
        /// <param name="bHasHeader">(optional) indicate the first row of a worksheet or definedName is a header row (default to true)</param>
        /// <remarks>Assume (require) the Excel Worksheet names are the same as the SQL Server "DestinationTableNames", also the columns will be matched</remarks>
        /// *********************************************************************************************************************************************************************************
        public static void importExcel(HttpPostedFile uploadFile, bool bImportOnlyFirstWorksheet = false, bool bImportHiddenWorksheets = true, bool bHasHeader = true)
        {
            if (uploadFile == null)
                return;

            ImportParameters importParam = new ImportParameters()
            {
                UploadFile = uploadFile,
                importOnlyFirstWorksheet = bImportOnlyFirstWorksheet,
                importHiddenWorksheets = bImportHiddenWorksheets,
                hasHeader = bHasHeader
            };
            importExcel(importParam);
        }

        /// ***************************************************************************************************************************************************
        /// <summary>(overload) import from an Excel file (xlsx &amp; xls format) to a SQL Server Table
        ///     <para>code example:</para>
        ///     <para>- ImportParameters importParam = new ImportParameter() {UploadFile = FileUpload1.PostedFile};</para>
        ///     <para>- ExcelLib.importExcel(importParam);</para>
        /// </summary>
        /// <param name="importParam">an ImportParameters object instance (required parameters: UploadFile (or ImportFile))</param>
        /// <example>
        ///     ImportParameters importParam = new ImportParameter() {UploadFile = FileUpload1.PostedFile};
        ///     ExcelLib.importExcel(importParam);
        /// </example>
        /// ***************************************************************************************************************************************************
        public static void importExcel(ImportParameters importParam)
        {
            // no importFile, no dataset!
            if (String.IsNullOrWhiteSpace(importParam.ImportFile) && importParam.UploadFile == null)
                return;

            DataSet ds = excelToDataSet(importParam);

            imp_SQLBulkImport(ds, importParam);
        }

        /// ***************************************************************************************************************************************************
        /// <summary>SQL bulk insert from a dataset </summary>
        /// <param name="ds">Input DataSet object to be bulk imported to a SQL Server Database</param>
        /// <param name="importParam">an ImportParameters object instance (require ImportParameters.selectColumns being populated)</param>
        /// ***************************************************************************************************************************************************
        public static void imp_SQLBulkImport(DataSet ds, ImportParameters importParam)
        {
            // check if any source data were generated; if not, do nothing
            if (ds.Tables.Count == 0 || ds.Tables[0].Rows.Count == 0)
                return;

            // sql bulk insert
            using (SqlConnection sqlConn = new SqlConnection(importParam.ConnectionString))
            {
                sqlConn.Open();

                // import all selectExcelTables to sql server
                foreach (SelectExcelTable excelTable in importParam.SelectExcelTables)
                {
                    // get sql table name
                    string dataTableName = String.IsNullOrWhiteSpace(excelTable.DataTableName) ? excelTable.Name : excelTable.DataTableName;
                    string sqlTableName = String.IsNullOrWhiteSpace(excelTable.DestinationTableName) ? dataTableName : excelTable.DestinationTableName;

                    // get a list of column names from targeted table (to be used for filtering out non-matched columns)
                    SqlDataAdapter da = new SqlDataAdapter("SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='" + sqlTableName + "' ORDER BY ORDINAL_POSITION ASC", sqlConn);
                    DataSet dsCol = new DataSet();
                    da.Fill(dsCol);
                    string[] aCol = (from col in dsCol.Tables[0].AsEnumerable()
                                     select col.Field<string>("COLUMN_NAME")).ToArray();

                    using (SqlBulkCopy bc = new SqlBulkCopy(sqlConn, SqlBulkCopyOptions.Default, null))
                    {
                        bc.BulkCopyTimeout = 28800; // seconds
                        // take worksheetName as the default, and use destinationTableName as overwrite
                        bc.DestinationTableName = sqlTableName;

                        // map all columns, so not all columns (including identity key) will be included
                        foreach (SelectColumn selCol in excelTable.SelectColumns)
                        {
                            // ColumnMappings are case sensitive, so the column names have to come directly from each source
                            string colName = String.IsNullOrWhiteSpace(selCol.ColumnAlias) ? selCol.ColumnName : selCol.ColumnAlias;
                            //int colIndex = Array.IndexOf(aCol, colName);
                            //int colIndex = Array.FindIndex(aCol, t => t.IndexOf(colName, StringComparison.InvariantCultureIgnoreCase) >= 0);
                            int colIndex = Array.FindIndex(aCol, t => t.ToLower() == colName.ToLower());
                            // skip if column is not in the database table
                            if (colIndex > -1)
                                bc.ColumnMappings.Add(colName, aCol[colIndex]);
                        }

                        bc.BatchSize = ds.Tables[dataTableName].Rows.Count;
                        bc.WriteToServer(ds.Tables[dataTableName]);
                    }
                }
            }
        }

        #endregion "Excel Import Functions"

        #region "Excel Update Functions"

        /// ***************************************************************************************************************************************************
        /// <summary>update Excel (only for Excel 2010+) - this method only works for web application</summary>
        /// <param name="uploadFile">HttpPostedFile Upload file path</param>
        /// <param name="setCellValues">Dictionary of CellAddress &amp; CellValue that would be inserted/updated to the worksheet </param>
        /// <param name="worksheetName"></param>
        /// ***************************************************************************************************************************************************
        public static void updateExcel(HttpPostedFile uploadFile, Dictionary<string, string> setCellValues, string worksheetName = "")
        {
            UpdateParameter updateParam = new UpdateParameter()
            {
                ClientUploadFile = uploadFile,
                WorksheetName = worksheetName,
                SetCellValues = setCellValues
            };

            updateExcel(updateParam);
        }

        /// ***************************************************************************************************************************************************
        /// <summary>update Excel (only for Excel 2010+), e.g.
        ///     <para>- UpdateParameter updateParameter() {</para>
        ///     <para>-  UploadFile = uploadFile,    // FileUpload object</para>
        ///     <para>-  WorksheetName = "Sheet1",  // worksheet name</para>
        ///     <para>-  SetCellValues = new Dictionary&lt;string, string&gt;() {{"A1","text1"}, {"B2","text2"}}</para>
        ///     <para>- };</para>
        /// </summary>
        /// <exception cref="ExcelLibException">3 exceptions:
        ///     <para>* Error: No Excel file provided.</para>
        ///     <para>* Error: No cell values provided.</para>
        ///     <para>* Error: Excel update does not support XLS (Excel 2003) format.</para>
        /// </exception>
        /// <param name="updateParam">See UpdateParameter for parameter details</param>
        /// <example><code>
        /// UpdateParameter updateParameter() {
        ///     UploadFile = uploadFile,    // FileUpload object
        ///     WorksheetName = "Sheet1",  // worksheet name
        ///     SetCellValues = new Dictionary<string, string>() {{"A1","text1"}, {"B2","text2"}}
        /// };
        /// </code></example>
        /// ***************************************************************************************************************************************************
        public static void updateExcel(UpdateParameter updateParam)
        {
            // no importFile, no dataset!
            if (String.IsNullOrWhiteSpace(updateParam.ServerUpdateFile) && updateParam.ClientUploadFile == null)
                throw new ExcelLibException("Error: No Excel file provided.");

            // generate importFile from uploadFile
            if (updateParam.ClientUploadFile != null)
            {
                updateParam.downloadFileName = Path.GetFileName(updateParam.ClientUploadFile.FileName);
                updateParam.ServerUpdateFile = uploadFile(updateParam.ClientUploadFile);
            }

            // if user does not overwrite this parameter, set default value to true if Excel file is from upload; false if Excel has already in the server
            if (updateParam.downloadExcelFile == null)
                updateParam.downloadExcelFile = !(updateParam.ClientUploadFile == null);

            // throw an exception if no cell values are provided
            if (updateParam.SetCellValues == null)
                throw new ExcelLibException("Error: No cell values provided.");

            if (Path.GetExtension(updateParam.ServerUpdateFile).ToLower() == ".xlsx")
            {
                // log import
                GU.writeLog("UPDATE", System.Reflection.MethodBase.GetCurrentMethod().Name, Path.GetFileName(updateParam.ServerUpdateFile));

                upd_populateCellValues(updateParam);
            }
            else    // extension == ".xls"
                throw new ExcelLibException("Error: Excel update does not support XLS (Excel 2003) format.");

            // download the updated Excel file. The Excel file in server will be deleted after download.
            if (updateParam.downloadExcelFile ?? false)
            {
                downloadExcelFile(updateParam.ServerUpdateFile, updateParam.downloadFileName);
            }
        }

        #endregion "Excel Update Functions"

        #region "More Generic Excel Function"

        /// *******************************************************************************************************************
        /// <summary>(overload) get Excel column Name for specifying cell position</summary>
        /// <param name="columnIndex">input : column order number</param>
        /// <returns>column alphabetic address</returns>
        //   This algorithm was found here:
        //   http://stackoverflow.com/questions/181596/how-to-convert-a-column-number-eg-127-into-an-excel-column-eg-aa
        /// *******************************************************************************************************************
        public static string getColumnAddress(int columnIndex)
        {
            // Given a column number, retrieve the corresponding
            // column string name:
            int value = 0;
            int remainder = 0;
            string result = string.Empty;
            value = columnIndex;

            while ((value > 0))
            {
                remainder = (value - 1) % 26;
                result = (char)(65 + remainder) + result;
                value = Convert.ToInt32(Math.Floor(Convert.ToDouble((value - remainder) / 26)));
            }

            return result;
        }

        /// ************************************************************************************************************
        /// <summary>(overload) get Excel column Name for specifying cell position (semantic overload) </summary>
        /// <param name="cellAddress">column address or cell reference, e.g. "A13", "AS123", "B"</param>
        /// ************************************************************************************************************
        public static string getColumnAddress(string cellAddress)
        {
            // Create a regular expression to match the column name portion of the cell name.
            //System.Text.RegularExpressions.Regex regex = new System.Text.RegularExpressions.Regex("[A-Za-z]+");

            //System.Text.RegularExpressions.Match match = regexAlpha.Match(cellAddress);
            //return match.Value;

            // regular expression is too slow
            char[] characters = cellAddress.ToCharArray();
            List<char> indRow = new List<char>();
            for (int i = 0; i < characters.Length; i++)
            {
                if (characters[i] >= 'A' && characters[i] <= 'z')
                    indRow.Add(characters[i]);
            }

            return new string(indRow.ToArray());
        }

        /// *************************************************************************************************
        /// <summary>get Excel column index for specifying cell position (semantic overload) </summary>
        /// <param name="cellAddress">column address or cell reference, e.g. "A13", "AS123", "B"</param>
        /// *************************************************************************************************
        public static int getColumnIndex(string cellAddress)
        {
            string columnName = getColumnAddress(cellAddress);

            if (String.IsNullOrEmpty(columnName))
            {
                throw new ArgumentNullException("columnName");
            }

            char[] characters = columnName.ToUpperInvariant().ToCharArray();

            int sum = 0;

            for (int i = 0; i <= characters.Length - 1; i++)
            {
                sum *= 26;
                sum += Convert.ToInt32(characters[i]) - Convert.ToInt32('A') + 1;
            }

            return sum;
        }

        /// **********************************************************************************************
        /// <summary>get Excel row index for specifying cell position (semantic overload) </summary>
        /// <param name="cellAddress">cell reference, e.g. "A13", "AS123", "123"</param>
        /// **********************************************************************************************
        public static int getRowIndex(string cellAddress)
        {
            // Create a regular expression to match the row index portion the cell name.
            //System.Text.RegularExpressions.Regex regex = new System.Text.RegularExpressions.Regex("\\d+");

            //System.Text.RegularExpressions.Match match = regexDigit.Match(cellAddress);
            //return int.Parse(match.Value);

            // regular expression is too slow
            char[] characters = cellAddress.ToCharArray();
            List<char> indRow = new List<char>();
            for (int i = 0; i < characters.Length; i++)
            {
                if (characters[i] >= '0' && characters[i] <= '9')
                    indRow.Add(characters[i]);
            }

            return int.Parse(new string(indRow.ToArray()));
        }

        /// ***********************************************************************
        /// <summary>get a list of Worksheet Names from an excel file</summary>
        /// ***********************************************************************
        public static List<string> getWorksheetNames(String fileName, bool excludeHiddenWorksheets = false)
        {
            // Open the spreadsheet document for read-only access.
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(GU.getPhysicalPath(fileName), false))
            {
                return getWorksheetNames(document.WorkbookPart, excludeHiddenWorksheets);
            }
        }

        /// ***********************************************************************
        /// <summary>get a list of Column Names from an excel worksheet</summary>
        /// ***********************************************************************
        public static List<string> getColumnNames(String fileName, string worksheetName, string firstHeaderCell = "A1")
        {
            // Open the spreadsheet document for read-only access.
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(GU.getPhysicalPath(fileName), false))
            {
                return getColumnNames(document.WorkbookPart, worksheetName, firstHeaderCell);
            }
        }

        /// *************************************************************************************************
        /// <summary>validate CellAddress format </summary>
        /// <param name="cellAddress">cell reference, e.g. A1, $A1, A$1, $A$1</param>
        /// *************************************************************************************************
        public static bool validateCellAddress(string cellAddress)
        {
            // the regex only check if less than 7 digit number, need to make sure it's equal to or less than 1048576
            if (regexCellAddress.IsMatch(cellAddress))
                return getRowIndex(cellAddress.Replace("$", "")) <= GenUtils.maxExcelRowNumber;
            else
                return false;
        }

        /// ***************************************************************************************************************
        /// <summary>validate CellRange format </summary>
        /// <param name="cellRange">cell range, e.g. A1, A1:B2, WS!A1, WS!$A$1:B2, 'WS''!'!$A$1:$B$10</param>
        /// ***************************************************************************************************************
        public static bool validateCellRange(string cellRange)
        {
            int indExclamation = cellRange.LastIndexOf('!');
            // see if worksheet/definedname exists before !
            if (indExclamation > -1)
            {
                // need to test worksheet/definedname before !
                string workSheet = cellRange.Left(indExclamation);
                // could be wrapped with single quotes
                if (!validateWorksheetName(workSheet, true))
                    return false;

                cellRange = cellRange.Right(cellRange.Length - indExclamation - 1);
            }

            string[] aCellAddress = cellRange.Split(':');

            // cannot be more than 2 cell addresses
            if (aCellAddress.GetUpperBound(0) > 1)
                return false;

            for (var i = 0; i < aCellAddress.Length; i++)
                if (!validateCellAddress(aCellAddress[i]))
                    return false;

            return true;
        }

        /// ************************************************************************************************
        /// <summary>validate Worksheet name format
        ///     <para>  * Cannot have one of these charaters : \ / ? * [ ]</para>
        ///     <para>  * at least one but less than 31 characters</description></item>
        ///     <para>  * if wrapped around single quotes, all single quotes need to be escaped (double single quotes)</para>
        ///     <para>  * does not support cross workbook reference (like [test.xlsx]Sheet1)</para>
        /// </summary>
        /// <param name="sheetName">worksheet name to be validated</param>
        /// <param name="bCheck4CellRange">(optional) check cell range (default to false) </param>
        /// <remarks><list type="bullet">
        ///     <item><description>Cannot have one of these characters : \ / ? * [ ]</description></item>
        ///     <item><description>at least one but less than 31 characters</description></item>
        ///     <item><description>if wrapped around single quotes, all single quotes need to be escaped (double single quotes)</description></item>
        ///     <item><description>does not support cross workbook reference (like [test.xlsx]Sheet1)</description></item>
        ///     </list>
        /// </remarks>
        /// ************************************************************************************************
        public static bool validateWorksheetName(string sheetName, bool bCheck4CellRange = false)
        {
            // if possible wrapped with single quotes
            if (bCheck4CellRange)
            {
                // any of left-most of right-most character is single quote, both ends should be single quotes
                if ((sheetName.Left(1) == "'" || sheetName.Right(1) == "'") && (sheetName.Left(1) != "'" || sheetName.Right(1) != "'"))
                    return false;

                bool bHasQuotes = false;
                // if there is single quote wrapper, remove them first
                if (sheetName.Left(1) == "'" || sheetName.Right(1) == "'")
                {
                    sheetName = sheetName.Substring(1, sheetName.Length - 2);
                    bHasQuotes = true;
                }

                // if there are non-alphenumeric (and non-underscore) character, it needs to be wrapped with single quotes
                if (!bHasQuotes && !regexWSName.IsMatch(sheetName))
                    return false;

                // if there is a single quote, it needs to be repeated (consider 'WS''''''WS'!A1:B2)
                if (sheetName.Replace("''", "").Contains("'"))
                    return false;
            }

            // at least one but less than 31 characters
            if (String.IsNullOrEmpty(sheetName) || sheetName.Length > 31)
                return false;

            // cannot have one of these characters : \ / ? * [ ]
            char[] invalidChars = new char[] { ':', '\\', '/', '?', '*', '[', ']' };
            if (invalidChars.Any(sheetName.Contains))
                return false;

            return true;
        }

        /// ****************************************************************************************************************************
        ///  <summary>upload file into server - return a server upload path (for later processing and deleting the file)
        ///     <para>* uploaded file will be save to a GUID file name with the same extension</para>
        ///     <para>* it will return a physical path (including file name) on the server where the file is saved</para>
        ///  </summary>
        /// ****************************************************************************************************************************
        public static string uploadFile(HttpPostedFile sfile)
        {
            if (sfile == null)
                return "";

            string sFileName = Path.GetFileName(sfile.FileName);
            // cast datetime to avoid file name conflict
            //string uploadPath = Path.GetFileNameWithoutExtension(sFileName) + "_" + DateTime.Now.ToOADate().ToString().Replace(".", "") + Path.GetExtension(sFileName);
            // using GUID() to write uploaded file on the server
            string uploadPath = Guid.NewGuid().ToString() + "." + Path.GetExtension(sFileName);
            uploadPath = GU.combinePath((string.IsNullOrEmpty(System.Configuration.ConfigurationManager.AppSettings["XLSX_TempFileFolder"])) ? GU.defaultTempFileFolder : System.Configuration.ConfigurationManager.AppSettings["XLSX_TempFileFolder"], uploadPath);
            uploadPath = GU.getPhysicalPath(uploadPath);

            try
            {
                if (File.Exists(uploadPath))
                    File.Delete(uploadPath);
                sfile.SaveAs(uploadPath);
            }
            catch (Exception ex)
            {
                GU.writeLog("ERROR", System.Reflection.MethodBase.GetCurrentMethod().Name, ex.Message);

                if (File.Exists(uploadPath))
                    File.Delete(uploadPath);
                return "";
            }

            return uploadPath;
        }

        /// ********************************************************************************************************
        /// <summary>download Excel file from a redirect (does not remove the file after downalod)</summary>
        /// <param name="excelPath">Excel file path; can be a physical or virtual path</param>
        /// ********************************************************************************************************
        public static void downloadExcelFileFromRedirect(string excelPath)
        {
            string excelVirtualPath = GU.getVirtualPath(excelPath);
            //string excelPhysicalPath = GU.getPhysicalPath(excelPath);

            HttpContext.Current.Response.Redirect(excelVirtualPath, false);
            HttpContext.Current.ApplicationInstance.CompleteRequest();

            //if (!isFileInUse(excelPhysicalPath))
            //    System.IO.File.Delete(excelPhysicalPath);
        }

        /// ***********************************************************************************************************
        /// <summary>download Excel file (and remove it after download)</summary>
        /// <param name="excelPath">Excel file path; can be a physical or virtual path</param>
        /// <param name="sSaveAsFileName">(optional) download file name (default to excelPath's file name)</param>
        /// ***********************************************************************************************************
        public static void downloadExcelFile(string excelPath, string sSaveAsFileName = null)
        {
            try
            {
                excelPath = GU.getPhysicalPath(excelPath);
                if (string.IsNullOrWhiteSpace(sSaveAsFileName))
                    sSaveAsFileName = Path.GetFileName(excelPath);

                HttpContext.Current.Response.Clear();
                HttpContext.Current.Response.BufferOutput = false;

                HttpContext.Current.Response.AddHeader("Pragma", "no-cache, no-store");
                HttpContext.Current.Response.AddHeader("Cache-Control", "no-cache, no-store, must-revalidate, max-age=0");
                HttpContext.Current.Response.AddHeader("Expires", "-1");

                // this is for Excel 2003 and under
                //HttpContext.Current.Response.ContentType = "application/vnd.ms-excel";
                // and this is for Excel 2007 and above
                HttpContext.Current.Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                HttpContext.Current.Response.AddHeader("Content-Disposition", "attachment;filename=\"" + sSaveAsFileName + "\"");
                //HttpContext.Current.Response.AddHeader("Content-Disposition", "inline; filename=\"" + sFile + "\"");

                // open and read the file, and copy it to Response.OutputStream
                using (System.IO.FileStream fs = System.IO.File.OpenRead(excelPath))
                {
                    byte[] b = new byte[1025];
                    Int32 n = new Int32();
                    n = -1;
                    // make sure the client is still connected, or cancel the download
                    while ((HttpContext.Current.Response.IsClientConnected) && (n != 0))
                    {
                        n = fs.Read(b, 0, b.Length);
                        if ((n != 0))
                            HttpContext.Current.Response.OutputStream.Write(b, 0, n);
                    }
                }
                HttpContext.Current.Response.Flush();

                // Prevents any other content from being sent to the browser
                HttpContext.Current.Response.SuppressContent = true;

                // causes the thread to skip past most of the events in the HttpApplication event pipeline and go straight to the final event, named HttpApplication.EventEndRequest
                HttpContext.Current.ApplicationInstance.CompleteRequest();
                //HttpContext.Current.Response.Close();

                if (!GU.isFileInUse(excelPath))
                    System.IO.File.Delete(excelPath);
            }
            catch (Exception ex)
            {
                GU.writeLog("ERROR", System.Reflection.MethodBase.GetCurrentMethod().Name, ex.Message);
            }
        }

        //
        /// ***********************************************************************************************************
        /// <summary>(overload) download zip file (and remove it after download)
        ///         Using the default folder path (GUID) from session variable as the zipFolderPath</summary>
        /// <param name="zipFileName">Zip file Name</param>
        /// <remarks>Using the default folder path (GUID) from session variable as the zipFolderPath </remarks>
        /// ***********************************************************************************************************
        public static void downloadZip(string zipFileName)
        {
            // default base folder + GUID from the session variable
            string zipFolderPath = Path.Combine((String.IsNullOrEmpty(System.Configuration.ConfigurationManager.AppSettings["XLSX_TempFileFolder"])) ? GU.defaultTempFileFolder : System.Configuration.ConfigurationManager.AppSettings["XLSX_TempFileFolder"], GU.getZipGUID());
            downloadZip(zipFolderPath, zipFileName);
        }

        /// **************************************************************************************************************************************************************
        /// <summary>(overload) download zip file (and remove it after download)
        ///         create a zip file from the zipFolderPath into the its parent folder</summary>
        /// <param name="zipFolderPath">Zip folder path, e.g. zipFolderPath="XLSX_Repository/Vendor Open-end" (it will use the folder name as zip file name)</param>
        /// <param name="zipFileName">Zip file Name</param>
        /// <remarks>create a zip file from the zipFolderPath into the its parent folder</remarks>
        /// **************************************************************************************************************************************************************
        public static void downloadZip(string zipFolderPath, string zipFileName)
        {
            // if not provide, use the default download file name
            if (String.IsNullOrWhiteSpace(zipFileName))
                zipFileName = String.IsNullOrEmpty(System.Configuration.ConfigurationManager.AppSettings["XLSX_DefaultDownloadFileName"]) ? "Excel_Export" : System.Configuration.ConfigurationManager.AppSettings["XLSX_DefaultDownloadFileName"];

            // make sure zipFileName ends with .zip
            zipFileName = GU.cleanupFileName(zipFileName.Replace(".ZIP", "", StringComparison.OrdinalIgnoreCase)) + ".zip";

            // get zipFilePath from zipFolderPath (e.g. "XLSX_Repository/Vendor Open-end_{GUID}" will be "XLSX_Repository/Vendor Open-end_{GUID}.zip")
            string zipFilePath = Path.Combine(zipFolderPath, zipFileName);

            // make sure it is a physical path first
            zipFolderPath = GU.getPhysicalPath(zipFolderPath);
            zipFilePath = GU.getPhysicalPath(zipFilePath);

            //--------------------------------------------
            // create zip file from a directory
            //--------------------------------------------
#if NET45
            // To use the ZipFile class, you must reference the System.IO.Compression.FileSystem assembly in your project.
            ZipFile.CreateFromDirectory(zipFolderPath, zipFilePath);
#endif
            // use Ionic.zip library (should be switched to ZipFile.CreateFromDirectory() after upgrade to 4.5)
            GU.createZipFile(zipFolderPath, zipFilePath);

            //--------------------------------------------
            // download zip file through a stream
            //--------------------------------------------
            HttpContext.Current.Response.Clear();
            HttpContext.Current.Response.BufferOutput = false;

            HttpContext.Current.Response.AddHeader("Pragma", "no-cache, no-store");
            HttpContext.Current.Response.AddHeader("Cache-Control", "no-cache, no-store, must-revalidate, max-age=0");
            HttpContext.Current.Response.AddHeader("Expires", "-1");

            HttpContext.Current.Response.ContentType = "application/zip";
            HttpContext.Current.Response.AddHeader("Content-Disposition", "inline; filename=\"" + zipFileName + "\"");

            // open and read the file, and copy it to Response.OutputStream
            using (System.IO.FileStream fs = System.IO.File.OpenRead(zipFilePath))
            {
                byte[] b = new byte[1025];
                Int32 n = new Int32();
                n = -1;
                while ((n != 0))
                {
                    n = fs.Read(b, 0, b.Length);
                    if ((n != 0))
                        HttpContext.Current.Response.OutputStream.Write(b, 0, n);
                }
            }

            // remove zip file and zipped directory
            try
            {
                // delete zip file
                System.IO.File.Delete(zipFilePath);

                // delete zip source folder
                DirectoryInfo di = new DirectoryInfo(zipFolderPath);
                FileInfo[] diar1 = di.GetFiles();
                foreach (FileInfo dra in diar1)
                    System.IO.File.Delete(dra.FullName);
                Directory.Delete(zipFolderPath);
            }
            catch (Exception ex)
            {
                GU.writeLog("ERROR", System.Reflection.MethodBase.GetCurrentMethod().Name, ex.Message);
            }

            HttpContext.Current.Response.Flush();

            // Prevents any other content from being sent to the browser
            HttpContext.Current.Response.SuppressContent = true;

            // causes the thread to skip past most of the events in the HttpApplication event pipeline and go straight to the final event, named HttpApplication.EventEndRequest
            HttpContext.Current.ApplicationInstance.CompleteRequest();
            //HttpContext.Current.Response.Close();

            // reset ZipGUID
            HttpContext.Current.Session["ExcelUtil_ZipGUID"] = null;
        }

#if USING_PACKAGING_FOR_ZIP

        //****************************************************************************************************
        // creat zip from System.IO.Packaging API (doesn't need Ionic.zip DLL)
        //  but it will generate [Content_Types].xml and the file names will be url-encoded (%20 as a space)
        //****************************************************************************************************
        private const long BUFFER_SIZE = 4096;
        public static void createZipFile(string directoryToZip, string zipFilePath)
        {
            // make sure it's a physical path
            zipFilePath = getPhysicalPath(zipFilePath);

            using (Package zip = System.IO.Packaging.Package.Open(zipFilePath, FileMode.OpenOrCreate))
            {
                List<string> sFileList = new List<string>();
                DirectoryInfo di = new DirectoryInfo(directoryToZip);
                FileInfo[] diar1 = di.GetFiles();

                foreach (FileInfo dra in diar1)
                {
                    Uri uri = PackUriHelper.CreatePartUri(new Uri(".\\" + Path.GetFileName(dra.FullName), UriKind.Relative));
                    if (zip.PartExists(uri))
                        zip.DeletePart(uri);

                    PackagePart part = zip.CreatePart(uri, "", CompressionOption.Normal);

                    using (FileStream inputStream = new FileStream(dra.FullName, FileMode.Open, FileAccess.Read))
                    {
                        using (Stream outputStream = part.GetStream())
                        {
                            long bufferSize = inputStream.Length < BUFFER_SIZE ? inputStream.Length : BUFFER_SIZE;
                            byte[] buffer = new byte[bufferSize];
                            int bytesRead = 0;
                            long bytesWritten = 0;
                            while ((InlineAssignHelper(ref bytesRead, inputStream.Read(buffer, 0, buffer.Length))) != 0)
                            {
                                outputStream.Write(buffer, 0, bytesRead);
                                bytesWritten += bufferSize;
                            }
                        }
                    }
                }
            }
        }

#endif

        #endregion "More Generic Excel Function"
    }

    /// **************************************************************************************
    /// <summary>Custom Exception Class for this library</summary>
    /// **************************************************************************************
    [Serializable]
    public class ExcelLibException : System.Exception
    {
        /// <summary>constructor</summary>
        public ExcelLibException() { }

        /// <summary>constructor</summary>
        public ExcelLibException(string message) : base(message) { }

        /// <summary>constructor</summary>
        public ExcelLibException(string message, Exception innerException) : base(message, innerException) { }

        //protected ExcelLibException(SerializationInfo info, StreamingContext context) : base(info, context) { }
    }
}