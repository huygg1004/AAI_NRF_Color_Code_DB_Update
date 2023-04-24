using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using ExcelDataReader;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;

namespace AAI_NRF_Color_Code_DB_Update.Models
{
    public class AAI_File_Upload_Process
    {
        private string connectionString;
        public const string DOCUMENT_CODE = "ItemTable";
        public const string FILE_EXTENSION_IN = "xls";
        public const string FILE_EXTENSION_OUT = "xls";

        private readonly string _inputFolder;
        private readonly string _outputFolder;
        private readonly string _workingFolder;
        private readonly string _failedFolder;
        private readonly string _failedReportedFolder;
        private readonly string _failedSentFolder;
        private readonly string _archiveFilePath;
        private readonly bool _archiveEnabled;
        private readonly string _connectionString;

        public AAI_File_Upload_Process(string tmpFilePath, string userSelectDatabase)
        {

            string BuyerShortCode = "AAI";

            _inputFolder = HttpContext.Current.Server.MapPath("~/AAI/upload");
            _workingFolder = tmpFilePath;
            _outputFolder = HttpContext.Current.Server.MapPath("~/AAI");
            _failedFolder = HttpContext.Current.Server.MapPath("~/AAI/xfailed");
            _failedReportedFolder = HttpContext.Current.Server.MapPath("~/AAI/xfailed/reported");
            _failedSentFolder = HttpContext.Current.Server.MapPath("~/AAI/xfailed");
            _archiveFilePath = HttpContext.Current.Server.MapPath("~/AAI/archive");

            switch (userSelectDatabase)
            {
                case "UAT":
                    _connectionString = System.Configuration.ConfigurationManager.ConnectionStrings["UATDB"].ConnectionString;
                    break;
                case "PROD":
                    _connectionString = System.Configuration.ConfigurationManager.ConnectionStrings["PRODDB"].ConnectionString;
                    break;
                default:
                    _connectionString = System.Configuration.ConfigurationManager.ConnectionStrings["LOCALDB"].ConnectionString;
                    break;
            }

        }

        public Dictionary<string, int> DoWork()
        {
            var inputFiles = (new DirectoryInfo(_inputFolder)).GetFiles($@"*.{FILE_EXTENSION_IN}").ToList();
            var CountOutputDictionary = ImportExcelDataToSql(inputFiles[0].FullName, "sample upc + nrf color report -", "tblItemMaster");

            return CountOutputDictionary;
        }
        public Dictionary<string, int> ImportExcelDataToSql(string filePath, string sheetName, string tableName)
        {
            System.Web.HttpContext.Current.Session["process1"] = "";

            Dictionary<string, int> OutputDictionary = new Dictionary<string, int>();

            // Open the Excel file using ExcelDataReader
            FileStream stream = File.Open(filePath, FileMode.Open, FileAccess.Read);
            IExcelDataReader excelReader = ExcelReaderFactory.CreateReader(stream);

            // Get the dataset containing the Excel data
            DataSet result = excelReader.AsDataSet(new ExcelDataSetConfiguration()
            {
                ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
                {
                    UseHeaderRow = false
                }
            });

            // Get the specified sheet from the dataset
            DataTable table = result.Tables[sheetName];

            // Get the column names from the first row of the sheet
            string[] columnNames = new string[table.Columns.Count];
            for (int i = 0; i < table.Columns.Count; i++)
            {
                columnNames[i] = table.Rows[0][i].ToString();
            }

            string UPC_Number = columnNames[0];
            string NRF_From_Stylemaster_Detail = columnNames[1];

            List<ItemMasterTableList> masterItemList = loadItemMaterTableList(_workingFolder);

            bool IsMissing = masterItemList.Any(item => string.IsNullOrEmpty(item.UPC) || string.IsNullOrEmpty(item.COLORCODE));
            if (IsMissing)
            {
                OutputDictionary.Add("Count Update", 0);
                OutputDictionary.Add("Count Insert", 0);
                return OutputDictionary;
            }



            var masterValueList = masterItemList;

            int totalRecord = 0;
            int countInsert = 0;
            int countUpdate = 0;

            //DATA INSERTION or UPDATE
            foreach (var masterItem in masterValueList)
            {
                SqlConnection connection = new SqlConnection(_connectionString);
                connection.Open();
                SqlCommand recordExist = new SqlCommand("SELECT count(1) FROM tblItemMaster Where BUYERLONGCODE='ARIELASSOCINT' AND ORGANIZATION='" + masterItem.ORGANIZATION + "'" + "AND LABEL ='" + masterItem.LABEL + "'" + "AND UPC ='" + masterItem.UPC + "'" + "AND COLORCODE ='" + masterItem.COLORCODE + "'", connection);

                Int32 countRecordExist = Convert.ToInt32(recordExist.ExecuteScalar());  //check contains key
                if (countRecordExist > 0)
                {
                    //Update Record
                    if (!UpdateDataInMasterTable(connection, masterItem.BUYERLONGCODE, masterItem.ORGANIZATION, masterItem.LABEL, masterItem.UPC, masterItem.COLORCODE, masterItem.LASTMODIFIEDBY))
                    {
                        string error1 = "Update excel record to db failed...SeqID: {0}, EAN: {1}";
                    }
                    countUpdate++;
                }
                else
                {
                    //Insert Record
                    if (!InsertDataIntoMasterTable(connection, masterItem.BUYERLONGCODE, masterItem.ORGANIZATION, masterItem.LABEL, masterItem.UPC, masterItem.COLORCODE, masterItem.LASTMODIFIEDBY))
                    {
                        string error2 = "Insert excel record to db failed...SeqID: {0}, EAN: {1}";
                    }
                    countInsert++;
                }
                connection.Close();
            }

            System.Web.HttpContext.Current.Session["process1"] += "Inserted:" + countInsert + " Updated:" + countUpdate + " Total Record:" + masterValueList.Count() + Environment.NewLine;

            OutputDictionary.Add("Count Update", countUpdate);
            OutputDictionary.Add("Count Insert", countInsert);


            // Close the Excel reader and stream
            excelReader.Close();
            stream.Close();

            return OutputDictionary;
        }

        //Data Model
        private class ItemMasterTableList
        {
            public string BUYERLONGCODE { get; set; }
            public string ORGANIZATION { get; set; }
            public string LABEL { get; set; }
            public string UPC { get; set; }
            public string COLORCODE { get; set; }
            public string LASTMODIFIEDBY { get; set; }
        }

        private List<ItemMasterTableList> loadItemMaterTableList(string inputFile)
        {
            List<ItemMasterTableList> MasterItemList = new List<ItemMasterTableList>();
            try
            {
                HSSFWorkbook workbook;
                using (FileStream file = new FileStream(inputFile, FileMode.Open, FileAccess.Read))
                {
                    workbook = new HSSFWorkbook(file);
                }

                var sheet = workbook.GetSheetAt(0);
                int SeqID = 0;
                for (var i = 1; i <= sheet.LastRowNum; i++)
                {
                    var row = sheet.GetRow(i);
                    if (row == null) continue;

                    string colSender = "AAI";
                    string colUPC = row.GetCell(0).ToString();
                    string colCOLORCODE = row.GetCell(1).ToString();

                    MasterItemList.Add(new ItemMasterTableList()
                    {
                        BUYERLONGCODE = "ARIELASSOCINT",
                        ORGANIZATION = "ARIELASSOCINT",
                        LABEL = "ARIELASSOCINT Item Master",
                        UPC = colUPC,
                        COLORCODE = colCOLORCODE,
                        LASTMODIFIEDBY = "TLO"
                    });

                }
                return MasterItemList;
            }
            catch (Exception e1)
            {
                return MasterItemList;
            }
        }

        private bool InsertDataIntoMasterTable(SqlConnection connection, string BUYERLONGCODE, string ORGANIZATION, string LABEL, string UPC, string COLORCODE, string LASTMODIFIEDBY)
        {
            try
            {
                using (var command = connection.CreateCommand())
                {
                    command.CommandType = CommandType.Text;
                    command.CommandText = @"insert into tblItemMaster (
                                            BUYERLONGCODE,ORGANIZATION, LABEL, UPC, COLORCODE, LASTMODIFIEDBY) values (
                                            @BUYERLONGCODE, @ORGANIZATION, @LABEL, @UPC, @COLORCODE, @LASTMODIFIEDBY)";
                    command.Parameters.AddWithValue("@BUYERLONGCODE", BUYERLONGCODE);
                    command.Parameters.AddWithValue("@ORGANIZATION", ORGANIZATION);
                    command.Parameters.AddWithValue("@LABEL", LABEL);
                    command.Parameters.AddWithValue("@UPC", (UPC == "#N/A") ? Convert.DBNull : UPC);
                    command.Parameters.AddWithValue("@COLORCODE", (COLORCODE == "#N/A") ? Convert.DBNull : COLORCODE);
                    command.Parameters.AddWithValue("@LASTMODIFIEDBY", LASTMODIFIEDBY);

                    //connection.Open();
                    var result = command.ExecuteScalar();
                    //connection.Close();
                }
            }
            catch
            {
                return false;
            }
            return true;
        }

        private bool UpdateDataInMasterTable(SqlConnection connection, string BUYERLONGCODE, string ORGANIZATION, string LABEL, string UPC, string COLORCODE, string LASTMODIFIEDBY)
        {

            try
            {
                using (var command = connection.CreateCommand())
                {
                    command.CommandType = CommandType.Text;
                    command.CommandText = @"UPDATE tblItemMaster SET BUYERLONGCODE=@BUYERLONGCODE,ORGANIZATION=@ORGANIZATION,LABEL = @LABEL, UPC=@UPC,
                                            COLORCODE=@COLORCODE, LASTMODIFIEDBY=@LASTMODIFIEDBY
                                            WHERE BUYERLONGCODE='ARIELASSOCINT' AND ORGANIZATION=@ORGANIZATION AND LABEL=@LABEL AND UPC = @UPC AND COLORCODE=@COLORCODE AND LASTMODIFIEDBY = @LASTMODIFIEDBY";

                    command.Parameters.AddWithValue("@BUYERLONGCODE", BUYERLONGCODE);
                    command.Parameters.AddWithValue("@ORGANIZATION", ORGANIZATION);
                    command.Parameters.AddWithValue("@LABEL", LABEL);
                    command.Parameters.AddWithValue("@UPC", (UPC == "#N/A") ? Convert.DBNull : UPC);
                    command.Parameters.AddWithValue("@COLORCODE", (COLORCODE == "#N/A") ? Convert.DBNull : COLORCODE);
                    command.Parameters.AddWithValue("@LASTMODIFIEDBY", LASTMODIFIEDBY);
                    var result = command.ExecuteScalar();
                }
            }
            catch
            {
                return false;
            }
            return true;
        }

        private string createDateFolder(string archivePath)
        {
            string thisYear = DateTime.Now.ToString("yyyy");
            string thisMth = DateTime.Now.ToString("MM");
            string thisday = DateTime.Now.ToString("dd");

            archivePath = Path.Combine(archivePath, thisYear);
            if (Directory.Exists(archivePath) == false)
            {
                Directory.CreateDirectory(archivePath);
            }
            archivePath = Path.Combine(archivePath, thisMth);
            if (Directory.Exists(archivePath) == false)
            {
                Directory.CreateDirectory(archivePath);
            }
            archivePath = Path.Combine(archivePath, thisday);
            if (Directory.Exists(archivePath) == false)
            {
                Directory.CreateDirectory(archivePath);
            }

            return archivePath;
        }

        public static bool IsNullOrWhiteSpace(String value)
        {
            if (value == null) return true;

            for (int i = 0; i < value.Length; i++)
            {
                if (!Char.IsWhiteSpace(value[i])) return false;
            }

            return true;
        }

        // Modified by Calvin (2020/1/13) - Helper function to log the missing field and quit the program
        private string failedWithMissingColumn(string msg, int rowNumber, string inputFile)
        {
            msg = msg + " at row " + rowNumber;
            //Logger.Error(msg, BuyerShortCode, CorrelationId);
            Console.WriteLine(msg);

            string failedFile = Path.Combine(_failedFolder, Path.GetFileName(inputFile));
            File.Move(inputFile, failedFile);

            System.Environment.Exit(-1);
            return "NULL";
        }
    }
}