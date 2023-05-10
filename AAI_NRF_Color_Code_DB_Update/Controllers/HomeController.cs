using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Threading.Tasks;
using System.IO;
using DataSolutions.Logging.Logger;
using AAI_NRF_Color_Code_DB_Update.Models;
using System.Data.SqlClient;
using System.Data;
using System.Web.UI;
using ExcelDataReader;

namespace AAI_NRF_Color_Code_DB_Update.Controllers
{
    public class HomeController : Controller
    {
        [HttpGet]
        public ActionResult Index()
        {
            return View();
        }


        [HttpPost]
        public ActionResult Index(HttpPostedFileBase UploadedFile)
        {
            TempData["MsgChangeStatus"] = "";
            if (UploadedFile?.ContentLength == null)
            {
                Response.Write("<script>alert('Please choose at least one file');</script>");
                return View("Index");
            }

            if (UploadedFile.ContentLength > 0)
            {
                string FileName = Path.GetFileName(UploadedFile.FileName);

                string FolderPath = Path.Combine(Server.MapPath("~/AAI/upload"), FileName);

                UploadedFile.SaveAs(FolderPath); //save the file to the folder

                var currentFullFileNamePath = Path.GetFullPath(FileName);
                System.IO.File.Copy(FolderPath, Path.Combine(Server.MapPath("~/AAI/upload/tmp"), FileName), true);

                var tempFilePath = Path.Combine(Server.MapPath("~/AAI/upload/tmp"), FileName);

                var userSelectDatabase = Request.Form["UserSelectDatabase"].ToString(); //this will get selected value from front end which database the user wants to insert in
                if (string.IsNullOrEmpty(userSelectDatabase))
                {
                    TempData["MsgChangeStatus"] = "No operation because UAT or PROD database is not selected yet";
                    return View("Index");
                }

                //Validating the file if any missing value in columns: UPC or NRF code or both data is missing in a row and display a system message
                using (var stream = new FileStream(FolderPath, FileMode.Open))
                {
                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                    {
                        bool isFirstRow = true; // Flag to skip the first row

                        do
                        {
                            while (reader.Read())
                            {
                                if (isFirstRow)
                                {
                                    isFirstRow = false;
                                    continue; // Skip the first row
                                }

                                string upc = reader.GetString(0);
                                string nrfCode = reader.GetString(1);

                                if (string.IsNullOrEmpty(upc) || string.IsNullOrEmpty(nrfCode))
                                {
                                    // Add code to display a system message
                                    TempData["MsgChangeStatus"] = "The file validation is failed because UPC / NRF code is missing";
                                    return View("Index");
                                }
                            }
                        } while (reader.NextResult()); // Move to the next sheet if any
                    }
                }


                //start the insertion/updating process
                var payload_AAINRF_Process_Cycle = new AAI_File_Upload_Process(tempFilePath, userSelectDatabase);

                try
                {
                    var CountsDictionary = payload_AAINRF_Process_Cycle.DoWork();
                    if (CountsDictionary["Count Update"] == 0 && CountsDictionary["Count Insert"] == 0)
                    {
                        ViewBag.InvalidFile = true;
                        return View("Index");
                    }

                    string OutputMessage = "";

                    if (CountsDictionary["Count Update"] > 0)
                    {
                        OutputMessage = $"Total {CountsDictionary["Count Update"].ToString()} AAI new records have been updated to {userSelectDatabase.ToString()} database successfully.";
                    }
                    else if (CountsDictionary["Count Insert"] > 0)
                    {
                        OutputMessage = $"Total {CountsDictionary["Count Insert"].ToString()} AAI new records have been inserted to {userSelectDatabase.ToString()} database successfully.";
                    }
                    else if (CountsDictionary["Count Update"] == 0 && CountsDictionary["Count Insert"] == 0)
                    {
                        return View("Index");
                    }

                    TempData["MsgChangeStatus"] += OutputMessage;

                    ViewBag.DeleteElements = true;

                    return View("Index");
                }
                catch (Exception ex)
                {
                    TempData["MsgChangeStatus"] += ex.ToString();
                    return View("Index");
                    throw;
                }
            }
           
            return View();
        }

        public ActionResult Delete(string userSelect)
        {
            string connectionString = "";
            if (userSelect == "" || userSelect == null)
            {
                return RedirectToAction("Index");
            }
            var userSelectDatabase = userSelect;

            if (userSelectDatabase == "UAT")
            {
                connectionString = System.Configuration.ConfigurationManager.ConnectionStrings["UATDB"].ConnectionString;
            }
            else if (userSelectDatabase == "PROD")
            {
                connectionString = System.Configuration.ConfigurationManager.ConnectionStrings["PRODDB"].ConnectionString;
            }
            else
            {
                connectionString = System.Configuration.ConfigurationManager.ConnectionStrings["LOCALDB"].ConnectionString;
            }

            SqlConnection connection = new SqlConnection(connectionString);

            string sqlStatement = "DELETE FROM tblItemMaster WHERE BUYERLONGCODE = 'ARIELASSOCINT';";

            try
            {
                connection.Open();
                SqlCommand cmd = new SqlCommand(sqlStatement, connection);
                cmd.CommandType = CommandType.Text;
                int numRows = cmd.ExecuteNonQuery();

                string deleteMessage = "";
                if (userSelectDatabase == "UAT")
                {
                    deleteMessage = $"Total {numRows.ToString()} AAI old records have been removed from UAT database successfully.";
                }
                else if (userSelectDatabase == "PROD")
                {
                    deleteMessage = $"Total {numRows.ToString()} AAI old records have been removed from PROD database successfully.";
                }

                TempData["DeleteStatus"] += deleteMessage;
                TempData["MsgChangeStatus"] = " ";
                System.Web.HttpContext.Current.Session["process1"] = "";

                // Return JSON object with number of records deleted
                return Json(new { numDeleted = numRows });
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message, "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                connection.Close();
            }

            return RedirectToAction("Index");
        }


    }
}