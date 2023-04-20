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
            if (UploadedFile?.ContentLength == null)
            {
                Response.Write("<script>alert('Please choose at least one file');</script>");
                return View("Index");
            }

            if (UploadedFile.ContentLength > 0)
            {
                string FileName = Path.GetFileName(UploadedFile.FileName);

                string FolderPath = Path.Combine(Server.MapPath("~/AAI/upload"), FileName);

                UploadedFile.SaveAs(FolderPath);

                var currentFullFileNamePath = Path.GetFullPath(FileName);
                System.IO.File.Copy(FolderPath, Path.Combine(Server.MapPath("~/AAI/upload/tmp"), FileName), true);

                var tempFilePath = Path.Combine(Server.MapPath("~/AAI/upload/tmp"), FileName);

                var userSelectDatabase = Request.Form["UserSelectDatabase"].ToString(); //this will get selected value

                var payload_AAINRF_Process_Cycle = new AAI_File_Upload_Process(tempFilePath, userSelectDatabase);
                try
                {
                    var CountsDictionary = payload_AAINRF_Process_Cycle.DoWork();

                    string OutputMessage = "";

                    if (CountsDictionary["Count Update"] > 0)
                    {
                        OutputMessage = $"Total {CountsDictionary["Count Update"].ToString()} AAI new records have been updated to {userSelectDatabase.ToString()} database successfully.";
                    }
                    else if (CountsDictionary["Count Insert"] > 0)
                    {
                        OutputMessage = $"Total {CountsDictionary["Count Insert"].ToString()} AAI new records have been inserted to {userSelectDatabase.ToString()} database successfully.";
                    }

                    TempData["MsgChangeStatus"] += OutputMessage;

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


        [HttpPost]
        public ActionResult Delete()
        {
            string connectionString = "";
            var userSelectDatabase = Request.Form["DeleteUserDBSelect"].ToString(); //this will get selected value

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