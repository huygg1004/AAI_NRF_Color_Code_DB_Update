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

                var payload_AAINRF_Process_Cycle = new AAI_File_Upload_Process(tempFilePath);
                try
                {
                    payload_AAINRF_Process_Cycle.DoWork();

                    TempData["MsgChangeStatus"] += "Records have been successfully inserted into DB";

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
        public ActionResult Delete()
        {
            string connectionString = "Server = localhost; Database = TLO20PSUAT; Trusted_Connection = True; ";

            SqlConnection connection = new SqlConnection(connectionString);

            string sqlStatement = "DELETE FROM tblItemMaster;";

            try
            {
                connection.Open();
                SqlCommand cmd = new SqlCommand(sqlStatement, connection);
                cmd.CommandType = CommandType.Text;
                int numRows = cmd.ExecuteNonQuery();
                TempData["DeleteStatus"] += "Records have been removed successfully from DB.";
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