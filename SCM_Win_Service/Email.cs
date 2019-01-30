using System;
using System.Configuration;
using System.Data;
using System.IO;
using System.Web;
using System.Net;
using System.Net.Mail;
using System.Reflection;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace SCM_Win_Service
{
    public static class Email
    {
        public static string[] networkCredentials = ConfigurationManager.AppSettings["NetworkCredentials"].ToString().Split(',');
        private static string fromAddress = ConfigurationManager.AppSettings["FromAddress"].ToString();
        private static string port = ConfigurationManager.AppSettings["EmailPort"].ToString();
        private static string hostName = ConfigurationManager.AppSettings["HostName"].ToString();

        private static string FilePath = ConfigurationManager.AppSettings["ExcelSavaPath"].ToString();

        public static void ExportDataSet(DataSet ds, string email, string buyerName)
        {
            string savedFilePath = string.Empty;

            try
            {
                savedFilePath = ExportDataSetToExcel(ds, buyerName);

                using (var eMail = new MailMessage())
                {
                    eMail.To.Add(email);
                    eMail.From = new MailAddress(fromAddress);
                    eMail.CC.Add(new MailAddress(fromAddress));

                    eMail.Subject = "List of SCM Data's";
                    Attachment attachment = new System.Net.Mail.Attachment(savedFilePath);
                    eMail.Attachments.Add(attachment);
                    eMail.Body = "";
                    SmtpClient MailClient = new SmtpClient();
                    MailClient.Host = hostName;
                    NetworkCredential NC = new NetworkCredential(networkCredentials[0].ToString(), networkCredentials[1].ToString());

                    MailClient.UseDefaultCredentials = true;
                    MailClient.Credentials = NC;
                    MailClient.EnableSsl = true;
                    MailClient.Port = Convert.ToInt32(port);
                    MailClient.Send(eMail);
                }
            }
            catch (Exception ex)
            {
                Logger.LogError(ex, "Exception from mail sending");
            }

            if (File.Exists(savedFilePath))
            {
                File.Delete(savedFilePath);
            }
        }

        private static string ExportDataSetToExcel(DataSet ds, string buyerName)
        {
            string savedFilePath = string.Empty;

            string fileName = "SCM_" + buyerName + "_" + System.DateTime.UtcNow.ToString("ddMMyyyyhhmmss");
            //string fileLocalPath = @"F:\CodeBase\Engineering\Velan_SCM_Windows_Service\SCM_Win_Service\SCM_Win_Service\EmailExcel\";

            string path = FilePath + fileName + ".xlsx";

            ApplicationClass ExcelApp = new ApplicationClass();
            Workbook xlWorkbook = ExcelApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);

            try

            {
                // Loop over DataTables in DataSet.
                DataTableCollection collection = ds.Tables;

                for (int i = collection.Count; i > 0; i--)
                {
                    Sheets xlSheets = null;
                    Worksheet xlWorksheet = null;

                    //Create Excel Sheets

                    xlSheets = ExcelApp.Sheets;

                    xlWorksheet = (Worksheet)xlSheets.Add(xlSheets[1],

                                   Type.Missing, Type.Missing, Type.Missing);



                    System.Data.DataTable table = collection[i - 1];

                    xlWorksheet.Name = table.TableName;

                    for (int j = 1; j < table.Columns.Count + 1; j++)
                    {
                        ExcelApp.Cells[1, j] = table.Columns[j - 1].ColumnName;
                    }

                    // Storing Each row and column value to excel sheet

                    for (int k = 0; k < table.Rows.Count; k++)
                    {
                        for (int l = 0; l < table.Columns.Count; l++)
                        {
                            ExcelApp.Cells[k + 2, l + 1] =

                            table.Rows[k].ItemArray[l].ToString();
                        }
                    }

                    ExcelApp.Columns.AutoFit();
                }

                ((Worksheet)ExcelApp.ActiveWorkbook.Sheets[ExcelApp.ActiveWorkbook.Sheets.Count]).Delete();

                ExcelApp.ActiveWorkbook.SaveAs(path);
            }
            catch (Exception ex)
            {
                Logger.LogError(ex, "Exception from converting to dataset into excel");
            }
            finally
            {
                ExcelApp = null;
                Marshal.ReleaseComObject(xlWorkbook);
                xlWorkbook.Close(false, path, null);
            }

            return path;
        }
    }
}
