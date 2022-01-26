using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Common;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace _Git_
{
    public partial class February : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

        }

        protected void UploadExcel_Click(object sender, EventArgs e)

        {

            if (FileUploadExcel.HasFile)
            {
                try
                {
                    //Save File that have uploaded
                    FileUploadExcel.SaveAs(@"C:\Users\BSPT0765\source\repos\_Git_\_Git_\ExcelFile\" + FileUploadExcel.FileName);
                    string path1 = @"C:\Users\BSPT0765\source\repos\_Git_\_Git_\ExcelFile\" + FileUploadExcel.FileName;
                    string path = (Server.MapPath("/ExcelFile/" + FileUploadExcel.FileName));
                    FileUploadExcel.SaveAs(path);

                    string ExcelConnString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + path1 + ";Extended Properties=Excel 4.0;";
                    OleDbConnection OleDBcon = new OleDbConnection(ExcelConnString);

                    //Using ole db to read file from excel 
                    OleDbCommand cmd = new OleDbCommand("Select * from [June$] WHERE [Warehouse] IS NOT NULL", OleDBcon);
                    OleDbDataAdapter objAdapter1 = new OleDbDataAdapter(cmd);
                    DataSet ds = new DataSet();
                    objAdapter1.Fill(ds);

                    OleDBcon.Open();

                    // Create DbDataReader to Data Worksheet  
                    DbDataReader dr = cmd.ExecuteReader();

                    // SQL Server Connection String  
                    string ConnStr = "Data Source=HMSB-M1LGPT0698\\SQLEXPRESS;Initial Catalog=git; Persist Security Info=True; User ID=Git;Password=Honda12345";

                    // Bulk Copy to SQL Server
                    SqlBulkCopy bulkInsert = new SqlBulkCopy(ConnStr);

                    try
                    {
                        bulkInsert.DestinationTableName = "feb";
                        
                        SqlBulkCopyColumnMapping warehouse_ = new SqlBulkCopyColumnMapping("Warehouse", "warehouse");
                        bulkInsert.ColumnMappings.Add(warehouse_);

                        SqlBulkCopyColumnMapping containercapacity_ = new SqlBulkCopyColumnMapping("Container Capacity", "containercapacity");
                        bulkInsert.ColumnMappings.Add(containercapacity_);

                        SqlBulkCopyColumnMapping date_ = new SqlBulkCopyColumnMapping("Date", "date");
                        bulkInsert.ColumnMappings.Add(date_);

                        SqlBulkCopyColumnMapping incoming_ = new SqlBulkCopyColumnMapping("Incoming", "incoming");
                        bulkInsert.ColumnMappings.Add(incoming_);

                        SqlBulkCopyColumnMapping outcoming_ = new SqlBulkCopyColumnMapping("Outgoing", "outgoing");
                        bulkInsert.ColumnMappings.Add(outcoming_);

                        SqlBulkCopyColumnMapping totalstorage_ = new SqlBulkCopyColumnMapping("Total Storage", "totalstorage");
                        bulkInsert.ColumnMappings.Add(totalstorage_);

                        SqlBulkCopyColumnMapping maxcapacity_ = new SqlBulkCopyColumnMapping("Max Capacity", "maximumcapacity");
                        bulkInsert.ColumnMappings.Add(maxcapacity_);

                        SqlBulkCopyColumnMapping accord_ = new SqlBulkCopyColumnMapping("ACCORD", "accord");
                        bulkInsert.ColumnMappings.Add(accord_);

                        SqlBulkCopyColumnMapping civic_ = new SqlBulkCopyColumnMapping("CIVIC", "civic");
                        bulkInsert.ColumnMappings.Add(civic_);

                        SqlBulkCopyColumnMapping crv_ = new SqlBulkCopyColumnMapping("CRV", "crv");
                        bulkInsert.ColumnMappings.Add(crv_);

                        SqlBulkCopyColumnMapping jazz_ = new SqlBulkCopyColumnMapping("JAZZ", "jazz");
                        bulkInsert.ColumnMappings.Add(jazz_);

                        SqlBulkCopyColumnMapping city_ = new SqlBulkCopyColumnMapping("CITY", "city");
                        bulkInsert.ColumnMappings.Add(city_);

                        SqlBulkCopyColumnMapping hrv_ = new SqlBulkCopyColumnMapping("HRV", "hrv");
                        bulkInsert.ColumnMappings.Add(hrv_);

                        SqlBulkCopyColumnMapping brv_ = new SqlBulkCopyColumnMapping("BRV", "brv");
                        bulkInsert.ColumnMappings.Add(brv_);

                        SqlBulkCopyColumnMapping jh_ = new SqlBulkCopyColumnMapping("JAZZ HEV", "jazzhev");
                        bulkInsert.ColumnMappings.Add(jh_);

                        SqlBulkCopyColumnMapping ch_ = new SqlBulkCopyColumnMapping("CITY HEV", "cityhev");
                        bulkInsert.ColumnMappings.Add(ch_);

                        SqlBulkCopyColumnMapping hh_ = new SqlBulkCopyColumnMapping("HRV HEV", "hrvhev");
                        bulkInsert.ColumnMappings.Add(hh_);





                        bulkInsert.WriteToServer(dr);

                    }

                    catch (Exception labelex)
                    {
                        Response.Write("<script>alert('Data Duplicated/Error')</script>");
                        Response.Write(labelex);
                    }

                    OleDBcon.Close();

                    Response.Write("<script>alert('Data Inserted Successfully.')</script>");
                    UploadLabel.Text = FileUploadExcel.FileName + " Uploaded";

                }

                catch (Exception)
                {
                    Response.Write("<script>alert('Error occured. Data may duplicated.')</script>");
                    UploadLabel.Text = "Error while uploading excel file.";
                }
            }

            else
            {
                UploadLabel.Text = "No file uploaded";
            }

        }
    }
}
