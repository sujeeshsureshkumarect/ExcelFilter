using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Data.OleDb;
using System.IO;

namespace ExcelFilter
{
    public partial class ExcelUpload : System.Web.UI.Page
    {
        public void Page_Load(object sender, EventArgs e)
        {
            if(!IsPostBack)
            {
               
            }
            
        }
        DataSet Excelsheet = new DataSet();

        public void btn_Upload_Click(object sender, EventArgs e)
        {
            //Coneection String by default empty  
            string ConStr = "";
            //Extantion of the file upload control saving into ext because   
            //there are two types of extation .xls and .xlsx of Excel   
            string ext = Path.GetExtension(flp_Upload.FileName).ToLower();
            //getting the path of the file   
            string path = Server.MapPath("~/Upload/" + flp_Upload.FileName);
            //saving the file inside the MyFolder of the server  
            flp_Upload.SaveAs(path);
           // Label1.Text = FileUpload1.FileName + "\'s Data showing into the GridView";
            //checking that extantion is .xls or .xlsx  
            if (ext.Trim() == ".xls")
            {
                //connection string for that file which extantion is .xls  
                ConStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + path + ";Extended Properties=\"Excel 8.0;HDR=Yes;IMEX=2\"";
            }
            else if (ext.Trim() == ".xlsx")
            {
                //connection string for that file which extantion is .xlsx  
                ConStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties=\"Excel 12.0;HDR=Yes;IMEX=2\"";
            }
            Excelsheet= Parse(flp_Upload.FileName, ext, path);

            Session["mydataset"] = Excelsheet;

            for (int i=0;i<Excelsheet.Tables.Count;i++)
            {
                drp_Sheet.Items.Add(new ListItem(Excelsheet.Tables[i].TableName.ToString(), i.ToString()));
            }

            for (int j = 0; j < Excelsheet.Tables[0].Columns.Count; j++)
            {
                drp_Column.Items.Add(new ListItem(Excelsheet.Tables[0].Columns[j].ColumnName.ToString(), j.ToString()));
            }
        }
        static DataSet Parse(string fileName,string ext,string path)
        {
            string connectionString = string.Format("provider=Microsoft.Jet.OLEDB.4.0; data source={0};Extended Properties=Excel 8.0;", fileName);

            if (ext.Trim() == ".xls")
            {
                //connection string for that file which extantion is .xls  
                connectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + path + ";Extended Properties=\"Excel 8.0;HDR=Yes;IMEX=2\"";
            }
            else if (ext.Trim() == ".xlsx")
            {
                //connection string for that file which extantion is .xlsx  
                connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties=\"Excel 12.0;HDR=Yes;IMEX=2\"";
            }
            DataSet data = new DataSet();

            foreach (var sheetName in GetExcelSheetNames(connectionString))
            {
                string str = sheetName.Substring(sheetName.Length - 1);
                if(str!="_")
                {
                    using (OleDbConnection con = new OleDbConnection(connectionString))
                    {
                        string founderMinus1 = sheetName.Remove(sheetName.Length - 1, 1);
                        var dataTable = new DataTable(founderMinus1);
                        string query = string.Format("SELECT * FROM [{0}]", sheetName);
                        con.Open();
                        OleDbDataAdapter adapter = new OleDbDataAdapter(query, con);
                        adapter.Fill(dataTable);
                        data.Tables.Add(dataTable);
                    }
                }               
            }
            return data;
        }

        static string[] GetExcelSheetNames(string connectionString)
        {
            OleDbConnection con = null;
            DataTable dt = null;
            con = new OleDbConnection(connectionString);
            con.Open();
            dt = con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });

            if (dt == null)
            {
                return null;
            }

            String[] excelSheetNames = new String[dt.Rows.Count];
            int i = 0;

            foreach (DataRow row in dt.Rows)
            {
                excelSheetNames[i] = row["TABLE_NAME"].ToString();
                i++;
            }
            return excelSheetNames;
        }

        public void btn_Retrieve_Click(object sender, EventArgs e)
        {
            DataSet ds = (DataSet)Session["mydataset"];

            DataTable dt = ds.Tables[Convert.ToInt32(drp_Sheet.SelectedItem.Value)];
            string column = drp_Column.SelectedItem.Text;

            string[] selectedColumns = new[] {column};

            DataTable dt1 = new DataView(dt).ToTable(false, selectedColumns);

            GridView1.DataSource = dt1;
            GridView1.DataBind();
        }

        protected void drp_Sheet_SelectedIndexChanged(object sender, EventArgs e)
        {
            DataSet ds = (DataSet)Session["mydataset"];
            drp_Column.Items.Clear();

            for (int j = 0; j < ds.Tables[Convert.ToInt32(drp_Sheet.SelectedItem.Value)].Columns.Count; j++)
            {
                drp_Column.Items.Add(new ListItem(ds.Tables[Convert.ToInt32(drp_Sheet.SelectedItem.Value)].Columns[j].ColumnName.ToString(), j.ToString()));
            }
        }
    }
}