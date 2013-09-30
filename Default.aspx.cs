using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data.SqlClient;
using System.Data;
using System.Configuration;
using System.Data.OleDb;
using System.Collections;

public partial class _Default : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {

    }
    protected void btnUpload_Click(object sender, EventArgs e)
    {
        if (fuCSV.HasFile)
        {
            string filepath = Request.PhysicalApplicationPath;
            fuCSV.PostedFile.SaveAs(filepath + "\\File.xlsx");
            DataSet ds = new DataSet();
            string strError = string.Empty;
            ds = GetExcelDataSet(filepath + "\\File.xlsx");

            if (ds != null && ds.Tables.Count > 0)
            {
                bool result = true;
                int ColumSize = ds.Tables[0].Columns.Count;
                DateTime dtCurrent = DateTime.Now;
                for (int rowIndex = 0; rowIndex < ds.Tables[0].Rows.Count; rowIndex++)
                {
                    ArrayList arr3 = new ArrayList();
                    for (int colIndex = 0; colIndex < ColumSize; colIndex++)
                    {
                        arr3.Add(ds.Tables[0].Rows[rowIndex].ItemArray[colIndex].ToString());
                    }
                    if (GetInsert(arr3, "FileImport", ref strError))
                    {
                        lblerror.Text = "File Uploaded Successfully";
                    }
                }
            }
        }
    }

    public DataSet GetExcelDataSet(string filename)
    {
        //This is Provider for normal Excel file 2003
        //OleDbConnection OleDbcnn = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + filename + ";Extended Properties='Excel 8.0;HDR=NO;IMEX=1;'");

        //This is Provider for normal Excel file 2007
        OleDbConnection OleDbcnn = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filename + ";Extended Properties='Excel 12.0;HDR=YES;'");
        try
        {
            DataSet ds = new DataSet();
            OleDbCommand oledbCmd;
            DataTable dt = new DataTable();
            ds = new DataSet();
            if (OleDbcnn.State == ConnectionState.Open)
            {
                OleDbcnn.Close();
            }
            //Open OLEDB connection
            OleDbcnn.Open();
            dt = OleDbcnn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
            for (int index = 0; dt.Rows.Count > index; index++)
            {
                if (Convert.ToString(dt.Rows[index].ItemArray[2]).ToLower() == "sheet1$")
                {
                    oledbCmd = new OleDbCommand("SELECT * FROM [Sheet1$]", OleDbcnn);
                    OleDbDataAdapter oledbDA = new OleDbDataAdapter(oledbCmd);
                    oledbDA.Fill(ds);//Fill data into dataset row by row
                    break;
                }
            }

            OleDbcnn.Close();
            return ds;
        }
        catch (Exception ex)
        {
            throw ex;
        }
        finally
        {
            if (OleDbcnn.State == ConnectionState.Open)
            {
                OleDbcnn.Close();
            }
        }
    }

    public bool GetInsert(ArrayList arr, string strProcName, ref string err)
    {
        bool result = false;
        string strcol = "";
        string strConnection = ConfigurationManager.AppSettings["DBConnection"];
        SqlConnection con = new SqlConnection(strConnection);
        try
        {
            if (con.State == ConnectionState.Open)
            {
                con.Close();
            }
            con.Open();

            SqlCommand strcmmand = new SqlCommand(strProcName, con);
            strcmmand.CommandType = CommandType.StoredProcedure;

            Guid idUser = Guid.NewGuid();
            strcmmand.Parameters.Add(new SqlParameter("@UserID", idUser));
            strcmmand.Parameters.Add(new SqlParameter("@Name", arr[0].ToString().Trim()));
            strcmmand.Parameters.Add(new SqlParameter("@Address", arr[1].ToString().Trim()));           

            strcmmand.ExecuteNonQuery();
            con.Close();
            result = true;
        }
        catch (Exception e)
        {
            err = err + e.Message + "<br/><br/><b>Column</b> " + strcol + "<br/>";
            result = false;

        }
        finally
        {
            if (con.State == ConnectionState.Open)
            {
                con.Close();
            }
        }
        return result;

    }
}