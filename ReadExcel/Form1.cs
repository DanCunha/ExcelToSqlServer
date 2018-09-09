using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data.OleDb;
using System.Data;
using System.Text;
using System.Collections.Generic;
using System.Data.SqlClient;

namespace ReadExcel
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btAbrir_Click(object sender, EventArgs e)
        {
            DataSet ds = new DataSet();

            OpenFileDialog openfileDialog1 = new OpenFileDialog();
            openfileDialog1.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
            if (openfileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            { 
                String filename = openfileDialog1.FileName;
                MessageBox.Show(ImportExcelToSql(filename));
            }
            else
                MessageBox.Show("Erro");        
        }

        public string ImportExcelToSql(string filename)
        {
            String constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filename + ";Extended Properties='Excel 12.0 XML;HDR=YES;';";
            OleDbConnection con = new OleDbConnection(constr);
            con.Open();

            // Get all Sheets in Excel File
            DataTable dtSheet = con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

            DataRow row = dtSheet.Rows[0];

            string sheetName = row["TABLE_NAME"].ToString();
            OleDbCommand oconn = new OleDbCommand("Select * From [" + sheetName + "]", con);

            OleDbDataAdapter sda = new OleDbDataAdapter(oconn);
            DataTable data = new DataTable();
            sda.Fill(data);

            dataGridView1.DataSource = data;

            string codigo = "";
            string protocolo = "";
            string pedido = "";
            string bairro = "";
            string query = "";

            foreach (DataRow dr in data.Rows)
            {
                protocolo = dr[1].ToString();
                bairro = dr[3].ToString();

                if (bairro.Length > 20)
                {
                    return "Tamanho do Bairro maior que 20: " + bairro;
                }

                if(String.IsNullOrEmpty(protocolo) || protocolo == "")
                {
                    return "Existe um protocolo em branco: " + bairro;
                }

                if (String.IsNullOrEmpty(bairro) || bairro == "")
                {
                    return "Existe um Bairro em branco: " + protocolo;
                }
            }

            foreach (DataRow dr in data.Rows)
            {
                //codigo = dr[0].ToString();
                protocolo = dr[1].ToString();
                //pedido = dr[2].ToString();
                bairro = dr[3].ToString();

                //query = "INSERT INTO PROTOCOLOS VALUES(" + codigo + ",'" + protocolo + "','" + pedido + "','" + bairro + "')";
                InsertSql(protocolo, bairro);
            }

            return "Atualização realizada com sucesso";
        }
        
        public void InsertSql(string protocolo, string bairro)
        {
            string connectionString  = "Data Source=NOTEBOOK-SAM;Initial Catalog=Desafio;Integrated Security=True;MultipleActiveResultSets=True;";

            #region
            //using (SqlConnection connection = new SqlConnection(connectionString))
            //{
            //    SqlCommand command = new SqlCommand(queryString, connection);
            //    command.Connection.Open();
            //    command.ExecuteNonQuery();
            //}
            #endregion

            try
            {
                SqlConnection connection = new SqlConnection(connectionString);
                SqlCommand cmd = new SqlCommand();
                cmd = new SqlCommand("SP_ALT_BAIRRO_CLIENTE", connection);
                cmd.Parameters.AddWithValue("@PROTOCOLO", protocolo);
                cmd.Parameters.AddWithValue("@BAIRRO", bairro);
                cmd.CommandType = CommandType.StoredProcedure;

                connection.Open();
                cmd.ExecuteNonQuery();
                connection.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erro: " + ex.Message);
            }
        }

        private DataSet ReadExcelFile(string connectionString)
        {
            DataSet ds = new DataSet();

            using (OleDbConnection conn = new OleDbConnection(connectionString))
            {
                conn.Open();
                OleDbCommand cmd = new OleDbCommand();
                cmd.Connection = conn;

                // Get all Sheets in Excel File
                DataTable dtSheet = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

                // Loop through all Sheets to get data
                foreach (DataRow dr in dtSheet.Rows)
                {
                    string sheetName = dr["TABLE_NAME"].ToString();

                    if (!sheetName.EndsWith("$"))
                        continue;

                    // Get all rows from the Sheet
                    cmd.CommandText = "SELECT * FROM [" + sheetName + "]";

                    DataTable dt = new DataTable();
                    dt.TableName = sheetName;

                    OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                    da.Fill(dt);

                    ds.Tables.Add(dt);
                }

                cmd = null;
                conn.Close();
            }

            return ds;
        }

        private string GetConnectionString(string filename)
        {
            Dictionary<string, string> props = new Dictionary<string, string>();

            // XLSX - Excel 2007, 2010, 2012, 2013
            props["Provider"] = "Microsoft.ACE.OLEDB.12.0;";
            props["Extended Properties"] = "Excel 12.0 XML";
            props["Data Source"] = filename;

            // XLS - Excel 2003 and Older
            //props["Provider"] = "Microsoft.Jet.OLEDB.4.0";
            //props["Extended Properties"] = "Excel 8.0";
            //props["Data Source"] = "C:\\MyExcel.xls";

            StringBuilder sb = new StringBuilder();

            foreach (KeyValuePair<string, string> prop in props)
            {
                sb.Append(prop.Key);
                sb.Append('=');
                sb.Append(prop.Value);
                sb.Append(';');
            }

            return sb.ToString();
        }

        private void AbrirExcel()
        {
            OpenFileDialog openfileDialog1 = new OpenFileDialog();
            if (openfileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                String filename = openfileDialog1.FileName;

                Excel.Application xlApp;
                Excel.Workbook xlWorkBook;
                Excel.Worksheet xlWorkSheet;
                Excel.Range range;

                string str;
                int rCnt;
                int cCnt;
                int rw = 0;
                int cl = 0;

                xlApp = new Excel.Application();
                xlWorkBook = xlApp.Workbooks.Open(filename, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

                range = xlWorkSheet.UsedRange;
                rw = range.Rows.Count;
                cl = range.Columns.Count;


                for (rCnt = 1; rCnt <= rw; rCnt++)
                {
                    for (cCnt = 1; cCnt <= cl; cCnt++)
                    {
                        str = (string)(range.Cells[rCnt, cCnt] as Excel.Range).Value2;
                        MessageBox.Show(str);
                    }
                }

                xlWorkBook.Close(true, null, null);
                xlApp.Quit();

                Marshal.ReleaseComObject(xlWorkSheet);
                Marshal.ReleaseComObject(xlWorkBook);
                Marshal.ReleaseComObject(xlApp);



                //var excelApp = new Microsoft.Office.Interop.Excel.Application();
                //excelApp.Visible = true;
                //excelApp.Workbooks.Open(btAbrir.Text);
            }
        }
    }
}
