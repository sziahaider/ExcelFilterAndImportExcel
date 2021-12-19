using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelDataReader;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using MaterialSkin.Controls;


namespace Excel
{
    public partial class frmImport : MaterialForm
    {
        public OleDbConnection con;

        //public void pintu(string s)
        //{
        //    con = new OleDbConnection("provider=Microsoft.ACE.OLEDB.12.0;Data source=" + textBox1.Text + ";Extended properties=\"Excel 8.0 Xml;HDR=Yes;\";");
        //}
        public frmImport()
        {
            InitializeComponent();

            MaterialSkin.MaterialSkinManager manager = MaterialSkin.MaterialSkinManager.Instance;
            manager.AddFormToManage(this);
            manager.Theme = MaterialSkin.MaterialSkinManager.Themes.LIGHT;
            manager.ColorScheme = new MaterialSkin.ColorScheme(MaterialSkin.Primary.Blue300,
                MaterialSkin.Primary.Blue500, MaterialSkin.Primary.Blue500, MaterialSkin.Accent.LightBlue400,
                MaterialSkin.TextShade.WHITE);


        }
       
        //Load function
        private void button2_Click(object sender, EventArgs e)
        {
            
        }

        //Choose File function
        private void button1_Click(object sender, EventArgs e)
        {
            
        }
        
        //Import Data function
        private void button3_Click(object sender, EventArgs e)
        {
         

        }
        SqlConnection sql = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlcommand"].ConnectionString);

        public object MaterialSkinManager { get; }

        //Import Data function
        public void ImportExcel()
        {
            if (materialComboBox1.SelectedItem.ToString()=="Rate")
            {
                try
                {
                    for (int j = 0; j < dataGridView1.Rows.Count; j++)
                    {
                        if (dataGridView1.Rows[j].Cells[0].Value.ToString() != "")
                        {
                            SqlCommand cmd = new SqlCommand(@"sp_insert_rate", sql);
                            cmd.CommandType = CommandType.StoredProcedure;
                            cmd.Parameters.AddWithValue("@p_channel", dataGridView1.Rows[j].Cells[0].Value.ToString());
                            cmd.Parameters.AddWithValue("@p_hour", dataGridView1.Rows[j].Cells[1].Value.ToString());
                            cmd.Parameters.AddWithValue("@p_monday", dataGridView1.Rows[j].Cells[2].Value.ToString());
                            cmd.Parameters.AddWithValue("@p_tuesday", dataGridView1.Rows[j].Cells[3].Value.ToString());
                            cmd.Parameters.AddWithValue("@p_wednesday", dataGridView1.Rows[j].Cells[4].Value.ToString());
                            cmd.Parameters.AddWithValue("@p_thursday", dataGridView1.Rows[j].Cells[5].Value.ToString());
                            cmd.Parameters.AddWithValue("@p_friday", dataGridView1.Rows[j].Cells[6].Value.ToString());
                            cmd.Parameters.AddWithValue("@p_saturday", dataGridView1.Rows[j].Cells[7].Value.ToString());
                            cmd.Parameters.AddWithValue("@p_sunday", dataGridView1.Rows[j].Cells[8].Value.ToString());
                            sql.Open();
                            cmd.ExecuteNonQuery();
                            sql.Close();
                        
                        }
                       
                    }
                    MessageBox.Show("Imported Successfully");
                }
                catch (Exception ex)
                {

                    MessageBox.Show(ex.Message);
                }
                finally
                {
                    if (sql.State == ConnectionState.Open)
                    {
                        sql.Close();
                    }
                }
               
            }
            else if (materialComboBox1.SelectedItem.ToString() == "Data")
            {
                try
                {
                    for (int j = 1; j < dataGridView1.Rows.Count; j++)
                    {
                        if (dataGridView1.Rows[j].Cells[0].Value.ToString() != "")
                        {
                            SqlCommand cmd = new SqlCommand(@"sp_insert_data", sql);
                            cmd.CommandType = CommandType.StoredProcedure;
                            cmd.Parameters.AddWithValue("@p_channel", dataGridView1.Rows[j].Cells[0].Value.ToString());
                            cmd.Parameters.AddWithValue("@p_date", dataGridView1.Rows[j].Cells[1].Value.ToString());
                            cmd.Parameters.AddWithValue("@p_day", dataGridView1.Rows[j].Cells[2].Value.ToString());
                            cmd.Parameters.AddWithValue("@p_caption", dataGridView1.Rows[j].Cells[3].Value.ToString());
                            cmd.Parameters.AddWithValue("@p_brand", dataGridView1.Rows[j].Cells[4].Value.ToString());
                            cmd.Parameters.AddWithValue("@p_subcategory", dataGridView1.Rows[j].Cells[5].Value.ToString());
                            cmd.Parameters.AddWithValue("@p_category", dataGridView1.Rows[j].Cells[6].Value.ToString());
                            cmd.Parameters.AddWithValue("@p_companyname", dataGridView1.Rows[j].Cells[7].Value.ToString());
                            cmd.Parameters.AddWithValue("@p_agencyName", dataGridView1.Rows[j].Cells[8].Value.ToString());
                            cmd.Parameters.AddWithValue("@p_adtype", dataGridView1.Rows[j].Cells[9].Value.ToString());
                            cmd.Parameters.AddWithValue("@p_adstart", dataGridView1.Rows[j].Cells[10].Value.ToString());
                            cmd.Parameters.AddWithValue("@p_adend", dataGridView1.Rows[j].Cells[11].Value.ToString());
                            cmd.Parameters.AddWithValue("@p_transmissionhour", dataGridView1.Rows[j].Cells[12].Value.ToString());
                            cmd.Parameters.AddWithValue("@p_timeband", dataGridView1.Rows[j].Cells[13].Value.ToString());
                            cmd.Parameters.AddWithValue("@p_addurationinmins", dataGridView1.Rows[j].Cells[14].Value.ToString());
                            cmd.Parameters.AddWithValue("@p_addurationinsec", dataGridView1.Rows[j].Cells[15].Value.ToString());
                            cmd.Parameters.AddWithValue("@p_slotposition", dataGridView1.Rows[j].Cells[16].Value.ToString());
                            cmd.Parameters.AddWithValue("@p_slotcount", dataGridView1.Rows[j].Cells[17].Value.ToString());
                            cmd.Parameters.AddWithValue("@p_slotposition2", dataGridView1.Rows[j].Cells[18].Value.ToString());
                            cmd.Parameters.AddWithValue("@p_slotcount2", dataGridView1.Rows[j].Cells[19].Value.ToString());
                            cmd.Parameters.AddWithValue("@p_soltpositionall", dataGridView1.Rows[j].Cells[20].Value.ToString());
                            cmd.Parameters.AddWithValue("@p_slotcountall", dataGridView1.Rows[j].Cells[21].Value.ToString());
                            cmd.Parameters.AddWithValue("@p_adslot", dataGridView1.Rows[j].Cells[22].Value.ToString());
                            cmd.Parameters.AddWithValue("@p_netamount", dataGridView1.Rows[j].Cells[23].Value.ToString());
                            cmd.Parameters.AddWithValue("@p_programname", dataGridView1.Rows[j].Cells[24].Value.ToString());
                            cmd.Parameters.AddWithValue("@p_programtype", dataGridView1.Rows[j].Cells[25].Value.ToString());
                            cmd.Parameters.AddWithValue("@p_programstart", dataGridView1.Rows[j].Cells[26].Value.ToString());
                            cmd.Parameters.AddWithValue("@p_programend", dataGridView1.Rows[j].Cells[27].Value.ToString());
                            cmd.Parameters.AddWithValue("@p_midbreak", dataGridView1.Rows[j].Cells[28].Value.ToString());
                            cmd.Parameters.AddWithValue("@p_breakstart", dataGridView1.Rows[j].Cells[29].Value.ToString());
                            cmd.Parameters.AddWithValue("@p_breakend", dataGridView1.Rows[j].Cells[30].Value.ToString());
                            cmd.Parameters.AddWithValue("@p_quality", dataGridView1.Rows[j].Cells[31].Value.ToString());
                            cmd.Parameters.AddWithValue("@p_clienttype", dataGridView1.Rows[j].Cells[32].Value.ToString());
                            cmd.Parameters.AddWithValue("@p_customcaptionname", dataGridView1.Rows[j].Cells[33].Value.ToString());
                            cmd.Parameters.AddWithValue("@p_customduration", dataGridView1.Rows[j].Cells[34].Value.ToString());
                            cmd.Parameters.AddWithValue("@p_commercialmarkid", dataGridView1.Rows[j].Cells[35].Value.ToString());
                            sql.Open();
                            cmd.ExecuteNonQuery();
                            sql.Close();
                         
                        }
                      
                    }
                    MessageBox.Show("Imported Successfully");
                }
                catch (Exception ex)
                {

                    MessageBox.Show(ex.Message);
                }
                finally
                {
                    if (sql.State == ConnectionState.Open)
                    {
                        sql.Close();
                    }
                }
            }
            else if (materialComboBox1.SelectedItem.ToString()=="Directories")
            {
                try
                {
                    for (int j = 0; j < dataGridView1.Rows.Count; j++)
                    {
                        if (dataGridView1.Rows[j].Cells[0].Value.ToString() != "")
                        {
                            SqlCommand cmd = new SqlCommand(@"sp_insert_directories", sql);
                            cmd.CommandType = CommandType.StoredProcedure;
                            cmd.Parameters.AddWithValue("@p_channel", dataGridView1.Rows[j].Cells[0].Value.ToString());
                            cmd.Parameters.AddWithValue("@p_companyname", dataGridView1.Rows[j].Cells[1].Value.ToString());
                            cmd.Parameters.AddWithValue("@p_brand", dataGridView1.Rows[j].Cells[2].Value.ToString());
                            cmd.Parameters.AddWithValue("@p_subcategory", dataGridView1.Rows[j].Cells[3].Value.ToString());
                            cmd.Parameters.AddWithValue("@p_category", dataGridView1.Rows[j].Cells[4].Value.ToString());
                            cmd.Parameters.AddWithValue("@p_agencyName", dataGridView1.Rows[j].Cells[5].Value.ToString());
                            cmd.Parameters.AddWithValue("@p_adtype", dataGridView1.Rows[j].Cells[6].Value.ToString());
                            cmd.Parameters.AddWithValue("@p_transmissionhour", dataGridView1.Rows[j].Cells[7].Value.ToString());
                            cmd.Parameters.AddWithValue("@p_timeband", dataGridView1.Rows[j].Cells[8].Value.ToString());
                            sql.Open();
                            cmd.ExecuteNonQuery();
                            sql.Close();
                            

                        }
                       
                    }
                    MessageBox.Show("Imported Successfully");
                }
                catch (Exception ex)
                {

                    MessageBox.Show(ex.Message);
                }
                finally
                {
                    if (sql.State == ConnectionState.Open)
                    {
                        sql.Close();
                    }
                }
            }


        }

        private void Form1_Load(object sender, EventArgs e)
        {
            materialComboBox1.Text = "Please select";
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void materialButton2_Click(object sender, EventArgs e)
        {
           
                string path = materialTextBox1.Text;
               // var stream = File.Open(path, FileMode.Open, FileAccess.Read);
                var stream = new FileStream(path,FileMode.Open,FileAccess.Read);
                var reader = ExcelReaderFactory.CreateReader(stream);
                var result = reader.AsDataSet();
                DataTable Exceldt = result.Tables[0];
                dataGridView1.DataSource = Exceldt;
            
  
        }

        private void materialButton3_Click(object sender, EventArgs e)
        {
            ImportExcel();
        }

        private void materialButton1_Click(object sender, EventArgs e)
        {
            OpenFileDialog openfiledialog1 = new OpenFileDialog();
            openfiledialog1.ShowDialog();
            openfiledialog1.Filter = "allfiles|*.xlxs";
            materialTextBox1.Text = openfiledialog1.FileName;
            Microsoft.Office.Interop.Excel.Application ExcelObj = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook theWorkbook = null;
            string strPath = materialTextBox1.Text;
            theWorkbook = ExcelObj.Workbooks.Open(strPath);
            Microsoft.Office.Interop.Excel.Sheets sheets = theWorkbook.Worksheets;
            for (int i = 1; i <= sheets.Count; i++)
            {
                Microsoft.Office.Interop.Excel.Worksheet worksheet = (Microsoft.Office.Interop.Excel.Worksheet)sheets.get_Item(i);//Get the reference of second worksheet
                string strWorksheetName = worksheet.Name;//Get the name of worksheet.
                materialComboBox1.Items.Add(strWorksheetName);
            }

            theWorkbook.Close(0);
        }

        private void materialButton5_Click(object sender, EventArgs e)
        {
            frmMain main = new frmMain();
            main.Show();
            this.Hide();

        }
    }
}
