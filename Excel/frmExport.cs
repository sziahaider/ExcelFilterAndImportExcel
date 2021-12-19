using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using MaterialSkin.Controls;

namespace Excel
{
    public partial class frmExport : MaterialForm
    {
        public frmExport()
        {
            InitializeComponent();
            MaterialSkin.MaterialSkinManager manager = MaterialSkin.MaterialSkinManager.Instance;
            manager.AddFormToManage(this);
            manager.Theme = MaterialSkin.MaterialSkinManager.Themes.LIGHT;
            manager.ColorScheme = new MaterialSkin.ColorScheme(MaterialSkin.Primary.Blue300,
                MaterialSkin.Primary.Blue500, MaterialSkin.Primary.Blue500, MaterialSkin.Accent.LightBlue400,
                MaterialSkin.TextShade.WHITE);
        }

        SqlConnection sql = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlcommand"].ConnectionString);
      
        private void materialButton1_Click_1(object sender, EventArgs e)
        {
            try
            {
                sql.Open();
                DataTable dt = new DataTable();
                SqlCommand cmd = new SqlCommand(@"sp_get_data", sql);
                cmd.CommandType = CommandType.StoredProcedure;
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                da.Fill(dt);
                dataGridView1.DataSource = dt;
                sql.Close();
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

        private void materialButton2_Click(object sender, EventArgs e)
        {
            try
            {
                sql.Open();
                DataTable dt = new DataTable();
                SqlCommand cmd = new SqlCommand(@"sp_get_rate", sql);
                cmd.CommandType = CommandType.StoredProcedure;
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                da.Fill(dt);
                dataGridView1.DataSource = dt;
                sql.Close();
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

        private void materialButton3_Click(object sender, EventArgs e)
        {
            try
            {
                sql.Open();
                DataTable dt = new DataTable();
                SqlCommand cmd = new SqlCommand(@"sp_get_directories", sql);
                cmd.CommandType = CommandType.StoredProcedure;
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                da.Fill(dt);
                dataGridView1.DataSource = dt;
                sql.Close();

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

        private void materialButton4_Click(object sender, EventArgs e)
        {
            try
            {
                if (dataGridView1.Rows.Count > 0)
                {
                    Microsoft.Office.Interop.Excel.Application XcelApp = new Microsoft.Office.Interop.Excel.Application();
                    XcelApp.Application.Workbooks.Add(Type.Missing);
                    for (int i = 1; i < dataGridView1.Columns.Count + 1; i++)
                    {
                        XcelApp.Cells[1, i] = dataGridView1.Columns[i - 1].HeaderText;
                    }
                    for (int i = 0; i < dataGridView1.Rows.Count; i++)
                    {
                        for (int j = 0; j < dataGridView1.Columns.Count; j++)
                        {
                            XcelApp.Cells[i + 2, j + 1] = dataGridView1.Rows[i].Cells[j].Value.ToString();
                        }
                    }
                    XcelApp.Columns.AutoFit();
                    XcelApp.Visible = true;
                }
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
        }

        private void materialButton5_Click_1(object sender, EventArgs e)
        {
            frmMain main = new frmMain();
            main.Show();
            this.Hide();
        }

        private void materialMaskedTextBox1_TextChanged(object sender, EventArgs e)
        {
            try
            {
                //this code is used to search Name on the basis of txttxtSearchItem.text
                ((DataTable)dataGridView1.DataSource).DefaultView.RowFilter = string.Format("Channel like '%{0}%'", textBox1.Text.Trim().Replace("'", "''"));
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }

        

        private void materialButton6_Click_1(object sender, EventArgs e)
        {
           
        }

        private void materialMaskedTextBox1_Click(object sender, EventArgs e)
        {

        }

        private void frmExport_Load(object sender, EventArgs e)
        {

        }

        private void materialMaskedTextBox1_Click_1(object sender, EventArgs e)
        {
           
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void TextBox2_TextChanged_1(object sender, EventArgs e)
        {
            try
            {
                //this code is used to search Name on the basis of txttxtSearchItem.text
                ((DataTable)dataGridView1.DataSource).DefaultView.RowFilter = string.Format("Caption like '%{0}%'", textBox2.Text.Trim().Replace("'", "''"));
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void TextBox3_TextChanged_1(object sender, EventArgs e)
        {
            try
            {
                //this code is used to search Name on the basis of txttxtSearchItem.text
                ((DataTable)dataGridView1.DataSource).DefaultView.RowFilter = string.Format("Brand like '%{0}%'", textBox3.Text.Trim().Replace("'", "''"));
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void TextBox4_TextChanged_1(object sender, EventArgs e)
        {
            try
            {
                //this code is used to search Name on the basis of txttxtSearchItem.text
                ((DataTable)dataGridView1.DataSource).DefaultView.RowFilter = string.Format("SubCategory like '%{0}%'", textBox4.Text.Trim().Replace("'", "''"));
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void TextBox5_TextChanged_1(object sender, EventArgs e)
        {
            try
            {
                //this code is used to search Name on the basis of txttxtSearchItem.text
                ((DataTable)dataGridView1.DataSource).DefaultView.RowFilter = string.Format("Category like '%{0}%'", textBox5.Text.Trim().Replace("'", "''"));
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void TextBox6_TextChanged_1(object sender, EventArgs e)
        {
            try
            {
                //this code is used to search Name on the basis of txttxtSearchItem.text
                ((DataTable)dataGridView1.DataSource).DefaultView.RowFilter = string.Format("CompanyName like '%{0}%'", textBox6.Text.Trim().Replace("'", "''"));
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void TextBox7_TextChanged_1(object sender, EventArgs e)
        {
            try
            {
                //this code is used to search Name on the basis of txttxtSearchItem.text
                ((DataTable)dataGridView1.DataSource).DefaultView.RowFilter = string.Format("AgencyName like '%{0}%'", textBox7.Text.Trim().Replace("'", "''"));
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void TextBox8_TextChanged_1(object sender, EventArgs e)
        {
            try
            {
                //this code is used to search Name on the basis of txttxtSearchItem.text
                ((DataTable)dataGridView1.DataSource).DefaultView.RowFilter = string.Format("AdType like '%{0}%'", textBox8.Text.Trim().Replace("'", "''"));
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void TextBox10_TextChanged_1(object sender, EventArgs e)
        {
            try
            {
                //this code is used to search Name on the basis of txttxtSearchItem.text
                ((DataTable)dataGridView1.DataSource).DefaultView.RowFilter = string.Format("TransmissionHour = '{0}'", textBox10.Text.Trim().Replace("'", "''"));
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void TextBox9_TextChanged_1(object sender, EventArgs e)
        {
            try
            {
                //this code is used to search Name on the basis of txttxtSearchItem.text
                ((DataTable)dataGridView1.DataSource).DefaultView.RowFilter = string.Format("ProgramName like '%{0}%'", textBox9.Text.Trim().Replace("'", "''"));
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void TextBox1_TextChanged_1(object sender, EventArgs e)
        {
       
            try
            {
                //this code is used to search Name on the basis of txttxtSearchItem.text
                ((DataTable)dataGridView1.DataSource).DefaultView.RowFilter = string.Format("Channel like '%{0}%'", textBox1.Text.Trim().Replace("'", "''"));
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            try
            {
                //this code is used to search Name on the basis of txttxtSearchItem.text
                ((DataTable)dataGridView1.DataSource).DefaultView.RowFilter = string.Format("Channel like '%{0}%'", textBox1.Text.Trim().Replace("'", "''"));
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            try
            {
                //this code is used to search Name on the basis of txttxtSearchItem.text
                ((DataTable)dataGridView1.DataSource).DefaultView.RowFilter = string.Format("Caption like '%{0}%'", textBox2.Text.Trim().Replace("'", "''"));
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            try
            {
                //this code is used to search Name on the basis of txttxtSearchItem.text
                ((DataTable)dataGridView1.DataSource).DefaultView.RowFilter = string.Format("Brand like '%{0}%'", textBox3.Text.Trim().Replace("'", "''"));
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {
            try
            {
                //this code is used to search Name on the basis of txttxtSearchItem.text
                ((DataTable)dataGridView1.DataSource).DefaultView.RowFilter = string.Format("SubCategory like '%{0}%'", textBox4.Text.Trim().Replace("'", "''"));
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void textBox5_TextChanged(object sender, EventArgs e)
        {
            try
            {
                //this code is used to search Name on the basis of txttxtSearchItem.text
                ((DataTable)dataGridView1.DataSource).DefaultView.RowFilter = string.Format("Category like '%{0}%'", textBox5.Text.Trim().Replace("'", "''"));
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void textBox6_TextChanged(object sender, EventArgs e)
        {
            try
            {
                //this code is used to search Name on the basis of txttxtSearchItem.text
                ((DataTable)dataGridView1.DataSource).DefaultView.RowFilter = string.Format("CompanyName like '%{0}%'", textBox6.Text.Trim().Replace("'", "''"));
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void textBox7_TextChanged(object sender, EventArgs e)
        {
            try
            {
                //this code is used to search Name on the basis of txttxtSearchItem.text
                ((DataTable)dataGridView1.DataSource).DefaultView.RowFilter = string.Format("AgencyName like '%{0}%'", textBox7.Text.Trim().Replace("'", "''"));
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void textBox8_TextChanged(object sender, EventArgs e)
        {
            try
            {
                //this code is used to search Name on the basis of txttxtSearchItem.text
                ((DataTable)dataGridView1.DataSource).DefaultView.RowFilter = string.Format("AdType like '%{0}%'", textBox8.Text.Trim().Replace("'", "''"));
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }

        private void textBox9_TextChanged(object sender, EventArgs e)
        {
            try
            {
                //this code is used to search Name on the basis of txttxtSearchItem.text
                ((DataTable)dataGridView1.DataSource).DefaultView.RowFilter = string.Format("TransmissionHour = '{0}'", textBox9.Text.Trim().Replace("'", "''"));
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void textBox10_TextChanged(object sender, EventArgs e)
        {
            try
            {
                //this code is used to search Name on the basis of txttxtSearchItem.text
                ((DataTable)dataGridView1.DataSource).DefaultView.RowFilter = string.Format("ProgramName like '%{0}%'", textBox10.Text.Trim().Replace("'", "''"));
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void materialButton6_Click(object sender, EventArgs e)
        {
            try
            {
                using (DataTable dt = new DataTable("Data"))
                {
                    using (SqlCommand cmd = new SqlCommand(@"sp_filter_by_date", sql))
                    {
                        sql.Open();
                        //DataTable dt = new DataTable();
                        //   SqlCommand cmd = new SqlCommand(@"sp_filter_by_date", sql);
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.AddWithValue("@p_fromdate", dateTimePicker1.Value);
                        cmd.Parameters.AddWithValue("@p_todate", dateTimePicker2.Value);
                        SqlDataAdapter da = new SqlDataAdapter(cmd);
                        da.Fill(dt);
                        dataGridView1.DataSource = dt;
                        sql.Close();
                    }
                }

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

        private void materialButton5_Click(object sender, EventArgs e)
        {
            frmMain main = new frmMain();
            main.Show();
            this.Hide();
        }
    }
}
