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
    public partial class frmLogin : MaterialForm
    {
        SqlConnection con = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlcommand"].ConnectionString);
        public frmLogin()
        {
            InitializeComponent();
            MaterialSkin.MaterialSkinManager manager = MaterialSkin.MaterialSkinManager.Instance;
            manager.AddFormToManage(this);
            manager.Theme = MaterialSkin.MaterialSkinManager.Themes.LIGHT;
            manager.ColorScheme = new MaterialSkin.ColorScheme(MaterialSkin.Primary.Blue300,
                MaterialSkin.Primary.Blue500, MaterialSkin.Primary.Blue500, MaterialSkin.Accent.LightBlue400,
                MaterialSkin.TextShade.WHITE);
        }

        private void frmLogin_Load(object sender, EventArgs e)
        {

        }

        private void materialLabel1_Click(object sender, EventArgs e)
        {

        }
        public static int loggedinuser;
        private void materialButton1_Click(object sender, EventArgs e)
        {
            DataTable dt = new DataTable();
            SqlCommand cmd = new SqlCommand("sp_login", con);
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.AddWithValue("@p_username", materialTextBox1.Text.Trim());
            cmd.Parameters.AddWithValue("@p_password", materialTextBox2.Text.Trim());
            SqlDataAdapter da = new SqlDataAdapter(cmd);
            da.Fill(dt);

            if (dt.Rows.Count > 0)
            {
                if (dt.Rows[0]["Username"].ToString() == materialTextBox1.Text.Trim() && dt.Rows[0]["Password"].ToString() == materialTextBox2.Text.Trim())
                {
                    loggedinuser = Convert.ToInt32(dt.Rows[0]["UID"].ToString());
                    frmMain main = new frmMain();
                    main.Show();
                    this.Hide();

                }

            }
            else
            {
                MessageBox.Show("Invalid User");
                this.Show();   
            }

        }
    }
}
