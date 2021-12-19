using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using MaterialSkin.Controls;
namespace Excel
{
    public partial class frmMain : MaterialForm

    {
        public frmMain()
        {
            InitializeComponent();
            MaterialSkin.MaterialSkinManager manager = MaterialSkin.MaterialSkinManager.Instance;
            manager.AddFormToManage(this);
            manager.Theme = MaterialSkin.MaterialSkinManager.Themes.LIGHT;
            manager.ColorScheme = new MaterialSkin.ColorScheme(MaterialSkin.Primary.Blue300,
                MaterialSkin.Primary.Blue500, MaterialSkin.Primary.Blue500, MaterialSkin.Accent.LightBlue400,
                MaterialSkin.TextShade.BLACK);
        }

        private void frmMain_Load(object sender, EventArgs e)
        {

        }

        private void materialButton3_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void materialButton1_Click(object sender, EventArgs e)
        {
            frmImport import = new frmImport();
            import.Show();
            this.Hide();
        }

        private void materialButton2_Click(object sender, EventArgs e)
        {
            frmExport export = new frmExport();
            export.Show();
            this.Hide();

        }
    }
}
