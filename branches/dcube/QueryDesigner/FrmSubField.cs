using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;

using System.Text;
using System.Windows.Forms;


namespace dCube
{
    public partial class FrmSubField : Form
    {
        public string Description
        {
            get { return txtDesc.Text; }
            set { txtDesc.Text = value; }
        }
        public int Index
        {
            get { return Convert.ToInt32(nudFrom.Value); }
            set { nudFrom.Value = value; }
        }
        public int Length
        {
            get { return Convert.ToInt32(nudLength.Value); }
            set { nudLength.Value = value; }
        }
        string THEME = "Breeze";
        public FrmSubField()
        {
            InitializeComponent();
            //ThemeResolutionService.ApplyThemeToControlTree(this, THEME);
        }

        private void btSave_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.OK;
            Close();
        }

        private void btCancel_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }



    }
}
