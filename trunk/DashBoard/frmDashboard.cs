using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace DashBoard
{
    public partial class frmDashboard : Form
    {
        public frmDashboard()
        {
            InitializeComponent();
        }

        private void createGadgetToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Control.uiGadget gadget = new Control.uiGadget();
            ucDashboard.AddPanel(gadget);
        }
    }
}
