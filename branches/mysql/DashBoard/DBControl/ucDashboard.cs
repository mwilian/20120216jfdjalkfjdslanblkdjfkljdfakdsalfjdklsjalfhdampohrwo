using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Text;
using System.Windows.Forms;

namespace DashBoard.Control
{
    public partial class ucDashboard : UserControl
    {
        public ucDashboard()
        {
            InitializeComponent();
        }
        public void AddPanel(Janus.Windows.UI.Dock.UIPanel x )
        {
            uiPanelManager1.Panels.Add(x);
        }
    }
}
