using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Janus.Windows.GridEX;
namespace QueryDesigner
{
    public partial class frmFieldChooser : Form
    {
        public frmFieldChooser()
        {
            InitializeComponent();
        }
        public void Show(GridEX grid, Form owner)
        {
            this.gridEXFieldChooserControl1.GridEX = grid;
            this.gridEXFieldChooserControl1.VisualStyleManager = grid.VisualStyleManager;
            Point location = grid.Location;
            location = grid.PointToScreen(location);
            location.X = owner.Bounds.Right + 4;
            this.Location = location;
            Rectangle screenBounds = Screen.GetBounds(grid);
            if (this.Bounds.Right > screenBounds.Right)
            {
                this.Left = screenBounds.Right - this.Width;
            }
            this.Show(owner);
            
        }
    }
}