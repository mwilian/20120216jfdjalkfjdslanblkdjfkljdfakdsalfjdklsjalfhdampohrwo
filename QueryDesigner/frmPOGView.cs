using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;

using System.Text;
using System.Windows.Forms;


using BUS;
using DTO;
using QueryBuilder;
using System.IO;

namespace dCube
{
    public partial class frmPOGView : Form
    {
        public String themname = "";
        public String database = "";
        private String sErr = "";
        private String _Code = "";

        public String Code
        {
            get { return _Code; }
            set { _Code = value; }
        }

        private bool flag;
        public frmPOGView()
        {
            InitializeComponent();
            DialogResult = DialogResult.Cancel;
        }

        private void Form_Load(object sender, EventArgs e)
        {
            POGControl pdControl = new POGControl();
            DataTable dt = pdControl.GetAll(ref sErr);
            //dgvFilter.MasterGridViewTemplate.AutoGenerateColumns = false;            
            LoadDataGrid(dt);
            flag = false;
            //dgvFilter.CurrentRow = null;
        }

        private void btCancel_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
            Close();
        }

        private void btSave_Click(object sender, EventArgs e)
        {
            if (dgvPOGView.CurrentRow != null && dgvPOGView.CurrentRow.RowIndex >= 0)
            {
                _Code = dgvPOGView.CurrentRow.Cells["Code"].Value.ToString();
                DialogResult = DialogResult.OK;
                Close();
            }
        }

        private void btnReresh_Click(object sender, EventArgs e)
        {
            POGControl pdControl = new POGControl();
            DataTable dt = pdControl.GetAll(ref sErr);
            //dgvFilter.MasterGridViewTemplate.AutoGenerateColumns = false;
            dgvPOGView.DataSource =  dt;
           //dgvPOGView.AutoSizeColumns();
            SaveLayout();
            flag = false;
            //dgvFilter.CurrentRow = null;
        }

        private void LoadDataGrid(DataTable dt)
        {
            dgvPOGView.DataSource = dt;
           //dgvPOGView.AutoSizeColumns();
            string path = Form_QD.__documentDirectory + "\\Layout\\" + dgvPOGView.SettingsKey + ".gxl";
            if (File.Exists(path))
            {
                FileStream fs = new FileStream(path, FileMode.Open);
                try { dgvPOGView.LoadLayoutFile(fs); }
                catch { fs.Close(); File.Delete(path); }
                fs.Close();
            }
        }


        private void SaveLayout()
        {
            string path = Form_QD.__documentDirectory + "\\Layout\\" + dgvPOGView.SettingsKey + ".gxl";
            try
            {
                FileStream fs = new FileStream(path, FileMode.OpenOrCreate, FileAccess.Write);
                dgvPOGView.SaveLayoutFile(fs);
                fs.Close();
            }
            catch (Exception ex)
            {
            }
        }

        private void dgvQDView_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                btSave_Click(null, null);
        }

        private void dgvQDView_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            
        }

        private void dgvQDView_GroupsChanged(object sender, Janus.Windows.GridEX.GroupsChangedEventArgs e)
        {
            SaveLayout();
        }

        private void dgvQDView_FilterApplied(object sender, EventArgs e)
        {
            SaveLayout();
        }

        private void dgvPOGView_RowDoubleClick(object sender, Janus.Windows.GridEX.RowActionEventArgs e)
        {
            btSave_Click(null, null);
        }


    }
}
