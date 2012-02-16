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

namespace QueryDesigner
{
    public partial class frmPOPView : Form
    {
        public String themname = "";
        public String database = "";
        private String sErr = "";
        private String _Code = "";
        private String _Description = "";
        private string _lookup = "";

        public string Lookup
        {
            get { return _lookup; }
            set { _lookup = value; }
        }
        public String Description
        {
            get { return _Description; }
            set { _Description = value; }
        }
        public String Code
        {
            get { return _Code; }
            set { _Code = value; }
        }

        private bool flag;
        public frmPOPView()
        {
            InitializeComponent();
            DialogResult = DialogResult.Cancel;
        }

        private void Form_Load(object sender, EventArgs e)
        {
            POPControl pdControl = new POPControl();
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
            if (dgvPOPView.CurrentRow != null && dgvPOPView.CurrentRow.RowIndex >= 0)
            {
                _Code = dgvPOPView.CurrentRow.Cells["Code"].Value.ToString();
                _Description = dgvPOPView.CurrentRow.Cells["Description"].Value.ToString();
                _lookup = dgvPOPView.CurrentRow.Cells["LookUp"].Value.ToString();
                DialogResult = DialogResult.OK;
                Close();
            }
        }

        private void btnReresh_Click(object sender, EventArgs e)
        {
            POPControl pdControl = new POPControl();
            DataTable dt = pdControl.GetAll(ref sErr);
            //dgvFilter.MasterGridViewTemplate.AutoGenerateColumns = false;
            dgvPOPView.DataSource =  dt;
           //dgvPOPView.AutoSizeColumns();
            SaveLayout();
            flag = false;
            //dgvFilter.CurrentRow = null;
        }

        private void LoadDataGrid(DataTable dt)
        {
            dgvPOPView.DataSource = dt;
           //dgvPOPView.AutoSizeColumns();
            string path = Form_QD.__documentDirectory + "\\Layout\\" + dgvPOPView.SettingsKey + ".gxl";
            if (File.Exists(path))
            {
                FileStream fs = new FileStream(path, FileMode.Open);
                try { dgvPOPView.LoadLayoutFile(fs); }
                catch { fs.Close(); File.Delete(path); }
                fs.Close();
            }
        }


        private void SaveLayout()
        {
            string path = Form_QD.__documentDirectory + "\\Layout\\" + dgvPOPView.SettingsKey + ".gxl";
            try
            {
                FileStream fs = new FileStream(path, FileMode.OpenOrCreate, FileAccess.Write);
                dgvPOPView.SaveLayoutFile(fs);
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
            btSave_Click(null, null);
        }

        private void dgvQDView_GroupsChanged(object sender, Janus.Windows.GridEX.GroupsChangedEventArgs e)
        {
            SaveLayout();
        }

        private void dgvQDView_FilterApplied(object sender, EventArgs e)
        {
            SaveLayout();
        }


    }
}
