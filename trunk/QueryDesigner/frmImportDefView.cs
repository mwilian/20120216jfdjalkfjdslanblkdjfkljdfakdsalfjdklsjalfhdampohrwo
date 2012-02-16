using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;

using System.Text;
using System.Windows.Forms;
using System.IO;

namespace QueryDesigner
{
    public partial class frmImportDefView : Form
    {
        string _code = "";
        string _connect = "";

        public string Connect
        {
            get { return _connect; }
            set { _connect = value; }
        }

        public string ReturnCode
        {
            get { return _code; }
            set { _code = value == null ? "" : value.Trim(); }
        }
        public frmImportDefView()
        {
            InitializeComponent();
        }

        private void dgvLookup_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {

            if (e.RowIndex >= 0 && e.ColumnIndex >= 0)
                ReturnCode = dgvQDADDView.GetRow(e.RowIndex).Cells["Code"].Value.ToString();
            DialogResult = DialogResult.OK;
            Close();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
            Close();
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.OK;
            if (dgvQDADDView.CurrentRow != null && dgvQDADDView.CurrentRow.RowIndex >= 0)
            {
                ReturnCode = dgvQDADDView.CurrentRow.Cells["Code"].Value.ToString();
            }
            Close();
        }

        private void frmLookup_Load(object sender, EventArgs e)
        {
            string sErr = "";
            BUS.IMPORT_SCHEMAControl ctr = new BUS.IMPORT_SCHEMAControl();
            DataTable dt = ctr.GetAll(_connect, ref sErr);
            dt.Columns["SCHEMA_ID"].ColumnName = "Code";
            //dgvQDADDView.AutoGenerateColumns = false;
            LoadDataGrid(dgvQDADDView, dt);
        }

        private void dgvQDADDView_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                btnOK_Click(null, null);
        }

        private void dgvQDADDView_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            
        }
        private void LoadDataGrid(Janus.Windows.GridEX.GridEX dgv, DataTable dt)
        {
            dgv.DataSource = dt;
           //dgv.AutoSizeColumns();
            string path = Form_QD.__documentDirectory + "\\Layout\\" + dgv.SettingsKey + ".gxl";
            if (File.Exists(path))
            {
                FileStream fs = new FileStream(path, FileMode.Open);
                try { dgv.LoadLayoutFile(fs); }
                catch { fs.Close(); File.Delete(path); }
                fs.Close();
            }
            dgv.Focus();
        }
        private void SaveLayout(Janus.Windows.GridEX.GridEX dgv)
        {
            string path = Form_QD.__documentDirectory + "\\Layout\\" + dgv.SettingsKey + ".gxl";
            try
            {
                FileStream fs = new FileStream(path, FileMode.OpenOrCreate, FileAccess.Write);
                dgv.SaveLayoutFile(fs);
                fs.Close();
            }
            catch (Exception ex)
            {
            }
        }

        private void dgvQDADDView_GroupsChanged(object sender, Janus.Windows.GridEX.GroupsChangedEventArgs e)
        {
            SaveLayout(dgvQDADDView);
        }

        private void dgvQDADDView_FilterApplied(object sender, EventArgs e)
        {
            SaveLayout(dgvQDADDView);
        }

        private void btnRefresh_Click(object sender, EventArgs e)
        {
            string sErr = "";
            BUS.IMPORT_SCHEMAControl ctr = new BUS.IMPORT_SCHEMAControl();
            DataTable dt = ctr.GetAll(_connect, ref sErr);
            dt.Columns["SCHEMA_ID"].ColumnName = "Code";
            //dgvQDADDView.AutoGenerateColumns = false;
            dgvQDADDView.DataSource = dt;
           //dgvQDADDView.AutoSizeColumns();
            SaveLayout(dgvQDADDView);
        }

        private void dgvQDADDView_SizingColumn(object sender, Janus.Windows.GridEX.SizingColumnEventArgs e)
        {
            SaveLayout(dgvQDADDView);
        }

        private void dgvQDADDView_RowDoubleClick(object sender, Janus.Windows.GridEX.RowActionEventArgs e)
        {
            btnOK_Click(null, null);
        }

        
    }
}
