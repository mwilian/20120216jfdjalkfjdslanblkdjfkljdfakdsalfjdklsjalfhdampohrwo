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
    public partial class Form_DTBView : Form
    {
        public String themname = "";
        private String sErr = "";
        public String Code_DTB = "";
        public String Description_DTB = "";

        private bool flag;
        string THEME = "Breeze";
        public Form_DTBView()
        {
            InitializeComponent();
            //ThemeResolutionService.ApplyThemeToControlTree(this, THEME);
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            if (themname != "")
            {
                //ThemeResolutionService.ApplyThemeToControlTree(this, themname);
            }
            flag = true;
            //dgvDTBView.DataSource = DBInfoList.GetDBInfoList();
            LoadDataGrid(dgvDTBView, DBInfoList.GetDBInfoList());
            //dgvFilter.RetrieveStructure();
            flag = false;
            //dgvFilter.CurrentRow = null;
        }

        private void LoadDataGrid(Janus.Windows.GridEX.GridEX dgv, DBInfoList dBInfoList)
        {
            dgv.DataSource = dBInfoList;
            //dgv.AutoSizeColumns();
            string path = Form_QD.__documentDirectory + "\\Layout\\" + dgv.SettingsKey + ".gxl";
            if (File.Exists(path))
            {
                FileStream fs = new FileStream(path, FileMode.Open);
                try { dgv.LoadLayoutFile(fs); }
                catch { fs.Close(); File.Delete(path); }
                fs.Close();
            }
        }

        private void btSave_Click(object sender, EventArgs e)
        {
            if (dgvDTBView.CurrentRow != null && dgvDTBView.CurrentRow.RowIndex >= 0)
            {
                Code_DTB = dgvDTBView.CurrentRow.Cells["DB"].Value.ToString();
                Description_DTB = dgvDTBView.CurrentRow.Cells["Description"].Value.ToString();
                DialogResult = DialogResult.OK;
                Close();
            }
        }

        private void btnReresh_Click(object sender, EventArgs e)
        {            //dgvFilter.MasterGridViewTemplate.AutoGenerateColumns = false;

            dgvDTBView.DataSource = DBInfoList.GetDBInfoList();
            // //dgvDTBView.AutoSizeColumns();
            //dgvFilter.RetrieveStructure();
            SaveLayout(dgvDTBView);
            flag = false;
            //dgvFilter.CurrentRow = null;
        }

        private void btCancel_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
            Close();
        }


        private void dgvFilter_MouseDoubleClick(object sender, MouseEventArgs e)
        {

        }

        private void dgvFilter_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                btSave_Click(null, null);
            }
        }

        private void dgvDTBView_GroupsChanged(object sender, Janus.Windows.GridEX.GroupsChangedEventArgs e)
        {
            SaveLayout(dgvDTBView);
        }

        private void SaveLayout(Janus.Windows.GridEX.GridEX dgv)
        {
            string path = Form_QD.__documentDirectory + "\\Layout\\" + dgv.SettingsKey + ".gxl";
            using (FileStream fs = new FileStream(path, FileMode.OpenOrCreate, FileAccess.Write))
            {
                dgv.SaveLayoutFile(fs);
            }
        }

        private void dgvDTBView_FilterApplied(object sender, EventArgs e)
        {
            SaveLayout(dgvDTBView);
        }

        private void dgvDTBView_ColumnMoved(object sender, Janus.Windows.GridEX.ColumnActionEventArgs e)
        {
            SaveLayout(dgvDTBView);
        }

        private void dgvDTBView_RowDoubleClick(object sender, Janus.Windows.GridEX.RowActionEventArgs e)
        {
            btSave_Click(null, null);
        }
    }
}
