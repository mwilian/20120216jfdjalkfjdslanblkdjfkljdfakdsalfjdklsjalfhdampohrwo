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
    public partial class Form_TableView : Form
    {
        public String themname = "";
        private String sErr = "";
        public String Code_DTB = "";
        public String Description_DTB = "";
        string _user = "";

        private bool flag;
        string THEME = "Breeze";
        public Form_TableView(string db, string user)
        {
            InitializeComponent();
            Code_DTB = db; _user = user;
            //ThemeResolutionService.ApplyThemeToControlTree(this, THEME);
        }

        private void dgvFilter_SelectionChanged(object sender, EventArgs e)
        {
            //RadGridView temp = (RadGridView)sender;
            //if (temp is RadGridView && flag == false && temp.SelectedRows.Count > 0)
            //{

            //    Code_DTB = temp.SelectedRows[0].Cells["Code"].Value.ToString();
            //    Description_DTB = temp.SelectedRows[0].Cells["Description"].Value.ToString();
            //    Close();
            //}

        }

        private void Form_TableView_Load(object sender, EventArgs e)
        {
            if (themname != "")
            {
                //ThemeResolutionService.ApplyThemeToControlTree(this, themname);
            }
            flag = true;
            //dgvTableView.MasterGridViewTemplate.AutoGenerateColumns = false;
            LoadDataGrid(dgvTableView, SchemaDefinition.GetTableList(Code_DTB, _user));
            flag = false;
            //dgvFilter.CurrentRow = null;


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
        private void LoadDataGrid(Janus.Windows.GridEX.GridEX dgv, BindingList<TableItem> dBInfoList)
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
        private void btCancel_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
            Close();
        }

        private void btnReresh_Click(object sender, EventArgs e)
        {
            //dgvTableView.MasterGridViewTemplate.AutoGenerateColumns = false;
            dgvTableView.DataSource = SchemaDefinition.GetTableList(Code_DTB, _user);
            //dgvTableView.AutoSizeColumns();
            SaveLayout(dgvTableView);
            flag = false;
        }

        private void btSave_Click(object sender, EventArgs e)
        {
            if (dgvTableView.CurrentRow != null && dgvTableView.CurrentRow.RowIndex >= 0)
            {
                Code_DTB = dgvTableView.CurrentRow.Cells["Code"].Value.ToString();
                Description_DTB = dgvTableView.CurrentRow.Cells["Description"].Value.ToString();
                DialogResult = DialogResult.OK;
                Close();
            }
        }

        private void dgvTableView_MouseDoubleClick(object sender, MouseEventArgs e)
        {

        }

        private void dgvTableView_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                btSave_Click(null, null);
        }

        private void dgvTableView_GroupsChanged(object sender, Janus.Windows.GridEX.GroupsChangedEventArgs e)
        {
            SaveLayout(dgvTableView);
        }

        private void dgvTableView_FilterApplied(object sender, EventArgs e)
        {
            SaveLayout(dgvTableView);
        }

        private void dgvTableView_RowDoubleClick(object sender, Janus.Windows.GridEX.RowActionEventArgs e)
        {
            btSave_Click(null, null);
        }
    }
}
