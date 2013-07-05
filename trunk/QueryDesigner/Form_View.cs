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
    public partial class Form_View : Form
    {
        public String themname = "";
        public String database = "";
        private String sErr = "";
        public LIST_QDInfo qdinfo = new LIST_QDInfo();

        private bool flag;
        string THEME = "Breeze";
        string _user = "";

        public string User
        {
            get { return _user; }
            set { _user = value; }
        }
        //public Form_View(string user)
        //{
        //    InitializeComponent();
        //    DialogResult = DialogResult.Cancel;
        //    _user = user;
        //    //ThemeResolutionService.ApplyThemeToControlTree(this, THEME);
        //}
        public Form_View(string aDatabase, string user)
        {
            database = aDatabase;
            _user = user;
            InitializeComponent();
            DialogResult = DialogResult.Cancel;
        }
        protected void LoadLayout()
        {
            try
            {
                string path = Form_QD.__documentDirectory + "\\Layout\\QDViewLayout.xml";
                if (File.Exists(path))
                {
                    using (Stream sr = new FileStream(path, FileMode.OpenOrCreate))
                    {
                        try
                        {
                            dgvQDView.LoadLayoutFile(sr);
                        }
                        catch { }
                    }

                }
                else if (File.Exists(Form_QD.__documentDirectory + "\\Layout\\QDViewLayoutDefault.xml"))
                {
                    using (Stream sr = new FileStream(Form_QD.__documentDirectory + "\\Layout\\QDViewLayoutDefault.xml", FileMode.OpenOrCreate))
                    {
                        try
                        {
                            dgvQDView.LoadLayoutFile(sr);
                        }
                        catch { }

                    }
                }
                else
                {
                    InitLayout();
                }
                UpdateLayout();
            }
            catch { ResetLayout(); }
        }
        protected virtual void UpdateLayout()
        { }
        protected void SaveLayout()
        {
            string path = Form_QD.__documentDirectory + "\\Layout\\QDViewLayout.xml";
            if (File.Exists(path))
                File.Delete(path);
            using (Stream sr = new FileStream(path, FileMode.OpenOrCreate))
            {
                try
                {
                    dgvQDView.SaveLayoutFile(sr);
                }
                catch { }
            }
        }
        protected void ResetLayout()
        {
            string path = Form_QD.__documentDirectory + "\\Layout\\QDViewLayoutDefault.xml";
            using (Stream sr = new FileStream(path, FileMode.OpenOrCreate))
            {

                if (File.Exists(path))
                {
                    try
                    {
                        dgvQDView.LoadLayoutFile(sr);
                    }
                    catch { }

                }
            }
            SaveLayout();
        }
        protected void InitLayout()
        {
            dgvQDView.RetrieveStructure();
            string path = Form_QD.__documentDirectory + "\\Layout\\QDViewLayoutDefault.xml";
            using (Stream sr = new FileStream(path, FileMode.OpenOrCreate))
            {
                try
                {
                    dgvQDView.SaveLayoutFile(sr);
                }
                catch { }
            }
        }
        private void dgvFilter_SelectionChanged(object sender, EventArgs e)
        {
            //RadGridView temp = (RadGridView)sender;
            //if (temp is RadGridView && flag == false && temp.SelectedRows.Count > 0)
            //{

            //    String qd_id = temp.SelectedRows[0].Cells["QD_ID"]._Value.ToString();
            //    String dtb = temp.SelectedRows[0].Cells["DTB"]._Value.ToString();
            //    LIST_QDControl ctr = new LIST_QDControl();

            //    qdinfo = ctr.Get_LIST_QD(dtb, qd_id, ref sErr);
            //    DialogResult = DialogResult.OK;
            //    Close();
            //}
        }

        private void Form_View_Load(object sender, EventArgs e)
        {
            if (themname != "")
            {
                //ThemeResolutionService.ApplyThemeToControlTree(this, themname);
            }
            flag = true;

            LIST_QDControl pdControl = new LIST_QDControl();
            DataTable dt = pdControl.GetAll_LIST_QD_USER(database, _user, ref sErr);
            //dgvFilter.MasterGridViewTemplate.AutoGenerateColumns = false;            
            LoadDataGrid(dt);
            LoadLayout();
            flag = false;
            //DialogResult = DialogResult.Cancel;
            //dgvFilter.CurrentRow = null;
        }

        private void btCancel_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
            Close();
        }

        private void btSave_Click(object sender, EventArgs e)
        {
            if (dgvQDView.CurrentRow != null && dgvQDView.CurrentRow.RowIndex >= 0)
            {
                String qd_id = dgvQDView.CurrentRow.Cells["QD_ID"].Value.ToString();
                String dtb = dgvQDView.CurrentRow.Cells["DTB"].Value.ToString();
                LIST_QDControl ctr = new LIST_QDControl();

                qdinfo = ctr.Get_LIST_QD(dtb, qd_id, ref sErr);
                DialogResult = DialogResult.OK;
                Close();
            }
        }

        private void btnReresh_Click(object sender, EventArgs e)
        {
            LIST_QDControl pdControl = new LIST_QDControl();
            DataTable dt = pdControl.GetAll_LIST_QD_USER(database, _user, ref sErr);
            //dgvFilter.MasterGridViewTemplate.AutoGenerateColumns = false;
            dgvQDView.DataSource = dt;
            //dgvQDView.AutoSizeColumns();
            ResetLayout();
            flag = false;
            //dgvFilter.CurrentRow = null;
        }

        private void LoadDataGrid(DataTable dt)
        {
            dgvQDView.DataSource = dt;
            //dgvQDView.AutoSizeColumns();

        }




        private void dgvQDView_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                btSave_Click(null, null);
        }

        private void dgvQDView_MouseDoubleClick(object sender, MouseEventArgs e)
        {

        }





        private void Form_View_FormClosed(object sender, FormClosedEventArgs e)
        {
            SaveLayout();
        }

        private void dgvQDView_RowDoubleClick(object sender, Janus.Windows.GridEX.RowActionEventArgs e)
        {
            btSave_Click(null, null);
        }


    }
}
