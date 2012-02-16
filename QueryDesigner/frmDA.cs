using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;
using Janus.Windows.GridEX;

namespace QueryDesigner
{
    public partial class frmDA : Form
    {
        BUS.LIST_DAControl ctr = new BUS.LIST_DAControl();
        string _sErr = "";
        string _processStatus = "";
        DataTable _dt;
        public frmDA()
        {
            InitializeComponent();
        }

        private void btnNew_Click(object sender, EventArgs e)
        {
            RefreshForm("");
            EnabledForm(true);
            _processStatus = "C";
        }

        private void btnView_Click(object sender, EventArgs e)
        {
            frmDAView frm = new frmDAView();
            if (frm.ShowDialog() == DialogResult.OK)
            {
                SetDataToForm(ctr.Get(frm.Code, ref _sErr));
                EnabledForm(false);
            }
        }

        private void btnEdit_Click(object sender, EventArgs e)
        {
            if (ctr.IsExist(txtCode.Text))
            {
                EnabledForm(true);
                txtCode.Enabled = false;
                _processStatus = "A";
            }
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            //string sErr = "";

            DTO.LIST_DAInfo inf = new DTO.LIST_DAInfo();

            if (_processStatus == "C")
            {
                if (!ctr.IsExist(txtCode.Text))
                {
                    ctr.Add(GetDataFromForm(inf), ref _sErr);
                    if (_sErr == "")
                    {
                        BUS.LIST_DAOGControl ctrDAOG = new BUS.LIST_DAOGControl();
                        foreach (DataRow row in _dt.Rows)
                            _sErr = ctrDAOG.InsertUpdate(new DTO.LIST_DAOGInfo(row));
                    }
                }
                else
                    _sErr = txtCode.Text.Trim() + " is exist!";
            }
            else if (_processStatus == "A")
            {
                _sErr = ctr.InsertUpdate(GetDataFromForm(inf));
                BUS.LIST_DAOGControl ctrDAOG = new BUS.LIST_DAOGControl();
                if (ctrDAOG.Deletes(txtCode.Text) == "")
                    foreach (DataRow row in _dt.Rows)
                        _sErr = ctrDAOG.InsertUpdate(new DTO.LIST_DAOGInfo(row));
            }
            if (_sErr == "")
            {
                _processStatus = "V";
                EnabledForm(false);
            }
            else lbErr.Text = _sErr;
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            if (ctr.IsExist(txtCode.Text))
            {
                _sErr = ctr.Delete(txtCode.Text);
                RefreshForm("");
                EnabledForm(false);
                _processStatus = "";
            }
        }

        private void btnCopy_Click(object sender, EventArgs e)
        {
            if (ctr.IsExist(txtCode.Text))
            {
                EnabledForm(true);
                txtCode.Focus();
                _processStatus = "C";
            }
        }


        private void frmPOD_Load(object sender, EventArgs e)
        {
            RefreshForm("");
            EnabledForm(false);
            DataTable dt = new DataTable("LIST_DOAG");
            DataColumn[] columns = new DataColumn[] { 
                new DataColumn("DAG_ID"), 
                new DataColumn("ROLE_ID")                
            };
            dt.Columns.AddRange(columns);
            dgvData.DataSource = dt;
            BUS.POGControl pogctr = new BUS.POGControl();
            DataTable dtx = pogctr.GetAll(ref _sErr);
            dgvData.RootTable.Columns["Code"].EditValueList = new Janus.Windows.GridEX.GridEXValueListItemCollection();
            foreach (DataRow row in dtx.Rows)
                dgvData.RootTable.Columns["Code"].EditValueList.Add(row["ROLE_ID"], row["ROLE_ID"].ToString());

        }
        private void lbErr_Click(object sender, EventArgs e)
        {
            if (lbErr.Text != "" && lbErr.Text != "...")
                MessageBox.Show(lbErr.Text);
        }
        private void SetDataToForm(DTO.LIST_DAInfo obj)
        {
            txtCode.Text = obj.DAG_ID;
            txtIE.Text = obj.EI;
            txtDescription.Text = obj.NAME;
            BUS.LIST_DAOGControl ctr = new BUS.LIST_DAOGControl();
            _dt = ctr.GetAll(obj.DAG_ID, ref _sErr);
            dgvData.DataSource = _dt;

        }
        private DTO.LIST_DAInfo GetDataFromForm(DTO.LIST_DAInfo obj)
        {
            obj.DAG_ID = txtCode.Text;
            obj.NAME = txtDescription.Text;
            obj.EI = txtIE.Text;
            _dt = dgvData.DataSource as DataTable;
            foreach (DataRow row in _dt.Rows)
            {
                row["DAG_ID"] = txtCode.Text;
            }
            return obj;
        }
        private void EnabledForm(bool value)
        {
            pContain.Enabled = value;
            txtCode.Enabled = true;
        }
        private void RefreshForm(string value)
        {
            foreach (Control x in pContain.Controls)
            {
                if (x is TextBox)
                    x.Text = value;
            }
            DataTable dt = dgvData.DataSource as DataTable;
            if (dt != null)
                dt.Rows.Clear();
        }

        private void btnGroup_Click(object sender, EventArgs e)
        {
            frmPOGView frm = new frmPOGView();
            if (frm.ShowDialog() == DialogResult.OK)
            {
                txtCode.Text = frm.Code;
            }
        }

        private void dgvData_LoadingRow(object sender, Janus.Windows.GridEX.RowLoadEventArgs e)
        {
            if (e.Row.RowType == RowType.Record)
            {
                //e.Row.BeginEdit();
                e.Row.Cells["Description"].Value = CalculateDetailValue(e.Row);
                //e.Row.Cells["DAG_ID"].Value = txtCode.Text;
                //e.Row.EndEdit();
            }
        }

        private object CalculateDetailValue(GridEXRow gridEXRow)
        {
            BUS.POGControl ctr = new BUS.POGControl();
            DTO.POGInfo inf = ctr.Get(gridEXRow.Cells["Code"].Value.ToString(), ref _sErr);
            return inf.ROLE_NAME;
        }
    }
}
