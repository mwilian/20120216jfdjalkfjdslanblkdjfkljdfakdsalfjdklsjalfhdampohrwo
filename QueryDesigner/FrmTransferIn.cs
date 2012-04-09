using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;

using System.Text;
using System.Windows.Forms;


namespace QueryDesigner
{
    public partial class FrmTransferIn : Form
    {
        string THEME = "Breeze";
        string _type = "";
        public FrmTransferIn(string type)
        {
            InitializeComponent();
            ////ThemeResolutionService.ApplyThemeToControlTree(this, THEME);
            _type = type;
        }
        DataTable dt = new DataTable();
        private void btnBrowse_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "Xml file (*.xml)|*.xml";
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                txtFileName.Text = ofd.FileName;
            }
        }

        private void btnLoad_Click(object sender, EventArgs e)
        {
            if (_type == "QD")
            {
                BUS.LIST_QDControl control = new BUS.LIST_QDControl();
                dt = control.ToTransferInStruct();
                dt.TableName = "Table";
                foreach (DataColumn col in dt.Columns)
                {
                    col.DataType = typeof(string);
                }
            }
            else if (_type == "QDADD")
            {
                BUS.LIST_QD_SCHEMAControl control = new BUS.LIST_QD_SCHEMAControl();
                dt = control.ToTransferInStruct();
                dt.TableName = "Table";
                foreach (DataColumn col in dt.Columns)
                {
                    col.DataType = typeof(string);
                }

            }
            else if (_type == "TASK")
            {
                BUS.LIST_TASKControl control = new BUS.LIST_TASKControl();
                dt = control.ToTransferInStruct();
                dt.TableName = "Table";
                foreach (DataColumn col in dt.Columns)
                {
                    col.DataType = typeof(string);
                }

            }
            else if (_type == "POD")
            {
                BUS.PODControl control = new BUS.PODControl();
                dt = control.ToTransferInStruct();
                dt.TableName = "Table";
                foreach (DataColumn col in dt.Columns)
                {
                    col.DataType = typeof(string);
                }

            }
            BUS.CommonControl commonCtr = new BUS.CommonControl();

            try
            {
                dt.ReadXml(txtFileName.Text);
                dt = commonCtr.ValidatedDataTransferIn(dt, _type);
                radGridView1.DataSource = dt;
                radGridView1.RetrieveStructure();
                radGridView1.RootTable.Columns["tmp_Validated"].Visible = false;
            }
            catch { }
        }

        private void btnImport_Click(object sender, EventArgs e)
        {
            string sErr = "";
            BUS.LIST_QDControl control = new BUS.LIST_QDControl();
            BUS.LIST_QD_SCHEMAControl controlADD = new BUS.LIST_QD_SCHEMAControl();
            BUS.LIST_TASKControl controlTASK = new BUS.LIST_TASKControl();
            BUS.PODControl controlPOD = new BUS.PODControl();
            try
            {
                foreach (DataRow row in dt.Rows)
                {
                    if ((bool)row["tmp_Validated"] == true)
                    {
                        if (_type == "QD")
                            control.TransferIn(row, ref sErr);
                        else if (_type == "QDADD")
                            sErr = controlADD.TransferIn(row);
                        else if (_type == "TASK")
                            sErr = controlTASK.TransferIn(row);
                        else if (_type=="POD")
                            sErr = controlPOD.TransferIn(row);
                    }
                }
                Close();
            }
            catch { }
        }

       

        private void radGridView1_LoadingRow(object sender, Janus.Windows.GridEX.RowLoadEventArgs e)
        {
            //e.Row
            if (e.Row != null && e.Row.DataRow != null)
            {
                DataRow row = ((DataRowView)e.Row.DataRow).Row;
                if ((bool)row["tmp_Validated"] == false)
                {
                    e.Row.RowStyle = new Janus.Windows.GridEX.GridEXFormatStyle();
                    e.Row.RowStyle.ForeColor = Color.Red;
                }
            }
        }

        //private void radGridView1_CreateRow(object sender, Telerik.WinControls.UI.GridViewCreateRowEventArgs e)
        //{
        //    if (e.RowInfo != null && e.RowInfo.DataBoundItem != null)
        //    {
        //        DataRow row = ((DataRowView)e.RowInfo.DataBoundItem).Row;
        //        if ((bool)row["tmp_Validated"] == false)
        //            e.RowElement.ForeColor = Color.Red;
        //    }
        //}
    }
}
