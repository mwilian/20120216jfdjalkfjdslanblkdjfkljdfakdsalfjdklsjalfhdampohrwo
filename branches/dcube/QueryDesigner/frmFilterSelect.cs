using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;



namespace dCube
{
    public partial class frmFilterSelect : Form
    {
        QueryBuilder.SQLBuilder _sqlBuidler;
        int _indexFilter = 0;
        string _connectString = "";

        public string ConnectString
        {
            get { return _connectString; }
            set { _connectString = value; }
        }
        public int IndexFilter
        {
            get { return _indexFilter; }
            set { _indexFilter = value; }
        }


        public QueryBuilder.SQLBuilder SqlBuidler
        {
            get { return _sqlBuidler; }
            set { _sqlBuidler = value; }
        }
        string _filterFrom = "";

        public string FilterFrom
        {
            get { return _filterFrom; }
            set { _filterFrom = value; }
        }
        string _filterTo = "";

        public string FilterTo
        {
            get { return _filterTo; }
            set { _filterTo = value; }
        }
        public frmFilterSelect(string connect, QueryBuilder.SQLBuilder sqlBuilder, int indexFilter)
        {
            InitializeComponent();
            _connectString = connect;
            _indexFilter = indexFilter;
            _sqlBuidler = new QueryBuilder.SQLBuilder(sqlBuilder);
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.OK;
            _filterFrom = txtFilterFrom.Text;
            _filterTo = txtFilterTo.Text;
            Close();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
            Close();
        }
        bool flag = true;



        private void frmFilterSelect_Load(object sender, EventArgs e)
        {
            QueryBuilder.Node node = _sqlBuidler.Filters[_indexFilter].Node;
            if (node.FType == "SPN")
            {
                _sqlBuidler.SelectedNodes.Clear();
                _sqlBuidler.SelectedNodes.Add(_sqlBuidler.Filters[_indexFilter].Node);
                _sqlBuidler.Filters.RemoveAt(_indexFilter);

                try
                {
                    _sqlBuidler.SelectedNodes[0].Agregate = "Min";
                    object min = _sqlBuidler.BuildObject("", _connectString);

                    _sqlBuidler.SelectedNodes[0].Agregate = "Max";
                    object max = _sqlBuidler.BuildObject("", _connectString);

                    int minyear = DateTime.Now.Year - 2;
                    int maxyear = DateTime.Now.Year + 2;

                    if (min != null)
                    {
                        minyear = ((int)min) / 1000;
                    }
                    if (max != null)
                    {
                        maxyear = ((int)max) / 1000;
                    }

                    DataTable dt = new DataTable();
                    DataColumn[] col = new DataColumn[] { new DataColumn("SELECTED", typeof(bool)), new DataColumn("VALUE"), new DataColumn("Lookup"), new DataColumn("Description") };
                    dt.Columns.AddRange(col);
                    for (int i = minyear; i <= maxyear; i++)
                    {
                        for (int j = 1; j < 13; j++)
                        {
                            DataRow row = dt.NewRow();
                            row["SELECTED"] = false;
                            row["VALUE"] = i.ToString("0000") + j.ToString("000");
                            row["Lookup"] = i.ToString("0000");
                            row["Description"] = j.ToString("000") + "/" + i.ToString("0000");
                            dt.Rows.Add(row);
                        }
                    }
                    dgvSelect.RootTable.Columns["VALUE"].Caption = "Code";
                    dgvSelect.DataSource = dt;
                }
                catch (Exception ex)
                { }
            }
            else
            {
                _sqlBuidler.SelectedNodes.Clear();

                _sqlBuidler.SelectedNodes.Add(_sqlBuidler.Filters[_indexFilter].Node);
                if (_sqlBuidler.Filters[_indexFilter].Node.NodeDesc != "")
                {
                    QueryBuilder.Node nodeNew = new QueryBuilder.Node(_sqlBuidler.Filters[_indexFilter].Node.NodeDesc, "Description");
                    nodeNew.AddMeToParent(_sqlBuidler.Filters[_indexFilter].Node.MyFamily);
                    _sqlBuidler.SelectedNodes.Add(nodeNew);
                }
                else
                    dgvSelect.RootTable.Columns["Description"].Visible = false;

                _sqlBuidler.Filters.RemoveAt(_indexFilter);
                _sqlBuidler.StrConnectDes = _connectString;

                try
                {
                    DataTable dt = _sqlBuidler.BuildDataTable("");
                    dgvSelect.RootTable.Columns["VALUE"].Caption = dt.Columns[0].ColumnName;
                    dgvSelect.RootTable.Columns["VALUE"].DataMember = dt.Columns[0].ColumnName;
                    if (dt.Columns.Count == 2)
                    {
                        dgvSelect.RootTable.Columns["Description"].Caption = dt.Columns[1].ColumnName;
                        dgvSelect.RootTable.Columns["Description"].DataMember = dt.Columns[1].ColumnName;
                    }
                    //dt.Columns[0].ColumnName = "VALUE";

                    DataColumn col = new DataColumn("SELECTED", typeof(bool));
                    dt.Columns.Add(col);
                    foreach (DataRow row in dt.Rows)
                    {
                        row["SELECTED"] = false;
                    }
                    dgvSelect.DataSource = dt;
                }
                catch (Exception ex)
                { }

            }
            dgvSelect.RootTable.Columns["VALUE"].Width = 100;
            //dgvSelect.RootTable.Columns["VALUE"].ReadOnly = true;
            txtFilterFrom.Text = _filterFrom;
            txtFilterTo.Text = _filterTo;
        }

        private void txtFilterFrom_Enter(object sender, EventArgs e)
        {
            flag = true;
        }

        private void txtFilterTo_Enter(object sender, EventArgs e)
        {
            flag = false;
        }


        private void dgvSelect_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            if (dgvSelect.CurrentRow != null && dgvSelect.CurrentRow.RowIndex >= 0)
            {
                if (flag) txtFilterFrom.Text = dgvSelect.CurrentRow.Cells["VALUE"].Value.ToString();
                else txtFilterTo.Text = dgvSelect.CurrentRow.Cells["VALUE"].Value.ToString();
                flag = !flag;
            }

            for (int i = 0; i < dgvSelect.RowCount; i++)
            {
                dgvSelect.GetRow(i).BeginEdit();
                dgvSelect.GetRow(i).Cells["SELECTED"].Value = false;
                dgvSelect.GetRow(i).EndEdit();
            }
        }

        private void dgvSelect_CellEdited(object sender, Janus.Windows.GridEX.ColumnActionEventArgs e)
        {
            string filterFrom = "";
            int dem = 0;
            if (dgvSelect.CurrentRow != null && dgvSelect.CurrentRow.RowIndex >= 0 && e.Column.Index > -1 && dgvSelect.RootTable.Columns[e.Column.Index].Key == "SELECTED")
            {
                //dgvSelect.CurrentRow.Cells["SELECTED"]._Value = !(bool)dgvSelect.CurrentRow.Cells["SELECTED"]._Value;
                for (int i = 0; i < dgvSelect.RowCount; i++)
                {
                    if (dgvSelect.GetRow(i).Cells["SELECTED"].Value != null && (bool)dgvSelect.GetRow(i).Cells["SELECTED"].Value == true)
                    {
                        filterFrom += "," + dgvSelect.GetRow(i).Cells["VALUE"].Value.ToString();
                        dem++;
                    }
                }
            }
            if (dem == 1)
            {
                if (flag) txtFilterFrom.Text = filterFrom.Substring(1);
                else txtFilterTo.Text = filterFrom.Substring(1);
                flag = !flag;
            }
            else if (dem > 1)
            {
                txtFilterFrom.Text = "" + filterFrom.Substring(1);
                txtFilterTo.Text = "";
            }
        }

    }
}
