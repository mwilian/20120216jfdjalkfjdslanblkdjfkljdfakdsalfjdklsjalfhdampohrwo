using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;

using System.Text;
using System.Windows.Forms;

namespace dCube
{
    public partial class FrmUserFunc : Form
    {
        BindingList<QueryBuilder.Node> _Nodes;
        QueryBuilder.Node _Node;

        public QueryBuilder.Node Node
        {
            get { return _Node; }
            set { _Node = value; }
        }

        public BindingList<QueryBuilder.Node> Nodes
        {
            get { return _Nodes; }
            set { _Nodes = value; }
        }
        public FrmUserFunc()
        {
            InitializeComponent();
        }
        public FrmUserFunc(BindingList<QueryBuilder.Node> nodes)
        {
            InitializeComponent();
            _Nodes = nodes;
        }
        public FrmUserFunc(BindingList<QueryBuilder.Node> nodes, QueryBuilder.Node node)
        {
            InitializeComponent();
            _Nodes = nodes;
            if (node != null)
            {
                _Node = node;
                btnNew.Visible = true;
                txtName.Enabled = false;
                int vitri = -1;
                for (int i = 0; i < nodes.Count; i++)
                {
                    if (node.Code == nodes[i].Code)
                    {
                        vitri = i;
                    }
                }
                if (vitri != -1)
                    nodes.RemoveAt(vitri);
            }
            else
                btnNew.Visible = false;
        }

        private void btSave_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.OK;
            string type = "";
            if (rdDate.Checked)
                type = "SD";
            else if (rdNum.Checked)
                type = "N";
            _Node = new QueryBuilder.Node("", txtName.Text, txtName.Text, type,"");
            _Node.Expresstion = txtMain.Text;
            Close();
        }

        private void btCancel_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void FrmUserFunc_Load(object sender, EventArgs e)
        {
            rdText.Checked = true;
            //dgvSelectNodes.MasterGridViewTemplate.AutoGenerateColumns = false;
            dgvSelectNodes.DataSource = _Nodes;
            if (_Node != null)
            {
                txtMain.Text = _Node.Expresstion;
                txtName.Text = _Node.Code;
                if (_Node.FType == "N")
                    rdNum.Select();
                else if (_Node.FType == "SND")
                    rdDate.Select();
                else
                    rdText.Select();
            }
        }

        private void dgvSelectNodes_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                if (dgvSelectNodes.CurrentRow != null)
                {

                    QueryBuilder.Node x = (QueryBuilder.Node)dgvSelectNodes.CurrentRow.DataRow;

                    if (x != null)
                        DoDragDrop("[" + x.MyCode + "]", DragDropEffects.All);

                }
            }
        }

        private void txtMain_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.Text))
                e.Effect = DragDropEffects.Copy;
            else
                e.Effect = DragDropEffects.None;

        }

        public void txtMain_DragDrop(object sender, DragEventArgs e)
        {
            // Get Start Position to Drop the Text 

            int i = txtMain.SelectionStart;
            String s = txtMain.Text.Substring(i);
            txtMain.Text = txtMain.Text.Substring(0, i);

            if (e.Data.GetDataPresent(typeof(System.String)))
            {
                txtMain.Text = txtMain.Text + (System.String)e.Data.GetData(typeof(System.String));
            }
            txtMain.Text = txtMain.Text + s;
            e.Effect = DragDropEffects.None;

        }

        private void btnNew_Click(object sender, EventArgs e)
        {
            txtName.Enabled = true;
            txtName.Text = "";
            txtMain.Text = "";
            if (_Node != null)
                Nodes.Add(_Node);
            btnNew.Visible = false;
            dgvSelectNodes.DataSource = Nodes;
        }

    }
}
