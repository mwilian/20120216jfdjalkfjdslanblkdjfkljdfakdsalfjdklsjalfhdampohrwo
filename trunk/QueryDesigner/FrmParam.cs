using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;

using System.Text;
using System.Windows.Forms;


namespace dCube
{
    public partial class FrmParam : Form
    {
        BindingList<QueryBuilder.Node> _Nodes;
        QueryBuilder.Filter _Filter;

        public QueryBuilder.Filter Filter
        {
            get { return _Filter; }
            set { _Filter = value; }
        }
        string THEME = "Breeze";
        public FrmParam()
        {
            InitializeComponent();
            //ThemeResolutionService.ApplyThemeToControlTree(this, THEME);
        }       
        private void btSave_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.OK;
            string type = "";
            if (rdDate.Checked)
                type = "SDN";
            else if (rdNum.Checked)
                type = "N";
            QueryBuilder.Node _Node = new QueryBuilder.Node("", "__" + txtName.Text, "__" + txtName.Text, type,"");
            //_Node.Expresstion = txtMain.Text;
            _Filter = new QueryBuilder.Filter(_Node);
            _Filter.FilterFrom = _Filter.ValueFrom = txtMain.Text;
            Close();
        }

        private void btCancel_Click(object sender, EventArgs e)
        {

        }

        private void FrmUserFunc_Load(object sender, EventArgs e)
        {
            rdText.Checked = true;

        }
    }
       
}
