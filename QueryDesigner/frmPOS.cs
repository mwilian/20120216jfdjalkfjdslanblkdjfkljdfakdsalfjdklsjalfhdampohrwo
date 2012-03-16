using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace QueryDesigner
{
    public partial class frmPOS : Form
    {
        string sErr = "";
        string _user = "";
        public frmPOS(string user)
        {
            InitializeComponent();
            _user = user;
        }

        private void btnClear_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                RemoveAt(i);
            }
            LoadDataGrid();
        }

        private void frmPOS_Load(object sender, EventArgs e)
        {
            //dataGridView1.AutoGenerateColumns = false;
            LoadDataGrid();
        }

        private void LoadDataGrid()
        {
            BUS.POSControl ctr = new BUS.POSControl();
            dataGridView1.DataSource = ctr.GetAll(ref sErr);
        }

       

        private void RemoveAt(int index)
        {
            DataRowView rview = dataGridView1.GetRow(index).DataRow as DataRowView;
            DTO.POSInfo inf = new DTO.POSInfo(rview.Row);
            if (_user != inf.USER_ID)
            {
                BUS.POSControl ctr = new BUS.POSControl();
                ctr.Delete(inf.USER_ID);
            }
        }

        private void dataGridView1_Click(object sender, EventArgs e)
        {
            if (dataGridView1.Row >= 0 && dataGridView1.Col == 0)
            {
                RemoveAt(dataGridView1.Row);
                LoadDataGrid();
            }
        }
    }
}
