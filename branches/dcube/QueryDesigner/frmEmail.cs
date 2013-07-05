using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;

namespace dCube
{
    public partial class frmEmail : Form
    {
        public frmEmail(string dtb)
        {
            InitializeComponent();
            _dtb = dtb;
        }
        string _dtb = "";
        private void frmEmail_Load(object sender, EventArgs e)
        {
            LoadDataGrid(_dtb);
        }

        private void LoadDataGrid(string dtb)
        {
            BUS.CommonControl ctr = new BUS.CommonControl();
            dgvList.DataSource = ctr.executeSelectQuery("Select * from LIST_EMAIL");
        }

        private void btSave_Click(object sender, EventArgs e)
        {
            BUS.CommonControl ctr = new BUS.CommonControl();

            DataTable dt = dgvList.DataSource as DataTable;
            if (dt != null)
            {
                ctr.executeNonQuery("Delete from LIST_EMAIL");
                string query = "";
                foreach (DataRow row in dt.Rows)
                {
                    query += " insert into LIST_EMAIL(Mail, _Name,Lookup) values('" + row["Mail"].ToString() + "','" + row["_Name"].ToString() + "','" + row["Lookup"].ToString() + "')";
                }
                ctr.executeNonQuery(query);

            }
            DialogResult = DialogResult.OK;
            Close();
        }

        private void btCancel_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
            Close();
        }

        private void btRefresh_Click(object sender, EventArgs e)
        {
            LoadDataGrid(_dtb);
        }

        private void btImport_Click(object sender, EventArgs e)
        {

        }

        private void btExport_Click(object sender, EventArgs e)
        {

        }

        private void cutToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void copyToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string kq = "";
            for (int i = dgvList.SelectedItems[0].Position; i < dgvList.SelectedItems.Count; i++)
            {
                string tmp = "";
                for (int j = 0; j < dgvList.GetRow(i).Cells.Count; j++)
                {
                    if (dgvList.GetRow(i).Cells[j].Value != null)
                        tmp += "\t" + dgvList.GetRow(i).Cells[j].Value.ToString();
                    else tmp += "\t";
                }
                kq += "\n" + tmp.Substring(1);
            }
            kq = kq.Substring(1);
            Clipboard.SetText(kq, TextDataFormat.Text);
        }

        private void pasteToolStripMenuItem_Click(object sender, EventArgs e)
        {

            DataTable dt = dgvList.DataSource as DataTable;
            IDataObject obj = Clipboard.GetDataObject();
            using (StreamReader sr = new StreamReader((System.IO.MemoryStream)obj.GetData(System.Windows.Forms.DataFormats.CommaSeparatedValue)))
            {
                string line = sr.ReadLine();
                while (line != null && line != "\0")
                {
                    if (line != "")
                        dt.Rows.Add(line.Split(','));

                    line = sr.ReadLine();
                }
            }
            dgvList.DataSource = dt;

        }

        private void deleteToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }
    }
}
