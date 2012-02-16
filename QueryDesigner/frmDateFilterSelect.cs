using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Text.RegularExpressions;



namespace QueryDesigner
{
    public partial class frmDateFilterSelect : Form
    {
        string _type = "";
        public string Type
        {
            get { return _type; }
            set { _type = value; }
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
        public frmDateFilterSelect(string type)
        {
            InitializeComponent();
            _type = type;
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.OK;

            if (_type == "SDN")
            {
                if (txtFilterFrom.MinDate != txtFilterFrom.Value)
                    _filterFrom = Convert.ToString(txtFilterFrom.Value.Year * 10000 + txtFilterFrom.Value.Month * 100 + txtFilterFrom.Value.Day);
                else _filterFrom = "";
                if (txtFilterTo.MinDate != txtFilterTo.Value)
                    _filterTo = Convert.ToString(txtFilterTo.Value.Year * 10000 + txtFilterTo.Value.Month * 100 + txtFilterTo.Value.Day);
                else _filterTo = "";
            }
            else
            {
                if (txtFilterFrom.MinDate != txtFilterFrom.Value)
                    _filterFrom = txtFilterFrom.Value.Year + "-" + txtFilterFrom.Value.Month + "-" + txtFilterFrom.Value.Day;
                else _filterFrom = "";
                if (txtFilterTo.MinDate != txtFilterTo.Value)
                    _filterTo = txtFilterTo.Value.Year + "-" + txtFilterTo.Value.Month + "-" + txtFilterTo.Value.Day;
                else _filterTo = "";
            }
            Close();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
            Close();
        }
        bool flag = true;


        private void monthCalendar_DateSelected(object sender, DateRangeEventArgs e)
        {

        }

        private void frmDateFilterSelect_Load(object sender, EventArgs e)
        {
            txtFilterFrom.SetToNullValue();
            txtFilterTo.SetToNullValue();
            if (Regex.IsMatch(FilterFrom, @"^[0-9]{8}$"))
            {
                txtFilterFrom.Value = DateTime.Parse(FilterFrom.Substring(0, 4) + "-" + FilterFrom.Substring(4, 2) + "-" + FilterFrom.Substring(6, 2));
            }
            else if (Regex.IsMatch(FilterFrom, @"^[0-9]{4}-[0-9]+[0-9]+$"))
            {
                txtFilterFrom.Value = DateTime.Parse(FilterFrom);
            }
            else if (FilterFrom == "C")
                txtFilterFrom.Value = DateTime.Today;

            if (Regex.IsMatch(FilterTo, @"^[0-9]{8}$"))
            {
                txtFilterTo.Value = DateTime.Parse(FilterTo.Substring(0, 4) + "-" + FilterTo.Substring(4, 2) + "-" + FilterTo.Substring(6, 2));
            }
            else if (Regex.IsMatch(FilterTo, @"^[0-9]{4}-[0-9]+[0-9]+$"))
            {
                txtFilterTo.Value = DateTime.Parse(FilterTo);
            }
            else if (FilterTo == "C")
                txtFilterTo.Value = DateTime.Today;

        }

        private void txtFilterFrom_Enter(object sender, EventArgs e)
        {
            flag = true;
        }

        private void txtFilterTo_Enter(object sender, EventArgs e)
        {
            flag = false;
        }

        private void monthCalendar1_MouseDown(object sender, MouseEventArgs e)
        {
            if (flag)
            {
                //txtFilterFrom.Text = monthCalendar1.SelectionStart.Day + "/" + monthCalendar1.SelectionStart.Month + "/" + monthCalendar1.SelectionStart.Year;
                txtFilterFrom.Value = monthCalendar1.SelectionStart;
            }
            else
            {
                //txtFilterTo.Text = monthCalendar1.SelectionStart.Day + "/" + monthCalendar1.SelectionStart.Month + "/" + monthCalendar1.SelectionStart.Year;
                txtFilterTo.Value = monthCalendar1.SelectionStart;
            }
            flag = !flag;
        }
    }
}
