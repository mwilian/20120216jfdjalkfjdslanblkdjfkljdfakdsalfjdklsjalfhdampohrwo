using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Management;

namespace dCube
{
    public partial class FrmLicense : Form
    {
        public string THEME = "";
        string _key = "newoppo123456789";
        string _iv = "12345678";
        string _padMode = "PKCS7";
        string _opMode = "CBC";
        string _appPath = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase).Replace("file:\\", "");
        //string _pathLicense = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase) + "\\License.bin";
        public FrmLicense()
        {
            InitializeComponent();
        }

        private void FrmLicense_Load(object sender, EventArgs e)
        {
            lbErr.Text = "";
            BUS.CommonControl ctr = new BUS.CommonControl();
            object data = ctr.executeScalar(@"SELECT SUN_DATA  FROM SSINSTAL WHERE INS_TB='LCS' and INS_KEY='QD'");

            if (data != null)// if (File.Exists(_pathLicense.Replace("file:\\", "")))
            {
                //StreamReader reader = new StreamReader(_pathLicense.Replace("file:\\", ""));
                string result = data.ToString();
                string kq = RC2.DecryptString(result, _key, _iv, _padMode, _opMode);
                string[] tmp = kq.Split(';');
                DTO.License license = new DTO.License();
                license.CompanyName = tmp[0];
                license.ExpiryDate = Convert.ToInt32(tmp[1]);
                license.Modules = tmp[2];
                license.NumUsers = Convert.ToInt32(tmp[3]);
                license.SerialNumber = tmp[4];
                license.Key = tmp[5];
                license.SerialCPU = tmp[6];
                SetDataToForm(license);
                //reader.Close();
            }
        }

        private void SetDataToForm(DTO.License license)
        {
            int year = license.ExpiryDate / 10000;
            int month = (license.ExpiryDate - year * 10000) / 100;
            int day = (license.ExpiryDate - year * 10000 - month * 100);
            txtCompany.Text = license.CompanyName;
            txtSerial.Text = license.SerialNumber;
            txtNumUser.Text = license.NumUsers.ToString();

            string qd = license.Modules.Substring(0, 1);
            string add = license.Modules.Substring(1, 1);
            string web = license.Modules.Substring(2, 1);
            string qdadd = license.Modules.Substring(3, 1);
            string task = "";
            if (license.Modules.Length > 4)
                task = license.Modules.Substring(4, 1);
            ckbQD.Checked = qd == "Y" ? true : false;
            ckbAddin.Checked = add == "Y" ? true : false;
            ckbWeb.Checked = web == "Y" ? true : false;
            ckbQDADD.Checked = qdadd == "Y" ? true : false;
            ckbTask.Checked = task == "Y" ? true : false;
            DateTime date = new DateTime(year, month, day);
            dtExpiryDate.Value = date;
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            DTO.License license = GetDataFromForm();
            string param = license.CompanyName + license.SerialNumber + license.NumUsers.ToString() + license.Modules + license.ExpiryDate.ToString() + license.SerialCPU;


            string tmp = RC2.EncryptString(param, _key, _iv, _padMode, _opMode);
            string key = Convert.ToBase64String(new System.Security.Cryptography.SHA1CryptoServiceProvider().ComputeHash(Encoding.ASCII.GetBytes(tmp)));
            if (key == license.Key)
            {
                String kq = license.CompanyName + ";" +
                            license.ExpiryDate + ";" +
                            license.Modules + ";" +
                            license.NumUsers + ";" +
                            license.SerialNumber + ";" +
                            license.Key + ";" +
                            license.SerialCPU;


                BUS.CommonControl ctr = new BUS.CommonControl();
                string query = @"if EXISTS(SELECT INS_KEY  FROM SSINSTAL WHERE INS_TB='LCS' and INS_KEY='QD') 
UPDATE SSINSTAL SET SUN_DATA = '{0}' WHERE INS_TB='LCS' and INS_KEY='QD'
else 
INSERT INTO SSINSTAL(INS_TB ,INS_KEY ,SUN_DATA) VALUES ( 'LCS' ,'QD' ,'{0}')";
                string result = RC2.EncryptString(kq, _key, _iv, _padMode, _opMode);
                query = string.Format(query, result);
                ctr.executeNonQuery(query);
                //StreamWriter writerStream = new StreamWriter(_pathLicense.Replace("file:\\", ""));
                //writerStream.WriteLine(result);
                //writerStream.Close();

                Close();
                DialogResult = DialogResult.OK;
            }
            else
                lbErr.Text = "Registry fail";
        }

        private DTO.License GetDataFromForm()
        {
            DTO.License result = new DTO.License();
            result.CompanyName = txtCompany.Text.Trim();
            result.SerialNumber = txtSerial.Text.Trim();
            result.NumUsers = Convert.ToInt32(txtNumUser.Text);
            string qd = ckbQD.Checked == true ? "Y" : " ";
            string add = ckbAddin.Checked == true ? "Y" : " ";
            string web = ckbWeb.Checked == true ? "Y" : " ";
            string qdadd = ckbQDADD.Checked == true ? "Y" : " ";
            string task = ckbTask.Checked == true ? "Y" : " ";
            result.Modules = qd + add + web + qdadd + task;
            DateTime dateExpire = dtExpiryDate.Value;
            result.ExpiryDate = dateExpire.Year * 10000 + dateExpire.Month * 100 + dateExpire.Day;
            result.Key = txtKey.Text.Trim();
            BUS.CommonControl ctr = new BUS.CommonControl();
            result.SerialCPU = ctr.executeScalar(@"SELECT   CONVERT(varchar(200), SERVERPROPERTY('servername'))").ToString(); //"BFEBFBFF000006FD";
            return result;
        }
        public static string GetProcessorId()
        {
            string strCPU = "";
            string Key = "Win32_DiskDrive";
            ManagementObjectSearcher searcher = new ManagementObjectSearcher("select * from " + Key + " where InterfaceType ='IDE'");
            try
            {
                foreach (ManagementObject share in searcher.Get())
                {
                    if (share.Properties.Count <= 0)
                    {
                        return "";
                    }
                    foreach (PropertyData PC in share.Properties)
                    {
                        if (PC.Name.Contains("SerialNumber") || PC.Name.Contains("SerialNumber"))
                        {
                            strCPU += Convert.ToString(PC.Value);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
            }
            return strCPU;
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void radGroupBox1_Click(object sender, EventArgs e)
        {
        }

        private void label5_Click(object sender, EventArgs e)
        {

        }


    }
}