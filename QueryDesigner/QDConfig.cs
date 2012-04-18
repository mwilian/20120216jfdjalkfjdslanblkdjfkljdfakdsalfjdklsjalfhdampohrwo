using System.IO;
using System.Text;
using System.Data;
namespace dCube
{


    public partial class QDConfig
    {
        public void SaveConfig(string filename)
        {
            StringBuilder sb = new StringBuilder();
            StringWriter sw = new StringWriter(sb);
            WriteXml(sw);
            sw.Close();
            string temp = RC2.EncryptString(sb.ToString(), Form_QD._key, Form_QD._iv, Form_QD._padMode, Form_QD._opMode);
            StreamWriter streamw = new StreamWriter(filename);
            streamw.Write(temp);
            streamw.Close();

        }
        public void Clear()
        {
            foreach (System.Data.DataTable dt in Tables)
            {
                dt.Rows.Clear();
            }
        }
        public void LoadConfig(string filename)
        {
            if (File.Exists(filename))
            {
                Clear();
                StreamReader sr = new StreamReader(filename);
                string result = sr.ReadToEnd();
                sr.Close();
                string kq = RC2.DecryptString(result, Form_QD._key, Form_QD._iv, Form_QD._padMode, Form_QD._opMode);
                StringReader stringReader = new StringReader(kq);
                ReadXml(stringReader);
                stringReader.Close();
            }
            if (DTB.Rows.Count == 0)
            {
                DTB.AddDTBRow("", "");
            }
            if (DIR.Rows.Count == 0)
            {
                DIR.AddDIRRow("", "");
            }
            if (SYS.Rows.Count == 0)
            {
                SYS.AddSYSRow("", "", "", false);

            }
        }
        public string GetConnection(ref string key, string type)
        {
            string _strConnectDes = "";
            if (this.DTB.Rows.Count > 0)
            {
                foreach (DataRow row in this.ITEM.Rows)
                {
                    if (key == "")
                    {
                        if (row["KEY"].ToString() == this.DTB.Rows[0][type].ToString())
                        {
                            _strConnectDes = row["CONTENT"].ToString();
                            //SQLBuilder.ConnID = row["KEY"].ToString();

                            key = row["KEY"].ToString();
                            return _strConnectDes;
                        }
                    }
                    else
                    {
                        //_connDes = strAP;
                        if (row["KEY"].ToString() == key && row["TYPE"].ToString() == type)
                        {
                            this.DTB.Rows[0]["AP"] = key;
                            _strConnectDes = row["CONTENT"].ToString();
                            //SQLBuilder.ConnID = strAP;
                            return _strConnectDes;
                        }
                    }
                }
            }
            return _strConnectDes;
        }
    }

}
