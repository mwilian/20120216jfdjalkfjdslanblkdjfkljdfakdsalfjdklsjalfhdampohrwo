using System;
using System.Data;
using System.Configuration;




using System.Collections.Generic;
using System.Xml;
using System.ComponentModel;
using System.IO;
using System.Windows.Forms;


namespace dCube.Configuration
{
    public class clsConfigurarion
    {
        Dictionary<string, string> _dictionary = new Dictionary<string, string>();
        BindingList<clsControlsInForm> _controlsInForm = new BindingList<clsControlsInForm>();
        DataTable _dtDictionary = new DataTable();
        public DataTable DtDictionary
        {
            get { return _dtDictionary; }
            set { _dtDictionary = value; }
        }
        public BindingList<clsControlsInForm> ControlsInForm
        {
            get { return _controlsInForm; }
            set { _controlsInForm = value; }
        }

        public Dictionary<string, string> Dictionary
        {
            get { return _dictionary; }
            set { _dictionary = value; }
        }
        public void GetDictionary(string filename)
        {
            XmlTextReader reader = null;

            try
            {
                reader = new XmlTextReader(filename);
                string aliascode = string.Empty;
                string origin = string.Empty;

                while (reader.Read())
                {
                    if (reader.NodeType == XmlNodeType.Element & reader.Name == "ITEM")
                    {
                        aliascode = reader.GetAttribute("Code");
                        origin = reader.GetAttribute("_Value");
                        _dictionary.Add(aliascode, origin);
                    }
                }

            }
            catch (Exception ex)
            {
                //throw new ArgumentException("Alias Dictionary Building problem " + ex.Message);
            }
            finally
            {
                if (!(reader == null))
                {
                    reader.Close();
                }
            }
        }
        public void GetDataTableDictionary(string filename)
        {
            DataColumn[] col = new DataColumn[] { new DataColumn("Code"), new DataColumn("_Value") };
            _dtDictionary.Columns.Clear();
            _dtDictionary.Rows.Clear();
            _dtDictionary.Columns.AddRange(col);
            XmlTextReader reader = null;

            try
            {
                reader = new XmlTextReader(filename);
                string aliascode = string.Empty;
                string origin = string.Empty;

                while (reader.Read())
                {
                    if (reader.NodeType == XmlNodeType.Element & reader.Name == "ITEM")
                    {
                        DataRow row = _dtDictionary.NewRow();
                        row["Code"] = reader.GetAttribute("Code");
                        row["_Value"] = reader.GetAttribute("_Value");

                        _dtDictionary.Rows.Add(row);
                    }
                }

            }
            catch (Exception ex)
            {
                //throw new ArgumentException("Alias Dictionary Building problem " + ex.Message);
            }
            finally
            {
                if (!(reader == null))
                {
                    reader.Close();
                }
            }
        }
        public void GetControlsInForm(string formCode)
        {
            XmlTextReader reader = null;

            try
            {
                // System.Resources.ResourceManager.
                //    StringReader stream = new StringReader(System.Convert.ToString(Properties.Resources.ResourceManager.GetObject(origin))); // if table = ICAS - load from origin CA
                // StringReader stream = new StringReader(Reporting.Proper
                reader = new XmlTextReader(formCode);
                string id = string.Empty;
                string type = string.Empty;
                string index = "";
                string languageCode = "";

                while (reader.Read())
                {
                    if (reader.NodeType == XmlNodeType.Element & reader.Name == "ITEM")
                    {
                        id = reader.GetAttribute("ID");
                        type = reader.GetAttribute("Type");
                        index = reader.GetAttribute("Index");
                        languageCode = reader.GetAttribute("LanguageCode");
                        clsControlsInForm control = new clsControlsInForm(id, type, index, languageCode);
                        _controlsInForm.Add(control);
                    }
                }

            }
            catch (Exception ex)
            {
                //   throw new ArgumentException("Alias Dictionary Building problem " + ex.Message);
            }
            finally
            {
                if (!(reader == null))
                {
                    reader.Close();
                }

            }
        }
        public clsConfigurarion()
        {


        }

        public string GetValueDictionary(string key)
        {
            if (_dictionary.ContainsKey(key))
            {

                return _dictionary[key];
            }
            else
            {
                return key;
            }
        }
        public static void SetLanguages(System.Windows.Forms.Form page, string title, string language)
        {
            Configuration.clsConfigurarion config = new Configuration.clsConfigurarion();
            config.GetDictionary(System.Windows.Forms.Application.StartupPath + "/Configuration/Forms.xml");
            config.GetControlsInForm(System.Windows.Forms.Application.StartupPath + "/Configuration/Form/" + config.Dictionary[page.Name]);
            config.GetDictionary(System.Windows.Forms.Application.StartupPath + "/Configuration/Languages/" + language);
            SetLanguage(config, ref page, title);

        }
        private static void SetLanguage(clsConfigurarion config, ref System.Windows.Forms.Form page, string title)
        {
            page.Text = config.GetValueDictionary(title);
            foreach (Configuration.clsControlsInForm x in config.ControlsInForm)
            {
                string codeLanguage = x.LanguageCode;
                string id = x.Id;
                System.Windows.Forms.Control[] temp = (page.Controls.Find(id, true));
                if (temp.Length > 0)
                {
                    System.Windows.Forms.Control tmp = temp[0];
                    if (tmp != null)
                    {
                        switch (x.Type)
                        {
                            case "Label":
                                System.Windows.Forms.Label tmpLabel = (System.Windows.Forms.Label)tmp;
                                tmpLabel.Text = config.GetValueDictionary(codeLanguage);
                                break;
                            //case "RadLabel":
                            //    Telerik.WinControls.UI.RadLabel tmpRadLabel = (Telerik.WinControls.UI.RadLabel)tmp;
                            //    tmpRadLabel.Text = config.GetValueDictionary(codeLanguage);
                            //    break;
                            //case "RadTabStrip":
                            //    Telerik.WinControls.UI.RadTabStrip tmpRadTabStrip = (Telerik.WinControls.UI.RadTabStrip)tmp;
                            //    tmpRadTabStrip.Items[x.Index].Text = config.GetValueDictionary(codeLanguage); ;

                            //    break;
                            //case "RadGrid":
                            //    Telerik.WinControls.UI.RadGridView tmpRadGrid = (Telerik.WinControls.UI.RadGridView)tmp;
                            //    tmpRadGrid.Columns[x.Index].HeaderText = config.GetValueDictionary(codeLanguage); ;

                            //    // tmpObj.Tabs = config.GetValueDictionary(codeLanguage);
                            //    break;
                            //case "RadMenu":
                            //    Telerik.WinControls.UI.RadMenu tmpMenu = tmp as Telerik.WinControls.UI.RadMenu;
                            //    tmpMenu.Items[x.Index].Text = config.GetValueDictionary(codeLanguage);
                            //    break;                            
                            //case "RadButton":
                            //    Telerik.WinControls.UI.RadButton tmpRadButton = tmp as Telerik.WinControls.UI.RadButton;
                            //    tmpRadButton.Text = config.GetValueDictionary(codeLanguage);
                            //    // tmpObj.Tabs = config.GetValueDictionary(codeLanguage);
                            //    break;
                            case "Button":
                                System.Windows.Forms.Button tmpButton = tmp as System.Windows.Forms.Button;
                                tmpButton.Text = config.GetValueDictionary(codeLanguage);
                                // tmpObj.Tabs = config.GetValueDictionary(codeLanguage);
                                break;
                        }
                    }
                }
            }
            //RadGrid1.Rebind();
            //RadGrid2.Rebind();
        }

    }
}
