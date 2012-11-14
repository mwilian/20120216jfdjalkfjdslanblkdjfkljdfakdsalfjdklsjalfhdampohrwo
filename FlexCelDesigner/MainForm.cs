using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Configuration;
using System.IO;
using System.Reflection;
using System.Text;
using System.Xml;
using FlexCel.Core;
using FlexCel.Report;
using FlexCel.XlsAdapter;


namespace TVCDesigner
{
    /// <summary>
    /// FlexCel Dsigner main form
    /// </summary>
    public partial class MainForm : System.Windows.Forms.Form
    {

        public MainForm(DataTable dtList, DataTable dtFilter, DataTable dtParam)
        {
            InitializeComponent();
            dt_list = dtList;
            dt_Filter = dtFilter;
            dt_Params = dtParam;
            FillBasicListView(null);

            try
            {
                LoadConfig();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private const int imgExtras = 0;
        private const int imgDataset = 1;
        private const int imgTable = 2;
        private const int imgColumn = 3;
        private const int imgOneExtra = 4;
        private const int imgUserDefined = 5;
        private const int imgUserDefinedTable = 6;
        private const int imgUserDefinedColumn = 7;
        private const int imgReportVarList = 8;
        private const int imgReportVar = 9;
        private const int imgFullConfig = 10;
        private const int imgReportExpList = 11;
        private const int imgReportExp = 12;
        private const int imgReportFormatList = 13;
        private const int imgReportFormat = 14;
        private const int imgConfig = 15;
        private const int imgConfigTag = 16;


        private readonly int FirstConfigRow = 10;
        private readonly int ConfigColTableName = 1;
        private readonly int ConfigColSourceName = 2;
        private readonly int ConfigColVarName = 11;
        private readonly int ConfigColExpName = 13;
        private readonly int ConfigColFormatName = 8;
        private readonly int RowXsd = 6;
        private readonly int ColXsd = 2;
        private DataTable dt_Filter;
        private DataTable dt_list;
        private DataTable dt_Params;
        private void miExit_Click(object sender, System.EventArgs e)
        {
            Close();
        }

        private void miAlwaysOnTop_Click(object sender, System.EventArgs e)
        {
            miAlwaysOnTop.Checked = !miAlwaysOnTop.Checked;
            TopMost = miAlwaysOnTop.Checked;
        }

        private void changeOpacity_Click(object sender, System.EventArgs e)
        {
            string s = ((MenuItem)sender).Text;
            Opacity = Convert.ToDouble(s.Substring(0, s.IndexOf("%"))) / 100;
            miOpacity.Text = "Opacity " + (Opacity * 100).ToString() + "%";

        }

        private void ReloadList(TreeNodeCollection nc)
        {
            if (nc == null) return;
            foreach (TreeNode n in nc)
            {
                ReloadList(n.Nodes);
                if (n.ImageIndex == imgColumn || n.ImageIndex == imgUserDefinedColumn)
                {
                    if (miUseColumnCaptions.Checked) n.Text = ((string[])n.Tag)[2]; else n.Text = ((string[])n.Tag)[1];
                }
            }
        }

        private void FillIfListView(TreeNode IfNode)
        {
            string strOpenParen = TFormulaMessages.TokenString(TFormulaToken.fmOpenParen);
            string strCloseParen = TFormulaMessages.TokenString(TFormulaToken.fmCloseParen);

            TImplementedFunctionList fil = new TImplementedFunctionList();
            foreach (TImplementedFunction fi in fil.Values)
            {
                TreeNode n = IfNode.Nodes.Add(fi.FunctionName.ToLower());
                string ParamList = string.Empty;
                int MinParamCount = fi.MinArgCount;
                if (MinParamCount <= 0) MinParamCount = 1;
                int MaxParamCount = fi.MaxArgCount;
                string Optional = MaxParamCount > MinParamCount ? "..." : "";
                ParamList = strOpenParen.PadRight(strOpenParen.Length + MinParamCount - 1, TFormulaMessages.TokenChar(TFormulaToken.fmFunctionSep)) + Optional + strCloseParen;

                string[] sc = { n.Text + ParamList };
                n.ImageIndex = imgOneExtra;
                n.SelectedImageIndex = n.ImageIndex;
                n.Tag = sc;
            }
        }

        private void FillBasicListView(TreeNode[] FormatCellNode)
        {
            string opentag = ReportTag.StrOpen;
            string closetag = ReportTag.StrClose;
            string septag = ReportTag.DbSeparator;

            string strOpenParen = ReportTag.StrOpenParen.ToString();
            string strSepArg = ReportTag.ParamDelim.ToString();
            string strCloseParen = ReportTag.StrCloseParen.ToString();

            tvFields.Nodes.Clear();
            #region SysTag
            TreeNode Config = tvFields.Nodes.Add(FldMessages.GetString(FldMsg.strConfig));
            Config.ImageIndex = imgConfig;
            Config.SelectedImageIndex = Config.ImageIndex;

            TreeNode ConfigSheet = Config.Nodes.Add(FldMessages.GetString(FldMsg.strConfig));
            ConfigSheet.ImageIndex = imgFullConfig;
            ConfigSheet.SelectedImageIndex = ConfigSheet.ImageIndex;

            foreach (ConfigTagEnum tag in Enum.GetValues(typeof(ConfigTagEnum)))
            {
                string s = ReportTag.ConfigTag(tag);
                TreeNode n = Config.Nodes.Add(s.ToUpper());
                string ParamList = ReportTag.ConfigTagParams(tag);
                string[] sc = { n.Text + ParamList, n.Text };
                n.ImageIndex = imgConfigTag;
                n.SelectedImageIndex = n.ImageIndex;
                n.Tag = sc;
            }

            TreeNode Extras = tvFields.Nodes.Add(FldMessages.GetString(FldMsg.strExtras));
            Extras.ImageIndex = imgExtras;
            Extras.SelectedImageIndex = Extras.ImageIndex;


            foreach (string s in ReportTag.TagTableKeys)
            {
                TreeNode n = Extras.Nodes.Add(s.ToLower());
                string ParamList = string.Empty;
                int ParamCount;
                if (ReportTag.TryGetTagParams(s, out ParamCount))
                {
                    if (ParamCount > 0)
                    {
                        ParamList = strOpenParen.PadRight(strOpenParen.Length + ParamCount - 1, strSepArg[0]) + strCloseParen;
                    }
                }
                string[] sc = { opentag + n.Text + ParamList + closetag, n.Text };
                n.ImageIndex = imgOneExtra;
                n.SelectedImageIndex = n.ImageIndex;
                n.Tag = sc;
                TValueType TagDef;
                if (ReportTag.TryGetTag(s, out TagDef))
                {
                    if (TagDef == TValueType.IF) FillIfListView(n);
                    if (TagDef == TValueType.FormatCell && FormatCellNode != null) FormatCellNode[0] = n;
                }
            }
            #endregion
            #region My Tags

            #region DataField
            TreeNode root = tvFields.Nodes.Add("DataField");
            root.ImageIndex = imgDataset;
            root.SelectedImageIndex = root.ImageIndex;

            if (dt_list != null)
            {
                foreach (DataRow row in dt_list.Rows)
                {
                    TreeNode temp = root.Nodes.Add(row["Code"].ToString());
                    temp.Text = row["Name"].ToString();
                    temp.ImageIndex = imgColumn;
                    temp.SelectedImageIndex = temp.ImageIndex;
                    string[] sc = { dt_list.TableName, row["Code"].ToString(), temp.Text };
                    temp.Tag = sc;
                }
            }
            TreeNode tmp = root.Nodes.Add("#ROWPOS");
            tmp.Text = "#ROWPOS";
            tmp.ImageIndex = imgColumn;
            tmp.SelectedImageIndex = tmp.ImageIndex;
            string[] src = new string[] { dt_list.TableName, tmp.Text, tmp.Text };
            tmp.Tag = src;

            tmp = root.Nodes.Add("#ROWCOUNT");
            tmp.Text = "#ROWCOUNT";
            tmp.ImageIndex = imgColumn;
            tmp.SelectedImageIndex = tmp.ImageIndex;
            src = new string[] { dt_list.TableName, tmp.Text, tmp.Text };
            tmp.Tag = src;

            #endregion DataField

            #region Params
            TreeNode rootParams = tvFields.Nodes.Add("DataParams");
            rootParams.ImageIndex = imgDataset;
            rootParams.SelectedImageIndex = rootParams.ImageIndex;

            if (dt_Params != null)
            {
                foreach (DataRow row in dt_Params.Rows)
                {
                    TreeNode temp = rootParams.Nodes.Add(row["Code"].ToString());
                    temp.Text = row["Name"].ToString();
                    temp.ImageIndex = imgColumn;
                    temp.SelectedImageIndex = temp.ImageIndex;
                    string[] sc = { dt_Params.TableName, row["Code"].ToString(), temp.Text };
                    temp.Tag = sc;
                }

                tmp = rootParams.Nodes.Add("#ROWPOS");
                tmp.Text = "#ROWPOS";
                tmp.ImageIndex = imgColumn;
                tmp.SelectedImageIndex = tmp.ImageIndex;
                src = new string[] { dt_Params.TableName, tmp.Text, tmp.Text };
                tmp.Tag = src;

                tmp = rootParams.Nodes.Add("#ROWCOUNT");
                tmp.Text = "#ROWCOUNT";
                tmp.ImageIndex = imgColumn;
                tmp.SelectedImageIndex = tmp.ImageIndex;
                src = new string[] { dt_Params.TableName, tmp.Text, tmp.Text };
                tmp.Tag = src;
            }
            #endregion
            #region Function

            TreeNode root1 = tvFields.Nodes.Add("Function");
            root1.ImageIndex = imgUserDefined;
            root1.SelectedImageIndex = root1.ImageIndex;
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(Properties.Resources.MyFunc);
            XmlElement element = xmlDoc.DocumentElement;
            foreach (XmlElement node in element.ChildNodes)
            {
                string code = node.GetAttribute("Code");
                tmp = root1.Nodes.Add(code);
                tmp.ImageIndex = imgUserDefinedColumn;
                tmp.SelectedImageIndex = tmp.ImageIndex;
                tmp.Text = node.GetAttribute("Code");
                src = new string[] { "", node.GetAttribute("Value"), tmp.Text };
                tmp.Tag = src;
            }

            #endregion Function

            #region Parameter
            TreeNode rootfilter = tvFields.Nodes.Add("Parameter");
            rootfilter.ImageIndex = imgReportExpList;
            rootfilter.SelectedImageIndex = rootfilter.ImageIndex;
            if (dt_Filter != null)
            {
                foreach (DataRow row in dt_Filter.Rows)
                {
                    tmp = rootfilter.Nodes.Add(row["Code"].ToString());
                    tmp.ImageIndex = imgReportExp;
                    tmp.SelectedImageIndex = tmp.ImageIndex;
                    tmp.Text = row["Name"].ToString();
                    src = new string[] { row["Code"].ToString() };
                    tmp.Tag = src;
                }
            }

            #endregion Parameter

            #endregion My Tags
        }

        private void OpenXsd(string FileName, DataSet ds)
        {
            ds.ReadXmlSchema(FileName);

            TreeNode dataSets = tvFields.Nodes.Add(ds.DataSetName);
            dataSets.ImageIndex = imgDataset;
            dataSets.SelectedImageIndex = dataSets.ImageIndex;


            foreach (DataTable dt in ds.Tables)
            {
                AddTable(dataSets, dt, dt.TableName, imgTable, imgColumn);
            }
        }

        private void AddDsNode(TreeNode dataTableNode, string TableName, string ColumnName, string ColumnCaption, int ColumnImageIndex)
        {
            string stc;
            if (miUseColumnCaptions.Checked) stc = ColumnCaption; else stc = ColumnName;

            TreeNode dataNode = dataTableNode.Nodes.Add(Convert.ToString(stc));
            string[] sa = { TableName, ColumnName, ColumnCaption };
            dataNode.Tag = sa;
            dataNode.ImageIndex = ColumnImageIndex;
            dataNode.SelectedImageIndex = dataNode.ImageIndex;
        }

        private void AddTable(TreeNode ParentNode, DataTable dt, string TableName, int TableImageIndex, int ColumnImageIndex)
        {
            TreeNode dataTableNode = ParentNode.Nodes.Add(Convert.ToString(TableName));
            string[] st = { TableName };
            dataTableNode.Tag = st;
            dataTableNode.ImageIndex = TableImageIndex;
            dataTableNode.SelectedImageIndex = dataTableNode.ImageIndex;

            AddDsNode(dataTableNode, TableName, ReportTag.StrFullDs, ReportTag.StrFullDs, ColumnImageIndex);
            AddDsNode(dataTableNode, TableName, ReportTag.StrFullDsCaptions, ReportTag.StrFullDsCaptions, ColumnImageIndex);
            AddDsNode(dataTableNode, TableName, ReportTag.StrRowCountColumn, ReportTag.StrRowCountColumn, ColumnImageIndex);
            AddDsNode(dataTableNode, TableName, ReportTag.StrRowPosColumn, ReportTag.StrRowPosColumn, ColumnImageIndex);

            foreach (DataColumn dc in dt.Columns)
            {
                AddDsNode(dataTableNode, TableName, dc.ColumnName, dc.Caption, ColumnImageIndex);
            }

        }


        private void OpenConfig(ExcelFile Workbook, DataSet ds, TreeNode[] FormatCellNode)
        {
            TreeNode UserDefined = tvFields.Nodes.Add(FldMessages.GetString(FldMsg.strUserDefined));
            UserDefined.ImageIndex = imgUserDefined;
            UserDefined.SelectedImageIndex = UserDefined.ImageIndex;

            TreeNode ReportVarList = tvFields.Nodes.Add(FldMessages.GetString(FldMsg.strReportVars));
            ReportVarList.ImageIndex = imgReportVarList;
            ReportVarList.SelectedImageIndex = ReportVarList.ImageIndex;

            TreeNode ReportExpList = tvFields.Nodes.Add(FldMessages.GetString(FldMsg.strReportExpressions));
            ReportExpList.ImageIndex = imgReportExpList;
            ReportExpList.SelectedImageIndex = ReportExpList.ImageIndex;

            TreeNode ReportFormatList = tvFields.Nodes.Add(FldMessages.GetString(FldMsg.strFormats));
            ReportFormatList.ImageIndex = imgReportFormatList;
            ReportFormatList.SelectedImageIndex = ReportFormatList.ImageIndex;

            for (int i = FirstConfigRow; i <= Workbook.RowCount; i++)
            {
                string TableName = Convert.ToString(Workbook.GetCellValue(i, ConfigColTableName));
                if (TableName.Length > 0)
                {
                    string SourceName = Workbook.GetCellValue(i, ConfigColSourceName).ToString();
                    DataTable dt = ds.Tables[SourceName];
                    if (dt != null)
                        AddTable(UserDefined, dt, TableName, imgUserDefinedTable, imgUserDefinedColumn);
                }

                string VarName = Convert.ToString(Workbook.GetCellValue(i, ConfigColVarName));
                if (VarName.Length > 0)
                {
                    TreeNode ReportVar = ReportVarList.Nodes.Add(VarName);
                    ReportVar.ImageIndex = imgReportVar;
                    ReportVar.SelectedImageIndex = ReportVar.ImageIndex;
                    string[] st = { VarName };
                    ReportVar.Tag = st;
                }

                string ExpName = Convert.ToString(Workbook.GetCellValue(i, ConfigColExpName));
                if (ExpName.Length > 0)
                {
                    TreeNode ReportExp = ReportExpList.Nodes.Add(ExpName);
                    ReportExp.ImageIndex = imgReportExp;
                    ReportExp.SelectedImageIndex = ReportExp.ImageIndex;
                    string[] st = { ExpName };
                    ReportExp.Tag = st;
                }

                string FormatName = Convert.ToString(Workbook.GetCellValue(i, ConfigColFormatName));
                if (FormatName.Length > 0)
                {
                    TreeNode ReportFormat = ReportFormatList.Nodes.Add(FormatName);
                    ReportFormat.ImageIndex = imgReportFormat;
                    ReportFormat.SelectedImageIndex = ReportFormat.ImageIndex;
                    string[] st = { FormatName };
                    ReportFormat.Tag = st;

                    if (FormatCellNode != null && FormatCellNode[0] != null)
                    {
                        ReportFormat = FormatCellNode[0].Nodes.Add(ReportFormat.Text);
                        ReportFormat.Tag = new string[] {ReportTag.StrOpen+((string[])FormatCellNode[0].Tag)[1]+
                                                        ReportTag.StrOpenParen + FormatName +
                                                        ReportTag.StrCloseParen+ ReportTag.StrClose};
                        ReportFormat.ImageIndex = imgReportFormat;
                        ReportFormat.SelectedImageIndex = ReportFormat.ImageIndex;
                    }
                }
            }

        }

        private void OpenFile()
        {
            try
            {
                if (openXls.FileName.Length == 0) return;
                XlsFile xls = new XlsFile();
                xls.Open(openXls.FileName);
                xls.ActiveSheetByName = ReportTag.StrOpen + ReportTag.StrConfigSheet + ReportTag.StrClose;

                TreeNode[] FormatCellNode = new TreeNode[1];
                FillBasicListView(FormatCellNode);

                string XsdName = Convert.ToString(xls.GetCellValue(RowXsd, ColXsd));
                if (string.IsNullOrEmpty(XsdName)) return;

                using (DataSet ds = new DataSet())
                {
                    {
                        XsdName = Path.Combine(Path.GetDirectoryName(openXls.FileName), XsdName);
                    }
                    OpenXsd(XsdName, ds);

                    OpenConfig(xls, ds, FormatCellNode);
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }

        }

        private void LoadConfig()
        {
            if (!Properties.Settings.Default.AlwaysOnTop) miAlwaysOnTop.PerformClick();
            if (!Properties.Settings.Default.UseColumnCaptions) miUseColumnCaptions.PerformClick();
            miOpacity.Text = "Opacity " + (Properties.Settings.Default.Opacity * 100).ToString() + "%";

            openXls.FileName = Properties.Settings.Default.FileName;
            if (openXls.FileName != null)
            {
                try
                {
                    OpenFile();
                }
                catch (Exception e)
                {
                    openXls.FileName = string.Empty;
                    MessageBox.Show(e.Message);
                    //Clear
                }
            }
            else
                openXls.FileName = String.Empty;
        }


        private object GetConfigSheet(ref MemoryStream xlsStream)
        {
            xlsStream = null;
            try
            {
                XlsFile Xls = new XlsFile();
                Assembly a = Assembly.GetExecutingAssembly();

                /*This will not work on Delphi, since it does not embed the xls files as resources
                using (Stream MemStream = a.GetManifestResourceStream("FlexCelDesigner.Config.xls"))
                {
                    Xls.Open(MemStream);
                }
                */

                System.Resources.ResourceManager rm = new System.Resources.ResourceManager("TVCDesigner.Config", a);
                byte[] WbData = (byte[])rm.GetObject("Config_xls");
                using (Stream MemStream = new MemoryStream(WbData))
                {
                    Xls.Open(MemStream);
                }

                StringBuilder textString = new StringBuilder();
                xlsStream = new MemoryStream();
                //Do NOT DISPOSE THE STREAM. We could, if we called BegingDrag or CopyToClipHere. We can't dispose it till we call those methods.
                {
                    Xls.CopyToClipboard(textString, xlsStream);
                    xlsStream.Position = 0;
                    DataObject data = new DataObject();
                    data.SetData(FlexCelDataFormats.Excel97, xlsStream);
                    data.SetData(DataFormats.Text, textString.ToString());
                    return data;
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            return null;

        }

        private object GetClipboardObj(TreeNode aNode, ref MemoryStream xlsStream)
        {
            if (aNode == null) return null;
            if (aNode.ImageIndex == imgFullConfig)
            {
                return GetConfigSheet(ref xlsStream);
            }
            string[] DragText = (aNode.Tag as string[]);
            if (DragText != null)
            {
                string FinalText = String.Empty;
                switch (aNode.ImageIndex)
                {
                    case imgColumn:
                        string opentag = ReportTag.StrOpen;
                        string closetag = ReportTag.StrClose;
                        string septag = ReportTag.DbSeparator;

                        string df;
                        // if (miUseColumnCaptions.Checked) 
                        df = DragText[1]; //else df = DragText[2];
                        FinalText = opentag + DragText[0] + septag + df + closetag;
                        if ((ModifierKeys & Keys.Alt) != 0) FinalText = DragText[2] + Environment.NewLine + FinalText;
                        break;
                    case imgUserDefinedColumn:

                        opentag = ReportTag.StrOpen;
                        closetag = ReportTag.StrClose;
                        septag = ReportTag.DbSeparator;

                        df = DragText[1];
                        FinalText = opentag + df + closetag;
                        if ((ModifierKeys & Keys.Alt) != 0) FinalText = DragText[2] + Environment.NewLine + FinalText;
                        break;

                    case imgTable:
                    case imgUserDefinedTable:
                        FinalText = ReportTag.RowFull1 + DragText[0] + ReportTag.RowFull2;
                        break;
                    case imgOneExtra:
                    case imgConfigTag:
                    case imgReportFormat:
                        FinalText = DragText[0];
                        break;
                    case imgReportVar:
                    case imgReportExp:
                        FinalText = ReportTag.StrOpen + DragText[0] + ReportTag.StrClose;
                        break;
                }
                if (FinalText.Length > 0) return FinalText;

            }
            return null;

        }

        private void tvFields_ItemDrag(object sender, System.Windows.Forms.ItemDragEventArgs e)
        {
            try
            {
                MemoryStream xlsStream = null;
                try
                {
                    object o = GetClipboardObj((TreeNode)e.Item, ref xlsStream);
                    if (o != null)
                        DoDragDrop(o, DragDropEffects.Copy);
                }
                finally
                {
                    if (xlsStream != null) xlsStream.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void miOpen_Click(object sender, System.EventArgs e)
        {
            if (openXls.ShowDialog() != DialogResult.OK) return;
            OpenFile();
        }

        private void miUseColumnCaptions_Click(object sender, System.EventArgs e)
        {
            miUseColumnCaptions.Checked = !miUseColumnCaptions.Checked;

            ReloadList(tvFields.Nodes);
        }

        private void tvFields_DragDrop(object sender, System.Windows.Forms.DragEventArgs e)
        {
            try
            {
                IDataObject DropFile = e.Data;
                string[] files = DropFile.GetData("FileNameW") as string[];
                if (files == null) files = DropFile.GetData("FileName") as string[];
                if (files == null || files.Length <= 0) return;
                openXls.FileName = files[0];
                OpenFile();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void tvFields_DragOver(object sender, System.Windows.Forms.DragEventArgs e)
        {
            try
            {
                IDataObject DropFile = e.Data;
                if (DropFile.GetDataPresent("FileNameW") || DropFile.GetDataPresent("FileName"))
                    e.Effect = DragDropEffects.Copy;
            }
            catch (Exception)
            {
            }
        }

        private void miCopy_Click(object sender, System.EventArgs e)
        {
            try
            {
                MemoryStream xlsStream = null;
                try
                {
                    object o = GetClipboardObj(tvFields.SelectedNode, ref xlsStream);
                    if (o != null)
                        Clipboard.SetDataObject(o, true);
                }
                finally
                {
                    if (xlsStream != null) xlsStream.Close();
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void File_Click(object sender, System.EventArgs e)
        {

        }

        private void MainForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            Properties.Settings.Default.AlwaysOnTop = miAlwaysOnTop.Checked;
            Properties.Settings.Default.UseColumnCaptions = miUseColumnCaptions.Checked;
            Properties.Settings.Default.Opacity = Opacity;
            Properties.Settings.Default.FileName = openXls.FileName;
            Properties.Settings.Default.Save();
        }

        private void MainForm_Load(object sender, EventArgs e)
        {

        }


    }
}
