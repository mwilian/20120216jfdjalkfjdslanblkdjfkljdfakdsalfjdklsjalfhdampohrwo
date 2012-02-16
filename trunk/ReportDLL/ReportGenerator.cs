using System;
using System.Data;
using FlexCel.Core;
using FlexCel.Report;
using FlexCel.XlsAdapter;
using System.IO;
using System.Globalization;
using System.Text.RegularExpressions;
using FlexCel.Render;
using System.Collections.Generic;
using System.Xml;
using System.Diagnostics;
using System.Collections;
using System.ComponentModel;
using System.Drawing.Printing;
using System.Windows.Forms;

namespace ReportDLL
{
    public class ReportGenerator
    {
        string _FilterPara = "FilterPara";
        FlexCelPrintDocument flexCelPrintDocument1 = new FlexCelPrintDocument();
        public string FilterPara
        {
            get { return _FilterPara; }
            set { _FilterPara = value; }
        }
        string _qdCode = "";
        public static string __connectString = string.Empty;
        public string QdCode
        {
            get { return _qdCode; }
            set { _qdCode = value; }
        }
        string _queryText = "";
        public string _sErr = "";
        public string _fileName = "";
        string _database = "";
        string _pathTemplate = string.Empty;
        string _pathReport = string.Empty;
        ExcelFile _xlsFile = null;
        DataSet _dtSet = null;
        string _name = string.Empty;
        PrintDialog printDialog1 = new PrintDialog();
        public string Name
        {

            get { return _name; }
            set { _name = value; }
        }


        private static void SetParameter(DataTable dt, string code, string value, string type)
        {
            dt.Columns.Add(code);
            if (value == "C")
            {
                if (type == "D")
                {
                    dt.Rows[0][code] = DateTime.Today.ToString("yyyy-MM-dd");
                }
                else if (type == "SDN")
                {
                    dt.Rows[0][code] = DateTime.Today.ToString("yyyyMMdd");
                }
                else if (type == "SPN")
                {
                    dt.Rows[0][code] = DateTime.Today.Year.ToString() + DateTime.Today.Month.ToString("000");
                }
                else
                    dt.Rows[0][code] = value;
            }
            else
                dt.Rows[0][code] = value;
        }
        private void LoadUdfs(ExcelFile Xls)
        {
            //Xls.AddUserDefinedFunction(TUserDefinedFunctionScope.Local, TUserDefinedFunctionLocation.Internal, new TT_XLB_EB());
        }
        private ExcelFile AddData(ExcelFile Xls)
        {
            LoadUdfs(Xls);
            try
            {
                TUnsupportedFormulaList a = Xls.RecalcAndVerify();
            }
            catch (Exception ex)
            {
            }
            //if (clsListValueTT_XLB_EB.Values.Count > 0)
            //{
            //    foreach (TPoint x in clsListValueTT_XLB_EB.Values.Keys)
            //    {
            //        Xls.SetCellValue(x.X, x.Y, clsListValueTT_XLB_EB.Values[x]);
            //    }
            //    clsListValueTT_XLB_EB.Values.Clear();
            //}
            Xls.AllowOverwritingFiles = true;
            //Xls.Save(filename);
            return Xls;
        }
        private ExcelFile Run_templatereport(FlexCelReport flexcelreport)
        {
            flexcelreport.SetUserFunction("DBEGIN", new DBEGIN());
            flexcelreport.SetUserFunction("DEND", new DEND());
            flexcelreport.SetUserFunction("STR2NUM", new STR2NUM());
            flexcelreport.SetUserFunction("NUM2ROMAN", new NUM2ROMAN());
            flexcelreport.SetUserFunction("SUNDATE2DATE", new SUNDATE2DATE());
            flexcelreport.SetUserFunction("PERIOD2STR", new PERIOD2STR());
            flexcelreport.SetUserFunction("NUM2STR", new NUM2STR());
            flexcelreport.SetUserFunction("Read_VN", new Read_VN());
            flexcelreport.SetUserFunction("Read_EN", new Read_EN());

            flexcelreport.SetUserFunction("PH", new PH());
            flexcelreport.SetUserFunction("PE", new PE());
            flexcelreport.SetUserFunction("PA", new PA());

            flexcelreport.SetUserFunction("YA", new YA());
            flexcelreport.SetUserFunction("YH", new YH());
            flexcelreport.SetUserFunction("YE", new YE());
            flexcelreport.SetUserFunction("YK", new YK());
            object misValue = System.Reflection.Missing.Value;

            String filename = _pathTemplate + _qdCode + ".template.xls";
            if (!File.Exists(filename))
            {
                _sErr = "Template Report is not exist!";
                return null;
            }
            ExcelFile result = new XlsFile(filename);
            flexcelreport.Run(result);
            return result;
        }

        #region Userfuntion
        /*  //private class TT_XLB_EB : TUserDefinedFunction
        //{
        //    public TT_XLB_EB() : base("TT_XLB_EB") { }
        //    public override object Evaluate(TUdfEventArgs arguments, object[] parameters)
        //    {
        //        #region Get Parameters
        //        int XF = 0;
        //        TFlxFormulaErrorValue Err = TFlxFormulaErrorValue.ErrValue;
        //        TFormula tmp = (TFormula)arguments.Xls.GetCellValue(arguments.Sheet, arguments.Row, arguments.Col, ref XF);
        //        QueryBuilder.SQLBuilder sqlBuilder = new SQLBuilder(processingMode.Balance);

        //        string formular = tmp.Text;
        //        object[] para = new object[parameters.Length - 1];
        //        TXls3DRange DescCell = new TXls3DRange();
        //        for (int i = 1; i < parameters.Length; i++)
        //        {
        //            if (i == 1)
        //            {
        //                if (!TryGetCellRange(parameters[i], out DescCell, out Err))
        //                {
        //                    break;
        //                }
        //                //formular = formular.Replace("{P}" + i, parameters[i].ToString());
        //            }
        //            else
        //            {
        //                TXls3DRange SourceCell;
        //                if (!TryGetCellRange(parameters[i], out SourceCell, out Err))
        //                    break;
        //                if (SourceCell.IsOneCell)
        //                {
        //                    string value = "";
        //                    if (!TryGetString(arguments.Xls, parameters[i], out value, out Err))
        //                        return Err;
        //                    sqlBuilder.ParaValueList[i - 1] = value;
        //                    //TCellAddress a = new TCellAddress(SourceCell.Top, SourceCell.Left, false, false);

        //                    //formular = formular.Replace(a.CellRef, value);
        //                    formular = formular.Replace("{P}" + (i - 1), value);
        //                }
        //            }
        //        }

        //        #endregion Get Parameters
        //        //formular = formular.Replace("$", "");
        //        Parsing.Formular2SQLBuilder(formular, ref sqlBuilder);
        //        string query = sqlBuilder.BuildSQLEx("");
        //        CommoControl control = new CommoControl();
        //        object result = sqlBuilder.BuildObject(query, ReportGenerator.__connectString);
        //        //formular = sqlBuilder.BuildTTformula();
        //        arguments.Xls.SetComment(DescCell.Top, DescCell.Left, formular);
        //        TPoint x = new TPoint(DescCell.Top, DescCell.Left);
        //        if (result != DBNull.Value)
        //        {
        //            clsListValueTT_XLB_EB.Values.Add(x, result);
        //            //arguments.Xls.SetCellValue(DescCell.Top, DescCell.Left, result.ToString());
        //        }

        //        else
        //        {
        //            result = 0;
        //            clsListValueTT_XLB_EB.Values.Add(x, result);
        //        }
        //        return result;
        //    }

        //}
        */
        class STR2NUM : TFlexCelUserFunction
        {
            public override object Evaluate(object[] parameters)
            {
                if (parameters == null || parameters.Length > 1)
                    throw new ArgumentException("Invalid number of params for user defined function \"MyUserFunction");
                string chuoi = parameters[0].ToString();
                decimal kq = 0;
                try
                {
                    kq = Convert.ToDecimal(chuoi);
                }
                catch (System.Exception ex)
                {

                }
                return kq;
            }

        }
        class NUM2ROMAN : TFlexCelUserFunction
        {
            public override object Evaluate(object[] parameters)
            {
                if (parameters == null || parameters.Length > 1)
                    throw new ArgumentException("Invalid number of params for user defined function \"MyUserFunction");
                int chuoi = Convert.ToInt32(parameters[0]);
                String kq = "";
                try
                {

                    kq = ToRoman(chuoi);
                }
                catch (System.Exception ex)
                {

                }
                return kq;
            }
            public static string ToRoman(int arabic)
            {
                //string result = "";
                //int[] arabic = new int[] { 1000, 900, 500, 400, 100, 90, 50, 40, 10, 9, 5, 4, 1 };
                //string[] roman = new string[] { "M", "CM", "D", "CD", "C", "XC", "L", "XL", "X", "IX", "V", "IV", "I" };
                //int i = 0;
                //while (n >= arabic[i])
                //{
                //    n = n - arabic[i];
                //    result = result + roman[i];

                //    i = i + 1;
                //}
                //return result;
                /* Arabic Roman relation
                * 1000 = M
                * 900 = CM
                * 500 = D
                *400 = CD
                *100 = C
                *90 = XC
                *50 = L
                *40 = XL
                *10 = L
                *9 = IX
                *5 = V
                *4 = IV
                *1 = 1
                */
                string result = "";
                for (int i = 0; i < arabic; i++)
                {
                    while (arabic >= 1000)
                    {//check for thousands place

                        result = result + "M";
                        arabic = arabic - 1000;
                    }
                    while (arabic >= 900)
                    {
                        //check for nine hundred place
                        result = result + "CM";
                        arabic = arabic - 900;
                    }
                    while (arabic >= 500)
                    {
                        //check for five hundred place
                        result = result + "D";
                        arabic = arabic - 500;
                    }
                    while (arabic >= 400)
                    {
                        //check for four hundred place
                        result = result + "CD";
                        arabic = arabic - 400;
                    }
                    while (arabic >= 100)
                    {
                        //check for one hundred place
                        result = result + "C";
                        arabic = arabic - 100;
                    }
                    while (arabic >= 90)
                    {
                        //check for ninety place
                        result = result + "XC";
                        arabic = arabic - 90;
                    }
                    while (arabic >= 50)
                    {
                        //check for fifty place
                        result = result + "L";
                        arabic = arabic - 50;
                    }
                    while (arabic >= 40)
                    {
                        // check for forty place
                        result = result + "XL";
                        arabic = arabic - 40;
                    }

                    while (arabic >= 10)
                    {
                        // check for tenth place
                        result = result + "X";
                        arabic = arabic - 10;
                    }
                    while (arabic >= 9)
                    {
                        //check for nineth place
                        result = result + "IX";
                        arabic = arabic - 9;
                    }
                    while (arabic >= 5)
                    {
                        //check for fifth place
                        result = result + "V";
                        arabic = arabic - 5;
                    }
                    while (arabic >= 4)
                    {
                        //check for fourth place
                        result = result + "IV";
                        arabic = arabic - 4;
                    }
                    while (arabic >= 1)
                    {
                        //check for first place
                        result = result + "I";
                        arabic = arabic - 1;
                    }
                }
                return result;
            }

        }
        class SUNDATE2DATE : TFlexCelUserFunction
        {
            public override object Evaluate(object[] parameters)
            {
                if (parameters == null || parameters.Length > 1)
                    throw new ArgumentException("Invalid number of params for user defined function \"MyUserFunction");
                String chuoi = parameters[0].ToString();
                String kq = "";
                try
                {
                    if (chuoi != "19000101")
                        kq = chuoi.Substring(6, 2) + "/" + chuoi.Substring(4, 2) + "/" + chuoi.Substring(0, 4);
                }
                catch (System.Exception ex)
                {

                }

                return kq;
            }

        }

        class PERIOD2STR : TFlexCelUserFunction
        {
            public override object Evaluate(object[] parameters)
            {
                if (parameters == null || parameters.Length > 1)
                    throw new ArgumentException("Invalid number of params for user defined function \"MyUserFunction");
                String chuoi = parameters[0].ToString();
                String kq = "";
                try
                {

                    kq = chuoi.Substring(5, 2) + "/" + chuoi.Substring(0, 4);

                }
                catch (System.Exception ex)
                {

                }

                return kq;
            }

        }
        class NUM2STR : TFlexCelUserFunction
        {
            public override object Evaluate(object[] parameters)
            {
                String kq = "";
                if (parameters == null || parameters.Length != 2)
                    throw new ArgumentException("Invalid number of params for user defined function \"NUM2STR\"");
                try
                {
                    Decimal para = Convert.ToDecimal(parameters[0]);
                    String chuoi = parameters[1].ToString();
                    string fm = chuoi.Replace("#", "").Replace("0", "");
                    switch (fm)
                    {
                        case ".,":
                        case ",":
                            System.Globalization.CultureInfo a = new System.Globalization.CultureInfo("de-DE");
                            kq = para.ToString(chuoi.Replace(",", "_").Replace(".", ",").Replace("_", "."), a);
                            break;
                        case ",.":
                        case ".":
                            System.Globalization.CultureInfo b = new System.Globalization.CultureInfo("en-US");
                            kq = para.ToString(chuoi, b);
                            break;
                    }
                }
                catch (System.Exception ex)
                {
                    throw new ArgumentException("Invalid number of params for user defined function \"NUM2STR");
                }

                return kq;
            }

        }
        class Read_VN : TFlexCelUserFunction
        {
            public override object Evaluate(object[] parameters)
            {
                if (parameters == null || parameters.Length > 1)
                    throw new ArgumentException("Invalid number of params for user defined function \"MyUserFunction");
                if (parameters[0] == null)
                    return "";
                String chuoi = parameters[0].ToString();
                String kq = "";
                try
                {
                    ReadVN readtv = new ReadVN();
                    kq = readtv.Convert(chuoi.Trim(), '.', " lẻ ");
                    kq[0].ToString().ToUpper();
                }
                catch (System.Exception ex)
                {

                }

                return kq;
            }

        }
        class Read_EN : TFlexCelUserFunction
        {
            public override object Evaluate(object[] parameters)
            {
                if (parameters == null || parameters.Length > 1)
                    throw new ArgumentException("Invalid number of params for user defined function \"MyUserFunction");
                if (parameters[0] == null)
                    return "";
                Double chuoi = Convert.ToDouble(parameters[0]);
                String kq = "";
                try
                {
                    ReadEN readtv = new ReadEN();
                    kq = readtv.NumberToWords(chuoi);
                }
                catch (System.Exception ex)
                {

                }

                return System.Globalization.CultureInfo.CurrentCulture.TextInfo.ToTitleCase(kq);
            }

        }
        class DBEGIN : TFlexCelUserFunction
        {
            public override object Evaluate(object[] parameters)
            {
                if (parameters == null || parameters.Length > 1)
                    throw new ArgumentException("Invalid number of params for user defined function \"MyUserFunction");
                String chuoi = parameters[0].ToString();
                String kq = "";
                try
                {
                    string periodParterm = @"^[0-9]{7}$";
                    string sunDateParterm = @"^[0-9]{8}$";
                    if (parameters[0].GetType() == typeof(DateTime))
                    {
                        DateTime date = Convert.ToDateTime(parameters[0]);
                        kq = "1/" + date.Month + "/" + date.Year;
                    }
                    else if (Regex.IsMatch(chuoi, periodParterm))
                    {
                        int year = Convert.ToInt32(chuoi) / 1000;
                        int month = Convert.ToInt32(chuoi) - year * 1000;
                        kq = "1/" + month + "/" + year;
                    }
                    else if (Regex.IsMatch(chuoi, sunDateParterm))
                    {
                        int year = Convert.ToInt32(chuoi) / 10000;
                        int month = Convert.ToInt32(chuoi) - year * 100;
                        kq = "1/" + month + "/" + year;
                    }
                }
                catch (System.Exception ex)
                {

                }

                return kq;
            }

        }
        class DEND : TFlexCelUserFunction
        {
            public override object Evaluate(object[] parameters)
            {
                if (parameters == null || parameters.Length > 1)
                    throw new ArgumentException("Invalid number of params for user defined function \"MyUserFunction");
                String chuoi = parameters[0].ToString();
                String kq = "";
                try
                {
                    string periodParterm = @"^[0-9]{7}$";
                    string sunDateParterm = @"^[0-9]{8}$";
                    if (parameters[0].GetType() == typeof(DateTime))
                    {
                        DateTime date = Convert.ToDateTime(parameters[0]);
                        kq = DateTime.DaysInMonth(date.Year, date.Month) + "/" + date.Month + "/" + date.Year;
                    }
                    else if (Regex.IsMatch(chuoi, periodParterm))
                    {
                        int year = Convert.ToInt32(chuoi) / 1000;
                        int month = Convert.ToInt32(chuoi) - year * 1000;
                        DateTime date = new DateTime(year, month, 1);
                        kq = DateTime.DaysInMonth(date.Year, date.Month) + "/" + date.Month + "/" + date.Year;
                    }
                    else if (Regex.IsMatch(chuoi, sunDateParterm))
                    {
                        int year = Convert.ToInt32(chuoi) / 10000;
                        int month = Convert.ToInt32(chuoi) - year * 100;
                        DateTime date = new DateTime(year, month, 1);
                        kq = DateTime.DaysInMonth(date.Year, date.Month) + "/" + date.Month + "/" + date.Year;
                    }
                }
                catch (System.Exception ex)
                {

                }

                return kq;
            }

        }

        class PH : TFlexCelUserFunction
        {
            public override object Evaluate(object[] parameters)
            {
                if (parameters == null || parameters.Length > 1)
                    throw new ArgumentException("Invalid number of params for user defined function \"MyUserFunction");
                String chuoi = parameters[0].ToString();
                String kq = "";
                try
                {
                    string periodParterm = @"^[0-9]{7}$";
                    if (Regex.IsMatch(chuoi, periodParterm))
                    {
                        int year = Convert.ToInt32(chuoi) / 1000;
                        int month = Convert.ToInt32(chuoi) - year * 1000;
                        month--;
                        if (month == 0)
                        {
                            year--;
                            month = 12;
                        }
                        kq = year.ToString() + month.ToString("000");
                    }
                }
                catch (System.Exception ex)
                {

                }

                return kq;
            }

        }
        class PA : TFlexCelUserFunction
        {
            public override object Evaluate(object[] parameters)
            {
                if (parameters == null || parameters.Length > 1)
                    throw new ArgumentException("Invalid number of params for user defined function \"MyUserFunction");
                String chuoi = parameters[0].ToString();
                String kq = "";
                try
                {
                    string periodParterm = @"^[0-9]{7}$";
                    if (Regex.IsMatch(chuoi, periodParterm))
                    {
                        int year = Convert.ToInt32(chuoi) / 1000;
                        int month = Convert.ToInt32(chuoi) - year * 1000;
                        //month--;
                        //if (month == 0)
                        //{
                        //    year--;
                        //    month = 12;
                        //}
                        kq = year.ToString() + month.ToString("000");
                    }
                }
                catch (System.Exception ex)
                {

                }

                return kq;
            }

        }
        class PE : TFlexCelUserFunction
        {
            public override object Evaluate(object[] parameters)
            {
                if (parameters == null || parameters.Length > 1)
                    throw new ArgumentException("Invalid number of params for user defined function \"MyUserFunction");
                String chuoi = parameters[0].ToString();
                String kq = "";
                try
                {
                    string periodParterm = @"^[0-9]{7}$";
                    if (Regex.IsMatch(chuoi, periodParterm))
                    {
                        int year = Convert.ToInt32(chuoi) / 1000;
                        int month = Convert.ToInt32(chuoi) - year * 1000;
                        month = 12;
                        kq = year.ToString() + month.ToString("000");
                    }
                }
                catch (System.Exception ex)
                {

                }

                return kq;
            }

        }

        class YA : TFlexCelUserFunction
        {
            public override object Evaluate(object[] parameters)
            {
                if (parameters == null || parameters.Length > 1)
                    throw new ArgumentException("Invalid number of params for user defined function \"MyUserFunction");
                String chuoi = parameters[0].ToString();
                String kq = "";
                try
                {
                    string periodParterm = @"^[0-9]{7}$";
                    if (Regex.IsMatch(chuoi, periodParterm))
                    {
                        int year = Convert.ToInt32(chuoi) / 1000;
                        int month = Convert.ToInt32(chuoi) - year * 1000;

                        year--;
                        month = 1;

                        kq = year.ToString() + month.ToString("000");
                    }
                }
                catch (System.Exception ex)
                {

                }

                return kq;
            }

        }
        class YE : TFlexCelUserFunction
        {
            public override object Evaluate(object[] parameters)
            {
                if (parameters == null || parameters.Length > 1)
                    throw new ArgumentException("Invalid number of params for user defined function \"MyUserFunction");
                String chuoi = parameters[0].ToString();
                String kq = "";
                try
                {
                    string periodParterm = @"^[0-9]{7}$";
                    if (Regex.IsMatch(chuoi, periodParterm))
                    {
                        int year = Convert.ToInt32(chuoi) / 1000;
                        int month = Convert.ToInt32(chuoi) - year * 1000;
                        year--;
                        month = 1;
                        kq = year.ToString() + month.ToString("000");
                    }
                }
                catch (System.Exception ex)
                {

                }

                return kq;
            }

        }
        class YH : TFlexCelUserFunction
        {
            public override object Evaluate(object[] parameters)
            {
                if (parameters == null || parameters.Length > 1)
                    throw new ArgumentException("Invalid number of params for user defined function \"MyUserFunction");
                String chuoi = parameters[0].ToString();
                String kq = "";
                try
                {
                    string periodParterm = @"^[0-9]{7}$";
                    if (Regex.IsMatch(chuoi, periodParterm))
                    {
                        int year = Convert.ToInt32(chuoi) / 1000;
                        int month = Convert.ToInt32(chuoi) - year * 1000;
                        month = 12;
                        year--;
                        kq = year.ToString() + month.ToString("000");
                    }
                }
                catch (System.Exception ex)
                {

                }

                return kq;
            }

        }
        class YK : TFlexCelUserFunction
        {
            public override object Evaluate(object[] parameters)
            {
                if (parameters == null || parameters.Length > 1)
                    throw new ArgumentException("Invalid number of params for user defined function \"MyUserFunction");
                String chuoi = parameters[0].ToString();
                String kq = "";
                try
                {
                    string periodParterm = @"^[0-9]{7}$";
                    if (Regex.IsMatch(chuoi, periodParterm))
                    {
                        int year = Convert.ToInt32(chuoi) / 1000;
                        int month = Convert.ToInt32(chuoi) - year * 1000;
                        month = 12;
                        kq = year.ToString() + month.ToString("000");
                    }
                }
                catch (System.Exception ex)
                {

                }

                return kq;
            }

        }
        #endregion


        #region Method
        private void GenTemplate(string path_template)
        {
            if (!File.Exists(path_template))
            {
                if (File.Exists(_pathTemplate + "-.template.xls"))
                {
                    XlsFile xlsTemp = new XlsFile(_pathTemplate + "-.template.xls");
                    xlsTemp.SetCellValue(xlsTemp.GetSheetIndex("<#Config>"), 10, 2, _qdCode, 0);
                    xlsTemp.SetCellValue(xlsTemp.GetSheetIndex("<#Config>"), 17, 2, "FilterPara", 0);

                    xlsTemp.Save(_pathTemplate + _qdCode + ".template.xls");
                }
                else
                {
                    byte[] file = ReportDLL.Properties.Resources.__template;
                    Stream afile = new FileStream(_pathTemplate + "-.template.xls", FileMode.CreateNew, FileAccess.Write);
                    afile.Write(file, 0, file.Length);
                    afile.Close();

                }
            }
        }
        /// <summary>
        /// Khoi tao mau bao cao
        /// </summary>
        /// <param name="xmlString">Chuoi dinh dang xml</param>
        /// <param name="qdCode">Ma cua bao cao</param>
        /// <param name="pathTemplate">Duong dan thu muc chua mau bao cao</param>
        /// <param name="pathReport">Duong dan thu muc chua bao cao dau ra</param>
        public ReportGenerator(string xmlString, string qdCode, string pathTemplate, string pathReport)
        {
            //_sqlBuilder = sqlBuilder;
            DataSet ds = new DataSet("DataSet");
            try
            {
                DataTable dt = GetDataTable(xmlString, "Record");
                dt.TableName = qdCode;
                DataTable dtFilter = new DataTable("FilterPara");
                ds.Tables.Add(dt);
                ds.Tables.Add(dtFilter);
                _dtSet = ds;
            }
            catch (Exception ex) { _sErr = ex.Message; _dtSet = null; }
            _qdCode = qdCode;
            _pathReport = pathReport;
            _pathTemplate = pathTemplate;
        }
        private bool LoadPreferences()
        {
            try
            {
                flexCelPrintDocument1.Workbook = _xlsFile;
                ExcelFile Xls = flexCelPrintDocument1.Workbook;
                Xls.PrintHeadings = false;
                Xls.PrintGridLines = false;
                //Xls.PrintPaperSize = TPaperSize.

                flexCelPrintDocument1.DefaultPageSettings.PaperSize = new PaperSize(Xls.PrintPaperDimensions.PaperName, Convert.ToInt32(Xls.PrintPaperDimensions.Width), Convert.ToInt32(Xls.PrintPaperDimensions.Height));
                //flexCelPrintDocument1.PrintPa
                flexCelPrintDocument1.DefaultPageSettings.Landscape = (Xls.PrintOptions & TPrintOptions.Orientation) == 0;
                return true;
            }
            catch (Exception ex) { throw ex; }
        }
        private bool DoSetup(PrintDocument doc)
        {
            printDialog1.Document = doc;
            printDialog1.PrinterSettings.DefaultPageSettings.PaperSize = doc.DefaultPageSettings.PaperSize;
            //printDialog1.PrinterSettings.
            bool Result = printDialog1.ShowDialog() == DialogResult.OK;
            //printDialog1.PrinterSettings.PaperSizes = doc.PrinterSettings.pga
            //Landscape.Checked = flexCelPrintDocument1.DefaultPageSettings.Landscape;
            return Result;
        }
        public void PrintBin()
        {
            try
            {
                CreateReport();
                LoadPreferences();
                //if (!LoadPreferences()) return;
                //if (!DoSetup(flexCelPrintDocument1)) return;
                flexCelPrintDocument1.Print();
            }
            catch (Exception ex)
            {
                _sErr = ex.Message;
            }
        }
        public void PrintSetup()
        {
            try
            {
                CreateReport();
                if (!LoadPreferences()) return;
                if (!DoSetup(flexCelPrintDocument1)) return;
                //flexCelPrintDocument1.Print();
            }
            catch (Exception ex)
            {
                _sErr = ex.Message;
            }
        }
        private DataTable GetDataTable(string xmlString, string code)
        {
            DataTable dt = new DataTable(code);
            XmlDocument root = new XmlDocument();
            root.LoadXml(xmlString);
            XmlElement doc = root.DocumentElement;
            int index = 0;
            foreach (XmlElement ele in doc.ChildNodes)
            {
                if (ele.Name == code)
                {
                    dt.Rows.Add(dt.NewRow());
                    foreach (XmlElement field in ele.ChildNodes)
                    {
                        if (!dt.Columns.Contains(field.Name))
                            dt.Columns.Add(field.Name);
                        dt.Rows[index][field.Name] = field.InnerText;
                    }
                    index++;
                }
            }
            return dt;
        }
        //public ReportGenerator(DataSet dtSet, string qdCode, string database, string connectString, string pathTemplate, string pathReport)
        //{
        //    //_sqlBuilder = sqlBuilder;
        //    _qdCode = qdCode;
        //    _database = database;
        //    _dtSet = dtSet;
        //    __connectString = connectString;
        //    _pathReport = pathReport;
        //    _pathTemplate = pathTemplate;
        //}

        /// <summary>
        /// Xuat ra bao cao
        /// </summary>
        /// <param name="path">Duong dan thu muc dau ra</param>
        /// <returns>ExcelFile de cho vao TVCPreviewer</returns>
        public ExcelFile ExportExcel(string path)
        {
            _xlsFile = CreateReport();
            return _xlsFile;
        }
        public string ExportExcelToPath(string path)
        {
            String filename = path + "\\" + _qdCode + ".xls";
            _xlsFile = CreateReport();
            _xlsFile.Save(filename);
            return filename;
        }
        public string ExportExcelToFile(string path, string filename)
        {
            //String filename = path + "\\" + _qdCode + ".xls";
            string file = path + filename;
            _xlsFile = CreateReport();
            _xlsFile.Save(file);
            return file;
        }
        private ExcelFile CreateReport()
        {
            if (_xlsFile != null)
                return _xlsFile;
            FlexCelReport flexcelreport = new FlexCelReport();
            GenTemplate(_pathTemplate + _qdCode + ".template.xls");
            if (_dtSet != null)
            {

                //flexcelreport.UserTable += new UserTableEventHandler(flexcelreport_UserTable);
                flexcelreport.AddTable(_dtSet);
                AddReportVariable(flexcelreport);

                ExcelFile rs = Run_templatereport(flexcelreport);
                rs = AddData(rs);
                _xlsFile = rs;
                if (File.Exists(_pathReport + _qdCode + ".xls"))
                    File.Delete(_pathReport + _qdCode + ".xls");
                rs.Save(_pathReport + _qdCode + ".xls");
                return rs;
            }
            return null;
        }
        public void DesignReport()
        {
            try
            {

                //      File.Delete(saveFileDialog1.FileName);
                string path_template = _pathTemplate + _qdCode + ".template.xls";
                string currentPath = System.Windows.Forms.Application.StartupPath + "\\";
                if (!File.Exists(path_template))
                {
                    XlsFile xlsTemp = new XlsFile(currentPath + "report.template.xls");
                    xlsTemp.SetCellValue(xlsTemp.GetSheetIndex("<#Config>"), 10, 2, _qdCode, 0);
                    xlsTemp.SetCellValue(xlsTemp.GetSheetIndex("<#Config>"), 17, 2, "FilterPara", 0);

                    xlsTemp.Save(path_template);
                }

                //flexCelReport1.AddTable(temp);
                //AutoRun(temp.Tables[0], flag_filter);
                //Set_TT_XLB_EB(path_template);
                Process.Start(path_template);
                //}
                //else
                //{
                //}



                //ListFieldData a = new ListFieldData();
                #region Parameter
                DataTable dt_filter = new DataTable();
                dt_filter.TableName = "parameter";
                dt_filter.Columns.Add("Name");
                dt_filter.Columns.Add("Code");

                if (_dtSet.Tables["FilterPara"].Rows.Count > 0)
                {
                    for (int i = 0; i < _dtSet.Tables["FilterPara"].Rows.Count; i++)
                    {
                        dt_filter.Rows.Add(new string[] { _dtSet.Tables["FilterPara"].Rows[i]["Description"].ToString() + "_From", "parameter." + _dtSet.Tables["FilterPara"].Rows[i]["Code"].ToString() + "_From" });
                        dt_filter.Rows.Add(new string[] { _dtSet.Tables["FilterPara"].Rows[i]["Description"].ToString() + "_To", "parameter." + _dtSet.Tables["FilterPara"].Rows[i]["Code"].ToString() + "_To" });
                    }
                    //a.dt_Filter = dt_filter;
                }

                DataTable dt_param = new DataTable();
                DataColumn[] cols = new DataColumn[] { new DataColumn("Code")
                    , new DataColumn("Name")};
                dt_param.Columns.AddRange(cols);
                dt_param.TableName = "params";

                dt_param.Rows.Add("Code", "Code");
                dt_param.Rows.Add("Description", "Description");
                dt_param.Rows.Add("ValueFrom", "ValueFrom");
                dt_param.Rows.Add("ValueTo", "ValueTo");
                dt_param.Rows.Add("IsNot", "IsNot");
                dt_param.Rows.Add("Operate", "Operate");
                #endregion Parameter
                #region Field
                DataTable dt_list = new DataTable();
                if (_dtSet.Tables[_qdCode].Rows.Count > 0)
                {
                    //CommoControl commo = new CommoControl();
                    //string connnectString = commo.CreateConnectString(Properties.Settings.Default.Server
                    //         , Properties.Settings.Default.User
                    //         , Properties.Settings.Default.Pass
                    //         , Properties.Settings.Default.DBName);
                    //DataTable rs = _sqlBuilder.BuildDataTableStruct(txt_sql.Text, _strConnectDes);

                    ////a.THEME = this.THEME;
                    dt_list.TableName = "data";
                    dt_list.Columns.Add("Name");
                    dt_list.Columns.Add("Code");

                    foreach (DataColumn colum in _dtSet.Tables[_qdCode].Columns)
                    {
                        string desc = colum.ColumnName;
                        //foreach (Node node in _sqlBuilder.SelectedNodes)
                        //{
                        //    if (node.MyCode == colum.ColumnName)
                        //    {
                        //        desc = node.Description;
                        //        break;
                        //    }
                        //}
                        dt_list.Rows.Add(new string[] { desc, colum.ColumnName });
                    }


                    //a.dt_list = dt;


                }
                else
                {

                    //a.THEME = this.ThemeName;

                    dt_list.Columns.Add("Name");
                    dt_list.Columns.Add("Code");


                    ArrayList arr = GetFieldName();
                    if (arr.Count > 0)
                    {

                        for (int i = 0; i < arr.Count; i++)
                        {
                            dt_list.Rows.Add(new string[] { arr[i].ToString(), arr[i].ToString() });
                        }
                        //a.dt_list = dt;


                    }

                }
                #endregion Field
                TVCDesigner.MainForm frm = new TVCDesigner.MainForm(dt_list, dt_filter, dt_param);

                //frm.BringToFront();

                frm.Show();
                //this.MinimizeBox = true;
            }
            catch (Exception ex)
            {
                _sErr = ex.Message;
            }
        }

        private ArrayList GetFieldName()
        {
            ArrayList arr = new ArrayList();
            //CommoControl commo = new CommoControl();
            //string connnectString = commo.CreateConnectString(Properties.Settings.Default.Server
            //            , Properties.Settings.Default.User
            //            , Properties.Settings.Default.Pass
            //            , Properties.Settings.Default.DBName);
            //DataTable dt = _sqlBuilder.BuildDataTable(txt_sql.Text, _strConnectDes);

            foreach (DataColumn colum in _dtSet.Tables[_qdCode].Columns)
            {
                arr.Add(colum.ColumnName);
            }

            return arr;
        }
        private void AddReportVariable(FlexCelReport flexcelreport)
        {
            flexcelreport.SetValue("Date", DateTime.Now.ToShortDateString());
            flexcelreport.SetValue("QDName", _name);
            flexcelreport.SetValue("QDCode", _qdCode);
            flexcelreport.SetValue("DB", _database);
        }
        public MemoryStream ExportPDF(string path)
        {
            MemoryStream ms = new MemoryStream();
            //if (_xlsFile == null)
            _xlsFile = CreateReport();

            try
            {
                using (FlexCelPdfExport pdf = new FlexCelPdfExport())
                {
                    pdf.Workbook = _xlsFile;

                    pdf.BeginExport(ms);
                    pdf.ExportAllVisibleSheets(false, "test");
                    pdf.EndExport();
                }
                return ms;
            }
            catch
            {
                return null;
            }

        }
        public MemoryStream ExportPDF(ExcelFile xlsFile)
        {
            MemoryStream ms = new MemoryStream();
            try
            {
                using (FlexCelPdfExport pdf = new FlexCelPdfExport())
                {
                    pdf.Workbook = xlsFile;

                    pdf.BeginExport(ms);
                    pdf.ExportAllVisibleSheets(false, "test");
                    pdf.EndExport();
                }
                return ms;
            }
            catch
            {
                return null;
            }

        }
        public string ExportPDFToPath(string path)
        {
            String filename = path + "\\" + _qdCode + ".pdf";
            //FileStream file = new FileStream(filename, FileMode.OpenOrCreate, FileAccess.ReadWrite);
            //if (_xlsFile == null)
            _xlsFile = CreateReport();

            try
            {
                using (FlexCelPdfExport pdf = new FlexCelPdfExport())
                {
                    pdf.Workbook = _xlsFile;
                    using (MemoryStream ms = new MemoryStream())
                    {
                        pdf.BeginExport(ms);
                        pdf.ExportAllVisibleSheets(false, "test");
                        pdf.EndExport();

                        pdf.Export(filename);
                    }
                }
                return filename;
            }
            catch
            {
                return "";
            }

        }
        public string ExportHTMLToPath(string path)
        {
            String filename = path + _qdCode + ".htm";
            //TextWriter file = new StringWriter(); ;
            if (_xlsFile == null)
                _xlsFile = CreateReport();
            try
            {
                using (FlexCelHtmlExport html = new FlexCelHtmlExport())
                {
                    html.Workbook = _xlsFile;
                    //html.HtmlFileFormat = THtmlFileFormat.MHtml;
                    html.AllowOverwritingFiles = true;
                    html.SavedImagesFormat = THtmlImageFormat.Png;
                    //html.HtmlVersion = THtmlVersion.XHTML_10;
                    //if (File.Exists(filename))
                    //    File.Delete(filename);
                    //string pathx = Path.GetDirectoryName(filename);
                    //string name = Path.GetFileNameWithoutExtension(filename);
                    //string ext = ".png";
                    //string fileimage = pathx + "\\" + name + "_image1" + ext;
                    //if (File.Exists(fileimage))
                    //    File.Delete(fileimage);
                    //fileimage = pathx + "\\" + name + "_image2" + ext;
                    //if (File.Exists(fileimage))
                    //    File.Delete(fileimage);

                    html.Export(filename, "images", "css\\" + _qdCode + ".css");

                }
                return filename;
            }
            catch
            {
                return "";
            }
        }
        public string ExportHTMLToFile(string path, string filename)
        {
            string filehtml = path + filename;
            //if (_xlsFile == null)
            _xlsFile = CreateReport();
            try
            {
                using (FlexCelHtmlExport html = new FlexCelHtmlExport())
                {
                    html.Workbook = _xlsFile;
                    html.Workbook = _xlsFile;
                    //html.HtmlFileFormat = THtmlFileFormat.MHtml;
                    html.AllowOverwritingFiles = true;
                    html.SavedImagesFormat = THtmlImageFormat.Png;
                    //html.HtmlVersion = THtmlVersion.XHTML_10;
                    html.Export(filehtml, "images", "css\\" + _qdCode + ".css");
                }
                return filehtml;
            }
            catch
            {
                return "";
            }
        }
        public TextWriter ExportHTML(string path)
        {
            String filename = path + _qdCode + ".html";
            TextWriter file = new StringWriter(); ;
            //if (_xlsFile == null)
            _xlsFile = CreateReport();
            try
            {
                using (FlexCelHtmlExport html = new FlexCelHtmlExport())
                {
                    html.Workbook = _xlsFile;
                    using (MemoryStream ms = new MemoryStream())
                    {
                        html.Workbook = _xlsFile;
                        html.Workbook = _xlsFile;
                        //html.HtmlFileFormat = THtmlFileFormat.MHtml;
                        html.AllowOverwritingFiles = true;
                        html.SavedImagesFormat = THtmlImageFormat.Png;
                        //html.HtmlVersion = THtmlVersion.XHTML_10;                      

                        html.Export(file, filename, null);

                    }
                }
                return file;
            }
            catch
            {
                return null;
            }

        }
        public TextWriter ExportHTML(string path, ExcelFile xlsFile)
        {
            //String filename = path + _pathTemplate + "\\" + _database + "\\" + _qdCode + ".html";
            TextWriter file = new StringWriter(); ;
            try
            {
                using (FlexCelHtmlExport html = new FlexCelHtmlExport())
                {
                    html.Workbook = xlsFile;
                    using (MemoryStream ms = new MemoryStream())
                    {
                        //if (File.Exists(filename))
                        //    File.Delete(filename);
                        html.Export(file, "", null);
                        //file.ToString();
                    }
                }
                return file;
            }
            catch
            {
                return null;
            }

        }
        /*public string ExportSWF(string path, string qdID)
        {
            String filename = path + _pathTemplate + "\\" + _database + "\\" + _qdCode + ".swf";
            //String filePDF = path + _pathTemplate + "\\" + _sqlBuilder.Database + "\\" + _qdCode + ".pdf";
            try
            {
                string filePDF = ExportPDFToFile(path, qdID);
                System.Diagnostics.Process r = new System.Diagnostics.Process();
                r.StartInfo.UseShellExecute = false;
                r.StartInfo.RedirectStandardOutput = true;
                r.StartInfo.CreateNoWindow = true;
                r.StartInfo.RedirectStandardError = true;
                r.StartInfo.WorkingDirectory = HttpContext.Current.Server.MapPath("~/");
                r.StartInfo.FileName = HttpContext.Current.Server.MapPath("~/PDF2SWF/PDF2SWF.exe");
                r.StartInfo.Arguments = filePDF + " -o " + filename + " -T 9 -f ";
                r.Start();
                r.WaitForExit();
                r.Close();
                //file.Close();
                return filename;
            }
            catch
            {

                return "";
            }
        }*/
        #endregion Method

        #region ReadNumber

        class ReadVN
        {
            private string[] strSo = { "không", "một", "hai", "ba", "bốn", "năm", "sáu", "bảy", "tám", "chín" };
            private string[] strDonViNho = { "linh", "lăm", "mười", "mươi", "mốt", "trăm" };
            private string[] strDonViLon = { "", "ngàn", "triệu", "tỷ" };
            private string[] strMainGroup;
            private string[] strSubGroup;
            private string Len1(string strA)
            {
                return strSo[int.Parse(strA)];
            }
            private string Len2(string strA)
            {
                if (strA.Substring(0, 1) == "0")
                {
                    return strDonViNho[0] + " " + Len1(strA.Substring(1, 1));
                }
                else if (strA.Substring(0, 1) == "1")
                {
                    if (strA.Substring(1, 1) == "5")
                    {
                        return strDonViNho[2] + " " + strDonViNho[1];
                    }
                    else if (strA.Substring(1, 1) == "0")
                    {
                        return strDonViNho[2];
                    }
                    else
                    {
                        return strDonViNho[2] + " " + Len1(strA.Substring(1, 1));
                    }
                }
                else
                {
                    if (strA.Substring(1, 1) == "5")
                    {
                        return Len1(strA.Substring(0, 1)) + " " + strDonViNho[3] + " " + strDonViNho[1];
                    }
                    else if (strA.Substring(1, 1) == "0")
                    {
                        return Len1(strA.Substring(0, 1)) + " " + strDonViNho[3];
                    }
                    else if (strA.Substring(1, 1) == "1")
                    {
                        return Len1(strA.Substring(0, 1)) + " " + strDonViNho[3] + " " + strDonViNho[4];
                    }
                    else
                    {
                        return Len1(strA.Substring(0, 1)) + " " + strDonViNho[3] + " " + Len1(strA.Substring(1, 1));
                    }
                }
            }
            private string Len3(string strA)
            {
                if ((strA.Substring(0, 3) == "000"))
                {
                    return null;
                }
                else if ((strA.Substring(1, 2) == "00"))
                {
                    return Len1(strA.Substring(0, 1)) + " " + strDonViNho[5];
                }
                else
                {
                    return Len1(strA.Substring(0, 1)) + " " + strDonViNho[5] + " " + Len2(strA.Substring(1, strA.Length - 1));
                }
            }
            /////////////////////
            private string FullLen(string strSend)
            {
                bool boKTNull = false;
                string strKQ = "";
                string strA = strSend.Trim();
                int iIndex = strA.Length - 9;
                int iPreIndex = 0;

                if (strSend.Trim() == "")
                {
                    return Len1("0");
                }
                //tra ve khong neu la khong
                for (int i = 0; i < strA.Length; i++)
                {
                    if (strA.Substring(i, 1) != "0")
                    {
                        break;
                    }
                    else if (i == strA.Length - 1)
                    {
                        return strSo[0];
                    }
                }
                int k = 0;
                while (strSend.Trim().Substring(k++, 1) == "0")
                {
                    strA = strA.Remove(0, 1);
                }
                //
                if (strA.Length < 9)
                {
                    iPreIndex = strA.Length;
                }
                //
                if ((strA.Length % 9) != 0)
                {
                    strMainGroup = new string[strA.Length / 9 + 1];
                }
                else
                {
                    strMainGroup = new string[strA.Length / 9];
                }
                //nguoc
                for (int i = strMainGroup.Length - 1; i >= 0; i--)
                {
                    if (iIndex >= 0)
                    {
                        iPreIndex = iIndex;
                        strMainGroup[i] = strA.Substring(iIndex, 9);
                        iIndex -= 9;
                    }
                    else
                    {
                        strMainGroup[i] = strA.Substring(0, iPreIndex);
                    }

                }
                /////////////////////////////////
                //tach moi maingroup thanh nhieu subgroup
                //xuoi
                for (int j = 0; j < strMainGroup.Length; j++)
                {
                    //gan lai gia tri
                    iIndex = strMainGroup[j].Length - 3;
                    if (strMainGroup[j].Length < 3)
                    {
                        iPreIndex = strMainGroup[j].Length;
                    }
                    ///
                    if ((strMainGroup[j].Length % 3) != 0)
                    {
                        strSubGroup = new string[strMainGroup[j].Length / 3 + 1];
                    }
                    else
                    {
                        strSubGroup = new string[strMainGroup[j].Length / 3];
                    }
                    for (int i = strSubGroup.Length - 1; i >= 0; i--)
                    {
                        if (iIndex >= 0)
                        {
                            iPreIndex = iIndex;
                            strSubGroup[i] = strMainGroup[j].Substring(iIndex, 3);
                            iIndex -= 3;
                        }
                        else
                        {
                            strSubGroup[i] = strMainGroup[j].Substring(0, iPreIndex);
                        }
                    }
                    //duyet subgroup de lay string
                    for (int i = 0; i < strSubGroup.Length; i++)
                    {
                        boKTNull = false;//phai de o day
                        if ((j == strMainGroup.Length - 1) && (i == strSubGroup.Length - 1))
                        {
                            if (strSubGroup[i].Length < 3)
                            {
                                if (strSubGroup[i].Length == 1)
                                {
                                    strKQ += Len1(strSubGroup[i]);
                                }
                                else
                                {
                                    strKQ += Len2(strSubGroup[i]);
                                }
                            }
                            else
                            {
                                strKQ += Len3(strSubGroup[i]);
                            }
                        }
                        else
                        {
                            if (strSubGroup[i].Length < 3)
                            {
                                if (strSubGroup[i].Length == 1)
                                {
                                    strKQ += Len1(strSubGroup[i]) + " ";
                                }
                                else
                                {
                                    strKQ += Len2(strSubGroup[i]) + " ";
                                }
                            }
                            else
                            {
                                if (Len3(strSubGroup[i]) == null)
                                {
                                    boKTNull = true;
                                }
                                else
                                {
                                    strKQ += Len3(strSubGroup[i]) + " ";
                                }
                            }
                        }
                        //dung
                        if (!boKTNull)
                        {
                            if (strSubGroup.Length - 1 - i != 0)
                            {
                                strKQ += strDonViLon[strSubGroup.Length - 1 - i] + " ";
                            }
                            else
                            {
                                strKQ += strDonViLon[strSubGroup.Length - 1 - i] + " ";
                            }

                        }
                    }
                    //dung
                    if (j != strMainGroup.Length - 1)
                    {
                        if (!boKTNull)
                        {
                            strKQ = strKQ.Substring(0, strKQ.Length - 1) + strDonViLon[3] + " ";
                        }
                        else
                        {
                            strKQ = strKQ.Substring(0, strKQ.Length - 1) + " " + strDonViLon[3] + " ";
                        }
                    }
                }
                //xoa ky tu trang
                strKQ = strKQ.Trim();
                //xoa dau , neu co
                if (strKQ.Substring(strKQ.Length - 1, 1) == ".")
                {
                    strKQ = strKQ.Remove(strKQ.Length - 1, 1);
                }
                return strKQ;

                ////////////////////////////////////


            }
            public string Convert(string strSend, char charInSeparator, string strOutSeparator)
            {
                if (strOutSeparator == "")
                {
                    return "Lỗi dấu phân cách đầu ra rỗng";
                }
                if (strSend == "")
                {
                    return Len1("0");
                }

                string[] strTmp = new string[2];
                try
                {

                    strTmp = strSend.Split(charInSeparator);
                    string strTmpRight = strTmp[1];
                    for (int i = strTmpRight.Length - 1; i >= 0; i--)
                    {
                        if (strTmpRight.Substring(i, 1) == "0")
                        {
                            strTmpRight = strTmpRight.Remove(i, 1);
                        }
                        else
                        {
                            break;
                        }
                    }
                    if (strTmpRight != "")
                    {
                        string strRight = "";
                        for (int i = 0; i < strTmpRight.Length; i++)
                        {
                            strRight += Len1(strTmpRight.Substring(i, 1)) + " ";
                        }


                        return FullLen(strTmp[0]) + " " + strOutSeparator + " " + strRight.TrimEnd();
                    }
                    else
                    {
                        return FullLen(strTmp[0]);
                    }
                }
                catch
                {
                    return FullLen(strTmp[0]);
                }

            }

        }
        class ReadEN
        {
            // Single-digit and small number names
            private string[] _smallNumbers = new string[] { "Zero", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine", "Ten", "Eleven", "Twelve", "Thirteen", "Fourteen", "Fifteen", "Sixteen", "Seventeen", "Eighteen", "Nineteen" };

            // Tens number names from twenty upwards
            private string[] _tens = new string[] { "", "", "Twenty", "Thirty", "Forty", "Fifty", "Sixty", "Seventy", "Eighty", "Ninety" };

            // Scale number names for use during recombination
            private string[] _scaleNumbers = new string[] { "", "Thousand", "Million", "Billion" };



            // Converts an integer value into English words
            public string NumberToWords(Double number)
            {
                // Zero rule
                if (number == 0)
                    return _smallNumbers[0];

                // Array to hold four three-digit groups
                int[] digitGroups = new int[4];

                // Ensure a positive number to extract from
                int positive = Math.Abs(Convert.ToInt32(number));

                // Extract the three-digit groups
                for (int i = 0; i < 4; i++)
                {
                    digitGroups[i] = positive % 1000;
                    positive /= 1000;
                }

                // Convert each three-digit group to words
                string[] groupText = new string[4];

                for (int i = 0; i < 4; i++)
                    groupText[i] = ThreeDigitGroupToWords(digitGroups[i]);

                // Recombine the three-digit groups
                string combined = groupText[0];
                bool appendAnd;

                // Determine whether an 'and' is needed
                appendAnd = (digitGroups[0] > 0) && (digitGroups[0] < 100);

                // Process the remaining groups in turn, smallest to largest
                for (int i = 1; i < 4; i++)
                {
                    // Only add non-zero items
                    if (digitGroups[i] != 0)
                    {
                        // Build the string to add as a prefix
                        string prefix = groupText[i] + " " + _scaleNumbers[i];

                        if (combined.Length != 0)
                            prefix += appendAnd ? " and " : ", ";

                        // Opportunity to add 'and' is ended
                        appendAnd = false;

                        // Add the three-digit group to the combined string
                        combined = prefix + combined;
                    }
                }

                // Negative rule
                if (number < 0)
                    combined = "Negative " + combined;

                return combined;
            }



            // Converts a three-digit group into English words
            private string ThreeDigitGroupToWords(int threeDigits)
            {
                // Initialise the return text
                string groupText = "";

                // Determine the hundreds and the remainder
                int hundreds = threeDigits / 100;
                int tensUnits = threeDigits % 100;

                // Hundreds rules
                if (hundreds != 0)
                {
                    groupText += _smallNumbers[hundreds] + " Hundred";

                    if (tensUnits != 0)
                        groupText += " and ";
                }

                // Determine the tens and units
                int tens = tensUnits / 10;
                int units = tensUnits % 10;

                // Tens rules
                if (tens >= 2)
                {
                    groupText += _tens[tens];
                    if (units != 0)
                        groupText += " " + _smallNumbers[units];
                }
                else if (tensUnits != 0)
                    groupText += _smallNumbers[tensUnits];

                return groupText;
            }
        }
        #endregion


        //public void Close()
        //{
        //    _xlsFile = null;
        //}
        //public bool IsClose()
        //{
        //    if (_xlsFile == null)
        //        return true;
        //    return false;
        //}
    }
}
