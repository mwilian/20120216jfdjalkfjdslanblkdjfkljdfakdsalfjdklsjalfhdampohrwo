using System;
using System.IO;
using System.Text;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;

using FlexCel.Core;
using System.Globalization;
using System.Xml;

namespace FlexCel.XlsAdapter
{
    internal class TXlsxBaseRecord : TBaseRecord
    {
        protected override TBaseRecord DoCopyTo(TSheetInfo SheetInfo)
        {
            return (TBaseRecord)MemberwiseClone();
        }

        internal override void SaveToStream(IDataStream Workbook, TSaveData SaveData, int Row)
        {
            //Won't save this one.
        }

        internal override int TotalSize()
        {
            return 0;
        }

        internal override int TotalSizeNoHeaders()
        {
            return 0;
        }

        internal override int GetId
        {
            get { return 0; }
        }
    }

    internal class TFutureStorageRecord : TXlsxBaseRecord
    {
        internal string Xml;

        internal TFutureStorageRecord(string aXml)
        {
            Xml = RemoveIncompatible(aXml);
        }

        private string RemoveIncompatible(string aXml)
        {
            if (string.IsNullOrEmpty(aXml)) return aXml;
            XmlDocument xml = new XmlDocument();
            {
                xml.LoadXml(aXml);
                XmlNodeList Choices = xml.GetElementsByTagName("Choice", TOpenXmlManager.MarkupCompatNamespace);
                if (Choices.Count == 0) return aXml;

                for (int i = Choices.Count - 1; i >= 0; i--)
                {
                    Choices[i].ParentNode.RemoveChild(Choices[i]);
                }

                //Alternate contents must have at least one choice
                XmlNodeList AlternateContent = xml.GetElementsByTagName("AlternateContent", TOpenXmlManager.MarkupCompatNamespace);
                for (int i = AlternateContent.Count - 1; i >= 0; i--)
                {
                    if (AlternateContent[i].ChildNodes.Count == 0)
                    {
                        AlternateContent[i].ParentNode.RemoveChild(AlternateContent[i]);
                    }
                    else if (AlternateContent[i].ChildNodes.Count == 1 && AlternateContent[i].ChildNodes[0].LocalName == "Fallback")
                    {
                        //move the fallback out of the alternatecontent
                        XmlNode ParentNode = AlternateContent[i].ParentNode;
                        XmlNodeList ch = AlternateContent[i].ChildNodes[0].ChildNodes;
                        foreach (XmlNode chChild in ch)
                        {
                            ParentNode.InsertAfter(chChild, AlternateContent[i]);
                        }
                        AlternateContent[i].ParentNode.RemoveChild(AlternateContent[i]);
                    }
                }

                return xml.OuterXml;
            }
        }
    }

    internal class TFutureStorage
    {
        private List<TFutureStorageRecord> FList;

        private TFutureStorage()
        {
            FList = new List<TFutureStorageRecord>();
        }

        internal static void Add(ref TFutureStorage Fs, TFutureStorageRecord R)
        {
            if (R == null) return;
            if (Fs == null) Fs = new TFutureStorage();
            Fs.FList.Add(R);
        }

        internal int Count
        {
            get
            {
                return FList.Count;
            }
        }

        internal TFutureStorageRecord this[int index] { get { return FList[index]; } }

        public override bool Equals(object obj)
        {
            TFutureStorage o2 = obj as TFutureStorage;
            if (o2 == null) return false;

 	        if (FList.Count != o2.FList.Count) return false;
            for (int i = 0; i < FList.Count; i++)
            {
                if (!TFutureStorageRecord.Equals(FList[i], o2.FList[i])) return false;
            }
            return true;
        }

        public override int GetHashCode()
        {
            return base.GetHashCode();
        }

        internal TFutureStorage Clone()
        {
            TFutureStorage Result = new TFutureStorage();
            foreach (TFutureStorageRecord R in FList)
            {
                Result.FList.Add(new TFutureStorageRecord(R.Xml));
            }

            return Result;
        }

        internal static TFutureStorage Clone(TFutureStorage aFutureStorage)
        {
            if (aFutureStorage == null) return null;
            return aFutureStorage.Clone();
        }
    }

    struct TXlsxAttribute
    {
        public string Namespace;
        public string Name;
        public string Value;

        public TXlsxAttribute(string aNamespace, string aName, string aValue)
        {
            Namespace = aNamespace;
            Name = aName;
            Value = aValue;
        }
    }

    internal class TRichTextRun
    {
        internal int FirstChar;
        internal TFlxFont Font;
    }

    internal class TxSSTRecord : TXlsxBaseRecord
    {
        internal string Text;
        internal TRTFRun[] RTFRuns;

        internal static TxSSTRecord LoadFromXml(TOpenXmlReader DataStream, IFlexCelFontList FontList)
        {
            TxSSTRecord Result = new TxSSTRecord();
            Result.Text = String.Empty;

            string StartElement = DataStream.RecordName();
            if (DataStream.IsSimpleTag) { DataStream.NextTag(); return Result; }

            if (!DataStream.NextTag()) return Result;

            List<TRichTextRun> Runs = new List<TRichTextRun>();
            while (!DataStream.AtEndElement(StartElement))
            {
                switch (DataStream.RecordName())
                {
                    case "r":
                        TRichTextRun run = new TRichTextRun();
                        run.FirstChar = Result.Text.Length;
                        run.Font = null;
                        LoadRichTextRun(DataStream, run, Result);
                        if (run.Font != null) Runs.Add(run);

                        break;

                    case "t":
                        Result.Text += DataStream.ReadValueAsString();
                        break;

                    default:
                        DataStream.GetXml();
                        break;
                }
            }

            Result.RTFRuns = new TRTFRun[Runs.Count];
            for (int i = 0; i < Runs.Count; i++)
            {
                Result.RTFRuns[i].FirstChar = Runs[i].FirstChar;
                Result.RTFRuns[i].FontIndex = FontList.AddFont(Runs[i].Font);
            }
            return Result;
        }


        internal static TRichString LoadRichStringFromXml(TOpenXmlReader DataStream, IFlexCelFontList FontList)
        {
            TxSSTRecord r = TxSSTRecord.LoadFromXml(DataStream, FontList);
            return new TRichString(r.Text, r.RTFRuns, FontList, true);
        }

        private static void LoadRichTextRun(TOpenXmlReader DataStream, TRichTextRun run, TxSSTRecord Result)
        {
            string StartElement = DataStream.RecordName();
            if (DataStream.IsSimpleTag) { DataStream.NextTag(); return; }

            if (!DataStream.NextTag()) return;

            while (!DataStream.AtEndElement(StartElement))
            {
                switch (DataStream.RecordName())
                {
                    case "t":
                        Result.Text += DataStream.ReadValueAsString();
                        break;

                    case "rPr":
                        run.Font = new TFlxFont(); //this is not against default font, but an empty font. If the default font had bold, there would be no way to remove it.
                        TXlsxFontReaderWriter.ReadFont(DataStream, run.Font);
                        break;

                    default:
                        DataStream.GetXml();
                        break;
                }
            }
        }

    }

    internal static class TXlsxFontReaderWriter
    {
        internal static void ReadFont(TOpenXmlReader DataStream, TFlxFont Font)
        {
            string StartElement = DataStream.RecordName();
            if (DataStream.IsSimpleTag) { DataStream.NextTag(); return; }

            if (!DataStream.NextTag()) return;
            while (!DataStream.AtEndElement(StartElement))
            {
                switch (DataStream.RecordName())
                {
                    case "b":
                        {
                            bool Value = DataStream.GetAttributeAsBool("val", true);
                            if (Value) Font.Style |= TFlxFontStyles.Bold; else Font.Style &= ~TFlxFontStyles.Bold;
                            break;
                        }

                    case "charset":
                        Font.CharSet = (byte)DataStream.GetAttributeAsInt("val", 0);
                        break;

                    case "color":
                        Font.Color = TXlsxColorReaderWriter.GetColor(DataStream);
                        break;

                    case "condense":
                        {
                            bool Value = DataStream.GetAttributeAsBool("val", true);
                            if (Value) Font.Style |= TFlxFontStyles.Condense; else Font.Style &= ~TFlxFontStyles.Condense;
                            break;
                        }

                    case "extend":
                        {
                            bool Value = DataStream.GetAttributeAsBool("val", true);
                            if (Value) Font.Style |= TFlxFontStyles.Extend; else Font.Style &= ~TFlxFontStyles.Extend;
                            break;
                        }

                    case "family":
                        Font.Family = (byte)DataStream.GetAttributeAsInt("val", 0);
                        break;

                    case "i":
                        {
                            bool Value = DataStream.GetAttributeAsBool("val", true);
                            if (Value) Font.Style |= TFlxFontStyles.Italic; else Font.Style &= ~TFlxFontStyles.Italic;
                            break;
                        }

                    case "outline":
                        {
                            bool Value = DataStream.GetAttributeAsBool("val", true);
                            if (Value) Font.Style |= TFlxFontStyles.Outline; else Font.Style &= ~TFlxFontStyles.Outline;
                            break;
                        }

                    case "name":
                    case "rFont":
                        Font.Name = DataStream.GetAttribute("val");
                        break;

                    case "scheme":
                        {
                            string s = DataStream.GetAttribute("val");
                            Font.Scheme = GetFontScheme(s);
                            break;
                        }

                    case "shadow":
                        {
                            bool Value = DataStream.GetAttributeAsBool("val", true);
                            if (Value) Font.Style |= TFlxFontStyles.Shadow; else Font.Style &= ~TFlxFontStyles.Shadow;
                            break;
                        }

                    case "strike":
                        {
                            bool Value = DataStream.GetAttributeAsBool("val", true);
                            if (Value) Font.Style |= TFlxFontStyles.StrikeOut; else Font.Style &= ~TFlxFontStyles.StrikeOut;
                            break;
                        }

                    case "sz":
                        Font.Size20 = (int) (Math.Round(DataStream.GetAttributeAsDouble("val", 10) * 20));
                        break;

                    case "u":
                        {
                            string s = DataStream.GetAttribute("val");
                            switch (s)
                            {
                                case "double": Font.Underline = TFlxUnderline.Double; break;
                                case "doubleAccounting": Font.Underline = TFlxUnderline.DoubleAccounting; break;
                                case "single": Font.Underline = TFlxUnderline.Single; break;
                                case "singleAccounting": Font.Underline = TFlxUnderline.SingleAccounting; break;
                                default: Font.Underline = TFlxUnderline.Single; break;
                            }
                            break;
                        }

                    case "vertAlign":
                        {
                            string s = DataStream.GetAttribute("val");
                            switch (s)
                            {
                                case "subscript": Font.Style |= TFlxFontStyles.Subscript; break;
                                case "superscript": Font.Style |= TFlxFontStyles.Superscript; break;
                                default: Font.Style &= ~TFlxFontStyles.Subscript & ~TFlxFontStyles.Superscript; break;
                            }
                            break;
                        }

                }

                DataStream.FinishTag();
            }
        }

        internal static TFontScheme GetFontScheme(string s)
        {
            switch (s)
            {
                case "major": return TFontScheme.Major; 
                case "minor": return TFontScheme.Minor;
                default: return TFontScheme.None;
            }
        }

        internal static void WriteFont(TOpenXmlWriter DataStream, TFlxFont Font, bool InStyles)
        {

            if ((Font.Style & TFlxFontStyles.Bold) != 0) DataStream.WriteElement("b", null);
            if ((Font.Style & TFlxFontStyles.Italic) != 0) DataStream.WriteElement("i", null);

            switch (Font.Underline)
            {
                case TFlxUnderline.Single:
                    DataStream.WriteElement("u", null); break;

                case TFlxUnderline.Double:
                    DataStream.WriteStartElement("u"); DataStream.WriteAtt("val", "double"); DataStream.WriteEndElement(); break;

                case TFlxUnderline.SingleAccounting:
                    DataStream.WriteStartElement("u"); DataStream.WriteAtt("val", "singleAccounting"); DataStream.WriteEndElement(); break;

                case TFlxUnderline.DoubleAccounting:
                    DataStream.WriteStartElement("u"); DataStream.WriteAtt("val", "doubleAccounting"); DataStream.WriteEndElement(); break;
            }

            if ((Font.Style & TFlxFontStyles.Condense) != 0) DataStream.WriteElement("condense", null);
            if ((Font.Style & TFlxFontStyles.Extend) != 0) DataStream.WriteElement("extend", null);

            if ((Font.Style & TFlxFontStyles.Outline) != 0) DataStream.WriteElement("outline", null);
            if ((Font.Style & TFlxFontStyles.Shadow) != 0) DataStream.WriteElement("shadow", null);
            if ((Font.Style & TFlxFontStyles.StrikeOut) != 0) DataStream.WriteElement("strike", null);

            if ((Font.Style & TFlxFontStyles.Subscript) != 0)
            {
                DataStream.WriteStartElement("vertAlign"); DataStream.WriteAtt("val", "subscript"); DataStream.WriteEndElement();
            }

            if ((Font.Style & TFlxFontStyles.Superscript) != 0)
            {
                DataStream.WriteStartElement("vertAlign"); DataStream.WriteAtt("val", "superscript"); DataStream.WriteEndElement();
            }


            DataStream.WriteStartElement("sz"); DataStream.WriteAtt("val", Font.Size20 / 20.0); DataStream.WriteEndElement();
            DataStream.WriteStartElement("color"); TXlsxColorReaderWriter.WriteColor(DataStream, Font.Color); DataStream.WriteEndElement();

            if (InStyles)
            {
                DataStream.WriteStartElement("name");
            }
            else
            {
                DataStream.WriteStartElement("rFont");
            }
            DataStream.WriteAtt("val", Font.Name);
            DataStream.WriteEndElement();

            if (Font.Family != 0) { DataStream.WriteStartElement("family"); DataStream.WriteAtt("val", Font.Family); DataStream.WriteEndElement(); }
            switch (Font.Scheme)
            {
                case TFontScheme.Minor:
                    DataStream.WriteStartElement("scheme"); DataStream.WriteAtt("val", "minor"); DataStream.WriteEndElement();
                    break;

                case TFontScheme.Major:
                    DataStream.WriteStartElement("scheme"); DataStream.WriteAtt("val", "major"); DataStream.WriteEndElement();
                    break;
            }

            if (Font.CharSet != 0) { DataStream.WriteStartElement("charset"); DataStream.WriteAtt("val", Font.CharSet); DataStream.WriteEndElement(); }
        }

    }

    internal static class TXlsxColorReaderWriter
    {
        internal static TExcelColor GetColor(TOpenXmlReader DataStream)
        {
            double tint = DataStream.GetAttributeAsDouble("tint", 0);
            if (DataStream.GetAttributeAsBool("auto", false)) return TExcelColor.Automatic;

            int theme = DataStream.GetAttributeAsInt("theme", -1);
            if (theme >= 0 && Enum.IsDefined(typeof(TThemeColor), theme))
            {
                return TExcelColor.FromTheme((TThemeColor)theme, tint);
            }

            unchecked
            {
                long RGB = DataStream.GetAttributeAsHex("rgb", -1);
                unchecked
                {
                    if (RGB >= 0) return TExcelColor.FromArgb((int)RGB, tint);
                }
            }

            int Index = DataStream.GetAttributeAsInt("indexed", -1);
            if (Index >= 0) return TExcelColor.FromIndex(Index - 7);

            return TExcelColor.Automatic;
        }

        internal static void WriteColor(TOpenXmlWriter DataStream, TExcelColor aColor)
        {
            if (aColor.Tint != 0) DataStream.WriteAtt("tint", aColor.Tint);
            switch (aColor.ColorType)
            {
                case TColorType.RGB:
                    DataStream.WriteAttHex("rgb", aColor.RGB, 6);
                    break;

                case TColorType.Theme:
                    DataStream.WriteAtt("theme", (int)aColor.Theme);
                    break;

                case TColorType.Indexed:
                    DataStream.WriteAtt("indexed", (int)aColor.InternalIndex + 7);
                    break;

                case TColorType.Automatic:
                default:
                    DataStream.WriteAtt("auto", true, false);
                    break;
            }
        }
    }

    internal static class TXlsxBorderReaderWriter
    {
        internal static TFlxBorders LoadFromXml(TOpenXmlReader DataStream)
        {
            Debug.Assert(DataStream.RecordName() == "border", "LoadFromXml should be called from the start of a border element");
            TFlxBorders Result = new TFlxBorders();

            if (DataStream.GetAttributeAsBool("diagonalUp", false)) Result.DiagonalStyle = TFlxDiagonalBorder.DiagUp;
            if (DataStream.GetAttributeAsBool("diagonalDown", false))
            {
                if (Result.DiagonalStyle == TFlxDiagonalBorder.DiagUp) Result.DiagonalStyle = TFlxDiagonalBorder.Both;
                else Result.DiagonalStyle = TFlxDiagonalBorder.DiagDown;
            } 


            string StartElement = DataStream.RecordName();
            if (DataStream.IsSimpleTag) { DataStream.NextTag(); return Result; }
            if (!DataStream.NextTag()) return Result;
            while (!DataStream.AtEndElement(StartElement))
            {
                switch (DataStream.RecordName())
                {
                    case "bottom":
                        LoadBorder(DataStream, ref Result.Bottom);
                        break;
                    case "diagonal":
                        LoadBorder(DataStream, ref Result.Diagonal);
                        break;
                    case "left":
                        LoadBorder(DataStream, ref Result.Left);
                        break;
                    case "right":
                        LoadBorder(DataStream, ref Result.Right);
                        break;
                    case "top":
                        LoadBorder(DataStream, ref Result.Top);
                        break;

                    case "horizontal":
                    case "vertical":
                    default:
                        DataStream.GetXml();
                        break;
                }
            }

            return Result;
        }

        private static void LoadBorder(TOpenXmlReader DataStream, ref TFlxOneBorder aBorder)
        {
            string Style = DataStream.GetAttribute("style");
            if (Style != null) aBorder.Style = GetBorderStyle(Style);

            string StartElement = DataStream.RecordName();
            if (DataStream.IsSimpleTag) { DataStream.NextTag(); return; }
            if (!DataStream.NextTag()) return;
            while (!DataStream.AtEndElement(StartElement))
            {
                if (DataStream.RecordName() == "color")
                {
                   aBorder.Color = TXlsxColorReaderWriter.GetColor(DataStream);
                    DataStream.FinishTag();
                }
                else
                {
                    DataStream.GetXml();
                }

            }
        }

        private static TFlxBorderStyle GetBorderStyle(string Style)
        {
            switch (Style)
            {
                case "dashDot": return TFlxBorderStyle.Dash_dot;
                case "dashDotDot": return TFlxBorderStyle.Dash_dot_dot;
                case "dashed": return TFlxBorderStyle.Dashed;
                case "dotted": return TFlxBorderStyle.Dotted;
                case "double": return TFlxBorderStyle.Double;
                case "hair": return TFlxBorderStyle.Hair;
                case "medium": return TFlxBorderStyle.Medium;
                case "mediumDashDot": return TFlxBorderStyle.Medium_dash_dot;
                case "mediumDashDotDot": return TFlxBorderStyle.Medium_dash_dot_dot;
                case "mediumDashed": return TFlxBorderStyle.Medium_dashed;
                case "slantDashDot": return TFlxBorderStyle.Slanted_dash_dot;
                case "thick": return TFlxBorderStyle.Thick;
                case "thin": return TFlxBorderStyle.Thin;
                default:
                    return TFlxBorderStyle.None;
            }
        }

        private static string GetBorderString(TFlxBorderStyle Border)
        {
            switch (Border)
            {
                case TFlxBorderStyle.Dash_dot: return "dashDot";
                case TFlxBorderStyle.Dash_dot_dot: return "dashDotDot";
                case TFlxBorderStyle.Dashed: return "dashed";
                case TFlxBorderStyle.Dotted: return "dotted";
                case TFlxBorderStyle.Double: return "double";
                case TFlxBorderStyle.Hair: return "hair";
                case TFlxBorderStyle.Medium: return "medium";
                case TFlxBorderStyle.Medium_dash_dot: return "mediumDashDot";
                case TFlxBorderStyle.Medium_dash_dot_dot: return "mediumDashDotDot";
                case TFlxBorderStyle.Medium_dashed: return "mediumDashed";
                case TFlxBorderStyle.Slanted_dash_dot: return "slantDashDot";
                case TFlxBorderStyle.Thick: return "thick";
                case TFlxBorderStyle.Thin: return "thin";
                default:
                    return "none";
            }
        }

        internal static void SaveToXml(TOpenXmlWriter DataStream, TFlxBorders border)
        {
            DataStream.WriteAtt("diagonalDown", border.DiagonalStyle == TFlxDiagonalBorder.DiagDown || border.DiagonalStyle == TFlxDiagonalBorder.Both, false);
            DataStream.WriteAtt("diagonalUp", border.DiagonalStyle == TFlxDiagonalBorder.DiagUp || border.DiagonalStyle == TFlxDiagonalBorder.Both, false);

            //order is important.
            WriteBorder(DataStream, "left", border.Left);
            WriteBorder(DataStream, "right", border.Right);
            WriteBorder(DataStream, "top", border.Top);
            WriteBorder(DataStream, "bottom", border.Bottom);

            WriteBorder(DataStream, "diagonal", border.Diagonal);

        }

        private static void WriteBorder(TOpenXmlWriter DataStream, string tag, TFlxOneBorder Border)
        {
            if (Border.Style == TFlxBorderStyle.None) return;
            DataStream.WriteStartElement(tag, false);
            DataStream.WriteAtt("style", GetBorderString(Border.Style));

            DataStream.WriteStartElement("color");
            TXlsxColorReaderWriter.WriteColor(DataStream, Border.Color);
            DataStream.WriteEndElement();
            
            DataStream.WriteEndElement();
        }

    }

    internal static class TXlsxFillReaderWriter
    {
        internal static TFlxFillPattern LoadFromXml(TOpenXmlReader DataStream)
        {
            TFlxFillPattern Result = new TFlxFillPattern();

            string StartElement = DataStream.RecordName();
            if (DataStream.IsSimpleTag) { DataStream.NextTag(); return Result; }
            if (!DataStream.NextTag()) return Result;
            while (!DataStream.AtEndElement(StartElement))
            {
                switch (DataStream.RecordName())
                {
                    case "patternFill":
                        LoadPattern(DataStream, ref Result);
                        break;

                    case "gradientFill":
                        LoadGradient(DataStream, ref Result);
                        break;

                    default:
                        DataStream.GetXml();
                        break;
                }
            }

            return Result;
        }

        #region Pattern style
        private static TFlxPatternStyle GetPattern(string Style)
        {
            switch (Style)
            {
                case "solid": return TFlxPatternStyle.Solid;
                case "mediumGray": return TFlxPatternStyle.Gray50;
                case "darkGray": return TFlxPatternStyle.Gray75;
                case "lightGray": return TFlxPatternStyle.Gray25;
                case "darkHorizontal": return TFlxPatternStyle.Horizontal;
                case "darkVertical": return TFlxPatternStyle.Vertical;

                case "darkDown": return TFlxPatternStyle.Down;
                case "darkUp": return TFlxPatternStyle.Up;

                case "darkGrid": return TFlxPatternStyle.Checker;
                case "darkTrellis": return TFlxPatternStyle.SemiGray75;

                case "lightHorizontal": return TFlxPatternStyle.LightHorizontal;
                case "lightVertical": return TFlxPatternStyle.LightVertical;

                case "lightDown": return TFlxPatternStyle.LightDown;
                case "lightUp": return TFlxPatternStyle.LightUp;

                case "lightGrid": return TFlxPatternStyle.Grid;
                case "lightTrellis": return TFlxPatternStyle.CrissCross;

                case "gray125": return TFlxPatternStyle.Gray16;
                case "gray0625": return TFlxPatternStyle.Gray8;

                case "none":
                default:
                    return TFlxPatternStyle.None;

            }
        }

        private static string GetStyleString(TFlxPatternStyle Pattern)
        {
            switch (Pattern)
            {
                case TFlxPatternStyle.Solid: return "solid";
                case TFlxPatternStyle.Gray50: return "mediumGray";
                case TFlxPatternStyle.Gray75: return "darkGray";
                case TFlxPatternStyle.Gray25: return "lightGray";
                case TFlxPatternStyle.Horizontal: return "darkHorizontal";
                case TFlxPatternStyle.Vertical: return "darkVertical";

                case TFlxPatternStyle.Down: return "darkDown";
                case TFlxPatternStyle.Up: return "darkUp";

                case TFlxPatternStyle.Checker: return "darkGrid";
                case TFlxPatternStyle.SemiGray75: return "darkTrellis";

                case TFlxPatternStyle.LightHorizontal: return "lightHorizontal";
                case TFlxPatternStyle.LightVertical: return "lightVertical";

                case TFlxPatternStyle.LightDown: return "lightDown";
                case TFlxPatternStyle.LightUp: return "lightUp";

                case TFlxPatternStyle.Grid: return "lightGrid";
                case TFlxPatternStyle.CrissCross: return "lightTrellis";

                case TFlxPatternStyle.Gray16: return "gray125";
                case TFlxPatternStyle.Gray8: return "gray0625";

                default: return "none";
            }
        }
        #endregion

        private static void LoadPattern(TOpenXmlReader DataStream, ref TFlxFillPattern Pattern)
        {
            string Style = DataStream.GetAttribute("patternType");
            if (Style != null) Pattern.Pattern = GetPattern(Style);

            string StartElement = DataStream.RecordName();
            if (DataStream.IsSimpleTag) { DataStream.NextTag(); return; }
            if (!DataStream.NextTag()) return;
            while (!DataStream.AtEndElement(StartElement))
            {
                switch (DataStream.RecordName())
                {
                    case "bgColor":
                        Pattern.BgColor = TXlsxColorReaderWriter.GetColor(DataStream);
                        DataStream.FinishTag();
                        break;
                    case "fgColor":
                        Pattern.FgColor = TXlsxColorReaderWriter.GetColor(DataStream);
                        DataStream.FinishTag();
                        break;

                    default:
                        DataStream.GetXml();
                        break;
                }
            }
        }

        private static void LoadGradient(TOpenXmlReader DataStream, ref TFlxFillPattern Pattern)
        {
            TExcelGradient Gradient;
            List<TGradientStop> Stops = new List<TGradientStop>();

            switch (DataStream.GetAttribute("type"))
            {
                case "path":
                    Gradient = new TExcelRectangularGradient(
                        null,
                        DataStream.GetAttributeAsDouble("top", 0),
                        DataStream.GetAttributeAsDouble("left", 0),
                        DataStream.GetAttributeAsDouble("bottom", 0),
                        DataStream.GetAttributeAsDouble("right", 0)
                        );
                    break;

                case "linear":
                default:
                    Gradient = new TExcelLinearGradient(
                        null,
                        DataStream.GetAttributeAsDouble("degree", 0));
                    break;
            }


            string StartElement = DataStream.RecordName();
            if (DataStream.IsSimpleTag) { DataStream.NextTag(); return; }
            if (!DataStream.NextTag()) return;
            while (!DataStream.AtEndElement(StartElement))
            {
                switch (DataStream.RecordName())
                {
                    case "stop":
                        Stops.Add(ReadGradientStop(DataStream, DataStream.GetAttributeAsDouble("position", 0)));
                        break;

                    default:
                        DataStream.GetXml();
                        break;
                }
            }

            Gradient.Stops = Stops.ToArray();
            Pattern.Gradient = Gradient;
        }

        private static TGradientStop ReadGradientStop(TOpenXmlReader DataStream, double Position)
        {
            TGradientStop Result = new TGradientStop(Position, ColorUtil.Empty);
            string StartElement = DataStream.RecordName();
            if (DataStream.IsSimpleTag) { DataStream.NextTag(); return Result; }
            if (!DataStream.NextTag()) return Result;
            while (!DataStream.AtEndElement(StartElement))
            {
                switch (DataStream.RecordName())
                {
                    case "color":
                        Result.Color = TXlsxColorReaderWriter.GetColor(DataStream);
                        DataStream.FinishTag();
                        break;

                    default:
                        DataStream.GetXml();
                        break;
                }
            }

            return Result;
        }

        internal static void SaveToXml(TOpenXmlWriter DataStream, TFlxFillPattern Pattern)
        {
            if (Pattern.Pattern == TFlxPatternStyle.Gradient) WriteGradient(DataStream, Pattern); else WritePattern(DataStream, Pattern);
        }

        private static void WritePattern(TOpenXmlWriter DataStream, TFlxFillPattern pat)
        {
            DataStream.WriteStartElement("patternFill");
            DataStream.WriteAtt("patternType", GetStyleString(pat.Pattern));
            if (!pat.FgColor.IsAutomatic)
            {
                DataStream.WriteStartElement("fgColor");
                TXlsxColorReaderWriter.WriteColor(DataStream, pat.FgColor);
                DataStream.WriteEndElement();
            }

            if (!pat.BgColor.IsAutomatic)
            {
                DataStream.WriteStartElement("bgColor");
                TXlsxColorReaderWriter.WriteColor(DataStream, pat.BgColor);
                DataStream.WriteEndElement();
            }

            DataStream.WriteEndElement();
        }

        private static void WriteGradient(TOpenXmlWriter DataStream, TFlxFillPattern pat)
        {
            DataStream.WriteStartElement("gradientFill");
            switch (pat.Gradient.GradientType)
            {
                case TGradientType.Linear:
                    TExcelLinearGradient lg = (TExcelLinearGradient)pat.Gradient;
                    DataStream.WriteAtt("type", "linear");
                    DataStream.WriteAtt("degree", lg.RotationAngle);

                    break;
                case TGradientType.Rectangular:
                        TExcelRectangularGradient rg = (TExcelRectangularGradient)pat.Gradient;
                        DataStream.WriteAtt("type", "path");
                        DataStream.WriteAtt("top", rg.Top);
                        DataStream.WriteAtt("left", rg.Left);
                        DataStream.WriteAtt("bottom", rg.Bottom);
                        DataStream.WriteAtt("right", rg.Right);

                    break;
            }

            WriteStops(DataStream, pat.Gradient.Stops);
            DataStream.WriteEndElement();
            
        }

        private static void WriteStops(TOpenXmlWriter DataStream, TGradientStop[] aGradientStop)
        {
            foreach (TGradientStop stop in aGradientStop)
            {
                DataStream.WriteStartElement("stop");
                DataStream.WriteAtt("position", stop.Position);

                DataStream.WriteStartElement("color");
                TXlsxColorReaderWriter.WriteColor(DataStream, stop.Color);
                DataStream.WriteEndElement();

                DataStream.WriteEndElement();
            }
        }
    }

  
    internal struct TArrayAndTableFmlaData
    {
        internal TXlsCellRange ArrayRange;
        internal TXlsCellRange TableRange;
    }

    internal struct TSharedFmlaData
    {
        internal int Row;
        internal int Col;
        internal TParsedTokenList SharedTokens;

        public TSharedFmlaData(int aRow, int aCol, TParsedTokenList aSharedTokens)
        {
            Row = aRow;
            Col = aCol;
            SharedTokens = aSharedTokens;
        }
    }
    
    internal static class TXlsxCellReader
    {
        internal static void ReadRow(ExcelFile Workbook, TSST SST, TVirtualReader VirtualReader, TOpenXmlReader DataStream, TSheet WorkSheet, int WorkingSheet, ref int row,
            Dictionary<int, TSharedFmlaData> SharedFormulas, ref bool HasMultiCellArrayFmlas, bool Dates1904)
        {
            row = DataStream.GetAttributeAsInt("r", (row + 1) + 1) - 1;  //row is 1 based in xml, the parameter here is 0-based.
            int XF = DataStream.GetAttributeAsInt("s", FlxConsts.DefaultFormatId);
            bool CustomXF = DataStream.GetAttributeAsBool("customFormat", false);
            double RowHeight = DataStream.GetAttributeAsDouble("ht", GetDefaultRowHeight(WorkSheet)); //This depends on the normal font height, or DefaultRowHeight if CustomSize is true in deafaultRowHeight.
            bool CustomHeight = DataStream.GetAttributeAsBool("customHeight", false);

            bool Collapsed = DataStream.GetAttributeAsBool("collapsed", false);
            bool Hidden = DataStream.GetAttributeAsBool("hidden", false);

            int OutlineLevel = DataStream.GetAttributeAsInt("outlineLevel", 0);
            bool Phonetic = DataStream.GetAttributeAsBool("ph", false);

            bool ThickTop = DataStream.GetAttributeAsBool("thickTop", false);
            bool ThickBot = DataStream.GetAttributeAsBool("thickBot", false);

            TRowRecord RowRecord = new TRowRecord(XF, CustomXF, (int)(RowHeight * 20),
                CustomHeight, Collapsed, Hidden, OutlineLevel, Phonetic, ThickTop, ThickBot);

            if (VirtualReader != null)
            { 
            }
            else
            {
                WorkSheet.Cells.AddRow(row, RowRecord);
            }

            int CurrentCol = -1;

            string StartElement = DataStream.RecordName();
            if (DataStream.IsSimpleTag) { DataStream.NextTag(); return; }
            if (!DataStream.NextTag()) return;
            while (!DataStream.AtEndElement(StartElement))
            {
                switch (DataStream.RecordName())
                {
                    case "c":
                        TArrayAndTableFmlaData ExtraData;
                        TCellRecord Cell = GetCell(Workbook, DataStream, WorkingSheet, row, ref CurrentCol, SST, Workbook, SharedFormulas, out ExtraData, Dates1904);
                        WorkSheet.Cells.AddCell(Cell, row, HasMultiCellArrayFmlas, VirtualReader); //When no multicell array formulas, no need to look for existing cells.

                        if (ExtraData.ArrayRange != null)
                        {
                            PopulateArrayFormula(WorkSheet, ref HasMultiCellArrayFmlas, ExtraData.ArrayRange, Cell, VirtualReader);
                        }
                        else
                            if (ExtraData.TableRange != null)
                            {
                                PopulateArrayFormula(WorkSheet, ref HasMultiCellArrayFmlas, ExtraData.TableRange, Cell, VirtualReader);
                            }

                        break;

                    default:
                        RowRecord.AddFutureStorage(new TFutureStorageRecord(DataStream.GetXml()));
                        break;
                }
            }
        }

        private static double GetDefaultRowHeight(TSheet Worksheet)
        {
            return Worksheet.DefRowHeight / 20.0;
        }

        private static void PopulateArrayFormula(TSheet WorkSheet, ref bool HasMultiCellArrayFmlas, TXlsCellRange Range, TCellRecord Cell, TVirtualReader VirtualReader)
        {
            if (!HasMultiCellArrayFmlas) HasMultiCellArrayFmlas = Range.RowCount > 1 || Range.ColCount > 1;

            TFormulaRecord f = Cell as TFormulaRecord;
            if (f == null) FlxMessages.ThrowException(FlxErr.ErrInternal);

            if (VirtualReader != null)
            {
                VirtualReader.AddArray(Range, f);
                return;
            }


            {
                for (int r = Range.Top; r <= Range.Bottom; r++)
                {
                    for (int c = Range.Left; c <= Range.Right; c++)
                    {
                        if (r != Range.Top || c != Range.Left)
                        {
                            TFormulaRecord f1 = new TFormulaRecord(f.Id, r, c, null, f.XF, f.CloneData(), null, null, (int)f.OptionFlags, false, 0, f.bx); //ArrayOptionFlags doesn't matter here, this doesn't have an array.
                            WorkSheet.Cells.AddCell(f1, r, VirtualReader);
                        }
                    }
                }
            }
        }

        internal static TCellRecord GetCell(ExcelFile Workbook, TOpenXmlReader DataStream, int WorkingSheet, int row, ref int Col,
            TSST aSST, IFlexCelFontList aFontList, Dictionary<int, TSharedFmlaData> SharedFormulas, out TArrayAndTableFmlaData ExtraData, bool Dates1904)
        {
            ExtraData = new TArrayAndTableFmlaData();
            //bool Phonetic = DataStream.GetAttributeAsBool("ph", false);
            string r = DataStream.GetAttribute("r");
            if (r != null && r.Trim().Length > 0)
            {
                if (TCellAddress.ReadSimpleCol(r, 0, out Col) < 0) FlxMessages.ThrowException(FlxErr.ErrInvalidRef, r); //We already know Row, so we won't read it again
                Col--; //1 based
                //if (row != Row - 1) XlsMessages.ThrowException(XlsErr.ErrExcelInvalid);
            }
            else
            {
                Col++;
            }

            int XF = DataStream.GetAttributeAsInt("s", FlxConsts.DefaultFormatId);

            string CellValue = null;
            TRichString CellRichString = null;
            TFormulaRecord formula = null;
            TFutureStorageRecord FutureStorage = null;

            string celltype = DataStream.GetAttribute("t"); //all attributes must be read before we move on.

            string StartElement = DataStream.RecordName();
            if (DataStream.IsSimpleTag) { DataStream.NextTag(); return new TBlankRecord(Col, XF); }
            if (!DataStream.NextTag()) XlsMessages.ThrowException(XlsErr.ErrExcelInvalid);
            while (!DataStream.AtEndElement(StartElement))
            {
                switch (DataStream.RecordName())
                {
                    case "f": formula = GetFormula(Workbook, WorkingSheet, row, Col, XF, DataStream, SharedFormulas, out ExtraData); break;
                    case "is": CellRichString = TxSSTRecord.LoadRichStringFromXml(DataStream, aFontList); break;
                    case "v": CellValue = DataStream.ReadValueAsString(); break;

                    default: FutureStorage = new TFutureStorageRecord(DataStream.GetXml());
                        break;
                }

            }

            if (formula == null)
            {
                TCellRecord Result = GetSimpleCell(Col, XF, CellValue, CellRichString, celltype, aSST, aFontList, Dates1904);
                if (Result != null) Result.AddFutureStorage(FutureStorage);
                return Result;
            }

            SetFormulaValue(formula, CellValue, celltype);
            if (formula != null) formula.AddFutureStorage(FutureStorage);
            return formula;
        }

        private static TCellRecord GetSimpleCell(int Col, int XF, string CellValue, TRichString CellRichString, string celltype,
            TSST SST, IFlexCelFontList FontList, bool Dates1904)
        {
            switch (celltype)
            {
                case "b": return new TBoolErrRecord(Col, XF, FlxConvert.ToXlsxBoolean(CellValue));
                case "e": return new TBoolErrRecord(Col, XF, (TFlxFormulaErrorValue)TFormulaMessages.StringToErrCode(CellValue, true));
                case "inlineStr":
                    {
                        if (CellRichString == null) CellRichString = new TRichString();
                        return new TLabelSSTRecord(Col, XF, SST, FontList, CellRichString);
                    }

                case "s": //sst
                    return new TLabelSSTRecord(Col, XF, Convert.ToInt64(CellValue), SST, FontList);
                
                case "str": //This is a formula result. It can happen here only if we are setting an array formula or table.
                    return new TxLabelRecord(CellValue, Col, XF);

                case "d": //added in 2010 and not documented propertly :(  Excel 2007 on its current state won't support it.
                    {
                        DateTime DateValue = Convert.ToDateTime(CellValue, CultureInfo.InvariantCulture);
                        double RealValue = FlxDateTime.ToOADate(DateValue, Dates1904);
                        if (TRKRecord.IsRK(RealValue)) return new TRKRecord(Col, XF, RealValue);
                        else return new TNumberRecord(Col, XF, RealValue); //.CreateFromData
                    }

                case "n":
                default: //default celltype is number.
                    {
                        double RealValue = Convert.ToDouble(CellValue, CultureInfo.InvariantCulture);
                        if (TRKRecord.IsRK(RealValue)) return new TRKRecord(Col, XF, RealValue);
                        else return new TNumberRecord(Col, XF, RealValue); //.CreateFromData
                    }

            }
        }

        private static void SetFormulaValue(TFormulaRecord formula, string CellValue, string celltype)
        {
            if (CellValue == null) { formula.FormulaValue = null; return; }
            switch (celltype)
            {
                case "b": formula.FormulaValue = FlxConvert.ToXlsxBoolean(CellValue); break;
                case "e": formula.FormulaValue = (TFlxFormulaErrorValue)TFormulaMessages.StringToErrCode(CellValue, true); break;
                case "inlineStr": XlsMessages.ThrowException(XlsErr.ErrExcelInvalid); break; //this shouldn't appear in formulas
                case "s": XlsMessages.ThrowException(XlsErr.ErrExcelInvalid); break; //this shouldn't appear in formulas
                case "str": formula.FormulaValue = CellValue; break;
                case "n":
                default: formula.FormulaValue = Convert.ToDouble(CellValue, CultureInfo.InvariantCulture); break;
            }
        }

        private static TFormulaRecord GetFormula(ExcelFile Workbook, int WorkingSheet, int Row, int Col, int XF, TOpenXmlReader DataStream, 
            Dictionary<int, TSharedFmlaData> SharedFormulas, out TArrayAndTableFmlaData ExtraData)
        {
            string FormulaType = DataStream.GetAttribute("t");
            ExtraData = new TArrayAndTableFmlaData();
            TParsedTokenList ArrayData = null;
            bool MasterSharedFormula = false;
            int si = DataStream.GetAttributeAsInt("si", -1);

            bool AlwaysRecalc = DataStream.GetAttributeAsBool("ca", false);
            int OptionFlags = 0;
            if (AlwaysRecalc) OptionFlags |= 1;

            int ArrayOptionFlags = 0;
            if (DataStream.GetAttributeAsBool("aca", false)) ArrayOptionFlags |= 1;

            bool bx = DataStream.GetAttributeAsBool("bx", false);


            switch (FormulaType)
            {
                case "array":
                    ExtraData.ArrayRange = ReadFormulaReferences(Workbook, DataStream);
                    break;

                case "dataTable":
                    {
                        ExtraData.TableRange = ReadFormulaReferences(Workbook, DataStream);
                        TXlsCellRange range = ExtraData.TableRange;
                        TFormulaRecord Result = new TFormulaRecord((int)xlr.FORMULA, Row, Col, null, XF,
                            new TParsedTokenList(new TBaseParsedToken[] { new TTableToken(range.Top, range.Left) }),
                            null, null, OptionFlags, false, ArrayOptionFlags, bx);

                        bool IsRowTable = DataStream.GetAttributeAsBool("dtr", false);
                        bool dt2d = DataStream.GetAttributeAsBool("dt2D", false);
                        bool del1 = DataStream.GetAttributeAsBool("del1", false);
                        bool del2 = DataStream.GetAttributeAsBool("del2", false);
                        TCellAddress FirstCell = ReadCellInput(DataStream, "r1");
                        TCellAddress SecondCell = ReadCellInput(DataStream, "r2");

                        int oflags = 0;
                        if (AlwaysRecalc) oflags |= 0x01;
                        if (IsRowTable) oflags |= 0x04;
                        if (dt2d) oflags |= 0x08;
                        if (del1) { oflags |= TTableRecord.FlagDeleted1; FirstCell= new TCellAddress(0, 0); }
                        if (del2) { oflags |= TTableRecord.FlagDeleted2; SecondCell = new TCellAddress(0, 0); }

                        int rir = FirstCell.Row - 1;
                        int cir = FirstCell.Col - 1;

                        int ric = SecondCell.Row - 1;
                        int cic = SecondCell.Col - 1;

                        Result.TableRecord = new TTableRecord((int)xlr.TABLE, oflags, range.Top, range.Left, range.Bottom, range.Right, rir, cir, ric, cic);
                        DataStream.FinishTag();
                        return Result;
                    }
                case "shared":
                    if (si >= 0) //if not, ignore it. it is a shared formula.
                    {
                        TSharedFmlaData SharedFmlaData;
                        if (SharedFormulas.TryGetValue(si, out SharedFmlaData))
                        {
                            TFormulaRecord Result = new TFormulaRecord((int)xlr.FORMULA, Row, Col, ExtraData.ArrayRange, XF, null, ArrayData, null, OptionFlags, false, ArrayOptionFlags, bx); //Dates1904 doesn't matter here.
                            Result.MixShared(SharedFmlaData.SharedTokens, Row, Col, false);
                            DataStream.FinishTag();
                            return Result;
                        }
                        
                        MasterSharedFormula = true;
                    }
                    break;

                case "normal":
                default:
                    break;
            }

            //This can't be read until all attributes have been read.
            string FormulaText = DataStream.ReadValueAsString();
            if (FormulaText == null || FormulaText.Length == 0) return null;

            TFormulaConvertTextToInternal Ps = new TFormulaConvertTextToInternal(Workbook, WorkingSheet, true, TFormulaMessages.TokenString(TFormulaToken.fmStartFormula) + FormulaText, true, MasterSharedFormula);
            Ps.SetReadingXlsx();
            Ps.SetStartForRelativeRefs(Row, Col);
            Ps.Parse();
            TParsedTokenList Data = Ps.GetTokens();

            if (MasterSharedFormula)
            {
                SharedFormulas.Add(si, new TSharedFmlaData(Row, Col, Data.Clone()));
            }

            if (ExtraData.ArrayRange != null)
            {
                ArrayData = Data;
                Data = new TParsedTokenList(new TBaseParsedToken[] { new TExp_Token(ExtraData.ArrayRange.Top, ExtraData.ArrayRange.Left) });
            }

            TFormulaRecord Result2 = new TFormulaRecord((int)xlr.FORMULA, Row, Col, ExtraData.ArrayRange, XF, Data, ArrayData, null, OptionFlags, false, ArrayOptionFlags, bx); //Dates1904 doesn't matter here.

            if (MasterSharedFormula)
            {
                Result2.MixShared(Data, Row, Col, false);
            }

            return Result2;
        }

        private static TCellAddress ReadCellInput(TOpenXmlReader DataStream, string CellId)
        {
            string s1 = DataStream.GetAttribute(CellId);
            if (s1 != null) return new TCellAddress(s1);
            return new TCellAddress(0, 0);
        }

        private static TXlsCellRange ReadFormulaReferences(ExcelFile Workbook, TOpenXmlReader DataStream)
        {
            ExcelFile LocalXls; //not needed here, there are no external refs here.
            int sheet1, sheet2, row1, col1, row2, col2;
            string range = DataStream.GetAttribute("ref");
            TCellAddress.ParseAddress(Workbook, range, -1, out LocalXls, out sheet1, out sheet2, out row1, out col1, out row2, out col2);
            if (sheet1 != -1 || sheet2 != -1) XlsMessages.ThrowException(XlsErr.ErrExcelInvalid);
            return new TXlsCellRange(row1 - 1, col1 - 1, row2 - 1, col2 - 1);
        }

    }

    internal sealed class TXlsxRichStringWriter
    {
        private TXlsxRichStringWriter() { }

        internal static void WriteRichText(TOpenXmlWriter DataStream, IFlexCelFontList FontList, TExcelString Se)
        {
            TRTFRun[] Runs = TRTFRun.ToRTFRunArray(Se.RichTextFormats); //This could be unlooped to avoid the creation of a temporary TRTFRun struct. But, as rich text is not that common, it is not worth the code duplication.
            string s = Se.Data;
            WriteRichText(DataStream, FontList, Runs, s);
        }

        internal static void WriteRichText(TOpenXmlWriter DataStream, IFlexCelFontList FontList, TRTFRun[] Runs, string s)
        {
            if (Runs.Length == 0)
            {
                WriteT(DataStream, s);
                return;
            }

            if (Runs[0].FirstChar > 0)
            {
                DataStream.WriteStartElement("r");
                WriteT(DataStream, s.Substring(0, Runs[0].FirstChar));
                DataStream.WriteEndElement();
            }

            for (int i = 0; i < Runs.Length; i++)
            {
                int StartPos = Runs[i].FirstChar;
                int EndPos = i < Runs.Length - 1 ? Runs[i + 1].FirstChar : s.Length;

                if (EndPos > StartPos)
                {
                    DataStream.WriteStartElement("r");
                    WriteRPr(DataStream, FontList, Runs[i].FontIndex);
                    WriteT(DataStream, s.Substring(StartPos, EndPos - StartPos));
                    DataStream.WriteEndElement();
                }
            }
        }

        internal static void WriteT(TOpenXmlWriter DataStream, string s)
        {
            DataStream.WriteElement("t", s);
        }

        private static void WriteRPr(TOpenXmlWriter DataStream, IFlexCelFontList FontList, int FontIndex)
        {
            TFlxFont fnt = FontList.GetFont(FontIndex);

            DataStream.WriteStartElement("rPr");
            TXlsxFontReaderWriter.WriteFont(DataStream, fnt, false);
            DataStream.WriteEndElement();
        }

        internal static void WriteRichOrPlainText(TOpenXmlWriter DataStream, IFlexCelFontList FontList, TExcelString Se)
        {
            if (Se.HasRichText)
            {
                WriteRichText(DataStream, FontList, Se);
            }
            else
            {
                WriteT(DataStream, Se.Data);
            }
        }

        internal static void WriteRichOrPlainText(TOpenXmlWriter DataStream, IFlexCelFontList FontList, TRichString rs)
        {
            WriteRichText(DataStream, FontList, rs.GetRuns(), rs.Value);
        }

    }

    internal static class TXlsxCellWriter
    {

        internal static void WriteRow(TOpenXmlWriter DataStream, TSheet Worksheet, TCells Cells, int Row, bool Dates1904)
        {
            if (!Cells.CellList.HasRow(Row)) return;
            DataStream.WriteStartElement("row");

            DataStream.WriteAtt("r", Row + 1);
            TRowRecord r = Cells.CellList[Row].RowRecord;

            if (r.XF != FlxConsts.DefaultFormatId) DataStream.WriteAtt("s", r.XF);
            DataStream.WriteAtt("customFormat", r.IsFormatted(), false);

            if (Worksheet.DefRowHeight != r.Height)
            {
                int mrh = r.Height;
                if (mrh > XlsConsts.MaxRowHeight)
                {
                    mrh = XlsConsts.MaxRowHeight;
                }
                DataStream.WriteAtt("ht", mrh / 20.0); //This depends on the normal font height, or DefaultRowHeight if CustomSize is true in deafaultRowHeight.
            }
            DataStream.WriteAtt("customHeight", !r.IsAutoHeight(), false);

            DataStream.WriteAtt("collapsed", r.IsCollapsed(), false);
            DataStream.WriteAtt("hidden", r.IsHidden(), false);

            int OutlineLevel = r.GetRowOutlineLevel();
            if (OutlineLevel > 0) DataStream.WriteAtt("outlineLevel", OutlineLevel);
            DataStream.WriteAtt("ph", r.Phonetic, false);

            DataStream.WriteAtt("thickTop", r.ThickTop, false);
            DataStream.WriteAtt("thickBot", r.ThickBottom, false);

            if (Row < Cells.CellList.Count)
            {
                WriteCells(DataStream, Row, Cells.CellList[Row], Cells.CellList, Dates1904);
            }

            DataStream.WriteFutureStorage(r.FutureStorage);
            DataStream.WriteEndElement();
        }

        private static void WriteCells(TOpenXmlWriter DataStream, int Row, TCellRecordList CellRecords, TCellList CellList, bool Dates1904)
        {
            for (int i = 0; i < CellRecords.Count; i++)
            {
                WriteCell(DataStream, Row, CellRecords[i], CellList, Dates1904);
            }
        }

        private static void WriteCell(TOpenXmlWriter DataStream, int Row, TCellRecord Cell, TCellList CellList, bool Dates1904)
        {
            DataStream.WriteStartElement("c");

            string CellRef = TCellAddress.EncodeColumn(Cell.Col +1) + Convert.ToString(Row +1, CultureInfo.InvariantCulture);
            DataStream.WriteAtt("r", CellRef);

            if (Cell.XF != FlxConsts.DefaultFormatId) DataStream.WriteAtt("s", Cell.XF);

            Cell.SaveToXlsx(DataStream, Row, CellList, Dates1904);
            
            DataStream.WriteFutureStorage(Cell.FutureStorage);
            DataStream.WriteEndElement();
        }

    }

    internal class TCustomXMLDataStorageList: IEnumerable<TCustomXMLDataStorage>
    {
        private List<TCustomXMLDataStorage> FList;

        public TCustomXMLDataStorageList()
        {
            FList = new List<TCustomXMLDataStorage>();
        }

        public void Clear()
        {
            FList.Clear();
        }

        public TCustomXMLDataStorage Add(TCustomXMLDataStorage data)
        {
            FList.Add(data);
            return data;
        }

        public int Count
        {
            get { return FList.Count; }
        }


        #region IEnumerable<TCustomXMLDataStorage> Members

        public IEnumerator<TCustomXMLDataStorage> GetEnumerator()
        {
            return FList.GetEnumerator();
        }

        #endregion

        #region IEnumerable Members

        IEnumerator IEnumerable.GetEnumerator()
        {
            return FList.GetEnumerator();
        }

        #endregion
    }

    internal class TCustomXMLDataStorage
    {
        internal string XML;
        internal Uri PartUri;
        internal List<TCustomXMLDataStorageProp> CustomXMLDataStorageProps;

        internal TCustomXMLDataStorage(Uri aPartUri, string aXml)
        {
            PartUri = aPartUri;
            XML = aXml;
            CustomXMLDataStorageProps = new List<TCustomXMLDataStorageProp>();
        }
    }

    internal class TCustomXMLDataStorageProp
    {
        internal string XML;
        internal Uri PartUri;

        internal TCustomXMLDataStorageProp(Uri aPartUri, string aXml)
        {
            PartUri = aPartUri;
            XML = aXml;
        }
    }

    

}
