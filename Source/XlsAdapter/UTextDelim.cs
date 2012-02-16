using System;
using System.IO;
using System.Text;
using FlexCel.Core;

namespace FlexCel.XlsAdapter
{

    internal class TextWithPos
    {
        TextReader Text;

        internal TextWithPos(TextReader aText)
        {
            Text = aText;
        }

        internal int GetNextChar()
        {
            return Text.Read();
        }

        internal string GetLine()
        {
            return Text.ReadLine();
        }
    }

    /// <summary>
    /// Helper class for reading/writing text delimited files.
    /// </summary>
    internal sealed class TextDelim
    {
        private TextDelim(){}

        #region Public interface.
        internal static void Read(TextReader InString, ExcelFile Workbook, char Delim, int FirstRow, int FirstCol,
            ColumnImportType[] ColumnFormats, string[] DateFormats)
        {
            if (Workbook.VirtualMode)
            {
                Workbook.OnVirtualCellStartReading(Workbook, new VirtualCellStartReadingEventArgs(Workbook));
            }
            int r=FirstRow;
            int c=FirstCol;
            TextWithPos sr=new TextWithPos(InString);
            StringBuilder s= new StringBuilder();
            {
                int ch = sr.GetNextChar();
                while (ch>=0)
                {
                    if (ch=='"') ReadQString(sr, s, ref ch);
                    else if (ch==Delim) 
                    {                                  
                        c++;
                        ch=sr.GetNextChar();
                        continue;
                    }
                    else if (ch==10) //there are 3 types of EOL: Win (13 10) Mac (10)  and Ms.Dos(13)
                    {
                        c=FirstCol;
                        r++;
                        ch=sr.GetNextChar();
                        continue;
                    } 
                    else if (ch==13) 
                    {
                        c = FirstCol;
                        r++;
                        ch = sr.GetNextChar();
                        if (ch == 10) ch = sr.GetNextChar();
                        continue;
                    }
                    else ReadNString(sr, Delim, s, ref ch);

                    if ((ColumnFormats != null) && (c - FirstCol < ColumnFormats.Length))
                    {
                        switch (ColumnFormats[c - FirstCol])
                        {
                            case ColumnImportType.Text: SetCellValue(Workbook, r, c, s.ToString()); break;
                            case ColumnImportType.Skip: break;
                            default: SetCellFromString(Workbook, r, c, s.ToString(), DateFormats); break;
                        } //case
                    }
                    else SetCellFromString(Workbook, r, c, s.ToString(), DateFormats);
                }
            }
        }

        internal static void SetCellValue(ExcelFile Workbook, int r, int c, string p)
        {
            if (Workbook.VirtualMode)
            {
                Workbook.OnVirtualCellRead(Workbook, new VirtualCellReadEventArgs(new CellValue(Workbook.ActiveSheet, r, c, p, -1)));
            }
            else
            {
                Workbook.SetCellValue(r, c, p);
            }   
        }

        internal static void SetCellFromString(ExcelFile Workbook, int r, int c, string p, string[] DateFormats)
        {
            if (Workbook.VirtualMode)
            {
                int XF = FlxConsts.DefaultFormatId;
                CellValue cv = new CellValue(Workbook.ActiveSheet, r, c, Workbook.ConvertString(new TRichString(p), ref XF, DateFormats), -1);
                if (cv.Value is DateTime) cv.Value = TExcelTypes.ConvertToAllowedObject(cv.Value, Workbook.OptionsDates1904);
                cv.XF = XF;
                Workbook.OnVirtualCellRead(Workbook, new VirtualCellReadEventArgs(cv));
            }
            else
            {
                Workbook.SetCellFromString(r, c, p, DateFormats);
            }
        }

        internal static void Write(TextWriter OutString, ExcelFile Workbook, char Delim, TXlsCellRange Range, bool ExportHiddenRowsOrColumns)
        {
            if (Range == null) Range = new TXlsCellRange(1, 1, Workbook.RowCount, Workbook.GetColCount(Workbook.ActiveSheet, false));
            for (int r=Range.Top; r<=Range.Bottom; r++)
            {
                if (!ExportHiddenRowsOrColumns && Workbook.GetRowHidden(r)) continue;
                for (int c = Range.Left; c <= Range.Right; c++)
                {
                    if (!ExportHiddenRowsOrColumns && Workbook.GetColHidden(c)) continue;
                    string s = Workbook.GetStringFromCell(r, c).ToString();
                    if ((s.IndexOf(Delim)>=0) || (s.IndexOf('"')>=0) || (s.IndexOf("\r")>0) || (s.IndexOf("\n")>0)) 
                    {
                        s=QuotedStr(s,"\"");
                    }
                    OutString.Write(s);
                    if (c<Range.Right) OutString.Write(Delim); 
                    else OutString.Write(TCompactFramework.NewLine);
                }
            }
        }

        internal static void Write(TextWriter OutString, ExcelFile Workbook, char Delim, bool ExportHiddenRowsOrColumns)
        {
            Write(OutString, Workbook, Delim, null, ExportHiddenRowsOrColumns);
        }

        #endregion

        private static string QuotedStr(string s, string quote)
        {
            return quote+s.Replace(quote, quote+quote)+quote;
        }


        private static void ReadQString(TextWithPos sr, StringBuilder s, ref int ch)
        {
            bool InQuote=false;
            s.Length=0;
            while (ch>=0)
            {
                ch=sr.GetNextChar();
				if (ch < 0) return;
                if ((ch!='"') && InQuote) return;
                if (InQuote || (ch!='"')) s.Append((char)ch);
                InQuote=(ch=='"') && !InQuote;
            }
        }
      
        private static void ReadNString(TextWithPos sr, char Delim, StringBuilder s, ref int ch)
        {
            s.Length=0;
            s.Append((char)ch);
            while (ch>=0)
            {
                ch=sr.GetNextChar();
                if ((ch==Delim)||(ch==13)||(ch==10)||(ch<0))  return;
                s.Append((char)ch);
            } //while
        }
    }

    /// <summary>
    /// Helper class for reading / writing fixed width text files.
    /// </summary>
    internal sealed class TextFixedWidth
    {
        private TextFixedWidth() { }

        #region Public interface.
        internal static void Read(TextReader InString, ExcelFile Workbook, int[] ColumnWidths, int FirstRow, int FirstCol, ColumnImportType[] ColumnFormats, string[] DateFormats)
        {
            if (Workbook.VirtualMode)
            {
                Workbook.OnVirtualCellStartReading(Workbook, new VirtualCellStartReadingEventArgs(Workbook));
            }

            int r = FirstRow - 1;
            TextWithPos sr = new TextWithPos(InString);
            {
                while (true)
                {
                    string line = sr.GetLine();
                    if (line == null) return;
                    r++;

                    int c = FirstCol;
                    int FirstChar = 0;
                    for (int i = 0; i < ColumnWidths.Length; i++)
                    {
                        if (line.Length <= FirstChar) break;
                        int cw = Math.Min(ColumnWidths[i], line.Length - FirstChar);
                        string s = line.Substring(FirstChar, cw);
                        FirstChar += cw;

                        if ((ColumnFormats != null) && (i < ColumnFormats.Length))
                        {
                            switch (ColumnFormats[i])
                            {
                                case ColumnImportType.Text: TextDelim.SetCellValue(Workbook, r, c, s.ToString()); c++; break;
                                case ColumnImportType.Skip: break;
                                default: TextDelim.SetCellFromString(Workbook, r, c, s.ToString(), DateFormats); c++; break;
                            } //case
                        }
                        else
                        {
                            TextDelim.SetCellFromString(Workbook, r, c, s.ToString(), DateFormats);
                            c++;
                        }

                    }
                }
            }
        }

        internal static void Write(TextWriter OutString, ExcelFile Workbook, TXlsCellRange Range, 
            int[] ColumnWidths, int CharactersForFirstColumn, bool ExportHiddenRowsOrColumns, bool ExportTextOutsideCells)
        {
            if (Range == null) Range = new TXlsCellRange(1, 1, Workbook.RowCount, Workbook.GetColCount(Workbook.ActiveSheet, false));

            for (int r = Range.Top; r <= Range.Bottom; r++)
            {
                if (!ExportHiddenRowsOrColumns && Workbook.GetRowHidden(r)) continue;

                int cIndex = 0;
                double FirstColWidth = Workbook.GetColWidth(Range.Left);
                string Remaining = string.Empty;

                int AcumColLen = 0;
                bool InMergedCell = false;
                THFlxAlignment MergedAlign = THFlxAlignment.general;
                TCellType MergedType = TCellType.Unknown;
                int OrigColLen = 0;

                for (int c = Range.Left; c <= Range.Right; c++)
                {
                    if (!ExportHiddenRowsOrColumns && Workbook.GetColHidden(c)) continue;
                    string s = Workbook.GetStringFromCell(r, c).ToString();
                    TFlxFormat fmt = null;
                    if (ExportTextOutsideCells) fmt = Workbook.GetCellVisibleFormatDef(r, c); 

                    if (string.IsNullOrEmpty(s)) s = Remaining;

                    int ColLen = 0;
                    if (ColumnWidths == null)
                    {
                        if (CharactersForFirstColumn <= 0) ColLen = s.Length;
                        else if (FirstColWidth <= 0) ColLen = 0;
                        else ColLen = (int)Math.Round((double)CharactersForFirstColumn * Workbook.GetColWidth(c) / FirstColWidth);
                    }
                    else
                    {
                        if (cIndex >= ColumnWidths.Length) break;
                        ColLen = ColumnWidths[cIndex];
                    }
                    
                    cIndex++;
                    if (InMergedCell) OrigColLen += ColLen; else OrigColLen = ColLen;

                    if (s.Length == 0)
                    {
                        AcumColLen += ColLen;
                        continue;
                    }

                    THFlxAlignment HAlign = THFlxAlignment.left;
                    TCellType CellType;
                    if (InMergedCell)
                    {
                        HAlign = MergedAlign;
                        CellType = MergedType;
                    }
                    else
                    {
                        Object CellVal = Workbook.GetCellValue(r, c);
                        TFormula fmla = CellVal as TFormula;
                        if (fmla != null) CellVal = fmla.Result;
                        CellType = TExcelTypes.ObjectToCellType(CellVal);
                        if (ExportTextOutsideCells && fmt != null && Remaining.Length == 0)
                        {
                            HAlign = GetDataAlign(CellType, fmt);
                        }
                    }

                    if (HAlign == THFlxAlignment.left)
                    {
                        OutString.Write(new string(' ', AcumColLen));
                    }
                    else
                    {
                        TXlsCellRange mr = Workbook.CellMergedBounds(r, c);
                        InMergedCell = mr.Right > c;
                        if (mr.Right > c)
                        {
                            AcumColLen += ColLen;
                            if (c == mr.Left)
                            {
                                Remaining = s;
                                MergedAlign = HAlign;
                                MergedType = CellType;
                            }
                            continue;
                        }
                        if (mr.Right > mr.Left)
                        {
                            s = Remaining;
                            Remaining = string.Empty;
                        }

                        MergedAlign = THFlxAlignment.left;
                        MergedType = TCellType.Unknown;
                        InMergedCell = false;
                        ColLen += AcumColLen;
                    }
                    AcumColLen = 0;

                    if (s.Length > ColLen)
                    {
                        if (ExportTextOutsideCells && HAlign == THFlxAlignment.right)
                        {
                            if (CellType == TCellType.Number) OutString.Write(new string('#', ColLen));
                            else OutString.Write(s.Substring(s.Length - ColLen));
                        }
                        else
                        {
                            OutString.Write(s.Substring(0, ColLen));
                        }
                        if (ExportTextOutsideCells && HAlign != THFlxAlignment.right) Remaining = s.Substring(ColLen);
                    }
                    else
                    {
                        int Pad = ColLen - s.Length;
                        if (ExportTextOutsideCells && Remaining.Length == 0)
                        {
                            Pad = TextAlign(OutString, HAlign, s.Length, ColLen, OrigColLen);
                        }
                        OutString.Write(s);
                        OutString.Write(new string(' ', Pad));
                        Remaining = string.Empty;
                    }
                }

                if (ExportTextOutsideCells && Remaining.Length > 0) OutString.Write(Remaining);
                Remaining = string.Empty;
                OutString.Write(TCompactFramework.NewLine);
            }
        }

        private static int TextAlign(TextWriter OutString, THFlxAlignment HAlign, int SLen, int ColLen, int OrigColLen)
        {
            int Pad = ColLen - SLen;

            switch (HAlign)
            {
                case THFlxAlignment.center:
                    int P2 = (OrigColLen - SLen) / 2;
                    if (P2 < 0)
                    {
                        OutString.Write(new string(' ', Pad));
                        return 0;
                    }

                    OutString.Write(new string(' ', ColLen - OrigColLen + P2));
                    return OrigColLen - P2 - SLen;

                case THFlxAlignment.right:
                    OutString.Write(new string(' ', Pad));
                    return 0;

            }
            return Pad;
        }

        public static THFlxAlignment GetDataAlign(TCellType CellType, TFlxFormat Fm)
        {
            switch (Fm.HAlignment)
            {
                case THFlxAlignment.general:
                    return GetGeneralAlign(CellType);

                case THFlxAlignment.center_across_selection:
                case THFlxAlignment.center:
                    return THFlxAlignment.center;

                case THFlxAlignment.right:
                    return THFlxAlignment.right;

            }
            return THFlxAlignment.left;
        }

        private static THFlxAlignment GetGeneralAlign(TCellType CellType)
        {
            switch (CellType)
            {
                case TCellType.Bool:
                case TCellType.Error:
                    return THFlxAlignment.center;

                case TCellType.String:
                case TCellType.Unknown:
                    return THFlxAlignment.left;

                case TCellType.Number:
                    return THFlxAlignment.right;

                default:
                    return THFlxAlignment.right;
            }
        }


        #endregion

    }
}
