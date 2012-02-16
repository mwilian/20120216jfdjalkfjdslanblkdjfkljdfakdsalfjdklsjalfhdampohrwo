using System;
using System.Text;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Data;
using System.Text.RegularExpressions;
using FlexCel.Core;

using System.Collections.Generic;

namespace FlexCel.Report
{
    /// <summary>
    /// All the things we can find on a cell.
    /// </summary>
    public enum TValueType
    {
        /// <summary>
        /// A constant value, like 2 or "a".
        /// </summary>
        /// <example>On the string: "&lt;#tag1&gt;hello&lt;#tag1&gt;" "Hello is a Const type.</example>
        Const, 
        /// <summary>
        /// A value from a Dataset. This tag is written as "&lt;#DataTable.Column&gt;
        /// </summary>
        /// <example>If you have a table "Customers" with a column named "LastName", &lt;#Customers.LastName&gt; will replace the value on the report.</example>
        DataSet, 
        /// <summary>
        /// A report variable, like &lt;#Variable&gt;
        /// </summary>
        Property,
        /// <summary>
        /// A dataset with "*"
        /// </summary>
        FullDataSet,
        /// <summary>
        /// Captions for a full dataset.
        /// </summary>
        FullDataSetCaptions,
        /// <summary>
        /// An #IF construct
        /// </summary>
        IF,
        /// <summary>
        /// An include Tag.
        /// </summary>
        Include,
        /// <summary>
        /// A #= tag
        /// </summary>
        Equal,
        /// <summary>
        /// A #Delete sheet
        /// </summary>
        DeleteSheet,
        /// <summary>
        /// #Config
        /// </summary>
        ConfigSheet,

        /// <summary>
        /// #Delete Range
        /// </summary>
        DeleteRange,

        /// <summary>
        /// #Delete row
        /// </summary>
        DeleteRow,

        /// <summary>
        /// Delete Column
        /// </summary>
        DeleteCol,

        /// <summary>
        /// Format cell
        /// </summary>
        FormatCell,

        /// <summary>
        /// Format row
        /// </summary>
        FormatRow,

        /// <summary>
        /// Format column
        /// </summary>
        FormatCol,

        /// <summary>
        /// Format range
        /// </summary>
        FormatRange,

        /// <summary>
        /// Horizontal page break
        /// </summary>
        HPageBreak,

        /// <summary>
        /// Vertical page break
        /// </summary>
        VPageBreak,

        /// <summary>
        /// This one will not really be stored.
        /// </summary>
        Comment,

        /// <summary>
        /// An expression that will be evaluated
        /// </summary>
        Evaluate,

        /// <summary>
        /// Only used inside images, to specify its zoom in % and its aspect ratio.
        /// </summary>
        ImgSize,

        /// <summary>
        /// Only used inside images, to specify its position inside the cell.
        /// </summary>
        ImgPos,

        /// <summary>
        /// Only used inside images, to specify how to modify the containing cells to hold the image.
        /// </summary>
        ImgFit,

        /// <summary>
        /// Only used inside images, to delete an image.
        /// </summary>
        ImgDelete,

        /// <summary>
        /// A lookup field. Parameters are: DataBase, KeyFields, KeyValues, ResultField.
        /// </summary>
        Lookup,

        /// <summary>
        /// An array of values. Useful for example as the KeyValues argument of the Lookup field.
        /// </summary>
        Array,

        /// <summary>
        /// A regular expression replace.
        /// </summary>
        Regex,

        /// <summary>
        /// A cell that will be merged.
        /// </summary>
        MergeRange,

        /// <summary>
        /// The contents of this cell should be entered as formula.
        /// </summary>
        Formula,

        /// <summary>
        /// Adjusts the column width to a specified value.
        /// </summary>
        ColumnWidth,

        /// <summary>
        /// Adjusts the row height to a specified value.
        /// </summary>
        RowHeight,

        /// <summary>
        /// Defines if the cell has HTML formatted data or not.
        /// </summary>
        Html,

        /// <summary>
        /// The parameter is considered a reference and will be changed when copying the ranges.
        /// </summary>
        Ref,

        /// <summary>
        /// Sets autofit settings for the whole sheet.
        /// </summary>
        AutofitSettings,

        /// <summary>
        /// Returns true if the expression is defined.
        /// </summary>
        Defined,

        /// <summary>
        /// Returns true if the format is defined.
        /// </summary>
        DefinedFormat,

        /// <summary>
        /// An expression that will be pre evaluated and used to modify the template.
        /// </summary>
        Preprocess,

        /// <summary>
        /// FlexCel should automatically calculate the page breaks in this sheet.
        /// </summary>
        AutoPageBreaks,

        /// <summary>
        /// Does an operation like sum or average over a dataset.
        /// </summary>
        Aggregate,

        /// <summary>
        /// Concatenates all values in a dataset as a list.
        /// </summary>
        List,

        /// <summary>
        /// Gets the value of a dataset given a row or a column.
        /// </summary>
        DbValue
    }


    /// <summary>
    /// Enumeration with the different kind of aggratations that can be done in a FlexCelReport.
    /// </summary>
    public enum TAggregateType
    {
        /// <summary>
        /// Sum all values.
        /// </summary>
        Sum,

        /// <summary>
        /// Average all values.
        /// </summary>
        Average,

        /// <summary>
        /// Max of all values.
        /// </summary>
        Max,

        /// <summary>
        /// Minimum of all values.
        /// </summary>
        Min
    }

    internal class TValueAndXF
    {
        internal int XF;
        internal TConfigFormat XFRow;
        internal TConfigFormat XFCol;
        internal object Value;

        internal TValueType Action;

        internal int FullDataSetColumnIndex;
        internal int FullDataSetColumnCount;
        internal TImageSizeParams ImageSize;
        internal TImagePosParams ImagePos;
        internal TImageFitParams ImageFit;
        internal bool ImageDelete;

        internal bool IsFormula;
        internal TIncludeHtml IncludeHtml;

        internal TAutofitInfo AutofitInfo;

        internal TDebugStack DebugStack;

        internal int AutoPageBreaksPercent;
        internal int AutoPageBreaksPageScale;

        internal ExcelFile Workbook; //Won't be initialized here. It is supplied externally, only when it is allowed to modify it.
        internal TWaitingRangeList WaitingRanges; //Won't be initialized here. It is supplied externally.
        internal TFormatRangeList FormatRangeList; //Won't be initialized here. It is supplied externally.
        internal TFormatRangeList FormatCellList; //Won't be initialized here. It is supplied externally.

        internal TValueAndXF()
        {
            Clear();
            WaitingRanges = null;
            FormatRangeList=null;
            FormatCellList=null;
            Workbook = null;
        }

        internal TValueAndXF(TDebugStack aDebugStack) : this()
        {
            DebugStack = aDebugStack;
        }

        internal void Clear()
        {
            AutoPageBreaksPercent = -1;
            AutoPageBreaksPageScale = -1;
            XF=-1;
            XFRow=null;
            XFCol=null;
            Value=null;
            Action=TValueType.Comment;
            FullDataSetColumnIndex=0;
            FullDataSetColumnCount=0;
            ImageSize=null;
            IsFormula = false;
            IncludeHtml = TIncludeHtml.Undefined;
        }

    }
    
    #region Section Values
    /// <summary>
    /// On a cell, we can have many Sections containing datasets, properties, etc. (Any <see cref="TValueType"/>)
    /// TOneSectionValue represents the minimum section we can find.
    /// </summary>
    internal class TOneSectionValue: IDisposable
    {

        internal TValueType ValueType;
        internal int StartFont;
        internal string FTagText;

        internal TOneSectionValue(string aTagText, TValueType aValueType, int aStartFont)
        {
            FTagText = aTagText;
            ValueType=aValueType;
            StartFont=aStartFont;
        }

        private TDebugItem AddDebugInfo(TValueAndXF val)
        {
            TDebugItem Result = null;
            if (val.DebugStack != null) 
            {
                Result = val.DebugStack.Add(FTagText, val.Value);
                val.DebugStack.IncLevel();
            }
            return Result;
        }

        internal void ResetVal(TValueAndXF val)
        {
            val.Value= null;
            val.Action=ValueType;
        }

        protected virtual void EvaluateInternal(int RowAbs, int ColAbs, int RowOfs, int ColOfs, TValueAndXF val)
        {
            ResetVal(val);
        }

        internal void Evaluate(int RowAbs, int ColAbs, int RowOfs, int ColOfs, TValueAndXF val)
        {
            TDebugItem dbg = AddDebugInfo(val);
            EvaluateInternal(RowAbs, ColAbs, RowOfs, ColOfs, val);
            if (dbg != null) 
            {
                dbg.Value = val.Value;
                val.DebugStack.DecLevel();
            }
        }

        /// <summary>
        /// If the section has a reference to other sections, return the children.
        /// </summary>
        /// <returns>Sections referenced by this section.</returns>
        internal virtual TOneCellValue Resolve(int RowAbs, int ColAbs, int RowOfs, int ColOfs, TDebugStack aDebugStack, int FullDataSetColumnIndex)
        {
            return null;
        }


        internal virtual int RecordCount()
        {
            return 1;
        }

        internal virtual TBand DataBand()
        {
            return null;
        }
        #region IDisposable Members

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
        }

        #endregion
    }

    internal class TSectionConst: TOneSectionValue
    {
        //Constant value. If it is another thing, the text used.
        internal object Value;

        internal TSectionConst(TOneCellValue aParent, object aValue): base("Constant", TValueType.Const, -1)
        {
            Value=aValue;
        }

        protected override void EvaluateInternal(int RowAbs, int ColAbs, int RowOfs, int ColOfs, TValueAndXF val)
        {
            val.Value = Value;
            val.Action=ValueType;
        }
    }

    internal class TSectionFormatRange: TOneSectionValue
    {
        internal TFormatRange FormatRange;

        internal TSectionFormatRange(string aTagText, TOneCellValue aParent, TFormatRange aFormatRange): base(aTagText, TValueType.FormatRange, -1)
        {
            FormatRange=aFormatRange;
        }
        
        protected override void EvaluateInternal(int RowAbs, int ColAbs, int RowOfs, int ColOfs, TValueAndXF val)
        {
            if (val.FormatRangeList!=null) val.FormatRangeList.Add(FormatRange);
            base.EvaluateInternal(RowAbs, ColAbs, RowOfs, ColOfs, val);
        }    
    }

    internal class TSectionFormatCell: TOneSectionValue
    {
        internal TFormatRange FormatCell;

        internal TSectionFormatCell(string aTagText, TOneCellValue aParent, TFormatRange aFormatCell): base(aTagText, TValueType.FormatCell, -1)
        {
            FormatCell=aFormatCell;
        }

        internal static TSectionFormatCell Create( 
                string aTagText, TOneCellValue aParent, TFormatRange aFormatRange)     
                
        {
            return (new TSectionFormatCell(aTagText, aParent, aFormatRange));
        }
        
        protected override void EvaluateInternal(int RowAbs, int ColAbs, int RowOfs, int ColOfs, TValueAndXF val)
        {
            if (val.FormatCellList!=null) val.FormatCellList.Add(FormatCell);
            base.EvaluateInternal(RowAbs, ColAbs, RowOfs, ColOfs, val);
        }    
    }

    internal class TSectionImgSize: TOneSectionValue
    {
        internal TImageSizeParams ImageSizeParams;
        private TOneCellValue ZoomStr;

        private TSectionImgSize(string aTagText, TOneCellValue aParent, double aZoom, double aAspectRatio, bool aBoundImage, TOneCellValue aZoomStr): base(aTagText, TValueType.ImgSize, -1)
        {
            ImageSizeParams= new TImageSizeParams(aZoom, aAspectRatio, aBoundImage);
            ZoomStr = aZoomStr;
        }

        internal static TSectionImgSize Create(ExcelFile Workbook, TStackData Stack, TBand aCurrentBand, FlexCelReport fr,
            TOneCellValue aParent, TRichString TagText, TRichString TagParams)
        {
            bool BoundImage = false;
            TOneCellValue ZoomStr = null;
            double Zoom = 0;
            double AspectRatio = 0;


            if (TagParams == null || TagParams.ToString().Trim().Length == 0)
            {
                BoundImage = true;
            }
            else
            {
                TRichString[] IsSections= new TRichString[2];
                TCellParser.ParseParams(TagText, TagParams, IsSections);

                int iZoom;
                if (TCompactFramework.TryParse(IsSections[0].ToString(), out iZoom))
                {
                    Zoom = iZoom;
                }
                else
                {
                    int XF5 = -1;
                    ZoomStr = 
                        TCellParser.GetCellValue(TCellParser.TryConvert(Workbook, IsSections[0], ref XF5), Workbook, Stack, XF5, aCurrentBand, fr);
                }

                AspectRatio = Convert.ToDouble(IsSections[1].ToString());

            }
            return new TSectionImgSize(TagText.ToString(), aParent, Zoom, AspectRatio, BoundImage, ZoomStr);


        }

        protected override void EvaluateInternal(int RowAbs, int ColAbs, int RowOfs, int ColOfs, TValueAndXF val)
        {
            if (ZoomStr != null)
            {
                TValueAndXF v1 = new TValueAndXF();
                ZoomStr.Evaluate(RowAbs, ColAbs, RowOfs, ColOfs, v1);
                val.ImageSize = new TImageSizeParams(Convert.ToInt32(v1.Value), ImageSizeParams.AspectRatio, ImageSizeParams.BoundImage);
            }
            else
            {
                val.ImageSize=ImageSizeParams;
            }
            base.EvaluateInternal(RowAbs, ColAbs, RowOfs, ColOfs, val);
        }    
    }

    internal class TSectionImgPos: TOneSectionValue
    {
        internal TImagePosParams ImagePosParams;

        private TSectionImgPos(string aTagText, TOneCellValue aParent, TImageVAlign aRowAlign, TImageHAlign aColAlign, TOneCellValue aRowOffs, TOneCellValue aColOffs): base(aTagText, TValueType.ImgPos, -1)
        {
            ImagePosParams= new TImagePosParams(aRowAlign, aColAlign, aRowOffs, aColOffs);
        }

        internal static TSectionImgPos Create(ExcelFile Workbook, TStackData Stack, TBand aCurrentBand, FlexCelReport fr, TOneCellValue aParent, TRichString TagText, TRichString TagParams)
        {
            TImageVAlign aRowAlign = TImageVAlign.None;
            TImageHAlign aColAlign = TImageHAlign.None;
            TOneCellValue aRowOffs = null; TOneCellValue aColOffs = null;

            TRichString[] Sections= new TRichString[4];
            TCellParser.ParseParams(TagText, TagParams, Sections, true);
            if (Sections[0] != null) aRowAlign = GetRowAlign(Sections[0].ToString());
            if (Sections[1] != null) aColAlign = GetColAlign(Sections[1].ToString());

            int XFR = -1; int XFC = -1;
            if (Sections[2] != null) aRowOffs = TCellParser.GetCellValue(TCellParser.TryConvert(Workbook, Sections[2], ref XFR), Workbook, Stack, XFR, aCurrentBand, fr);
            if (Sections[3] != null) aColOffs = TCellParser.GetCellValue(TCellParser.TryConvert(Workbook, Sections[3], ref XFC), Workbook, Stack, XFC, aCurrentBand, fr);

            return new TSectionImgPos(TagText.ToString(), aParent, aRowAlign, aColAlign, aRowOffs, aColOffs );
        }

        private static TImageVAlign GetRowAlign(string s)
        {
            if (s== null || s.Trim().Length == 0) return TImageVAlign.None;
            if (String.Equals(s, ReportTag.StrAlignTop, StringComparison.InvariantCultureIgnoreCase)) return TImageVAlign.Top;
            if (String.Equals(s, ReportTag.StrAlignCenter, StringComparison.InvariantCultureIgnoreCase)) return TImageVAlign.Center;
            if (String.Equals(s, ReportTag.StrAlignBottom, StringComparison.InvariantCultureIgnoreCase)) return TImageVAlign.Bottom;

            FlxMessages.ThrowException(FlxErr.ErrInvalidImgPosParameter, s, ReportTag.StrAlignTop, ReportTag.StrAlignCenter, ReportTag.StrAlignBottom);
            return TImageVAlign.None; //just to compile.
        }

        private static TImageHAlign GetColAlign(string s)
        {
            if (s== null || s.Trim().Length == 0) return TImageHAlign.None;
            if (String.Equals(s, ReportTag.StrAlignLeft, StringComparison.InvariantCultureIgnoreCase)) return TImageHAlign.Left;
            if (String.Equals(s, ReportTag.StrAlignCenter, StringComparison.InvariantCultureIgnoreCase)) return TImageHAlign.Center;
            if (String.Equals(s, ReportTag.StrAlignRight, StringComparison.InvariantCultureIgnoreCase)) return TImageHAlign.Right;

            FlxMessages.ThrowException(FlxErr.ErrInvalidImgPosParameter, s, ReportTag.StrAlignLeft, ReportTag.StrAlignCenter, ReportTag.StrAlignRight);
            return TImageHAlign.None; //just to compile.
        }


        protected override void EvaluateInternal(int RowAbs, int ColAbs, int RowOfs, int ColOfs, TValueAndXF val)
        {
            val.ImagePos=ImagePosParams;
            base.EvaluateInternal(RowAbs, ColAbs, RowOfs, ColOfs, val);
        }    
    }

    internal class TSectionImgFit: TOneSectionValue
    {
        internal TImageFitParams ImageFitParams;

        private TSectionImgFit(string aTagText, TOneCellValue aParent, TAutofitGrow aFitInRows, TAutofitGrow aFitInCols, TOneCellValue aRowMargin, TOneCellValue aColMargin): base(aTagText, TValueType.ImgFit, -1)
        {
            ImageFitParams= new TImageFitParams(aFitInRows, aFitInCols, aRowMargin, aColMargin);
        }

        internal static TSectionImgFit Create(ExcelFile Workbook, TStackData Stack, TBand aCurrentBand, FlexCelReport fr, TOneCellValue aParent, TRichString TagText, TRichString TagParams)
        {
            TAutofitGrow aFitInRows = TAutofitGrow.None;
            TAutofitGrow aFitInCols = TAutofitGrow.None;
            TOneCellValue aRowMargin = null; TOneCellValue aColMargin = null;

            TRichString[] Sections= new TRichString[4];
            TCellParser.ParseParams(TagText, TagParams, Sections, true);
            if (Sections[0] != null) aFitInRows = GetFitInRow(Sections[0].ToString());
            if (Sections[1] != null) aFitInCols = GetFitInCol(Sections[1].ToString());
            int XFR = -1; int XFC = -1;
            if (Sections[2] != null) aRowMargin = TCellParser.GetCellValue(TCellParser.TryConvert(Workbook, Sections[2], ref XFR), Workbook, Stack, XFR, aCurrentBand, fr);
            if (Sections[3] != null) aColMargin = TCellParser.GetCellValue(TCellParser.TryConvert(Workbook, Sections[3], ref XFC), Workbook, Stack, XFC, aCurrentBand, fr);

            return new TSectionImgFit(TagText.ToString(), aParent, aFitInRows, aFitInCols, aRowMargin, aColMargin);
        }

        private static TAutofitGrow GetFitInRow(string s)
        {
            if (s== null || s.Trim().Length == 0) return TAutofitGrow.None;
            if (String.Equals(s, ReportTag.StrInRow, StringComparison.InvariantCultureIgnoreCase)) return TAutofitGrow.All;
            if (String.Equals(s, ReportTag.StrDontGrow, StringComparison.InvariantCultureIgnoreCase)) return TAutofitGrow.DontGrow;
            if (String.Equals(s, ReportTag.StrDontShrink, StringComparison.InvariantCultureIgnoreCase)) return TAutofitGrow.DontShrink;

            FlxMessages.ThrowException(FlxErr.ErrInvalidImgFitParameter, s, ReportTag.StrInRow);
            return TAutofitGrow.None; //just to compile.
        }

        private static TAutofitGrow GetFitInCol(string s)
        {
            if (s== null || s.Trim().Length == 0) return TAutofitGrow.None;
            if (String.Equals(s, ReportTag.StrInCol, StringComparison.InvariantCultureIgnoreCase)) return TAutofitGrow.All;
            if (String.Equals(s, ReportTag.StrDontGrow, StringComparison.InvariantCultureIgnoreCase)) return TAutofitGrow.DontGrow;
            if (String.Equals(s, ReportTag.StrDontShrink, StringComparison.InvariantCultureIgnoreCase)) return TAutofitGrow.DontShrink;

            FlxMessages.ThrowException(FlxErr.ErrInvalidImgFitParameter, s, ReportTag.StrInCol);
            return TAutofitGrow.None; //just to compile.
        }


        protected override void EvaluateInternal(int RowAbs, int ColAbs, int RowOfs, int ColOfs, TValueAndXF val)
        {
            val.ImageFit=ImageFitParams;
            base.EvaluateInternal(RowAbs, ColAbs, RowOfs, ColOfs, val);
        }    
    }


    internal class TSectionImgDelete: TOneSectionValue
    {
        internal TSectionImgDelete(string aTagText, TOneCellValue aParent): base(aTagText, TValueType.ImgDelete, -1)
        {
        }

        protected override void EvaluateInternal(int RowAbs, int ColAbs, int RowOfs, int ColOfs, TValueAndXF val)
        {
            val.ImageDelete=true;
            base.EvaluateInternal(RowAbs, ColAbs, RowOfs, ColOfs, val);
        }    
    }


    internal class TSectionDeleteRow: TOneSectionValue
    {
        private TDeleteRowWaitingRange DeleteRow;
        internal TSectionDeleteRow(string aTagText, TOneCellValue aParent, TDeleteRowWaitingRange aDeleteRow): base(aTagText, TValueType.DeleteRow, -1)
        {
            DeleteRow = aDeleteRow;
        }
        
        protected override void EvaluateInternal(int RowAbs, int ColAbs, int RowOfs, int ColOfs, TValueAndXF val)
        {
            if (val.WaitingRanges != null)
            {
                TDeleteRowWaitingRange NewDeleteRow = new TDeleteRowWaitingRange(RowAbs + 1, DeleteRow.Left, DeleteRow.LastCol);
                val.WaitingRanges.Add(NewDeleteRow);
            }
            base.EvaluateInternal(RowAbs, ColAbs, RowOfs, ColOfs, val);
        }    
    }

    internal class TSectionDeleteCol: TOneSectionValue
    {
        private TDeleteColWaitingRange DeleteCol;

        internal TSectionDeleteCol(string aTagText, TOneCellValue aParent, TDeleteColWaitingRange aDeleteCol): base(aTagText, TValueType.DeleteCol, -1)
        {
            DeleteCol = aDeleteCol;
        }
        
        protected override void EvaluateInternal(int RowAbs, int ColAbs, int RowOfs, int ColOfs, TValueAndXF val)
        {
            if (val.WaitingRanges != null)
            {
                TDeleteColWaitingRange NewDeleteCol = new TDeleteColWaitingRange(ColAbs  + 1, DeleteCol.Top, DeleteCol.LastRow);
                val.WaitingRanges.Add(NewDeleteCol);
            }
            base.EvaluateInternal(RowAbs, ColAbs, RowOfs, ColOfs, val);
        }    
    }

    internal class TSectionDeleteRange: TOneSectionValue
    {
        internal TDeleteRangeWaitingRange DeleteRange;

        internal TSectionDeleteRange(string aTagText, TOneCellValue aParent, TDeleteRangeWaitingRange aDeleteRange): base(aTagText, TValueType.DeleteRange, -1)
        {
            DeleteRange=aDeleteRange;
        }
        
        protected override void EvaluateInternal(int RowAbs, int ColAbs, int RowOfs, int ColOfs, TValueAndXF val)
        {
            if (val.WaitingRanges != null)
                val.WaitingRanges.Add(DeleteRange);
            base.EvaluateInternal(RowAbs, ColAbs, RowOfs, ColOfs, val);
        }    
    }

    internal enum TRowColReportAction
    {
        Show,
        Hide,
        Autofit,
        Size
    }

    internal abstract class TSectionSize : TOneSectionValue
    {
        protected double RowColSize;
        protected int AdjustmentFixed;
        protected TAutofitGrow AutofitGrow;
        protected double MinSize;
        protected double MaxSize;
        protected TRowColReportAction RowColReportAction;
        protected TXlsCellRange MergedRange;

        internal TSectionSize(string aTagText, TOneCellValue aParent, TValueType ValueType, TRichString[] Params): base(aTagText, ValueType, -1)
        {
            double Size = 0;

            string Param = FlxConvert.ToString(Params[0]).Trim();
            if (String.Equals(Param, ReportTag.StrShow, StringComparison.InvariantCultureIgnoreCase)) RowColReportAction = TRowColReportAction.Show;
            else
                if (String.Equals(Param, ReportTag.StrHide, StringComparison.InvariantCultureIgnoreCase)) RowColReportAction = TRowColReportAction.Hide;
            else
                if (String.Equals(Param, ReportTag.StrAutofit, StringComparison.InvariantCultureIgnoreCase)) 
            {
                RowColReportAction = TRowColReportAction.Autofit;
                
                string Param2 = FlxConvert.ToString(Params[1]).Trim();
                if (Param2 != null && Param2.Length > 0) 
                {
                    if (TCompactFramework.ConvertToNumber(Param2, CultureInfo.InvariantCulture, out Size))
                    {
                        RowColSize = Size / 100;
                    }
                    else FlxMessages.ThrowException(FlxErr.ErrInvalidRowColParameters2, Param2);
                }
                else RowColSize = 0;

                string Param3 = FlxConvert.ToString(Params[2]).Trim();
                if (Param3 != null && Param3.Length > 0) 
                {
                    if (TCompactFramework.ConvertToNumber(Param3, CultureInfo.InvariantCulture, out Size))
                    {
                        AdjustmentFixed = Convert.ToInt32(Size);
                    }
                    else FlxMessages.ThrowException(FlxErr.ErrInvalidRowColParameters3, Param3);
                }
                else AdjustmentFixed = 0;

                AutofitGrow = TAutofitGrow.All;
                GetAutoFitGrow(FlxConvert.ToString(Params[2]).Trim(), ref AutofitGrow, out MinSize);
                GetAutoFitGrow(FlxConvert.ToString(Params[3]).Trim(), ref AutofitGrow, out MaxSize);

            }
            else
                    if (TCompactFramework.ConvertToNumber(Param, CultureInfo.InvariantCulture, out Size))
            {
                RowColReportAction = TRowColReportAction.Size;
                RowColSize = Size; 
            }
            else FlxMessages.ThrowException(FlxErr.ErrInvalidRowColParameters, Param, ReportTag.StrShow, ReportTag.StrHide, ReportTag.StrAutofit);

        }

        private static void GetAutoFitGrow(string Param, ref TAutofitGrow AutofitGrow, out double Height)
        {
            Height = 0;

            if (String.Equals(Param, ReportTag.StrDontShrink, StringComparison.InvariantCultureIgnoreCase)) AutofitGrow = TAutofitGrow.DontShrink;
            else if (String.Equals(Param, ReportTag.StrDontGrow, StringComparison.InvariantCultureIgnoreCase)) AutofitGrow = TAutofitGrow.DontGrow;

            else if (Param != null && Param.Length > 0)  
            {
                double Size = 0;
                if (TCompactFramework.ConvertToNumber(Param, CultureInfo.InvariantCulture, out Size))
                {
                    Height = Size;
                }
                else FlxMessages.ThrowException(FlxErr.ErrInvalidRowColParameters4, Param, ReportTag.StrDontShrink, ReportTag.StrDontGrow);
            }
        }

        protected void CalcAutofitGrow(ExcelFile xls)
        {
            switch (AutofitGrow)
            {
                case TAutofitGrow.DontGrow:
                    MaxSize = -1;
                    break;
                case TAutofitGrow.DontShrink:
                    MinSize = -1;
                    break;
            }

        }
    
    }

    internal class TSectionRowHeight: TSectionSize
    {
        internal TSectionRowHeight(string aTagText, TOneCellValue aParent, TRichString[] Params): base(aTagText, aParent, TValueType.RowHeight, Params)
        {
            if (RowColReportAction == TRowColReportAction.Autofit) //for others cases we don't care about merged.
            {

            }
        }

    
        protected override void EvaluateInternal(int RowAbs, int ColAbs, int RowOfs, int ColOfs, TValueAndXF val)
        {
            if (val.Workbook != null)
            {
                switch (RowColReportAction)
                {
                    case TRowColReportAction.Show:
                        val.Workbook.SetRowHidden(RowAbs + 1 + RowOfs, false);
                        break;
                    case TRowColReportAction.Hide:
                        val.Workbook.SetRowHidden(RowAbs + 1 + RowOfs, true);
                        break;
                    case TRowColReportAction.Autofit:
                        CalcAutofitGrow(val.Workbook);
                        
                        if (MergedRange == null) MergedRange = val.Workbook.CellMergedBounds(RowAbs + 1, ColAbs + 1);
                        bool IsMerged = MergedRange.RowCount > 1;
                        val.Workbook.MarkRowForAutofit(RowAbs + 1 + RowOfs, true, (float)RowColSize, AdjustmentFixed, (int)Math.Round(MinSize * 20), (int)Math.Round(MaxSize * 20), IsMerged); //Should be delayed until all data has been filled.
                        if (val.AutofitInfo != null)
                        {
                            if (val.AutofitInfo.AutofitType == TAutofitType.None) val.AutofitInfo.AutofitType = TAutofitType.OnlyMarked;
                        }
                        break;
                    default:
                        val.Workbook.SetRowHeight(RowAbs + 1 + RowOfs, (int)Math.Round(RowColSize * 20));
                        break;
                }
            }
            base.EvaluateInternal(RowAbs, ColAbs, RowOfs, ColOfs, val);
        }    
    }

    internal class TSectionColWidth: TSectionSize
    {
        internal TSectionColWidth(string aTagText, TOneCellValue aParent, TRichString[] Params): base(aTagText, aParent, TValueType.ColumnWidth, Params)
        {
        }

        protected override void EvaluateInternal(int RowAbs, int ColAbs, int RowOfs, int ColOfs, TValueAndXF val)
        {
            if (val.Workbook != null)
            {
                switch (RowColReportAction)
                {
                    case TRowColReportAction.Show:
                        val.Workbook.SetColHidden(ColAbs + 1 + ColOfs, false);
                        break;
                    case TRowColReportAction.Hide:
                        val.Workbook.SetColHidden(ColAbs + 1 + ColOfs, true);
                        break;
                    case TRowColReportAction.Autofit:
                        CalcAutofitGrow(val.Workbook);

                        if (MergedRange == null) MergedRange = val.Workbook.CellMergedBounds(RowAbs + 1, ColAbs + 1);
                        bool IsMerged = MergedRange.ColCount > 1;
                        val.Workbook.MarkColForAutofit(ColAbs + 1 + ColOfs, true, (float)RowColSize, AdjustmentFixed, (int)Math.Round(MinSize * 256), (int)Math.Round(MaxSize * 256), IsMerged); //Should be delayed until all data has been filled.
                        if (val.AutofitInfo != null)
                        {
                            if (val.AutofitInfo.AutofitType == TAutofitType.None) val.AutofitInfo.AutofitType = TAutofitType.OnlyMarked;
                        }
                        break;
                    default:
                        val.Workbook.SetColWidth(ColAbs + 1 + ColOfs, (int)Math.Round(RowColSize * 256));
                        break;
                }
            }
            base.EvaluateInternal(RowAbs, ColAbs, RowOfs, ColOfs, val);
        }    
    }


    internal class TSectionAutofitSettings: TOneSectionValue
    {
        TAutofitType FAutofitType;
        bool FKeepAutofit;
        float FAdjustment;
        int FAdjustmentFixed;
        TAutofitMerged FMergedMode;

        internal TSectionAutofitSettings(string aTagText, TOneCellValue aParent, string GlobalAutofit, string GlobalKeepAutofit, string GlobalAdjustment, string GlobalAdjustmentFixed, string MergedCellsMode): base(aTagText, TValueType.AutofitSettings, -1)
        {
            if (String.Equals(GlobalAutofit, ReportTag.StrAutofitOn, StringComparison.InvariantCultureIgnoreCase)) FAutofitType = TAutofitType.Sheet;
            else
                if (String.Equals(GlobalAutofit, ReportTag.StrAutofitOff, StringComparison.InvariantCultureIgnoreCase)) {/*nothing*/}
            else FlxMessages.ThrowException(FlxErr.ErrInvalidGlobalAutofit, GlobalAutofit, ReportTag.StrAutofitOn, ReportTag.StrAutofitOff);

            if (String.Equals(GlobalKeepAutofit, ReportTag.StrKeepAutofit, StringComparison.InvariantCultureIgnoreCase)) FKeepAutofit = true;
            else
                if (String.Equals(GlobalKeepAutofit, ReportTag.StrDontKeepAutofit, StringComparison.InvariantCultureIgnoreCase)) FKeepAutofit = false;
            else FlxMessages.ThrowException(FlxErr.ErrInvalidGlobalAutofit, GlobalKeepAutofit, ReportTag.StrKeepAutofit, ReportTag.StrDontKeepAutofit);

            if (GlobalAdjustment == null || GlobalAdjustment.Length == 0)
            {
                FAdjustment = 1;
            }
            else
            {
                double adj=0;
                if (TCompactFramework.ConvertToNumber(GlobalAdjustment, CultureInfo.InvariantCulture, out adj))
                {
                    FAdjustment = (float)adj/100;
                }
                else FlxMessages.ThrowException(FlxErr.ErrInvalidGlobalAdjustment, GlobalAdjustment);
            }

            if (GlobalAdjustmentFixed == null || GlobalAdjustmentFixed.Length == 0)
            {
                FAdjustmentFixed = 0;
            }
            else
            {
                double adj=0;
                if (TCompactFramework.ConvertToNumber(GlobalAdjustmentFixed, CultureInfo.InvariantCulture, out adj))
                {
                    FAdjustmentFixed = Convert.ToInt32(adj);
                }
                else FlxMessages.ThrowException(FlxErr.ErrInvalidGlobalAdjustmentFixed, GlobalAdjustmentFixed);
            }

            if (MergedCellsMode == null || MergedCellsMode.Length == 0) FMergedMode = TAutofitMerged.OnLastCell;
            else
            {
                if (String.Equals(MergedCellsMode, ReportTag.StrAutofitModeNone, StringComparison.InvariantCultureIgnoreCase)) FMergedMode = TAutofitMerged.None;
                else
                    if (String.Equals(MergedCellsMode, ReportTag.StrAutofitModeBalanced, StringComparison.InvariantCultureIgnoreCase)) FMergedMode = TAutofitMerged.Balanced;
                else
                {
                    if (!GetMergedMode(MergedCellsMode, ReportTag.StrAutofitModeFirst, "+", TAutofitMerged.OnFirstCell, TAutofitMerged.OnSecondCell, TAutofitMerged.OnThirdCell, TAutofitMerged.OnFourthCell, TAutofitMerged.OnFifthCell))
                    {
                        if (!GetMergedMode(MergedCellsMode, ReportTag.StrAutofitModeLast, "-", TAutofitMerged.OnLastCell, TAutofitMerged.OnLastCellMinusOne, TAutofitMerged.OnLastCellMinusTwo, TAutofitMerged.OnLastCellMinusThree, TAutofitMerged.OnLastCellMinusFour))
                        {
                            FlxMessages.ThrowException(FlxErr.ErrInvalidAutoFitMerged, MergedCellsMode, ReportTag.StrAutofitModeFirst, ReportTag.StrAutofitModeLast, ReportTag.StrAutofitModeNone, ReportTag.StrAutofitModeBalanced);
                        }
                    }
                }
            }
        }

        private bool GetMergedMode(string MergedCellsMode, string Tag, string plus, params TAutofitMerged[] Modes)
        {
            if (!MergedCellsMode.ToUpper(CultureInfo.InvariantCulture).StartsWith(Tag)) return false;
            string Level = MergedCellsMode.Substring(Tag.Length).Trim();
            if (Level.Length == 0)
            {
                FMergedMode = Modes[0];
                return true;
            }

            if (!Level.StartsWith(plus))
                FlxMessages.ThrowException(FlxErr.ErrInvalidAutoFitMerged, MergedCellsMode, ReportTag.StrAutofitModeFirst, ReportTag.StrAutofitModeLast, ReportTag.StrAutofitModeNone, ReportTag.StrAutofitModeBalanced);
            Level = Level.Substring(1);

            double dlvl=0;
            if (TCompactFramework.ConvertToNumber(Level, CultureInfo.InvariantCulture, out dlvl))
            {
                int  iLevel = Convert.ToInt32(dlvl);
                if (iLevel < 0 || iLevel >= Modes.Length)
                {
                    FlxMessages.ThrowException(FlxErr.ErrInvalidAutoFitMerged, MergedCellsMode, ReportTag.StrAutofitModeFirst, ReportTag.StrAutofitModeLast, ReportTag.StrAutofitModeNone, ReportTag.StrAutofitModeBalanced);
                }
                FMergedMode = Modes[iLevel];
                return true;
            }
            
            FlxMessages.ThrowException(FlxErr.ErrInvalidAutoFitMerged, MergedCellsMode, ReportTag.StrAutofitModeFirst, ReportTag.StrAutofitModeLast, ReportTag.StrAutofitModeNone, ReportTag.StrAutofitModeBalanced);
            return false; //just to compile
        }
    
        protected override void EvaluateInternal(int RowAbs, int ColAbs, int RowOfs, int ColOfs, TValueAndXF val)
        {
            base.EvaluateInternal(RowAbs, ColAbs, RowOfs, ColOfs, val);
            if (val.AutofitInfo != null)
            {
                val.AutofitInfo.AutofitType = FAutofitType;
                val.AutofitInfo.GlobalAdjustment = FAdjustment;
                val.AutofitInfo.GlobalAdjustmentFixed = FAdjustmentFixed;
                val.AutofitInfo.KeepAutofit = FKeepAutofit;
                val.AutofitInfo.MergedMode = FMergedMode;
            }
        }    
    }


    internal class TSectionMergeRange: TOneSectionValue
    {
        private int Top;
        private int Left;
        private int Bottom;
        private int Right;

        private TOneCellValue RangeValue;

        internal TSectionMergeRange(ExcelFile Workbook, TStackData Stack, TBand CurrentBand, FlexCelReport fr, string aTagText, TOneCellValue aParent, string RangeStr): base(aTagText, TValueType.MergeRange, -1)
        {
            if (FixedNamedRange(Workbook, RangeStr)) return;

            int XF1 = -1;
            RangeValue = TCellParser.GetCellValue(TCellParser.TryConvert(Workbook,new TRichString(RangeStr), ref XF1), Workbook, Stack, XF1, CurrentBand, fr);
        }

        private bool FixedNamedRange(ExcelFile Workbook, string RangeStr)
        {
            TXlsNamedRange XlsRange = Workbook.GetNamedRange(RangeStr, -1, Workbook.ActiveSheet);
            if (XlsRange == null)
                XlsRange = Workbook.GetNamedRange(RangeStr, -1, 0);

            if (XlsRange != null)
            {
                Top = XlsRange.Top;
                Left = XlsRange.Left;
                Bottom = XlsRange.Bottom;
                Right = XlsRange.Right;
                return true;
            }

            return false;
        }
    
        protected override void EvaluateInternal(int RowAbs, int ColAbs, int RowOfs, int ColOfs, TValueAndXF val)
        {
            if (RangeValue != null)
            {
                TValueAndXF val1 = new TValueAndXF();
                RangeValue.Evaluate(RowAbs,ColAbs,0,0, val1);
                string RangeStr = FlxConvert.ToString(val1.Value);

                if (!FixedNamedRange(val.Workbook, RangeStr)) //might happen in a report parameter.
                {
                    string[] Addresses = RangeStr.Split(TFormulaMessages.TokenChar(TFormulaToken.fmRangeSep));
                    if (Addresses == null || (Addresses.Length != 2 && Addresses.Length != 1))
                        FlxMessages.ThrowException(FlxErr.ErrInvalidRef, RangeStr);
                    TCellAddress FirstCell = new TCellAddress(Addresses[0]);
                    Top = FirstCell.Row;
                    Left = FirstCell.Col;
                    if (Addresses.Length > 1) FirstCell = new TCellAddress(Addresses[1]);
                    Bottom = FirstCell.Row;
                    Right = FirstCell.Col;
                }
            }

            if (val.Workbook != null)
                val.Workbook.MergeCells(Top+RowOfs, Left+ColOfs, Bottom+RowOfs, Right+ColOfs);
            base.EvaluateInternal(RowAbs, ColAbs, RowOfs, ColOfs, val);
        }    
    }

    internal class TSectionHPageBreak: TOneSectionValue
    {
        internal TSectionHPageBreak(string aTagText, TOneCellValue aParent): base(aTagText, TValueType.HPageBreak, -1)
        {
        }
    
        protected override void EvaluateInternal(int RowAbs, int ColAbs, int RowOfs, int ColOfs, TValueAndXF val)
        {
            if (val.Workbook != null)
                val.Workbook.InsertHPageBreak(RowAbs + 1 + RowOfs, true);
            base.EvaluateInternal(RowAbs, ColAbs, RowOfs, ColOfs, val);
        }    
    }

    internal class TSectionVPageBreak: TOneSectionValue
    {
        internal TSectionVPageBreak(string aTagText, TOneCellValue aParent): base(aTagText, TValueType.VPageBreak, -1)
        {
        }
    
        protected override void EvaluateInternal(int RowAbs, int ColAbs, int RowOfs, int ColOfs, TValueAndXF val)
        {
            if (val.Workbook != null)
                val.Workbook.InsertVPageBreak(ColAbs + 1 + ColOfs, true);
            base.EvaluateInternal(RowAbs, ColAbs, RowOfs, ColOfs, val);
        }    
    }


    internal class TSectionAutoPageBreaks : TOneSectionValue
    {
        int PageScale;
        int PercentOfUsedPage;

        internal TSectionAutoPageBreaks(string aTagText, TOneCellValue aParent, string aPercentOfUsedPage, string aPageScale)
            : base(aTagText, TValueType.AutoPageBreaks, -1)
        {
            PercentOfUsedPage = GetValue(aPercentOfUsedPage, 20, FlxErr.ErrInvalidAutoPageBreaksPercent);
            PageScale = GetValue(aPageScale, 95, FlxErr.ErrInvalidAutoPageBreaksPageScale);
        }

        private static int GetValue(string StrValue, int DefaultPercent, FlxErr ErrorCode)
        {
            int Result = -1;

            if (StrValue == null || StrValue.Length == 0)
            {
                Result = DefaultPercent;
            }
            else
            {
                double adj = 0;
                if (TCompactFramework.ConvertToNumber(StrValue, CultureInfo.InvariantCulture, out adj))
                {
                    Result = (int)adj;
                }
                else FlxMessages.ThrowException(ErrorCode, StrValue);
            }

            if (Result < 0 || Result > 100)
                FlxMessages.ThrowException(ErrorCode, StrValue);

            return Result;
        }

        internal static TSectionAutoPageBreaks Create(TRichString aTagText, TOneCellValue aParent, TRichString TagParams)
        {
            List<TRichString> Sections = new List<TRichString>();
            TCellParser.ParseParams(aTagText, TagParams, Sections);
            if (Sections.Count > 2)
                FlxMessages.ThrowException(FlxErr.ErrTooMuchArgs, aTagText);

            string aPercentOfUsedPage; string aPageScale;
            if (Sections.Count < 1) aPercentOfUsedPage = null; else aPercentOfUsedPage = FlxConvert.ToString(Sections[0]);
            if (Sections.Count < 2) aPageScale = null; else aPageScale = FlxConvert.ToString(Sections[1]);
            return new TSectionAutoPageBreaks(aTagText.ToString(), aParent, aPercentOfUsedPage, aPageScale);


        }


        protected override void EvaluateInternal(int RowAbs, int ColAbs, int RowOfs, int ColOfs, TValueAndXF val)
        {
            base.EvaluateInternal(RowAbs, ColAbs, RowOfs, ColOfs, val);
            val.AutoPageBreaksPercent = PercentOfUsedPage;
            val.AutoPageBreaksPageScale = PageScale;
        }
    }


    internal class TSectionDataSet: TOneSectionValue
    {
        //DataSet
        private TBand FBand;
        private int DataSetColumn;

        private TSectionDataSet(string aTagText, TOneCellValue aParent, TBand aBand, int aDataSetColumn, int aStartFont, TValueType aValueType): base(aTagText, aValueType, aStartFont)
        {
            FBand=aBand;
            DataSetColumn=aDataSetColumn;
        }

        internal static TBand FindDetailBand(TBand CurrentBand, string DataSetName)
        {
            if (CurrentBand == null || CurrentBand.DetailBands == null) return null;
            for (int i=0; i< CurrentBand.DetailBands.Count; i++)
            {
                TBand DetailBand = CurrentBand.DetailBands[i];
                if (DetailBand == null || (DetailBand.DataSource == null)) continue;
                if (String.Equals(DetailBand.DataSourceName, DataSetName, StringComparison.CurrentCultureIgnoreCase)) return DetailBand;
                //If it is not a direct child we cannot use this, since the band might be filtered different than the one we are searching for.
                //TBand Result = FindDetailBand(DetailBand, DataSetName);
                //if (Result != null) return Result;
            }
            return null;

        }

        internal static void ParseDbName(TRichString aDataSetValue, out string DataSetName, out string Column, out TRichString DefaultValue, bool CanHaveEmptyCol)
        {
            string aDataSetMember = aDataSetValue.ToString();

            int sepPos=aDataSetMember.LastIndexOf(ReportTag.DbSeparator);
            if (sepPos<0) 
            {
                if (!CanHaveEmptyCol) FlxMessages.ThrowException(FlxErr.ErrMemberNotFound, aDataSetMember);

                DataSetName = aDataSetMember;
                Column = null;
                DefaultValue = null;
                return;
            }
            DataSetName=aDataSetMember.Substring(0, sepPos);

            string ColumnPlusDefault=aDataSetMember.Substring(sepPos+1);
            int defaultSepPos = ColumnPlusDefault.IndexOf(ReportTag.ParamDelim);
            Column = defaultSepPos < 0? ColumnPlusDefault: ColumnPlusDefault.Substring(0, defaultSepPos);

            DefaultValue =  defaultSepPos < 0? null: aDataSetValue.Substring(sepPos + 1 + defaultSepPos + 1);
        }

        internal static TOneSectionValue Create(ExcelFile Workbook, TStackData Stack, int XF1, TOneCellValue aParent, TRichString aDataSetValue, 
            TBand aCurrentBand, int aStartFont, FlexCelReport fr, bool CanAddDataSets)
        {
            string aTagText = aDataSetValue.ToString();
            
            string DataSetName; string Column; TRichString DefaultValue;
            ParseDbName(aDataSetValue, out DataSetName, out Column, out DefaultValue, false);
            
            
            if (String.Equals(Column, ReportTag.StrFullDsCaptions, StringComparison.InvariantCultureIgnoreCase))  //It doesn't need a current band.
            {
                return new TSectionDataSetFullCaptions(aTagText, aParent, DataSetName, aStartFont, fr);
            }

            TBand CurrentBand = aCurrentBand;
            while (CurrentBand != null && !String.Equals(CurrentBand.DataSourceName, DataSetName, StringComparison.CurrentCultureIgnoreCase))
                CurrentBand = CurrentBand.SearchBand;

            //ROWCOUNT can (and should) be accessed outside the named range. so the following checks are moved down.
            //if (CurrentBand==null) FlxMessages.ThrowException(FlxErr.ErrDataSetNotFoundInExpression, DataSetName, aDataSetMember);           
            // if (CurrentBand.DataSource==null) FlxMessages.ThrowException(FlxErr.ErrDataSetNotInRange, aDataSetMember);
            
            if (String.Equals(Column, ReportTag.StrRowCountColumn, StringComparison.InvariantCultureIgnoreCase))
            {
                if (CurrentBand == null || CurrentBand.DataSource == null)
                {
                    CurrentBand = FindDetailBand(aCurrentBand, DataSetName); //This is an optimization. If one of the children bands is the one we are searching for, we do not need to create a "pseudo band"
                }
                if (CurrentBand == null || CurrentBand.DataSource == null)
                {
                    TDataSourceInfo DsInfo = fr.GetDataTable(DataSetName, aTagText);  
                    CurrentBand = new TBand(DsInfo.CreateDataSource(aCurrentBand, fr.ExtraRelations, fr.StaticRelations), aCurrentBand, new TXlsCellRange(1,1,1,1), DataSetName, TBandType.Ignore, false, DataSetName);
                    aCurrentBand.DetailBands.Add(CurrentBand);
                }
    
                return new TSectionDataSet(aTagText, aParent, CurrentBand, (int)TPseudoColumn.RowCount, aStartFont, TValueType.DataSet);
            }

            if (CurrentBand == null || CurrentBand.DataSource == null)
            {
                if (CanAddDataSets)
                {
                    TDataSourceInfo DsInfo = fr.GetDataTable(DataSetName, aTagText);
                    CurrentBand = new TBand(DsInfo.CreateDataSource(aCurrentBand, fr.ExtraRelations, fr.StaticRelations), aCurrentBand, new TXlsCellRange(1, 1, 1, 1), DataSetName, TBandType.Ignore, false, DataSetName);
                    aCurrentBand.DetailBands.Add(CurrentBand);
                }
            }

            bool IsRowPos = String.Equals(Column, ReportTag.StrRowPosColumn, StringComparison.InvariantCultureIgnoreCase);
            bool IsFullDs = String.Equals(Column, ReportTag.StrFullDs, StringComparison.InvariantCultureIgnoreCase);

            //Use the default value if supplied.
            int ColumnIndex = -1;
            if (CurrentBand != null) ColumnIndex = CurrentBand.DataSource.GetColumnWithoutException(Column);
            if (CurrentBand == null || CurrentBand.DataSource == null || (ColumnIndex < 0 && !IsRowPos && !IsFullDs))			
            {
                if (DefaultValue == null) 
                {
                    if (CurrentBand==null) FlxMessages.ThrowException(FlxErr.ErrDataSetNotFoundInExpression, DataSetName, aTagText);           
                    if (CurrentBand.DataSource==null) FlxMessages.ThrowException(FlxErr.ErrDataSetNotInRange, aTagText);
                    FlxMessages.ThrowException(FlxErr.ErrColumNotFound, Column, CurrentBand.DataSource.Name);
                }
                return new TSectionEqual(aTagText, aParent, TCellParser.GetCellValue(TCellParser.TryConvert(Workbook, DefaultValue, ref XF1), Workbook, Stack, XF1, CurrentBand, fr), aStartFont); 
            }

            if (IsRowPos)
            {
                return new TSectionDataSet(aTagText, aParent, CurrentBand, (int)TPseudoColumn.RowPos, aStartFont, TValueType.DataSet);
            }
            
            if (IsFullDs)
            {
                return new TSectionDataSet(aTagText, aParent, CurrentBand, 0, aStartFont, TValueType.FullDataSet);
            }

            return new TSectionDataSet(aTagText, aParent, CurrentBand, ColumnIndex, aStartFont, TValueType.DataSet);
            
        }

        protected override void EvaluateInternal(int RowAbs, int ColAbs, int RowOfs, int ColOfs, TValueAndXF val)
        {
            int DsColumn = DataSetColumn < 0? DataSetColumn: DataSetColumn + val.FullDataSetColumnIndex;
            object o = FBand.DataSource.GetValue(DsColumn);
            string so= (o as String);
            if (so!=null && so.IndexOf((char)13)>=0)
                o=so.Replace((char)13,' ');

            if (ValueType == TValueType.FullDataSet) 
            {
                val.FullDataSetColumnCount = FBand.DataSource.ColumnCount;
            }
            val.Value= o;
            val.Action=ValueType;
        }

        internal override int RecordCount()
        {
            return FBand.RecordCount;
        }

        internal override TBand DataBand()
        {
            return FBand;
        }

    }
    internal class TSectionDataSetFullCaptions: TOneSectionValue
    {
        //DataSet
        private TFlexCelDataSource FDataSource;

        internal TSectionDataSetFullCaptions(string aTagText, TOneCellValue aParent, string aTableName, int aStartFont, FlexCelReport fr): base(aTagText, TValueType.FullDataSetCaptions, aStartFont)
        {
            TDataSourceInfo di = fr.TryGetDataTable(aTableName);
            if (di == null) FlxMessages.ThrowException(FlxErr.ErrDataSetNotFoundInExpression, aTableName, string.Empty);
            VirtualDataTable FTable= di.Table;
            if (FTable==null)
                FlxMessages.ThrowException(FlxErr.ErrDataSetNotFoundInExpression, aTableName, string.Empty);
            FDataSource= new TFlexCelDataSource(aTableName, FTable, fr.ExtraRelations, fr.StaticRelations, null, string.Empty, fr);
        }

        protected override void EvaluateInternal(int RowAbs, int ColAbs, int RowOfs, int ColOfs, TValueAndXF val)
        {
            base.EvaluateInternal(RowAbs, ColAbs, RowOfs, ColOfs, val);
            if (FDataSource.ColumnCount>0) val.Value=FDataSource.ColumnCaption(val.FullDataSetColumnIndex); else val.Value=null;
        
            val.FullDataSetColumnCount = FDataSource.ColumnCount;
        }

        protected override void Dispose(bool disposing)
        {
            try
            {
                if (disposing)
                {
                    if (FDataSource != null) FDataSource.Dispose();
                }
            }
            finally
            {
                base.Dispose(disposing);
            }
        }
    }

    internal class TSectionProperty: TOneSectionValue
    {
        internal object Value;

        internal TSectionProperty(string aTagText, TOneCellValue aParent, object aValue, int aStartFont): base(aTagText, TValueType.Property, aStartFont)
        {
            object o=aValue;
            string so= (o as String);
            if (so!=null && so.IndexOf((char)13)>=0)
                o=so.Replace((char)13,' ');
            Value=o;
        }

        protected override void EvaluateInternal(int RowAbs, int ColAbs, int RowOfs, int ColOfs, TValueAndXF val)
        {
            val.Value = Value;
            val.Action=ValueType;
        }
    }

    internal class TSectionLookup: TOneSectionValue
    {
        //DataSet
        private TOneCellValue FValues;
        private int FResultField;
        private TDataSourceInfo FDInfo;
        private string SearchKeys;

        private TSectionLookup(string aTagText, TOneCellValue aParent, string aTableName, string aKeys, TOneCellValue aValues,
            string aResultField, int aStartFont, FlexCelReport fr, string TagText): base(aTagText, TValueType.Lookup, aStartFont)
        {
            FValues=aValues;

            FDInfo = fr.TryGetDataTable(aTableName);
            if (FDInfo==null) FlxMessages.ThrowException(FlxErr.ErrDataSetNotFoundInExpression, aTableName, TagText);

            FResultField = FDInfo.Table.GetColumn(aResultField);
            if (FResultField < 0) FlxMessages.ThrowException(FlxErr.ErrColumNotFound, aResultField, aTableName);

            SearchKeys = aKeys;
        }

        protected override void Dispose(bool disposing)
        {
            try
            {
                if (disposing)
                {
                    if (FValues != null) FValues.Dispose();
                }
            }
            finally
            {
                base.Dispose(disposing);
            }
        }

        internal static TSectionLookup Create(
            ExcelFile Workbook, TStackData Stack, TBand CurrentBand, FlexCelReport fr,
            TOneCellValue aParent, TRichString TagText, TRichString TagParams)
        {
            TRichString[] LoSections = new TRichString[4];
            int XF5 = -1;
            TCellParser.ParseParams(TagText, TagParams, LoSections);
            int StartFont = -1;
            if (TagText.RTFRunCount > 0 && TagText.RTFRun(0).FirstChar == 0) StartFont = TagText.RTFRun(0).FontIndex;

            return new TSectionLookup(TagText.ToString(), aParent, LoSections[0].ToString(), LoSections[1].ToString(),
                TCellParser.GetCellValue(TCellParser.TryConvert(Workbook, LoSections[2], ref XF5), Workbook, Stack, XF5, CurrentBand, fr),
                LoSections[3].ToString(), StartFont, fr, TagText.ToString());

        }

        protected override void EvaluateInternal(int RowAbs, int ColAbs, int RowOfs, int ColOfs, TValueAndXF val)
        {
            val.Action=ValueType;

            TValueAndXF dVal= new TValueAndXF(val.DebugStack);
            FValues.Evaluate(RowAbs, ColAbs, RowOfs, ColOfs, dVal);

            object drs=null;
            
            object[] keys=(dVal.Value as Object[]);
                if (keys!=null)
                    drs=FDInfo.Table.Lookup(FResultField, SearchKeys, keys); 
                else 
                    drs=FDInfo.Table.Lookup(FResultField, SearchKeys, new object[]{dVal.Value});

            if (drs==null) val.Value=null;
            else
            {
                object o= drs;
                string so= (o as String);
                if (so!=null && so.IndexOf((char)13)>=0)
                    o=so.Replace((char)13,' ');
                val.Value=o;
            }
        }
    }

    internal class TSectionUserFunction: TOneSectionValue
    {
        internal TFlexCelUserFunction FFunction;
        internal TOneCellValue[] FParams;

        internal TSectionUserFunction(string aTagText, TOneCellValue aParent, TFlexCelUserFunction aFunction, TOneCellValue[] aParams, int aStartFont): base(aTagText, TValueType.Array, aStartFont)
        {
            if (aFunction==null) FlxMessages.ThrowException(FlxErr.ErrUndefinedUserFunction);
            FFunction=aFunction;
            FParams=aParams;
        }

        internal static TSectionUserFunction Create(
            ExcelFile Workbook, TStackData Stack, TBand CurrentBand, FlexCelReport fr,
            TFlexCelUserFunction aFunction,
            TOneCellValue aParent, TRichString TagText, TRichString TagParams, int aStartFont)
        {
            List<TRichString> Sections= new List<TRichString>();
            int XF5=-1;
            TCellParser.ParseParams(TagText, TagParams, Sections);
            TOneCellValue[] ValueArray= new TOneCellValue[Sections.Count];
            for (int i=0; i< ValueArray.Length; i++)
                ValueArray[i]= TCellParser.GetCellValue(TCellParser.TryConvert(Workbook,(TRichString) Sections[i], ref XF5), Workbook, Stack, XF5, CurrentBand, fr);

            return new TSectionUserFunction(TagText.ToString(), aParent, aFunction, ValueArray, aStartFont);
        }

        protected override void Dispose(bool disposing)
        {
            try
            {
                if (disposing)
                {
                    foreach (TOneCellValue c in FParams)
                    {
                        if (c != null) c.Dispose();
                    }
                }
            }
            finally
            {
                base.Dispose(disposing);
            }
        }

        protected override void EvaluateInternal(int RowAbs, int ColAbs, int RowOfs, int ColOfs, TValueAndXF val)
        {
            object[] Result= new object[FParams.Length];
            TValueAndXF OneVal= new TValueAndXF();
            for (int i=0; i<Result.Length;i++)
            {
                OneVal.Clear();
                OneVal.DebugStack = val.DebugStack;
                FParams[i].Evaluate(RowAbs, ColAbs, RowOfs, ColOfs, OneVal);
                Result[i]=OneVal.Value;
            }
            val.Value=FFunction.Evaluate(Result);
            val.Action=ValueType;
        }
    }
    

    internal abstract class TSectionAggDataSet: TOneSectionValue
    {
        protected int DataSetColumn;
        protected TBand FBand;
        protected TRPNExpression Expression;
        protected TRPNExpression Filter;
        protected string FilterString;

        protected TSectionAggDataSet(string aTagText, TOneCellValue aParent, TBand aBand, int aDataSetColumn, 
            TRPNExpression aExpression, TRPNExpression aFilter, string aFilterString,
            TValueType aValueType): base(aTagText, aValueType, -1)
        {
            DataSetColumn = aDataSetColumn;
            FBand = aBand;
            Expression = aExpression;
            Filter = aFilter;
            FilterString = aFilterString;
        }

        protected object GetValue(TFlexCelDataSource ds, int RowAbs, int ColAbs, int RowOfs, int ColOfs, TValueAndXF val)
        {
            if (Expression != null)
            {
                return Expression.Evaluate(RowAbs, ColAbs, RowOfs, ColOfs, val.DebugStack, val.FullDataSetColumnIndex);
            }
            else
            {
                int DsColumn = GetColIndex(val);
                return ds.GetValue(DsColumn);
            }
        }

        protected int GetColIndex(TValueAndXF val)
        {
            int DsColumn = DataSetColumn < 0 ? DataSetColumn : DataSetColumn + val.FullDataSetColumnIndex;
            return DsColumn;
        }

        protected bool FilterRecord(TFlexCelDataSource ds, int RowAbs, int ColAbs, int RowOfs, int ColOfs, TValueAndXF val)
        {
            if (Filter == null) return false;
            object o = Filter.Evaluate(RowAbs, ColAbs, RowOfs, ColOfs, val.DebugStack, val.FullDataSetColumnIndex);
            
            bool R;
            if (TBaseParsedToken.ExtToBool(o, out R)) return !R;  //filtered records are the ones where result is false.

            FlxMessages.ThrowException(FlxErr.ErrFilterMustReturnABooleanValue, FilterString);
            return false; //just to compile.
        }

    }

    internal class TSectionAggregate: TSectionAggDataSet
    {
        private TAggregateType AggregateType;

        private TSectionAggregate(string aTagText, TOneCellValue aParent, TBand aBand, int aDataSetColumn, 
            TRPNExpression aExpression, TRPNExpression aFilter, string aFilterString, TAggregateType aAggregateType,
            TValueType aValueType): base(aTagText, aParent, aBand, aDataSetColumn, aExpression, aFilter, aFilterString, aValueType)
        {
            AggregateType = aAggregateType;
        }

        internal static TOneSectionValue Create(ExcelFile Workbook, TStackData Stack, int XF1, TOneCellValue aParent, 
            TRichString aTagText, TRichString aTagParams, TBand aCurrentBand, FlexCelReport fr)
        {
            TRichString[] Sections = new TRichString[4];
            TCellParser.ParseParams(aTagText, aTagParams, Sections, true);

            TAggregateType AggType = GetAggType(FlxConvert.ToString(Sections[0]));

            string DataSetName; string Column; TRichString DefaultValue;
            TSectionDataSet.ParseDbName(Sections[1], out DataSetName, out Column, out DefaultValue, Sections[2] != null);
            
            TDataSourceInfo DsInfo = fr.GetDataTable(DataSetName, Sections[1].ToString());  
            TBand CurrentBand = new TBand(DsInfo.CreateDataSource(aCurrentBand, fr.ExtraRelations, fr.StaticRelations), aCurrentBand, new TXlsCellRange(1,1,1,1), DataSetName, TBandType.Ignore, false, DataSetName);
            aCurrentBand.DetailBands.Add(CurrentBand);

            TRPNExpression Expr = null;
            if (Sections[2] != null && Sections[2].ToString().Trim().Length > 0)
            {
                Expr = new TRPNExpression(Sections[2].ToString(), Workbook, CurrentBand, fr, Stack);
            }

            TRPNExpression Filt = null;
            if (Sections[3] != null && Sections[3].ToString().Trim().Length > 0)
            {
                Filt = new TRPNExpression(Sections[3].ToString(), Workbook, CurrentBand, fr, Stack);
            }


            int ColumnIndex = -1;			
            if (Expr == null)
                ColumnIndex = CurrentBand.DataSource.GetColumn(Column);

            return new TSectionAggregate(Convert.ToString(aTagText), aParent, CurrentBand, ColumnIndex, Expr, Filt, FlxConvert.ToString(Sections[2]), AggType, TValueType.Aggregate);          
        }

        private static TAggregateType GetAggType(string s)
        {
            if (String.Equals(s, ReportTag.StrAggSum, StringComparison.InvariantCultureIgnoreCase)) return TAggregateType.Sum;
            if (String.Equals(s, ReportTag.StrAggAvg, StringComparison.InvariantCultureIgnoreCase)) return TAggregateType.Average;
            if (String.Equals(s, ReportTag.StrAggMax, StringComparison.InvariantCultureIgnoreCase)) return TAggregateType.Max;
            if (String.Equals(s, ReportTag.StrAggMin, StringComparison.InvariantCultureIgnoreCase)) return TAggregateType.Min;

            FlxMessages.ThrowException(FlxErr.ErrInvalidAggParameter, FlxConvert.ToString(s));
            return TAggregateType.Sum; //just to compile.
        }

        protected override void EvaluateInternal(int RowAbs, int ColAbs, int RowOfs, int ColOfs, TValueAndXF val)
        {
            val.Action=ValueType;

            TFlexCelDataSource ds = FBand.DataSource;

            double? dn;
            if (Expression == null && Filter == null && ds.TryAggregate(AggregateType, GetColIndex(val), out dn))
            {
                if (!dn.HasValue) val.Value = null; else val.Value = dn.Value;
                return;
            }

            double d = 0;
            ds.First();

            int count = 0;
            bool First = true;
            while (!ds.Eof())
            {
                try
                {
                    if (FilterRecord(ds, RowAbs, ColAbs, RowOfs, ColOfs, val)) continue;

                    count++;

                    double v;

                    object o = GetValue(ds, RowAbs, ColAbs, RowOfs, ColOfs, val);

                    if (Convert.IsDBNull(o)) continue;
                    if (!TBaseParsedToken.ExtToDouble(o, out v)) 
                    {
                        val.Value = TFlxFormulaErrorValue.ErrNum;
                        return;
                    }
                
                    switch (AggregateType)
                    {
                        case TAggregateType.Sum:
                        case TAggregateType.Average:
                            d += v;
                            break;

                        case TAggregateType.Min:
                            if (First)
                            {
                                d = v;
                                First = false;
                            }
                            if (v < d) d = v;
                            break;

                        case TAggregateType.Max:
                            if (First)
                            {
                                d = v;
                                First = false;
                            }
                            if (v > d) d = v;
                            break;
                    }
                }
                finally
                {
                    ds.Next();
                }
            }
                
            if (AggregateType == TAggregateType.Average)
            {
                if (count <= 0)
                {
                    val.Value = TFlxFormulaErrorValue.ErrDiv0;
                    return;
                }

                d /= count;
            }

            val.Value= d;
        }
    }


    internal class TSectionList: TSectionAggDataSet
    {
        private string Delim;

        private TSectionList(string aTagText, TOneCellValue aParent, TBand aBand, int aDataSetColumn, 
            TRPNExpression aExpression, TRPNExpression aFilter, string aFilterString, string aDelim,
            TValueType aValueType): base(aTagText, aParent, aBand, aDataSetColumn, aExpression, aFilter, aFilterString, aValueType)
        {
            Delim = aDelim;
        }

        internal static TOneSectionValue Create(ExcelFile Workbook, TStackData Stack, int XF1, TOneCellValue aParent, 
            TRichString aTagText, TRichString aTagParams, TBand aCurrentBand, FlexCelReport fr)
        {
            TRichString[] Sections = new TRichString[4];
            TCellParser.ParseParams(aTagText, aTagParams, Sections, true);

            string DataSetName; string Column; TRichString DefaultValue;
            TSectionDataSet.ParseDbName(Sections[0], out DataSetName, out Column, out DefaultValue, Sections[2] != null);

#if(!COMPACTFRAMEWORK || FRAMEWORK20)
            string Delim = Sections[1]==null? " ": Convert.ToString(TCellParser.UnQuote(Sections[1]), CultureInfo.InvariantCulture);
#else
            string Delim = Sections[1]==null? " ": Convert.ToString(TCellParser.UnQuote(Sections[1]));
#endif
            TDataSourceInfo DsInfo = fr.GetDataTable(DataSetName, Sections[0].ToString());  
            TBand CurrentBand = new TBand(DsInfo.CreateDataSource(aCurrentBand, fr.ExtraRelations, fr.StaticRelations), aCurrentBand, new TXlsCellRange(1,1,1,1), DataSetName, TBandType.Ignore, false, DataSetName);
            if (aCurrentBand != null)
            {
                aCurrentBand.DetailBands.Add(CurrentBand);
            }

            TRPNExpression Expr = null;
            if (Sections[2] != null && Sections[2].ToString().Trim().Length > 0)
            {
                Expr = new TRPNExpression(Sections[2].ToString(), Workbook, CurrentBand, fr, Stack);
            }

            TRPNExpression Filt = null;
            if (Sections[3] != null && Sections[3].ToString().Trim().Length > 0)
            {
                Filt = new TRPNExpression(Sections[3].ToString(), Workbook, CurrentBand, fr, Stack);
            }


            int ColumnIndex = -1;			
            if (Expr == null)
                ColumnIndex = CurrentBand.DataSource.GetColumn(Column);

            return new TSectionList(Convert.ToString(aTagText), aParent, CurrentBand, ColumnIndex, Expr, Filt, FlxConvert.ToString(Sections[2]), Delim, TValueType.Aggregate);          
        }


        protected override void EvaluateInternal(int RowAbs, int ColAbs, int RowOfs, int ColOfs, TValueAndXF val)
        {
            val.Action=ValueType;

            TFlexCelDataSource ds = FBand.DataSource;

            ds.First();
            StringBuilder Result = new StringBuilder();

            bool First = true;
            while (!ds.Eof())
            {
                try
                {
                    if (FilterRecord(ds, RowAbs, ColAbs, RowOfs, ColOfs, val)) continue;

                    object o = GetValue(ds, RowAbs, ColAbs, RowOfs, ColOfs, val);

                    if (Convert.IsDBNull(o) || o == null) continue;
                
                    if (!First && Delim != null) Result.Append(Delim);
                    Result.Append(o.ToString());
                    First = false;

                }
                finally
                {
                    ds.Next();
                }
            }
                
            val.Value = Result.ToString();
        }
    }


    internal class TSectionDbValue: TOneSectionValue
    {
        protected TBand FBand;
        protected TRPNExpression DataSetColumn;
        protected TRPNExpression DataSetRow;
        private TRichString DefaultValue;

        protected TSectionDbValue(string aTagText, TOneCellValue aParent, TBand aBand, 
            TRPNExpression aDataSetRow, TRPNExpression aDataSetColumn, TRichString aDefaultValue,
            TValueType aValueType): base(aTagText, aValueType, -1)
        {
            FBand = aBand;
            DataSetColumn = aDataSetColumn;
            DataSetRow = aDataSetRow;
            DefaultValue = aDefaultValue;
        }

        internal static TOneSectionValue Create(ExcelFile Workbook, TStackData Stack, int XF1, TOneCellValue aParent, TRichString aTagText, TRichString aTagParams, TBand aCurrentBand, FlexCelReport fr)
        {
            TRichString[] Sections = new TRichString[4];
            TCellParser.ParseParams(aTagText, aTagParams, Sections, true);
    
            if (Sections[0] == null || Sections[1] == null || Sections[2] == null) FlxMessages.ThrowException(FlxErr.ErrMissingArgs, aTagText.ToString());
            string DataSetName = Sections[0].ToString().Trim();
            string RowExpr = Sections[1].ToString().Trim();
            string ColExpr = Sections[2].ToString().Trim();

            if (DataSetName.Length == 0 || ColExpr.Length == 0) FlxMessages.ThrowException(FlxErr.ErrMissingArgs, aTagText.ToString());

            TBand CurrentBand = aCurrentBand;
            while (CurrentBand!=null && !String.Equals(CurrentBand.DataSourceName, DataSetName, StringComparison.CurrentCultureIgnoreCase))
                CurrentBand=CurrentBand.SearchBand;

            if (CurrentBand == null || CurrentBand.DataSource == null) //It is not inside a range. No problem, since we are specifying row, so we will just use the full dataset for the row.
            {
                TDataSourceInfo DsInfo = fr.GetDataTable(DataSetName, DataSetName);
                CurrentBand = new TBand(DsInfo.CreateDataSource(aCurrentBand, fr.ExtraRelations, fr.StaticRelations), aCurrentBand, new TXlsCellRange(1, 1, 1, 1), DataSetName, TBandType.Ignore, false, DataSetName);
                aCurrentBand.DetailBands.Add(CurrentBand);
            }

            TRPNExpression RExpr = string.IsNullOrEmpty(RowExpr)? null: new TRPNExpression(RowExpr, Workbook, aCurrentBand, fr, Stack);
            TRPNExpression CExpr = new TRPNExpression(ColExpr, Workbook, aCurrentBand, fr, Stack);

            return new TSectionDbValue(Convert.ToString(aTagText), aParent, CurrentBand, RExpr, CExpr, Sections[3], TValueType.Aggregate);          
        }


        protected override void EvaluateInternal(int RowAbs, int ColAbs, int RowOfs, int ColOfs, TValueAndXF val)
        {
            val.Action=ValueType;

            TFlexCelDataSource ds = FBand.DataSource;

            object orow = DataSetRow == null? (double)FBand.DataSource.Position:
                DataSetRow.Evaluate(RowAbs, ColAbs, RowOfs, ColOfs, val.DebugStack, val.FullDataSetColumnIndex);
            if (!(orow is double)) FlxMessages.ThrowException(FlxErr.ErrInvalidRow, FlxConvert.ToString(orow));
            int r = Convert.ToInt32(orow);

            int c = -1;
            object ocol = DataSetColumn.Evaluate(RowAbs, ColAbs, RowOfs, ColOfs, val.DebugStack, val.FullDataSetColumnIndex);
            if ((ocol is double)) 
                c = Convert.ToInt32(ocol);
            else 
                c = ds.GetColumn(FlxConvert.ToString(ocol));

            if (r < 0 || r >= ds.RecordCount || c < 0 || c > ds.ColumnCount)
            {
                if (DefaultValue != null)
                {
                    val.Value = DefaultValue;
                    return;
                }
            }
            
            val.Value = ds.GetValueForRow(r, c);
        }
    }


    internal class TSectionArray: TOneSectionValue
    {
        internal TOneCellValue[] Values;

        internal TSectionArray(string aTagText, TOneCellValue aParent, TOneCellValue[] aValues): base(aTagText, TValueType.Array, -1)
        {
            Values=aValues;
        }

        internal static TSectionArray Create(
            ExcelFile Workbook, TStackData Stack, TBand CurrentBand, FlexCelReport fr,
            TOneCellValue aParent, TRichString TagText, TRichString TagParams)
        {
            List<TRichString> Sections= new List<TRichString>();
            int XF5=-1;
            TCellParser.ParseParams(TagText, TagParams, Sections);
            TOneCellValue[] ValueArray= new TOneCellValue[Sections.Count];
            for (int i=0; i< ValueArray.Length; i++)
                ValueArray[i]= TCellParser.GetCellValue(TCellParser.TryConvert(Workbook,(TRichString) Sections[i], ref XF5), Workbook, Stack, XF5, CurrentBand, fr);

            return new TSectionArray(TagText.ToString(), aParent, ValueArray);
        }

        protected override void EvaluateInternal(int RowAbs, int ColAbs, int RowOfs, int ColOfs, TValueAndXF val)
        {
            object[] Result= new object[Values.Length];
            TValueAndXF OneVal= new TValueAndXF();
            for (int i=0; i<Result.Length;i++)
            {
                OneVal.Clear();
                OneVal.DebugStack = val.DebugStack;
                Values[i].Evaluate(RowAbs, ColAbs, RowOfs, ColOfs, OneVal);
                Result[i]=OneVal.Value;
            }
            val.Value=Result;
            val.Action=ValueType;
        }
    }

    internal class TSectionIf: TOneSectionValue
    {
        private string ConditionText;
        private TRPNExpression IfCondition;
        private TOneCellValue IfTrue;
        private TOneCellValue IfFalse;

        internal TSectionIf (string aTagText, TOneCellValue aParent, string aConditionText, TRPNExpression Section1, TOneCellValue Section2, TOneCellValue Section3): base(aTagText, TValueType.IF, -1)
        {
            IfCondition=Section1;
            IfTrue=Section2;
            IfFalse=Section3;
            ConditionText = aConditionText;
        }

        internal static TSectionIf Create(
            ExcelFile Workbook, TStackData Stack, int XF, TBand CurrentBand, FlexCelReport fr,
            TOneCellValue aParent, TRichString TagText, TRichString TagParams)
        {
            TRichString[] IfSections = new TRichString[3];
            TCellParser.ParseParams(TagText, TagParams, IfSections);
            //If it is not a quoted string, try to convert it to a number, formula, etc.
            int XF1=XF; int XF2=XF;
            return new TSectionIf(TagText.ToString(), aParent, IfSections[0].ToString(),
                new TRPNExpression(IfSections[0].ToString(), Workbook, CurrentBand, fr, Stack), 
                TCellParser.GetCellValue(TCellParser.TryConvert(Workbook, IfSections[1], ref XF1), Workbook, Stack, XF1, CurrentBand, fr), 
                TCellParser.GetCellValue(TCellParser.TryConvert(Workbook, IfSections[2], ref XF2), Workbook, Stack, XF2, CurrentBand, fr));                            		

        }

        protected override void Dispose(bool disposing)
        {
            try
            {
                if (disposing)
                {
                    if (IfTrue != null) IfTrue.Dispose();
                    if (IfFalse != null) IfFalse.Dispose();
                }
            }
            finally
            {
                base.Dispose(disposing);
            }
        }

        protected override void EvaluateInternal(int RowAbs, int ColAbs, int RowOfs, int ColOfs, TValueAndXF val)
        {
            base.EvaluateInternal(RowAbs, ColAbs, RowOfs, ColOfs, val);
            TOneCellValue Result= Resolve(RowAbs, ColAbs, RowOfs, ColOfs, val.DebugStack, val.FullDataSetColumnIndex);
            if (val.DebugStack != null)
            {
                val.DebugStack.Add(ConditionText, Result == IfTrue);
            }
            Result.Evaluate(RowAbs, ColAbs, RowOfs, ColOfs, val);
        }

        internal override TOneCellValue Resolve(int RowAbs, int ColAbs, int RowOfs, int ColOfs, TDebugStack aDebugStack, int FullDataSetColumnIndex)
        {
            if (IfCondition.IsTrue(RowAbs,ColAbs,RowOfs,ColOfs,aDebugStack,FullDataSetColumnIndex)) return IfTrue; else return IfFalse; 
        }
    }

    internal class TSectionEvaluate: TOneSectionValue
    {
        private TRPNExpression EvaluateData;

        internal TSectionEvaluate(string aTagText, TOneCellValue aParent, TRPNExpression aEvaluateData): base(aTagText, TValueType.Evaluate, -1)
        {
            EvaluateData=aEvaluateData;
        }

        protected override void EvaluateInternal(int RowAbs, int ColAbs, int RowOfs, int ColOfs, TValueAndXF val)
        {
            val.Value = EvaluateData.Evaluate(RowAbs, ColAbs, RowOfs, ColOfs, val.DebugStack, val.FullDataSetColumnIndex);
            val.Action=ValueType;
        }
    }

    internal class TSectionEqual: TOneSectionValue
    {
        private TOneCellValue Value;
        internal TSectionEqual(string aTagText, TOneCellValue aParent, TOneCellValue aValue, int aStartFont): base(aTagText, TValueType.Equal, aStartFont)
        {
            Value=aValue;
        }

        internal static TSectionEqual Create(
            ExcelFile Workbook, TStackData Stack, int XF, TBand CurrentBand, FlexCelReport fr,
            string aTagText, TOneCellValue aParent, TRichString TagParams, int aStartFont)     
        {
            string CellStr=TagParams.ToString();
            int Aws=Workbook.ActiveSheet;
            try
            {
                string csep=TFormulaMessages.TokenString(TFormulaToken.fmExternalRef);
                int ShPos=CellStr.IndexOf(csep);
                if (ShPos>=0)
                {
                    Workbook.ActiveSheetByName=CellStr.Substring(0,ShPos);
                    CellStr=CellStr.Substring(ShPos+1);
                }
                TCellAddress addr= new TCellAddress(CellStr);
                string CellHash = Workbook.SheetName.ToUpper(CultureInfo.InvariantCulture) + csep + addr.CellRef;
                if (Stack.UsedRefs.ContainsKey(CellHash)) FlxMessages.ThrowException(FlxErr.ErrCircularReference, CellHash);
                Stack.UsedRefs.Add(CellHash, CellHash);
                
                return new TSectionEqual(aTagText, aParent, 
                    TCellParser.GetCellValue(Workbook.GetCellValue(addr.Row, addr.Col), Workbook, Stack, XF, CurrentBand, fr), aStartFont);
            }
            finally
            {
                Workbook.ActiveSheet=Aws;
            }

        }

        protected override void Dispose(bool disposing)
        {
            try
            {
                if (disposing)
                {
                    if (Value != null) Value.Dispose();
                }
            }
            finally
            {
                base.Dispose(disposing);
            }
        }

        protected override void EvaluateInternal(int RowAbs, int ColAbs, int RowOfs, int ColOfs, TValueAndXF val)
        {
            val.Action=ValueType;
            Value.Evaluate(RowAbs, ColAbs, RowOfs, ColOfs, val);
        }

        internal override TOneCellValue Resolve(int RowAbs, int ColAbs, int RowOfs, int ColOfs, TDebugStack aDebugStack, int FullDataSetColumnIndex)
        {
            return Value;
        }
    }


    internal class TSectionInclude: TOneSectionValue
    {
        //Data about the include
        internal TInclude Value;

        private bool CopyRowFormats;
        private bool CopyColFormats;

        internal TSectionInclude(string aTagText, TOneCellValue aParent, TInclude aValue, bool aCopyRowFormats, bool aCopyColFormats): base(aTagText, TValueType.Include, -1)
        {
            Value=aValue;
            CopyRowFormats = aCopyRowFormats;
            CopyColFormats = aCopyColFormats;
        }

        internal static TSectionInclude Create(
            ExcelFile Workbook, TStackData Stack, TBand CurrentBand, FlexCelReport fr, 
            TOneCellValue aParent, TRichString TagText, TRichString TagParams)     

        {
            TRichString[] IncSections= new TRichString[5];
            TCellParser.ParseParams(TagText, TagParams, IncSections, true);
            for (int i = 0; i < 3; i++)
            {
                if (IncSections[i] == null) FlxMessages.ThrowException(FlxErr.ErrNullParameter, "<#include>[" + i.ToString(CultureInfo.InvariantCulture) + "]");

                //This will evaluate already set expressions, not for example an include that changes with a dataset.
                TOneCellValue cv = TCellParser.GetCellValue(IncSections[i].ToString(), Workbook, Stack, -1, CurrentBand, fr);
                if (cv != null)
                {
                    TValueAndXF v =new TValueAndXF();
                    cv.Evaluate(0, 0, 0, 0, v);
                    if (v != null) IncSections[i] = new TRichString(Convert.ToString(v.Value));
                }
            }

            GetIncludeEventArgs ea=new GetIncludeEventArgs(Workbook, IncSections[0].ToString(), null);
            fr.OnGetInclude(ea);

            if (ea.IncludeData==null) 
                ea.IncludeData=TInclude.OpenInclude(ea, Workbook);
            else
                ea.FileName=String.Empty;

            string Static = IncSections[3] == null? String.Empty: IncSections[3].ToString(CultureInfo.InvariantCulture);
            bool StaticInc = String.Equals(Static, ReportTag.StrStaticInclude, StringComparison.InvariantCultureIgnoreCase);
            bool DynamicInc = Static.Length == 0 || String.Equals(Static, ReportTag.StrDynamicInclude, StringComparison.InvariantCultureIgnoreCase);
            if (!StaticInc && !DynamicInc) FlxMessages.ThrowException(FlxErr.ErrInvalidIncludeStatic, Static, ReportTag.StrStaticInclude, ReportTag.StrDynamicInclude);
            
            string RowCol = IncSections[4] == null? String.Empty: IncSections[4].ToString(CultureInfo.InvariantCulture);
            bool CopyRows = String.Equals(RowCol, ReportTag.StrCopyRows, StringComparison.InvariantCultureIgnoreCase);
            bool CopyCols = String.Equals(RowCol, ReportTag.StrCopyCols, StringComparison.InvariantCultureIgnoreCase);
            bool CopyRowsAndCols = String.Equals(RowCol, ReportTag.StrCopyRowsAndCols, StringComparison.InvariantCultureIgnoreCase);

            if (RowCol.Length > 0 && !CopyRows && !CopyCols && !CopyRowsAndCols) FlxMessages.ThrowException(FlxErr.ErrInvalidIncludeRowCol, RowCol, ReportTag.StrCopyRows, ReportTag.StrCopyCols, ReportTag.StrCopyRowsAndCols);


            return new TSectionInclude(TagText.ToString(), aParent, 
                new TInclude(ea.IncludeData, IncSections[1].ToString(), 
                TCellParser.GetBandType(IncSections[2].ToString(), TagText.ToString()), 
                CurrentBand, TagParams.ToString(), fr.NestedIncludeLevel+1, 
                fr.GetDataTables(),ea.FileName, StaticInc, fr), CopyRows || CopyRowsAndCols, CopyCols || CopyRowsAndCols); 

        }

        protected override void Dispose(bool disposing)
        {
            try
            {
                if (disposing)
                {
                    if (Value != null) Value.Dispose();
                    Value = null;
                }
            }
            finally
            {
                base.Dispose(disposing);
            }
        }

        protected override void EvaluateInternal(int RowAbs, int ColAbs, int RowOfs, int ColOfs, TValueAndXF val)
        {
            base.EvaluateInternal(RowAbs, ColAbs, RowOfs, ColOfs, val);
            if (val.WaitingRanges != null)
                val.WaitingRanges.Add(new TIncludeWaitingRange(Value, RowAbs + 1, ColAbs + 1, CopyRowFormats, CopyColFormats));
        }
    }

    internal class TSectionRegex: TOneSectionValue
    {
        internal Regex  FRegex;
        internal TOneCellValue FMatchString;
        internal TOneCellValue FReplaceString;

        internal TSectionRegex(string aTagText, TOneCellValue aParent, string aRegex, RegexOptions options, TOneCellValue aMatchString, TOneCellValue aReplaceString): base(aTagText, TValueType.Regex, -1)
        {
            FRegex=new Regex(aRegex, options);
            FMatchString=aMatchString;
            FReplaceString=aReplaceString;
        }

        internal static TSectionRegex Create(
            ExcelFile Workbook, TStackData Stack, int XF, TBand CurrentBand, FlexCelReport fr,
            TOneCellValue aParent, TRichString TagText, TRichString TagParams)
        {
            List<TRichString> Sections= new List<TRichString>();
            TCellParser.ParseParams(TagText, TagParams, Sections);
            if (Sections.Count<3) FlxMessages.ThrowException(FlxErr.ErrMissingArgs, TagText);
            if (Sections.Count>4) FlxMessages.ThrowException(FlxErr.ErrTooMuchArgs, TagText);

            string rx = FlxConvert.ToString(TCellParser.TryConvert(Workbook,(TRichString) Sections[1], ref XF));
            RegexOptions rxopt = RegexOptions.None;

            int XF5=-1;
            TOneCellValue Match = TCellParser.GetCellValue(TCellParser.TryConvert(Workbook,(TRichString) Sections[2], ref XF5), Workbook, Stack, XF5, CurrentBand, fr);
            TOneCellValue Replace = null;
            if (Sections.Count>3) Replace = TCellParser.GetCellValue(TCellParser.TryConvert(Workbook,(TRichString) Sections[3], ref XF5), Workbook, Stack, XF5, CurrentBand, fr);

            string ic = ((TRichString)Sections[0]).ToString(CultureInfo.CurrentCulture).Trim();
            if (ic=="1") rxopt = RegexOptions.IgnoreCase;
            else if (ic!="0") FlxMessages.ThrowException(FlxErr.ErrInvalidValue, FlxParam.IgnoreCase.ToString(), ic, 0, 1);

            return new TSectionRegex(TagText.ToString(), aParent, rx, rxopt, Match, Replace);
        }

        protected override void Dispose(bool disposing)
        {
            try
            {
                if (disposing)
                {
                    if (FMatchString != null) FMatchString.Dispose();
                    if (FReplaceString != null) FReplaceString.Dispose();
                }
            }
            finally
            {
                base.Dispose (disposing);
            }
        }

        protected override void EvaluateInternal(int RowAbs, int ColAbs, int RowOfs, int ColOfs, TValueAndXF val)
        {
            string Result = null;
            TValueAndXF m= new TValueAndXF(val.DebugStack);
            FMatchString.Evaluate(RowAbs, ColAbs, RowOfs, ColOfs, m);

            if (FReplaceString!=null)
            {
                TValueAndXF r= new TValueAndXF(val.DebugStack);
                FReplaceString.Evaluate(RowAbs, ColAbs, RowOfs, ColOfs, r);

                Result = FRegex.Replace(FlxConvert.ToString(m.Value), FlxConvert.ToString(r.Value));
            }
            else
            {
                Result = FRegex.Match(FlxConvert.ToString(m.Value)).ToString();
            }

            val.Value= Result;
            val.Action=ValueType;
        }
    }


    internal class TSectionFormula: TOneSectionValue
    {
        internal TSectionFormula(string aTagText, TOneCellValue aParent): base(aTagText, TValueType.Formula, -1)
        {
        }
    
        protected override void EvaluateInternal(int RowAbs, int ColAbs, int RowOfs, int ColOfs, TValueAndXF val)
        {
            val.IsFormula = true;
            base.EvaluateInternal(RowAbs, ColAbs, RowOfs, ColOfs, val);
        }    
    }

    internal class TSectionRef: TOneSectionValue
    {
        int RowRel;
        int ColRel;
        TCellAddress[] Ranges;
        string TagId;
        bool RowDollar;
        bool ColDollar;

        internal TSectionRef(ExcelFile Workbook, TOneCellValue aParent, string TagText, List<TRichString> Params): base(TagText, TValueType.Ref, -1)
        {
            if (Params.Count <= 0)
                FlxMessages.ThrowException(FlxErr.ErrMissingArgs, TagText);
            if (Params.Count > 3) 
                FlxMessages.ThrowException(FlxErr.ErrTooMuchArgs, TagText);

            RowRel = 0;
            ColRel = 0;
            Ranges = null;
            TagId = String.Empty;

            if (Params.Count == 1 || Params.Count == 3)
            {
                //This is a hard one. At parse time, we do not know which named range will be used.
                //For example, you might have 2 local named ranges called "Sheet1!Data" and "Sheet2!Data".
                //The expression could be on the config sheet, and be used on both sheets, and we need
                //to use the different ranges when evaluating. So we will make a "SemiParse" here, to avoid
                //doing everything at runtime.

                TagId = FlxConvert.ToString(Params[0]);
                
                Ranges = new TCellAddress[Workbook.SheetCount + 1];
          
                bool Found = false;
                for (int i=1; i<=Workbook.NamedRangeCount;i++)
                {
                    TXlsNamedRange Nr= Workbook.GetNamedRange(i);
                    if (String.Equals(TagId, Nr.Name, StringComparison.CurrentCultureIgnoreCase))
                    {
                        Ranges[Nr.NameSheetIndex] = new TCellAddress(Workbook.GetSheetName(Nr.SheetIndex), Nr.Top, Nr.Left, false, false);
                        Found = true;
                    }
                }
                if (!Found)FlxMessages.ThrowException(FlxErr.ErrInvalidRefTag, TagId);

                if (Params.Count == 3)
                {
                    string Ra = FlxConvert.ToString(Params[1]).Trim();
                    if (String.Equals(Ra, TFormulaMessages.TokenString(TFormulaToken.fmTrue), StringComparison.InvariantCultureIgnoreCase))
                    {
                        RowDollar = true;
                    }
                    else
                        if (!String.Equals(Ra, TFormulaMessages.TokenString(TFormulaToken.fmFalse), StringComparison.InvariantCultureIgnoreCase))
                    {
                        FlxMessages.ThrowException(FlxErr.ErrInvalidRefTag3, Ra, TFormulaMessages.TokenString(TFormulaToken.fmTrue), TFormulaMessages.TokenString(TFormulaToken.fmFalse));
                    }

                    Ra = FlxConvert.ToString(Params[2]).Trim();
                    if (String.Equals(Ra, TFormulaMessages.TokenString(TFormulaToken.fmTrue), StringComparison.InvariantCultureIgnoreCase))
                    {
                        ColDollar = true;
                    }
                    else
                        if (!String.Equals(Ra, TFormulaMessages.TokenString(TFormulaToken.fmFalse), StringComparison.InvariantCultureIgnoreCase))
                    {
                        FlxMessages.ThrowException(FlxErr.ErrInvalidRefTag3, Ra, TFormulaMessages.TokenString(TFormulaToken.fmTrue), TFormulaMessages.TokenString(TFormulaToken.fmFalse));
                    }
                }
            }
        
            if (Params.Count == 2)
            {
                double dRow = 0;
                if (!TCompactFramework.ConvertToNumber(FlxConvert.ToString(Params[0]).Trim(), CultureInfo.InvariantCulture, out dRow))
                    FlxMessages.ThrowException(FlxErr.ErrInvalidRefTag2, FlxConvert.ToString(Params[0]));

                RowRel = (int)Math.Floor(dRow);

                double dCol = 0;
                if (!TCompactFramework.ConvertToNumber(FlxConvert.ToString(Params[1]).Trim(), CultureInfo.InvariantCulture, out dCol))
                    FlxMessages.ThrowException(FlxErr.ErrInvalidRefTag2, FlxConvert.ToString(Params[1]));

                ColRel = (int)Math.Floor(dCol);
            }

        }
    
        protected override void EvaluateInternal(int RowAbs, int ColAbs, int RowOfs, int ColOfs, TValueAndXF val)
        {
            val.Action = ValueType;

            TCellAddress Address;
            if (Ranges == null)
            {
                Address = new TCellAddress(RowAbs + RowOfs + RowRel + 1, ColAbs + ColOfs + ColRel + 1);
            }
            else
            {
                TCellAddress Addr = Ranges[val.Workbook.ActiveSheet];
                if (Addr == null) Addr = Ranges[0];  //Global name.
                if (Addr == null) FlxMessages.ThrowException(FlxErr.ErrInvalidRefTag, TagId);
                
                string aSheet = String.Empty;
                if (!String.Equals(Addr.Sheet, val.Workbook.SheetName, StringComparison.CurrentCultureIgnoreCase)) aSheet = Addr.Sheet;
                int ROfs= RowDollar? 0: RowOfs;
                int COfs= ColDollar? 0: ColOfs;
                Address = new TCellAddress(aSheet, ROfs + Addr.Row, COfs + Addr.Col, RowDollar, ColDollar);
                
            }
            val.Value = Address.CellRef;
        } 
    }

    internal class TSectionHtml: TOneSectionValue
    {
        TIncludeHtml IncludeHtml;
        internal TSectionHtml(string aTagText, TOneCellValue aParent, bool aValue): base(aTagText, TValueType.Html, -1)
        {
            if (aValue) IncludeHtml = TIncludeHtml.Yes; else IncludeHtml = TIncludeHtml.No;
        }
    
        protected override void EvaluateInternal(int RowAbs, int ColAbs, int RowOfs, int ColOfs, TValueAndXF val)
        {
            val.IncludeHtml = IncludeHtml;
            base.EvaluateInternal(RowAbs, ColAbs, RowOfs, ColOfs, val);
        }    
    }

    internal class TSectionDefined: TOneSectionValue
    {
        private bool ValueIsDefined;

        internal TSectionDefined(TOneCellValue aParent, TRichString RTagText, TRichString TagParams, TStackData Stack, FlexCelReport fr, TBand CurrentBand): base(RTagText.ToString(), TValueType.Defined, -1)
        {
            List<TRichString> Params = new List<TRichString>();
            string TagText = RTagText.ToString();
            TCellParser.ParseParams(RTagText, TagParams, Params);
            if (Params.Count < 1)
                FlxMessages.ThrowException(FlxErr.ErrMissingArgs, TagText);
            if (Params.Count > 2) 
                FlxMessages.ThrowException(FlxErr.ErrTooMuchArgs, TagText);

            bool Global = true;
            if (Params.Count >= 2)
            {
                string GlobalStr = FlxConvert.ToString(Params[1]);
                if (String.Equals(GlobalStr, ReportTag.StrDefinedGlobal, StringComparison.InvariantCultureIgnoreCase)) Global = true;
                else
                    if (String.Equals(GlobalStr, ReportTag.StrDefinedLocal, StringComparison.InvariantCultureIgnoreCase)) Global = false;
                else FlxMessages.ThrowException(FlxErr.ErrInvalidDefinedGlobal, GlobalStr, ReportTag.StrDefinedLocal, ReportTag.StrDefinedGlobal);
            }

            ValueIsDefined = IsDefined(FlxConvert.ToString(Params[0]), Global, Stack, fr, CurrentBand);
        }
    
        protected override void EvaluateInternal(int RowAbs, int ColAbs, int RowOfs, int ColOfs, TValueAndXF val)
        {
            val.Action = ValueType;
            val.Value = ValueIsDefined;
        } 

        private static bool IsDefined(string TagText, bool Global, TStackData Stack, FlexCelReport fr, TBand CurrentBand)
        {
            if (TagText.IndexOf(ReportTag.DbSeparator)<0)
            {
                //Search order: User expression parameter, defined exp, values, user defined function.
                string VarNamePlusDefault=TagText;
                int defaultSepPos = VarNamePlusDefault.IndexOf(ReportTag.ParamDelim);
                string VarName = defaultSepPos < 0? VarNamePlusDefault: VarNamePlusDefault.Substring(0, defaultSepPos);

                TValueList ValueList=null; if (fr!=null) ValueList=fr.ValueList;
                TExpressionList ExpressionList=null; if (fr!=null) ExpressionList=fr.ExpressionList;
                TUserFunctionList UserFunctionList=null; if (fr!=null) UserFunctionList=fr.UserFunctionList;


                if (Stack.ExpParams != null && Stack.ExpParams.ContainsKey(VarName)) return true;
                else
                    if (ExpressionList != null && ExpressionList.ContainsKey(VarName)) return true;
                else
                    if (ValueList != null && ValueList.ContainsKey(VarName)) return true;
                else
                    if (UserFunctionList != null && UserFunctionList.ContainsKey(VarName)) return true;
                else
                    return false;

            }
            else return IsDefinedDbField(TagText, Global, CurrentBand, fr);
        }

        private static bool IsDefinedDbField(string TagText, bool Global, TBand aCurrentBand, FlexCelReport fr)
        {
            TBand CurrentBand = aCurrentBand;
            int sepPos=TagText.LastIndexOf(ReportTag.DbSeparator);
            if (sepPos<0) return false;
            string DataSetName=TagText.Substring(0, sepPos);
            
            string ColumnPlusDefault=TagText.Substring(sepPos+1);
            int defaultSepPos = ColumnPlusDefault.IndexOf(ReportTag.ParamDelim);
            string Column = defaultSepPos < 0? ColumnPlusDefault: ColumnPlusDefault.Substring(0, defaultSepPos);
            
            int ColumnId = -1;

            if (Global)
            {
                TDataSourceInfo ds = FindGlobalBand(DataSetName, fr);
                if (ds == null) return false;
                ColumnId = ds.Table.GetColumn(Column);
            }
            else
            {
                while (CurrentBand!=null && !String.Equals(CurrentBand.DataSourceName, DataSetName, StringComparison.CurrentCultureIgnoreCase))
                    CurrentBand=CurrentBand.SearchBand;

                if (CurrentBand == null || CurrentBand.DataSource == null) return false;
                ColumnId = CurrentBand.DataSource.GetColumnWithoutException(Column);
            }

            if (String.Equals(Column, ReportTag.StrFullDsCaptions, StringComparison.InvariantCultureIgnoreCase))  return true;
            if (String.Equals(Column, ReportTag.StrRowCountColumn, StringComparison.InvariantCultureIgnoreCase))  return true;
            if (String.Equals(Column, ReportTag.StrRowPosColumn, StringComparison.InvariantCultureIgnoreCase)) return true;
            if (String.Equals(Column, ReportTag.StrFullDs, StringComparison.InvariantCultureIgnoreCase)) return true;

            return ColumnId >= 0;

        }

        private static TDataSourceInfo FindGlobalBand(string DataTableName, FlexCelReport fr)
        {
            return fr.TryGetDataTable(DataTableName);
        }
    }

    internal class TSectionDefinedFormat : TOneSectionValue
    {
        private TOneCellValue FmtDef;
        TFormatList FmtList;

        private TSectionDefinedFormat(string aTagText, TOneCellValue aParent, TOneCellValue aFmtDef, TFormatList aFmtList)
            : base(aTagText, TValueType.DefinedFormat, -1)
        {
            FmtDef = aFmtDef;
            FmtList = aFmtList;
        }

        internal static TSectionDefinedFormat Create(string aTagText, TOneCellValue aParent, TOneCellValue aFmtDef, TFormatList aFmtList)
        {
            return new TSectionDefinedFormat(aTagText, aParent, aFmtDef, aFmtList);
        }

        protected override void EvaluateInternal(int RowAbs, int ColAbs, int RowOfs, int ColOfs, TValueAndXF val)
        {
            TValueAndXF val1 = new TValueAndXF();
            FmtDef.Evaluate(RowAbs, ColAbs, RowOfs, ColOfs, val1);
            bool IsDefined = FmtList.ContainsKey(FlxConvert.ToString(val1.Value));

            val.Action = ValueType;
            val.Value = IsDefined;
        }
    }


    #endregion

    /// <summary>
    /// A list of TOneSectionValue. Has a parsed version of what is on a cell.
    /// </summary>
    public class TOneCellValue: IDisposable
    {
        private List<TOneSectionValue> FList; 

        /// <summary>
        /// format of the cell.
        /// </summary>
        internal int XF;
        internal TFormatList FormatListRow;
        internal TFormatList FormatListCol;
        internal TOneCellValue XFDefRow;
        internal TOneCellValue XFDefCol;
 
        internal int Col;

        internal ExcelFile Workbook;

        internal bool IsPreprocess;

        internal TOneCellValue(ExcelFile aWorkbook)
        {
            FList= new List<TOneSectionValue>();
            Workbook=aWorkbook;
            IsPreprocess = false;
        }

        internal void Add(TOneSectionValue value)
        {
            FList.Add(value);
        }

        internal TOneSectionValue this[int index]
        {
            get
            {
                return FList[index];
            }
            set
            {
                FList[index]=value;
            }
        }

        internal int Count
        {
            get
            {
                return FList.Count;
            }
        }

        private static TConfigFormat CalcXF(TFormatList FormatList, TOneCellValue XFDef)
        {
            if (FormatList!=null)
            {
                TValueAndXF valr = new TValueAndXF();
                XFDef.Evaluate(0,0,0,0,valr);
                return FormatList.GetValue(FlxConvert.ToString(valr.Value));
            }
            else
            {
                return null;
            }
        }

        internal void Evaluate(int RowAbs, int ColAbs, int RowOfs, int ColOfs, TValueAndXF val)
        {
            Evaluate(RowAbs, ColAbs, RowOfs, ColOfs, val, true);
        }

        internal void Evaluate(int RowAbs, int ColAbs, int RowOfs, int ColOfs, TValueAndXF val, bool AllowObjectsAsResult)
        {
            if (XF >= 0) val.XF=XF;

            val.XFRow = CalcXF(FormatListRow, XFDefRow);
            val.XFCol = CalcXF(FormatListCol, XFDefCol);

            val.Value=null;
            if (FList.Count <= 0) return;

            int NotNullCount=0;
            object LastVal=null;
            
            StringBuilder sb= new StringBuilder();
            List<TRTFRun> RTFList = new List<TRTFRun>();

            foreach (TOneSectionValue Sec in FList)
            {
                val.Value=null;
                Sec.Evaluate(RowAbs, ColAbs, RowOfs, ColOfs, val);
                if (val.Value!=null)
                {
                    NotNullCount++;
                    LastVal=val.Value;

                    TRichString rs = val.Value as TRichString;
                    if (rs !=null)
                    {
                        for (int i=0; i< rs.RTFRunCount; i++)
                        {
                            TRTFRun RTFRun;
                            RTFRun.FontIndex=rs.RTFRun(i).FontIndex;
                            RTFRun.FirstChar=rs.RTFRun(i).FirstChar+sb.Length;
                            RTFList.Add(RTFRun);
                        }
                    }
                    else  //Not rich string
                        if (Sec.StartFont>=0)
                    {
                        TRTFRun RTFRun;
                        RTFRun.FontIndex=Sec.StartFont;
                        RTFRun.FirstChar=sb.Length;
                        RTFList.Add(RTFRun);
                        }
                    sb.Append(FlxConvert.ToStringWithArrays(val.Value));
                }
            } 

            if (NotNullCount<=1 && AllowObjectsAsResult &&  !(LastVal is string))
            {
                val.Value=LastVal;
                return; //This is the only case where we might not want to return a string/rich string
            }

            if (RTFList.Count<=0) val.Value= sb.ToString(); 
            else val.Value= new TRichString(sb.ToString(), RTFList, Workbook);
        }
        #region IDisposable Members

        /// <summary>
        /// Frees the resources of the cell.
        /// </summary>
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        /// <summary>
        /// Disposes the resources of the cell.
        /// </summary>
        /// <param name="isDisposing"></param>
        protected virtual void Dispose(bool isDisposing)
        {
            if (isDisposing)
            {
                foreach (TOneSectionValue Sec in FList)
                {
                    Sec.Dispose();
                }

                if (XFDefRow != null) XFDefRow.Dispose();
                if (XFDefCol != null) XFDefCol.Dispose();
            }
        }

        #endregion
    }

    /// <summary>
    /// A list of TOneCellValue. Contains all the cells, images and comments on a row
    /// </summary>
    internal class TOneRowValue: IDisposable
    {
        internal TOneCellValue[] Cols;
        internal int ColCount;

        internal TOneCellValue[] Comments;

        #region IDisposable Members

        public void Dispose()
        {
            if (Cols != null) foreach (TOneCellValue c in Cols)
            {
                if (c != null) c.Dispose();
            }
            if (Comments != null) foreach (TOneCellValue c in Comments)
            {
                if (c != null) c.Dispose();
            }
            GC.SuppressFinalize(this);
        }

        #endregion
    }

    internal sealed class TCellParser
    {
        private TCellParser(){}

        internal static int MinPositive(int a, int b)
        {
            if (a>=0)
            {
                if (b>=0) return Math.Min(a,b);
                else return a;
            }
            else 
                if (b>=0) return b;
            else return -1;
        }

        private static void AddString(object value, string s, int z, ref int p, TOneCellValue Result)
        {
            object v=s.Substring(p, z-p);
            TRichString rs = value as TRichString;
            if (rs!=null) 
            {
                v=new TRichString((string)v, rs, p);
            }
            Result.Add(new TSectionConst(Result, v));
            p=z;
        }

        internal static void GetSection(string TagText, ref int i, bool BreakOnBrace)
        {
            /*
             * There are 3 things to consider here:
             *    1) Quotes:  a "" means a single quote, and everything inside quotes does not modify parenthesis or brace level.
             *    2) Parenthesis: Whatever is inside them until the closing one is not taken into account.
             *    3) <# and #> 
             */

            int BraceLevel=0;
            int ParenLevel=0;
            bool InQuote=false;

            while(i<TagText.Length)
            {
                if (TagText[i]==ReportTag.StrQuote)
                {
                    if (!InQuote) InQuote=true;
                    else
                        if (i+1<TagText.Length && TagText[i+1]==ReportTag.StrQuote) {i++;} // this is an escaped ""
                    else
                        InQuote=false;
                }
                else
                    if (!InQuote)
                {
                    if (TagText[i]==ReportTag.StrOpenParen) 
                        ParenLevel++;
                    else
                        if (TagText[i]==ReportTag.StrCloseParen)
                    {
                        ParenLevel--;
                        if (ParenLevel<0) FlxMessages.ThrowException(FlxErr.ErrUnexpectedChar, TagText[i], i, TagText);
                    }

                    else 
                        if (ParenLevel== 0)
                    {
                        if (TagText.IndexOf(ReportTag.StrOpen,i)==i) 
                            BraceLevel++;
                        else
                            if (TagText.IndexOf(ReportTag.StrClose,i)==i && BraceLevel>0) //If bracelevel=0 we are on the main string, this could be a greater than sign.
                        {
                            BraceLevel--;
                            if (BraceLevel==0 && BreakOnBrace) break;
                            if (BraceLevel<0) FlxMessages.ThrowException(FlxErr.ErrUnexpectedChar, ReportTag.StrClose, i, TagText);
                        }
                        else if ((BraceLevel==0)&&(TagText[i]==ReportTag.ParamDelim))
                            break;
                    }
                }
                i++;
            }
            if (InQuote) FlxMessages.ThrowException(FlxErr.ErrUnterminatedString, TagText);
            if (ParenLevel>0) FlxMessages.ThrowException(FlxErr.ErrMissingParen, TagText);
            if (BraceLevel>0) FlxMessages.ThrowException(FlxErr.ErrMissingEOT, TagText);
        }

        internal static object UnQuote(TRichString Section)
        {
            if (Section == null) return Section;
            if ((Section.Length>2)&&(Section[0]==ReportTag.StrQuote)&&(Section[Section.Length-1]==ReportTag.StrQuote))
                return Section.Substring(1,Section.Length-2);
            else
                return Section;
        }

        internal static object TryConvert(ExcelFile Xls, TRichString Section, ref int XF)
        {
            if ((Section.Length>2)&&(Section[0]==ReportTag.StrQuote)&&(Section[Section.Length-1]==ReportTag.StrQuote))
                return Section.Substring(1,Section.Length-2);
            else
                return  Xls.ConvertString(Section, ref XF);
        }

        internal static void ParseTag(TRichString TagText, out string TagId, out TRichString Params)
        {
            string sTagText = TagText.ToString();
            int Tlen=sTagText.IndexOf(ReportTag.StrOpenParen);
            if (Tlen<0) Tlen=sTagText.Length;
            TagId=sTagText.Substring(0, Tlen).Trim().ToUpper(CultureInfo.InvariantCulture);

            Params=new TRichString();
            if (Tlen+1<sTagText.Length)
            {
                Params=TagText.Substring(Tlen).Trim();
                if (Params[0]!=ReportTag.StrOpenParen || Params[Params.Length-1]!=ReportTag.StrCloseParen) 
                    FlxMessages.ThrowException(FlxErr.ErrMissingParen, Params);
                Params=Params.Substring(1,Params.Length-2);
            }
        }

        private static bool ParseTag(TRichString TagText, out string TagId, out TRichString Params, ref TValueType TagType)
        {
            ParseTag(TagText, out TagId, out Params);
            return ReportTag.TryGetTag(TagId, out TagType);
        }

        internal static void ParseParams(TRichString TagText, TRichString Params, TRichString[] Sections)
        {                         
            ParseParams(TagText, Params, Sections, false);
        }

        internal static void ParseParams(TRichString TagText, TRichString Params, TRichString[] Sections, bool AllowLessParameters)
        {                         
            int i=0;                        
            for (int k=0;k<Sections.Length;k++)
            {
                int d=i;
                GetSection(Params.ToString(), ref i, false);

                if (i>Params.Length)
                {
                    if (AllowLessParameters) return;
                    FlxMessages.ThrowException(FlxErr.ErrMissingArgs, TagText);
                }
                Sections[k]=Params.Substring(d,i-d);
                i++; 
            }
            if (i<Params.Length) 
                FlxMessages.ThrowException(FlxErr.ErrTooMuchArgs, TagText);

        }

        /// <summary>
        /// For undetermined param count.
        /// </summary>
        internal static void ParseParams(TRichString TagText, TRichString Params, List<TRichString > Sections)
        {
            int i = 0;
            do
            {
                int d = i;
                GetSection(Params.ToString(), ref i, false);

                if (i <= Params.Length)
                    Sections.Add(Params.Substring(d, i - d));
                i++;
            }
            while (i <= Params.Length);
        }

        internal static TBandType GetBandType(string bt, string TagText)
        {
            if (String.Equals(bt, ReportTag.ColFull1, StringComparison.InvariantCultureIgnoreCase))
            {
                return TBandType.ColFull;
            }
            if (String.Equals(bt, ReportTag.RowFull1, StringComparison.InvariantCultureIgnoreCase))
            {
                return TBandType.RowFull;
            }
            if (String.Equals(bt, ReportTag.ColRange1, StringComparison.InvariantCultureIgnoreCase))
            {
                return TBandType.ColRange;
            }
            if (String.Equals(bt, ReportTag.RowRange1, StringComparison.InvariantCultureIgnoreCase))
            {
                return TBandType.RowRange;
            }

            FlxMessages.ThrowException(FlxErr.ErrUnknownRangeType, bt, TagText);
            return TBandType.Static; //Just to compile.
        }

        private static TValueList GetExpParams(TExpression Exp, TRichString TagText, TRichString TagParams,
            ExcelFile Workbook, TStackData Stack, TBand CurrentBand, FlexCelReport fr)
        {
            List<TRichString> Sections= new List<TRichString>();
            int XF5=-1;
            ParseParams(TagText, TagParams, Sections);

            TValueList Result = null;

            if (Exp.Parameters != null && Exp.Parameters.Length >0)
            {
                Result = new TValueList();
                if (Sections.Count < Exp.Parameters.Length)
                {
                    FlxMessages.ThrowException(FlxErr.ErrMissingArgs, TagText);
                }
                if (Sections.Count > Exp.Parameters.Length)
                {
                    FlxMessages.ThrowException(FlxErr.ErrTooMuchArgs, TagText);
                }

                for (int i=0; i<Exp.Parameters.Length; i++)
                {
                    Result.Add(
                        Exp.Parameters[i].Trim(), 
                        GetCellValue(TryConvert(Workbook, Sections[i], ref XF5), Workbook, Stack, XF5, CurrentBand, fr));
                }
            }
            else
            {
                if (TagParams.Trim().Length > 0)
                {
                    FlxMessages.ThrowException(FlxErr.ErrTooMuchArgs, TagText);
                }
            }
            return Result;
        }

        internal static TOneCellValue GetCellValue(object value, ExcelFile Workbook, TStackData Stack, int XF,
            TBand CurrentBand, FlexCelReport fr)
        {
            return GetCellValue(value, Workbook, Stack, XF, CurrentBand, fr, false);
        }

        internal static TOneCellValue GetCellValue(object value, ExcelFile Workbook, TStackData Stack, int XF,
            TBand CurrentBand, FlexCelReport fr, bool CanAddDataSets)
        {
            TFormatList FormatList=null; if (fr!=null) FormatList=fr.FormatList;
            TValueList ValueList=null; if (fr!=null) ValueList=fr.ValueList;
            TExpressionList ExpressionList=null; if (fr!=null) ExpressionList=fr.ExpressionList;
            TUserFunctionList UserFunctionList=null; if (fr!=null) UserFunctionList=fr.UserFunctionList;


            TOneCellValue Result= new TOneCellValue(Workbook);
            Result.XF=XF;
            Result.FormatListRow=null;
            Result.FormatListCol=null;
            Result.XFDefRow=null;
            Result.XFDefCol=null;

            if (!(value is string) && !(value is TRichString))
            {
                Result.Add(new TSectionConst(Result, value));
                return Result;
            }

            string s= value.ToString();
            TRichString rs= (value as TRichString);
            if (rs==null) rs= new TRichString(s);

            int p=0;
            do
            {
                int p1= s.IndexOf(ReportTag.StrOpen, p);
                int p3= -1;//s.IndexOf(ReportTag.StrOpen20b, p);

                int z=MinPositive(p1, p3);

                if (z!=p)  //copy the rest of the string
                {
                    if (z<0) z=s.Length;
                    AddString(value, s, z, ref p, Result);

                }
                else  //Start of a tag.
                {
                    int z2=z; //The ending pos
                    int TagLen=0;
                    int EndTagLen=0;
                    if (z==p1)
                    {
                        GetSection(s, ref z2, true);
                        //z2=s.IndexOf(ReportTag.StrClose, z);
                        EndTagLen=1;
                        TagLen=ReportTag.StrOpen.Length;
                        if ((z+ReportTag.StrOpen.Length<s.Length)&&(s[z+ReportTag.StrOpen.Length]=='#')) TagLen++;  //for <## tags
                    }
                    else 
                    {
                        z2=s.Length-1;
                        TagLen=2; //ReportTag.StrOpen20b.Length;
                    }
              
                    if (z2<0) //No ending tag, it will be considered text. Note that z CAN be>0
                    {
                        AddString(value, s, s.Length, ref p, Result);
                    }

                        ////////////////////////////////////////T A G /////////////////////////
                    else   //We found a tag.
                    {
                        TRichString TagText=rs.Substring(p+TagLen,z2-(p+TagLen)-EndTagLen+1); 
                        TRichString TagParams=null;
                        string TagId=String.Empty;
                        TValueType TagType=TValueType.Comment;
                        if (ParseTag(TagText, out TagId, out TagParams, ref TagType))
                        {
                            switch (TagType)
                            {
                                case TValueType.IF:
                                    Result.Add(
                                        TSectionIf.Create(Workbook, Stack, XF, CurrentBand, fr, Result, TagText, TagParams));
                                    break;

                                case TValueType.Evaluate:
                                    Result.Add(new TSectionEvaluate(TagText.ToString(), Result,
                                        new TRPNExpression(TagParams.ToString(), Workbook, CurrentBand, fr, Stack)));
                                    break;

                                case TValueType.Equal:
                                    Result.Add(
                                        TSectionEqual.Create(Workbook, Stack, XF, CurrentBand, fr, TagText.ToString(), Result, TagParams, -1));
                                    break;
                                
                                case TValueType.Comment:
                                    //Do nothing. It's a Comment.
                                    break;

                                case TValueType.Preprocess:
                                    Result.IsPreprocess = true;
                                    //Do not return anything.
                                    break;

                                case TValueType.Defined:
                                    Result.Add(new TSectionDefined(Result, TagText, TagParams, Stack, fr, CurrentBand));
                                    break;

                                case TValueType.DefinedFormat:
                                    if (FormatList != null)
                                    {
                                        int XF5 = -1;
                                        TOneCellValue FormatDef = GetCellValue(TagParams.ToString(), Workbook, Stack, XF5, CurrentBand, fr);
                                        Result.Add(TSectionDefinedFormat.Create(TagText.ToString(), Result, FormatDef, FormatList));
                                    }
                                    break;

                                case TValueType.FormatCell:
                                    if (FormatList!=null)
                                    {
                                        int XF5 = -1;
                                        TOneCellValue FormatDef = GetCellValue(TagParams.ToString(), Workbook, Stack, XF5, CurrentBand, fr);
                                        Result.Add(TSectionFormatCell.Create(TagText.ToString(), Result, new TFormatRange(FormatList, null, FormatDef)));
                                    }
                                    break;

                                case TValueType.Html:
                                    TRichString[] HtmlSections= new TRichString[1];
                                    ParseParams(TagText, TagParams, HtmlSections);
                                    bool IsHtml = false;
                                    if (String.Equals(HtmlSections[0].ToString(), TFormulaMessages.TokenString(TFormulaToken.fmTrue), StringComparison.InvariantCultureIgnoreCase))
                                    {
                                        IsHtml = true;
                                    }
                                    else
                                        if (!String.Equals(HtmlSections[0].ToString(), TFormulaMessages.TokenString(TFormulaToken.fmFalse), StringComparison.InvariantCultureIgnoreCase))
                                    {
                                        FlxMessages.ThrowException(FlxErr.ErrInvalidHtmlParam, HtmlSections[0].ToString(), TFormulaMessages.TokenString(TFormulaToken.fmFalse), TFormulaMessages.TokenString(TFormulaToken.fmTrue));
                                    }


                                    Result.Add(new TSectionHtml(TagText.ToString(), Result, IsHtml));
                                    break;
                                    
                                case TValueType.FormatRange:
                                    if (FormatList!=null)
                                    {
                                        TRichString[] FormatSections = new TRichString[2];
                                        ParseParams(TagText, TagParams, FormatSections);
                                        int XF5 = -1;
                                        TOneCellValue RangeDef = GetCellValue(FormatSections[0].ToString(), Workbook, Stack, XF5, CurrentBand, fr);
                                        TOneCellValue FormatDef = GetCellValue(FormatSections[1].ToString(), Workbook, Stack, XF5, CurrentBand, fr);
                                        Result.Add(new TSectionFormatRange(TagText.ToString(), Result, new TFormatRange(FormatList, RangeDef, FormatDef)));
                                    }
                                    break;


                                case TValueType.FormatRow:
                                    if (FormatList!=null)
                                    {
                                        int XF5 = -1;
                                        TOneCellValue FormatDef = GetCellValue(TagParams.ToString(), Workbook, Stack, XF5, CurrentBand, fr);
                                        Result.XFDefRow = FormatDef;
                                        Result.FormatListRow = FormatList;
                                        Result.Add(new TOneSectionValue(TagText.ToString(), TagType, -1));
                                    }
                                    break;

                                case TValueType.FormatCol:
                                    if (FormatList!=null)
                                    {
                                        int XF5 = -1;
                                        TOneCellValue FormatDef = GetCellValue(TagParams.ToString(), Workbook, Stack, XF5, CurrentBand, fr);
                                        Result.XFDefCol = FormatDef;
                                        Result.FormatListCol = FormatList;
                                        Result.Add(new TOneSectionValue(TagText.ToString(), TagType, -1));
                                    }
                                    break;

                                case TValueType.RowHeight:
                                    TRichString[] RHSections= new TRichString[5];
                                    ParseParams(TagText, TagParams, RHSections, true);
                                    Result.Add(new TSectionRowHeight(TagText.ToString(), Result, RHSections));
                                    break;

                                case TValueType.ColumnWidth:
                                    TRichString[] CWSections= new TRichString[4];
                                    ParseParams(TagText, TagParams, CWSections, true);
                                    Result.Add(new TSectionColWidth(TagText.ToString(), Result, CWSections));
                                    break;

                                case TValueType.AutofitSettings:
                                    TRichString[] AFSections= new TRichString[5];
                                    ParseParams(TagText, TagParams, AFSections, true);
                                    
                                    Result.Add(new TSectionAutofitSettings(TagText.ToString(), Result, AFSections[0].ToString().Trim(), FlxConvert.ToString(AFSections[1]).Trim(),FlxConvert.ToString(AFSections[2]).Trim(), FlxConvert.ToString(AFSections[3]).Trim(), FlxConvert.ToString(AFSections[4]).Trim()));
                                    break;

                                case TValueType.DeleteRange:
                                    int XF6 = -1;
                                    TRichString[] DrSections= new TRichString[2];
                                    ParseParams(TagText, TagParams, DrSections);
                                    Result.Add(new TSectionDeleteRange(TagText.ToString(), Result, new TDeleteRangeWaitingRange(GetCellValue(DrSections[0].ToString(), Workbook, Stack, XF6, CurrentBand, fr),GetBandType(DrSections[1].ToString(), TagText.ToString()))));
                                    break;

                                case TValueType.DeleteRow:
                                    {
                                        bool FullRange = GetDeleteRowColParams(TagText, TagParams);

                                        if (FullRange)
                                            Result.Add(new TSectionDeleteRow(TagText.ToString(), Result, new TDeleteRowWaitingRange(0, 1, FlxConsts.Max_Columns + 1)));
                                        else
                                            Result.Add(new TSectionDeleteRow(TagText.ToString(), Result, new TDeleteRowWaitingRange(0, CurrentBand.CellRange.Left, CurrentBand.CellRange.Right)));
                                    }
                                    break;

                                case TValueType.DeleteCol:
                                    {
                                        bool FullRange = GetDeleteRowColParams(TagText, TagParams);
                                        if (FullRange)
                                            Result.Add(new TSectionDeleteCol(TagText.ToString(), Result, new TDeleteColWaitingRange(0, 1, FlxConsts.Max_Rows + 1)));
                                        else
                                            Result.Add(new TSectionDeleteCol(TagText.ToString(), Result, new TDeleteColWaitingRange(0, CurrentBand.CellRange.Top, CurrentBand.CellRange.Bottom)));
                                    }
                                    break;

                                case TValueType.MergeRange:
                                    TRichString[] MrSections= new TRichString[1];
                                    ParseParams(TagText, TagParams, MrSections);
                                    Result.Add(new TSectionMergeRange(Workbook, Stack, CurrentBand, fr, TagText.ToString(), Result, MrSections[0].ToString()));
                                    break;

                                case TValueType.Formula:
                                    Result.Add(new TSectionFormula(TagText.ToString(), Result));
                                    break;

                                case TValueType.Ref:
                                    List<TRichString> RefParams= new List<TRichString>();
                                    ParseParams(TagText, TagParams, RefParams);
                                    Result.Add(new TSectionRef(Workbook ,Result, TagText.ToString(), RefParams));
                                    break;

                                case TValueType.ImgSize:
                                    Result.Add(TSectionImgSize.Create(Workbook, Stack, CurrentBand, fr, Result, TagText, TagParams));
                                    break;

                                case TValueType.ImgPos:
                                    Result.Add(TSectionImgPos.Create(Workbook, Stack, CurrentBand, fr, Result, TagText, TagParams));
                                    break;
                                
                                case TValueType.ImgFit:
                                    Result.Add(TSectionImgFit.Create(Workbook, Stack, CurrentBand, fr, Result, TagText, TagParams));
                                    break;

                                case TValueType.ImgDelete:
                                    Result.Add(new TSectionImgDelete(FlxConvert.ToString(TagText), Result));
                                    break;

                                case TValueType.Lookup:
                                    Result.Add(
                                        TSectionLookup.Create(Workbook, Stack, CurrentBand, fr, Result, TagText, TagParams));
                                    break;

                                case TValueType.Array:
                                    Result.Add(
                                        TSectionArray.Create(Workbook, Stack, CurrentBand, fr, Result, TagText, TagParams));
                                    break;


                                case TValueType.Include:
                                {
                                    if (fr!=null) 
                                    {
                                        Result.Add(
                                            TSectionInclude.Create(Workbook, Stack, CurrentBand, fr, Result, TagText, TagParams));
                                    }
                                    else
                                        Result.Add(new TOneSectionValue(TagText.ToString(), TagType, -1));

                                    break;
                                }

                                case TValueType.Regex:
                                    Result.Add(
                                        TSectionRegex.Create(Workbook, Stack, XF, CurrentBand, fr, Result, TagText, TagParams));
                                    break;

                                case TValueType.HPageBreak:
                                    Result.Add(
                                        new TSectionHPageBreak(TagText.ToString(), Result));
                                    break;

                                case TValueType.VPageBreak:
                                    Result.Add(
                                        new TSectionVPageBreak(TagText.ToString(), Result));
                                    break;

                                case TValueType.AutoPageBreaks:
                                    Result.Add(TSectionAutoPageBreaks.Create(TagText, Result, TagParams));
                                    break;

                                case TValueType.Aggregate:
                                    Result.Add(TSectionAggregate.Create(Workbook, Stack, XF, Result, TagText, TagParams, CurrentBand, fr));
                                    break;

                                case TValueType.List:
                                    Result.Add(TSectionList.Create(Workbook, Stack, XF, Result, TagText, TagParams, CurrentBand, fr));
                                    break;

                                case TValueType.DbValue:
                                    Result.Add(TSectionDbValue.Create(Workbook, Stack, XF, Result, TagText, TagParams, CurrentBand, fr));
                                    break;

                                default:
                                    Result.Add(new TOneSectionValue(TagText.ToString(), TagType, -1));
                                    break;
                            }
                        }
                        else  //Normal tag (database/dataset)
                        {
                            int StartFont=-1;
                            if (TagText.RTFRunCount>0 && TagText.RTFRun(0).FirstChar==0) StartFont=TagText.RTFRun(0).FontIndex;

                            if (TagText.ToString().IndexOf(ReportTag.DbSeparator)<0 || TagParams.Length>0)
                            {
                                //Search order: User expression parameter, defined exp, values, user defined function.
                                string VarNamePlusDefault=TagText.ToString();
                                int defaultSepPos = VarNamePlusDefault.IndexOf(ReportTag.ParamDelim);
                                string VarName = defaultSepPos < 0? VarNamePlusDefault: VarNamePlusDefault.Substring(0, defaultSepPos);

                                if (Stack.ExpParams != null && Stack.ExpParams.ContainsKey(VarName)) 
                                    Result.Add(new TSectionEqual(TagText.ToString(), Result, (TOneCellValue)Stack.ExpParams[VarName], StartFont));
                                else
                                    if (ExpressionList != null && ExpressionList.ContainsKey(TagId))                               
                                {
                                    if (Stack.UsedRefs.ContainsKey(TagId)) FlxMessages.ThrowException(FlxErr.ErrCircularReference, TagId);
                                    Stack.UsedRefs.Add(TagId, TagId);

                                    TExpression Exp = ExpressionList[TagId];
                                    TValueList ExpParams = GetExpParams(Exp, TagText, TagParams, Workbook, Stack, CurrentBand, fr);
                                    Result.Add(new TSectionEqual(TagText.ToString(), Result, GetCellValue(Exp.Value, Workbook, new TStackData(Stack.UsedRefs, ExpParams), XF, CurrentBand, fr), StartFont));

                                    Stack.UsedRefs.Remove(TagId);
                                }
                                else
                                    if (ValueList != null && ValueList.ContainsKey(VarName)) 
                                    Result.Add(new TSectionProperty(TagText.ToString(), Result, ValueList[VarName], StartFont));
                                else
                                    if (UserFunctionList != null && UserFunctionList.ContainsKey(TagId))
                                    Result.Add(TSectionUserFunction.Create(Workbook, Stack, CurrentBand, fr, 
                                        (TFlexCelUserFunction) UserFunctionList[TagId], Result, TagText, TagParams, StartFont));

                                else
                                    if (defaultSepPos >= 0)
                                {
                                    TRichString DefaultValue = TagText.Substring(defaultSepPos + 1);
                                    Result.Add(new TSectionEqual(TagText.ToString(), Result, TCellParser.GetCellValue(TCellParser.TryConvert(Workbook, DefaultValue, ref XF), Workbook, Stack, XF, CurrentBand, fr), StartFont)); 
                                }

                                else
                                    FlxMessages.ThrowException(FlxErr.ErrPropertyNotFound, VarName);

                            }
                            else               
                                Result.Add(TSectionDataSet.Create(Workbook, Stack, XF, Result, TagText, CurrentBand, StartFont, fr, CanAddDataSets));
                        }
                        p=z2+1;
                    }
                    
                }
            }
            while (p<s.Length);
            return Result;
        }

        private static bool GetDeleteRowColParams(TRichString TagText, TRichString TagParams)
        {
            List<TRichString> RParams = new List<TRichString>();
            ParseParams(TagText, TagParams, RParams);
            bool FullRange = true;
            if (RParams.Count == 1)
            {
                string p = FlxConvert.ToString(RParams[0]);
                if (p == null || p.Trim().Length == 0 || String.Equals(p, ReportTag.StrFullDelete, StringComparison.InvariantCultureIgnoreCase))
                {
                    FullRange = true;
                }
                else
                    if (String.Equals(p, ReportTag.StrRelativeDelete, StringComparison.InvariantCultureIgnoreCase))
                    {
                        FullRange = false;
                    }
                    else
                        FlxMessages.ThrowException(FlxErr.ErrInvalidRowColDelete, p, ReportTag.StrRelativeDelete, ReportTag.StrFullDelete);
            }
            else
                if (RParams.Count > 1) FlxMessages.ThrowException(FlxErr.ErrTooMuchArgs, TagText.ToString());
            return FullRange;
        }

    }    
}


