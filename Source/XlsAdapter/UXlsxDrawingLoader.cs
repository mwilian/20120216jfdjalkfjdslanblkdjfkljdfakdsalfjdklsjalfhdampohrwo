using System;
using System.IO;
using System.Text;
using System.Collections.Generic;
using System.Diagnostics;

using FlexCel.Core;
using System.Drawing;
using real = System.Single;
using System.Globalization;

namespace FlexCel.XlsAdapter
{
    /// <summary>
    /// A class to load DrawingML files
    /// </summary>
    internal class TXlsxDrawingLoader
    {
        #region Variables
        private ExcelFile xls;
        private TOpenXmlReader DataStream;
        #endregion

        #region Constructors
        internal TXlsxDrawingLoader(TOpenXmlReader aDataStream, ExcelFile axls)
        {
            DataStream = aDataStream;
            xls = axls;
        }
        #endregion

        #region Drawing
        internal void ReadDrawing(string relId, TSheet Sheet, int WorkingSheet)
        {
            DataStream.SelectFromCurrentPartAndPush(relId, TOpenXmlManager.SpreadsheetDrawingNamespace, false);
            while (DataStream.NextTag())
            {
                switch (DataStream.RecordName())
                {
                    case "wsDr":
                        ReadWsDr(Sheet, WorkingSheet);
                        break;

                    default: //shouldn't happen
                        DataStream.GetXml();
                        break;
                }
            }

            DataStream.PopPart();
        }

        private void ReadWsDr(TSheet Sheet, int WorkingSheet)
        {
            if (DataStream.IsSimpleTag) { DataStream.NextTag(); return; }
            string StartElement = DataStream.RecordName();
            if (!DataStream.NextTag()) return;
            while (!DataStream.AtEndElement(StartElement))
            {
                switch (DataStream.RecordName())
                {
                    case "twoCellAnchor":
                        ReadCellAnchor(Sheet, WorkingSheet, 2);
                        break;

                    case "oneCellAnchor":
                        ReadCellAnchor(Sheet, WorkingSheet, 1);
                        break;

                    case "absoluteAnchor":
                        ReadCellAnchor(Sheet, WorkingSheet, 0);
                        break;

                    default:
                        DataStream.GetXml();
                        break;
                }
            }
        }

        private void ReadCellAnchor(TSheet Sheet, int WorkingSheet, int CellCount)
        {
            TFlxAnchorType AnchorType = GetFlxAnchor(DataStream.GetAttribute("editAs"));

            if (DataStream.IsSimpleTag) { DataStream.NextTag(); return; }
            string StartElement = DataStream.RecordName();
            if (!DataStream.NextTag()) return;

            TObjectProperties ObjProps = new TObjectProperties();
            ObjProps.ShapeOptions = new TShapeProperties();
            ObjProps.ShapeOptions.ShapeOptions = new TShapeOptionList();

            int Row1 = 0; int Col1 = 0; TDrawingCoordinate Row1Offs = new TDrawingCoordinate(); TDrawingCoordinate Col1Offs = new TDrawingCoordinate();
            int Row2 = 0; int Col2 = 0; TDrawingCoordinate Row2Offs = new TDrawingCoordinate(); TDrawingCoordinate Col2Offs = new TDrawingCoordinate();
            TDrawingCoordinate cx = new TDrawingCoordinate(); TDrawingCoordinate cy = new TDrawingCoordinate();
            TDrawingCoordinate px = new TDrawingCoordinate(); TDrawingCoordinate py = new TDrawingCoordinate();
            while (!DataStream.AtEndElement(StartElement))
            {
                switch (DataStream.RecordName())
                {
                    case "from": if (CellCount >= 1) ReadMarker(ref Row1, ref Col1, ref Row1Offs, ref Col1Offs); break;
                    case "to": if (CellCount >= 2) ReadMarker(ref Row2, ref Col2, ref Row2Offs, ref Col2Offs); break;
                    case "pos": if (CellCount <= 0) ReadPos(out px, out py); break;
                    case "ext": if (CellCount <= 1) ReadExt(out cx, out cy); break;

                    case "clientData": ReadClientData(ObjProps); break;
                    default:
                        if (ReadEG_ObjectChoices(ObjProps)) break;

                        DataStream.GetXml();
                        break;
                }
            }

            IRowColSize RowColSizer = new RowColSize(xls.HeightCorrection, xls.WidthCorrection, Sheet);
            if (CellCount == 0)
            {
                TClientAnchor a = new TClientAnchor(AnchorType, 
                    0, 0, 0, 0,
                    (int)Math.Round(py.Pixels), (int)Math.Round(px.Pixels),
                    RowColSizer).Dec();

                Row1 = a.Row2;
                Row1Offs = TDrawingCoordinate.FromPixels(a.Dy2Pix(RowColSizer));
                Col1 = a.Col2;
                Col1Offs = TDrawingCoordinate.FromPixels(a.Dx2Pix(RowColSizer));
            }

            if (CellCount <= 1)
            {
                switch (CellCount)
                {
                    case 0: AnchorType = TFlxAnchorType.DontMoveAndDontResize; break;
                    case 1: AnchorType = TFlxAnchorType.MoveAndDontResize; break;
                }
                if (cx.Emu <= 0 || cy.Emu <= 0) return;
                ObjProps.Anchor = new TClientAnchor(AnchorType, 
                    Row1, (int)Math.Round(Row1Offs.Pixels), Col1, (int)Math.Round(Col1Offs.Pixels),
                    (int)Math.Round(cy.Pixels), (int)Math.Round(cx.Pixels),
                    RowColSizer).Dec();
            }
            else
            {
                if (Row1 <= 0 || Row2 <= 0) return;
                ObjProps.Anchor = new TClientAnchor(AnchorType, Row1, (int)Math.Round(Row1Offs.Pixels), Col1, (int)Math.Round(Col1Offs.Pixels),
                    Row2, (int)Math.Round(Row2Offs.Pixels), Col2, (int)Math.Round(Col2Offs.Pixels),
                    RowColSizer, false).Dec();
            }
            Sheet.Drawing.AddObject(xls, Sheet, WorkingSheet, ObjProps, null, true);
        }

        private void ReadPos(out TDrawingCoordinate px, out TDrawingCoordinate py)
        {
            px = DataStream.GetAttributeAsDrawingCoord("x", new TDrawingCoordinate(0));
            py = DataStream.GetAttributeAsDrawingCoord("y", new TDrawingCoordinate(0));

            DataStream.FinishTag();
        }

        private void ReadExt(out TDrawingCoordinate cx, out TDrawingCoordinate cy)
        {
            cx = DataStream.GetAttributeAsDrawingCoord("cx", new TDrawingCoordinate(0));
            cy = DataStream.GetAttributeAsDrawingCoord("cy", new TDrawingCoordinate(0));

            DataStream.FinishTag();
        }

        private void ReadMarker(ref int Row, ref int Col, ref TDrawingCoordinate RowOffs, ref TDrawingCoordinate ColOffs)
        {
            if (DataStream.IsSimpleTag) { DataStream.NextTag(); return; }
            string StartElement = DataStream.RecordName();
            if (!DataStream.NextTag()) return;

            while (!DataStream.AtEndElement(StartElement))
            {
                switch (DataStream.RecordName())
                {
                    case "col":   Col = DataStream.ReadValueAsInt() + 1; break;
                    case "colOff": ColOffs = TOpenXmlReader.GetAttributeAsDrawingCoord(DataStream.ReadValueAsString()); break;
                    case "row": Row = DataStream.ReadValueAsInt() + 1; break;
                    case "rowOff": RowOffs = TOpenXmlReader.GetAttributeAsDrawingCoord(DataStream.ReadValueAsString()); break;
                    default:
                        DataStream.GetXml();
                        break;
                }
            }
        }

        private TFlxAnchorType GetFlxAnchor(string anchorType)
        {
            switch (anchorType)
            {
                case "absolute": return TFlxAnchorType.DontMoveAndDontResize;
                case "oneCell": return TFlxAnchorType.MoveAndDontResize;

                default: return TFlxAnchorType.MoveAndResize;
            }       
        }

        private void ReadClientData(TImageProperties ImgProps)
        {
            ImgProps.Lock = DataStream.GetAttributeAsBool("fLocksWithSheet", true);
            ImgProps.Print = DataStream.GetAttributeAsBool("fPrintsWithSheet", true);
            DataStream.FinishTag();
        }

        private bool ReadEG_ObjectChoices(TObjectProperties ObjProps)
        {
            ObjProps.Published = true; 
            ObjProps.Macro = null;
            switch (DataStream.RecordName())
            {
                case "sp": ReadSp(ObjProps); return true;
                case "grpSp": ReadGrpSp(ObjProps); return true;
                case "graphicFrame": ReadGraphicFrame(ObjProps); return true;
                case "cxnSp": ReadCxnSp(ObjProps); return true;
                case "pic": ReadPic(ObjProps); return true;
                //case "contentPart": ReadContentPart();
            }
            return false;
        }

        private void ReadCxnSp(TObjectProperties ObjProps)
        {
            if (DataStream.IsSimpleTag) { DataStream.NextTag(); return; }
            string StartElement = DataStream.RecordName();
            if (!DataStream.NextTag()) return;

            while (!DataStream.AtEndElement(StartElement))
            {
                switch (DataStream.RecordName())
                {
                    case "nvCxnSpPr": ReadNvCxnSpPr(ObjProps); break;
                    case "spPr": ReadSpPr(ObjProps); break;
                    case "style": ReadStyle(ObjProps); break;
                    default: DataStream.GetXml(); break;
                }
            }
        }

        private void ReadNvCxnSpPr(TObjectProperties ObjProps)
        {
            if (DataStream.IsSimpleTag) { DataStream.NextTag(); return; }
            string StartElement = DataStream.RecordName();
            if (!DataStream.NextTag()) return;

            while (!DataStream.AtEndElement(StartElement))
            {
                switch (DataStream.RecordName())
                {
                    case "cNvPr": ReadCNvPr(ObjProps); break;
                    case "cNvCxnSpPr": ReadCNvCxnSpPr(); break;
                    default: DataStream.GetXml(); break;
                }
            }
        }

        private void ReadCNvCxnSpPr()
        {
            DataStream.GetXml();
        }

        private void ReadGraphicFrame(TObjectProperties ObjProps)
        {
            string SaveNamespace = DataStream.DefaultNamespace;
            DataStream.DefaultNamespace = TOpenXmlManager.DrawingNamespace;
            try
            {

                ObjProps.Macro = DataStream.GetAttribute("macro");
                ObjProps.Published = DataStream.GetAttributeAsBool("fPublished", false);
                if (DataStream.IsSimpleTag) { DataStream.NextTag(); return; }
                string StartElement = DataStream.RecordName();
                if (!DataStream.NextTag()) return;

                while (!DataStream.AtEndElement(StartElement))
                {
                    switch (DataStream.RecordName())
                    {
                        case "nvGraphicFramePr": ReadNvGraphicFramePr(ObjProps); break;
                        case "xfrm": ReadXfrm(ObjProps); break;
                        case "graphic": ReadGraphic(); break;
                        default: DataStream.GetXml(); break;
                    }
                }
            }
            finally
            {
                DataStream.DefaultNamespace = SaveNamespace;
            }
        }

        private void ReadNvGraphicFramePr(TObjectProperties ObjProps)
        {
            if (DataStream.IsSimpleTag) { DataStream.NextTag(); return; }
            string StartElement = DataStream.RecordName();
            if (!DataStream.NextTag()) return;

            while (!DataStream.AtEndElement(StartElement))
            {
                switch (DataStream.RecordName())
                {
                    case "cNvPr": ReadCNvPr(ObjProps); break;
                    case "cNvGraphicFramePr": ReadCNvGraphicFramePr(); break; 
                    default: DataStream.GetXml(); break;
                }
            }
        }

        private void ReadCNvGraphicFramePr()
        {
            if (DataStream.IsSimpleTag) { DataStream.NextTag(); return; }
            string StartElement = DataStream.RecordName();
            if (!DataStream.NextTag()) return;

            while (!DataStream.AtEndElement(StartElement))
            {
                switch (DataStream.RecordName())
                {
                    case "graphicFrameLocks": ReadGraphicFrameLocks(); break;
                    //case "extLst": ReadExtLst(); break;
                    default: DataStream.GetXml(); break;
                }
            }            
        }

        private void ReadGraphicFrameLocks()
        {
            DataStream.GetXml();
        }

        private void ReadGraphic()
        {
            if (DataStream.IsSimpleTag) { DataStream.NextTag(); return; }
            string StartElement = DataStream.RecordName();
            if (!DataStream.NextTag()) return;

            while (!DataStream.AtEndElement(StartElement))
            {
                switch (DataStream.RecordName())
                {
                    case "graphicData": ReadGraphicData(); break;
                    default: DataStream.GetXml(); break;
                }
            }
        }

        private void ReadGraphicData()
        {
            string GraphicNamespace = DataStream.GetAttribute("uri");

            if (GraphicNamespace == TOpenXmlManager.ChartNamespace)
            {
                ReadChart();                
            }
            else
            {
                DataStream.FinishTagAndIgnoreChildren();
            }
        }

        private TFlxChart ReadChart()
        {
            TFlxChart Result = null;

            if (DataStream.IsSimpleTag) { DataStream.NextTag(); return Result; }
            string StartElement = DataStream.RecordName();
            if (!DataStream.NextTag()) return Result;

            while (!DataStream.AtEndElement(StartElement))
            {
                switch (DataStream.RecordName())
                {
                    /*case TOpenXmlManager.ChartNamespace + ":chart":
                        TXlsxChartReader ChartReader = new TXlsxChartReader(DataStream, xls);
                        string relId = DataStream.GetRelationship("id");
                        ChartReader.ReadChart(relId, Result);
                        DataStream.FinishTag();
                        break;*/

                    default: DataStream.GetXml(); break;
                }
            }

            return Result;
        }

        private void ReadPic(TObjectProperties ObjProps)
        {
            ObjProps.Macro = DataStream.GetAttribute("macro");
            ObjProps.Published = DataStream.GetAttributeAsBool("fPublished", false);

            if (DataStream.IsSimpleTag) { DataStream.NextTag(); return; }
            string StartElement = DataStream.RecordName();
            if (!DataStream.NextTag()) return;

            while (!DataStream.AtEndElement(StartElement))
            {
                switch (DataStream.RecordName())
                {
                    case "nvPicPr": ReadNvPicPr(ObjProps); break;
                    case "blipFill": ObjProps.BlipFill = ReadBlipFill(); break;
                    case "spPr": ReadSpPr(ObjProps); break;
                    case "style": ReadStyle(ObjProps); break;
                    default: DataStream.GetXml(); break;
                }
            }
        }

        private void ReadNvPicPr(TObjectProperties ObjProps)
        {
            if (DataStream.IsSimpleTag) { DataStream.NextTag(); return; }
            string StartElement = DataStream.RecordName();
            if (!DataStream.NextTag()) return;

            while (!DataStream.AtEndElement(StartElement))
            {
                switch (DataStream.RecordName())
                {
                    case "cNvPr": ReadCNvPr(ObjProps); break;
                    case "cNvPicPr": ReadCNvPicPr(); break;
                    case "nvPr": ReadNvPr(ObjProps); break;
                    default: DataStream.GetXml(); break;
                }
            }
        }

        private void ReadCNvPr(TObjectProperties ObjProps)
        {
            string SaveNamespace = DataStream.DefaultNamespace;
            DataStream.DefaultNamespace = TOpenXmlManager.DrawingNamespace;
            try
            {
                int id = DataStream.GetAttributeAsInt("id", -1);
                if (id < 0) XlsMessages.ThrowException(XlsErr.ErrExcelInvalid);
                ObjProps.ShapeName = DataStream.GetAttribute("name");
                ObjProps.AltText = DataStream.GetAttribute("descr");
                ObjProps.FHidden = DataStream.GetAttributeAsBool("hidden", false);
                DataStream.GetAttribute("title");

                if (DataStream.IsSimpleTag) { DataStream.NextTag(); return; }
                string StartElement = DataStream.RecordName();
                if (!DataStream.NextTag()) return;

                while (!DataStream.AtEndElement(StartElement))
                {
                    switch (DataStream.RecordName())
                    {
                        case "hlinkClick": ObjProps.HLinkClick = ReadHlink(); break;
                        case "hlinkHover": ObjProps.HLinkHover = ReadHlink(); break;
                        case "extLst":
                        default: //ReadExtLst(); 
                            DataStream.GetXml();
                            break;
                    }
                }
            }
            finally
            {
                DataStream.DefaultNamespace = SaveNamespace;
            }
        }

        private TDrawingHyperlink ReadHlink()
        {
            return new TDrawingHyperlink(DataStream.GetXml());
        }
        
        private void ReadCNvPicPr()
        {
            DataStream.GetAttributeAsBool("preferRelativeResize", true);

            if (DataStream.IsSimpleTag) { DataStream.NextTag(); return; }
            string StartElement = DataStream.RecordName();
            if (!DataStream.NextTag()) return;

            while (!DataStream.AtEndElement(StartElement))
            {
                switch (DataStream.RecordName())
                {
                    case "picLocks": ReadPicLocks(); break;

                    case "extLst":
                    default: 
                      //ReadExtLst();
                        DataStream.GetXml();
                        break;
                }
            }
        }

        private void ReadPicLocks()
        {
            ReadAGLocking();
            if (DataStream.IsSimpleTag) { DataStream.NextTag(); return; }
            string StartElement = DataStream.RecordName();
            if (!DataStream.NextTag()) return;

            while (!DataStream.AtEndElement(StartElement))
            {
                //ReadExtLst(Globals);
                DataStream.GetXml();
            }            
        }

        private void ReadAGLocking()
        {
            DataStream.GetAttributeAsBool("noGrp", false);
            DataStream.GetAttributeAsBool("noSelect", false);
            DataStream.GetAttributeAsBool("noRot", false);
            DataStream.GetAttributeAsBool("noChangeAspect", false);
            DataStream.GetAttributeAsBool("noMove", false);
            DataStream.GetAttributeAsBool("noResize", false);
            DataStream.GetAttributeAsBool("noEditPoints", false);
            DataStream.GetAttributeAsBool("noAdjustHandles", false);
            DataStream.GetAttributeAsBool("noChangeArrowheads", false);
            DataStream.GetAttributeAsBool("noChangeShapeType", false);
        }
        
        private void ReadNvPr(TImageProperties ImgProps)
        {
            DataStream.GetXml();
        }

        private void ReadSpPr(TObjectProperties ObjProps)
        {
            string SaveNamespace = DataStream.DefaultNamespace;
            DataStream.DefaultNamespace = TOpenXmlManager.DrawingNamespace;
            try
            {
                ObjProps.ShapeOptions.ShapeOptions.SetLong(TShapeOption.bWMode, (long)GetBwMode(DataStream.GetAttribute("bwMode")), 0);

                if (DataStream.IsSimpleTag) { DataStream.NextTag(); return; }
                string StartElement = DataStream.RecordName();
                if (!DataStream.NextTag()) return;

                while (!DataStream.AtEndElement(StartElement))
                {
                    TFillStyle Fill;
                    if (ReadEG_Fill(out Fill))
                    {
                        if (Fill != null) ObjProps.ShapeFill = new TShapeFill(!(Fill is TNoFill), Fill);
                        continue;
                    }

                    switch (DataStream.RecordName())
                    {
                        case "xfrm": ReadXfrm(ObjProps); break;

                        // "EG_Geometry": 
                        case "custGeom": ReadCustGeom(ObjProps); break;
                        case "prstGeom": ReadPrstGeom(ObjProps); break;

                        case "ln": ReadLn(ObjProps); break;
                        
                        case "effectLst": 
                        case "effectDag": ObjProps.FEffectProperties = ReadEG_EffectProperties(); break;

                        case "scene3d": ReadScene3d(ObjProps.ShapeOptions.ShapeOptions); break;
                        case "sp3d": ReadSp3d(ObjProps.ShapeOptions.ShapeOptions); break;
                        case "extLst":
                        
                        default: DataStream.GetXml(); break;
                    }
                }
            }
            finally
            {
                DataStream.DefaultNamespace = SaveNamespace;
            }
        }

        #region Xfrm
        private void ReadXfrm(TObjectProperties ObjProps)
        {
            TShapeProperties ShProps = ObjProps.ShapeOptions;
            ShProps.ShapeOptions.Set1616(TShapeOption.Rotation, DataStream.GetAttributeAsAngle("rot", 0), 0);
            ShProps.FlipH = DataStream.GetAttributeAsBool("flipH", false);
            ShProps.FlipV = DataStream.GetAttributeAsBool("flipV", false);

            if (DataStream.IsSimpleTag) { DataStream.NextTag(); return; }
            string StartElement = DataStream.RecordName();
            if (!DataStream.NextTag()) return;

            while (!DataStream.AtEndElement(StartElement))
            {
                switch (DataStream.RecordName())
                {
                    case "off": ObjProps.Offs = ReadPoint(); DataStream.FinishTag(); break;
                    case "ext": ObjProps.Ext = ReadSize(); DataStream.FinishTag(); break;
                    case "chOff": ObjProps.ChOffs = ReadPoint(); DataStream.FinishTag(); break;
                    case "chExt": ObjProps.ChExt = ReadSize(); DataStream.FinishTag(); break;
                    default: DataStream.GetXml(); break;
                }
            }

        }

        private TDrawingPoint ReadPoint()
        {
            return new TDrawingPoint(DataStream.GetAttributeAsDrawingCoord("x", new TDrawingCoordinate(0)), DataStream.GetAttributeAsDrawingCoord("y", new TDrawingCoordinate(0)));
        }

        private Size ReadSize()
        {
            return new Size(DataStream.GetAttributeAsInt("cx", 0), DataStream.GetAttributeAsInt("cy", 0));
        }
        #endregion

        #region Geometry

        private void ReadCustGeom(TObjectProperties ObjProps)
        {
            TXlsxShapeReader shp = new TXlsxShapeReader(DataStream);
            ObjProps.ShapeOptions.ShapeGeom = shp.ReadShapeDef(null);
        }

        private void ReadPrstGeom(TObjectProperties ObjProps)
        {
            string prst = DataStream.GetAttribute("prst");
            ObjProps.ShapeOptions.ShapeType = GetAutoShapeKind(prst);

            TXlsxShapeReader shp = new TXlsxShapeReader(DataStream);
            ObjProps.ShapeOptions.ShapeGeom = shp.ReadShapeDef(prst);
            ReadGeomGuideList(ObjProps);
        }

        #region Autoshape definitions
        private TShapeType GetAutoShapeKind(string st)
        {
            TShapeType sp;
            if (TDrawingPresetGeom.FromString.TryGetValue(st, out sp)) return sp;
            return TShapeType.NotPrimitive;
        }
        #endregion

        private void ReadGeomGuideList(TObjectProperties ObjProps)
        {
            foreach (TShapeGuide guide in ObjProps.ShapeOptions.ShapeGeom.AvList)
            {
                ReadGd(ObjProps.ShapeOptions, guide);
            }
        }

        private void ReadGd(TShapeProperties ShProps, TShapeGuide guide)
        {
            long val = 0;
            long def;
            TShapeVal FmlaVal = guide.Fmla as TShapeVal;
            if (FmlaVal != null)
            {
                double vald = FmlaVal.ConstantVal;
                val = TShapePresets.ConvertAdjustToBiff8(ShProps.ShapeType, vald, out def);
            }
            else return;

            TShapeOption so;
            if (TShapePresets.GetBiff8Adjust(guide.Name, out so))
            {
                ShProps.ShapeOptions.SetLong(so, val, def);
            }
        }



        private void ReadLn(TObjectProperties ObjProps)
        {
             TLineStyle ls = ReadLineStyle();
             if (ObjProps.FShapeLine != null) ObjProps.FShapeLine.LineStyle = ls;
             else ObjProps.FShapeLine = new TShapeLine(true, ls);
        }

        private TEffectProperties ReadEG_EffectProperties()
        {
            return new TEffectProperties(DataStream.GetXml());
        }

        private void ReadScene3d(TShapeOptionList ShapeOptions)
        {
            DataStream.GetXml();
        }

        private void ReadSp3d(TShapeOptionList ShapeOptions)
        {
            DataStream.GetXml();
        }


        private TBwMode GetBwMode(string tag)
        {
           switch (tag)
            {
                case "clr": return TBwMode.Color;
                case "auto": return TBwMode.Automatic;
                case "gray": return TBwMode.GrayScale;
                case "ltGray": return TBwMode.LightGrayScale;
                case "invGray": return TBwMode.InverseGray;
                case "grayWhite": return TBwMode.GrayOutline;
                case "blackGray": return TBwMode.BlackTextLine;
                case "blackWhite": return TBwMode.HighContrast;
                case "black": return TBwMode.Black;
                case "white": return TBwMode.White;
                case "hidden": return TBwMode.DontShow;
            }

            return TBwMode.Color;
        }
        #endregion

        private void ReadStyle(TObjectProperties ObjProps)
        {
            string SaveNamespace = DataStream.DefaultNamespace;
            DataStream.DefaultNamespace = TOpenXmlManager.DrawingNamespace;
            try
            {
                if (DataStream.IsSimpleTag) { DataStream.NextTag(); return; }
                string StartElement = DataStream.RecordName();
                if (!DataStream.NextTag()) return;

                while (!DataStream.AtEndElement(StartElement))
                {
                    switch (DataStream.RecordName())
                    {
                        case "lnRef" : ReadLnRef(ObjProps); break;
                        case "fillRef": ReadFillRef(ObjProps); break;
                        case "effectRef" : ReadEffectRef(ObjProps); break;
                        case "fontRef": ReadFontRef(ObjProps); break;

                        default: DataStream.GetXml(); break;
                    }
                }
            }
            finally
            {
                DataStream.DefaultNamespace = SaveNamespace;
            }
        }

        private void ReadEffectRef(TObjectProperties ObjProps)
        {
            int Idx = DataStream.GetAttributeAsInt("idx", 0);
            TDrawingColor dc = GetDrawingColor();

            ObjProps.FShapeEffects = new TShapeEffects((TFormattingType)Idx, dc);
        }

        private void ReadFontRef(TObjectProperties ObjProps)
        {
            TFontScheme Idx = TXlsxFontReaderWriter.GetFontScheme(DataStream.GetAttribute("idx"));
            TDrawingColor dc = GetDrawingColor();

            ObjProps.ShapeOptions.ShapeThemeFont = new TShapeFont(Idx, dc);
        }

        private TShapeFill GetShapeFill(int idx, TDrawingColor dc, TShapeFill OldShapeFill)
        {
            bool UsesBk = idx >= 1000;
            if (UsesBk) idx -= 1000;
            TFillStyle OldShapeFillStyle = OldShapeFill == null ? null : OldShapeFill.FillStyle;
            if (idx == 0) return new TShapeFill(OldShapeFillStyle, false, (TFormattingType)idx, dc, UsesBk);
            return new TShapeFill(OldShapeFillStyle, !(OldShapeFillStyle is TNoFill), (TFormattingType)(idx - 1), dc, UsesBk);
        }


        private void ReadLnRef(TObjectProperties ObjProps)
        {
            int Idx = DataStream.GetAttributeAsInt("idx", 0);
            TDrawingColor dc = GetDrawingColor();

            if (ObjProps.FShapeLine == null)
            {
                ObjProps.FShapeLine = new TShapeLine(Idx != 0 && Idx != 1000, null, dc, GetFormattingType(Idx));
            }
            else
            {
                ObjProps.FShapeLine.ThemeStyle = GetFormattingType(Idx);
                ObjProps.FShapeLine.ThemeColor = dc;
            }
        }

        private TFormattingType GetFormattingType(int Idx)
        {
            if (Idx >= 1000) Idx -= 1000;
            if (Idx < 1) Idx = 1;
            return (TFormattingType)Idx - 1;
        }

        private void ReadFillRef(TObjectProperties ObjProps)
        {
            int Idx = DataStream.GetAttributeAsInt("idx", 0);
            TDrawingColor dc = GetDrawingColor();
            ObjProps.FShapeFill = GetShapeFill(Idx, dc, ObjProps.FShapeFill);
        }

        private TBlipFill ReadBlipFill()
        {
            TFillStyleList FillList = new TFillStyleList();
            string SaveDefNamespace = DataStream.DefaultNamespace;
            DataStream.DefaultNamespace = TOpenXmlManager.DrawingNamespace;
            try
            {
                ReadBlipFill(FillList);
            }
            finally
            {
                DataStream.DefaultNamespace = SaveDefNamespace;
            }
            if (FillList.Count != 1) return null;

            return FillList[0] as TBlipFill;
        }


        #endregion

        #region Group
        private void ReadGrpSp(TObjectProperties ObjProps)
        {
            ObjProps.ShapeOptions.ObjectType = TObjectType.Group;
            if (DataStream.IsSimpleTag) { DataStream.NextTag(); return; }
            string StartElement = DataStream.RecordName();
            if (!DataStream.NextTag()) return;

            while (!DataStream.AtEndElement(StartElement))
            {
                switch (DataStream.RecordName())
                {
                    case "nvGrpSpPr": ReadNvGrpSpPr(ObjProps); break;
                    case "grpSpPr": ReadGrpSpPr(ObjProps); break;
                    default:
                        TObjectProperties ChildProps = new TObjectProperties();
                        ChildProps.ShapeOptions = new TShapeProperties();
                        ChildProps.ShapeOptions.ShapeOptions = new TShapeOptionList();
                        ObjProps.FGroupedShapes.Add(ChildProps);
                        if (ReadEG_ObjectChoices(ChildProps)) break;

                        DataStream.GetXml();
                        break;
                }
            }

        }

        private void ReadGrpSpPr(TObjectProperties ObjProps)
        {
            string SaveNamespace = DataStream.DefaultNamespace;
            DataStream.DefaultNamespace = TOpenXmlManager.DrawingNamespace;
            try
            {

                if (DataStream.IsSimpleTag) { DataStream.NextTag(); return; }
                string StartElement = DataStream.RecordName();
                if (!DataStream.NextTag()) return;

                while (!DataStream.AtEndElement(StartElement))
                {
                    switch (DataStream.RecordName())
                    {
                        case "xfrm": ReadXfrm(ObjProps); break;
                        default:
                            DataStream.GetXml();
                            break;
                    }
                }
            }
            finally
            {
                DataStream.DefaultNamespace = SaveNamespace;
            }
        }

        private void ReadNvGrpSpPr(TObjectProperties ObjProps)
        {
            if (DataStream.IsSimpleTag) { DataStream.NextTag(); return; }
            string StartElement = DataStream.RecordName();
            if (!DataStream.NextTag()) return;

            while (!DataStream.AtEndElement(StartElement))
            {
                switch (DataStream.RecordName())
                {

                    case "cNvPr": ReadCNvPr(ObjProps); break;
                    case "cNvGrpSpPr": ReadCNvGrpSpPr(ObjProps); break;
                    default:
                        DataStream.GetXml();
                        break;
                }
            }
        }

        private void ReadCNvGrpSpPr(TObjectProperties ObjProps)
        {
            DataStream.GetXml();
        }

        #endregion

        #region Shape
        private void ReadSp(TObjectProperties ObjProps)
        {
            ObjProps.Macro = DataStream.GetAttribute("macro");
            ObjProps.Published = DataStream.GetAttributeAsBool("fPublished", false);
            ObjProps.FTextProperties.LockText = DataStream.GetAttributeAsBool("fLocksText", true);
            ObjProps.FLinkedFmla = DataStream.GetAttribute("textlink");

            if (DataStream.IsSimpleTag) { DataStream.NextTag(); return; }
            string StartElement = DataStream.RecordName();
            if (!DataStream.NextTag()) return;

            while (!DataStream.AtEndElement(StartElement))
            {
                switch (DataStream.RecordName())
                {
                    case "nvSpPr": ReadNvSpPr(ObjProps); break;
                    case "spPr": ReadSpPr(ObjProps); break;
                    case "style": ReadStyle(ObjProps); break;
                    case "txBody": ReadTxBody(ObjProps); break;
                    default: DataStream.GetXml(); break;
                }
            }
        }

        private void ReadNvSpPr(TObjectProperties ObjProps)
        {

            if (DataStream.IsSimpleTag) { DataStream.NextTag(); return; }
            string StartElement = DataStream.RecordName();
            if (!DataStream.NextTag()) return;

            while (!DataStream.AtEndElement(StartElement))
            {
                switch (DataStream.RecordName())
                {
                    case "cNvPr": ReadCNvPr(ObjProps); break;
                    case "cNvSpPr": ReadCNvSpPr(ObjProps); break;
                    case "nvPr": ReadNvPr(ObjProps); break;
                    default: DataStream.GetXml(); break;
                }
            }

        }

        private void ReadCNvSpPr(TObjectProperties ObjProps)
        {
            DataStream.GetXml();
        }

        private void ReadTxBody(TObjectProperties ObjProps)
        {
            string SaveNamespace = DataStream.DefaultNamespace;
            DataStream.DefaultNamespace = TOpenXmlManager.DrawingNamespace;
            try
            {
                List<TDrawingTextParagraph> paragraphs = new List<TDrawingTextParagraph>();

                if (DataStream.IsSimpleTag) { DataStream.NextTag(); return; }
                string StartElement = DataStream.RecordName();
                if (!DataStream.NextTag()) return;

                while (!DataStream.AtEndElement(StartElement))
                {
                    switch (DataStream.RecordName())
                    {
                        case "bodyPr": ReadBodyPr(ObjProps); break;
                        case "lstStyle": ReadLstStyle(ObjProps); break;
                        case "p": ReadP(paragraphs); break;

                        default: DataStream.GetXml(); break;
                    }
                }

                ObjProps.FTextExt = new TDrawingRichString(paragraphs.ToArray());
            }
            finally
            {
                DataStream.DefaultNamespace = SaveNamespace;
            }
        }

        private void ReadLstStyle(TObjectProperties ObjProps)
        {
            ObjProps.LstStyle = DataStream.GetXml();
        }

        private void ReadBodyPr(TObjectProperties ObjProps)
        {
            TDrawingCoordinate l = DataStream.GetAttributeAsDrawingCoord("lIns", new TDrawingCoordinate(91440));
            TDrawingCoordinate t = DataStream.GetAttributeAsDrawingCoord("tIns", new TDrawingCoordinate(45720));
            TDrawingCoordinate r = DataStream.GetAttributeAsDrawingCoord("rIns", new TDrawingCoordinate(91440));
            TDrawingCoordinate b = DataStream.GetAttributeAsDrawingCoord("bIns", new TDrawingCoordinate(45720));
            ObjProps.BodyPr = new TBodyPr(l,t,r,b,
                DataStream.GetXml());
        }

        private void ReadP(List<TDrawingTextParagraph> paragraphs)
        {
            if (DataStream.IsSimpleTag) { DataStream.NextTag(); AddEmptyParagraph(paragraphs); return; }
            string StartElement = DataStream.RecordName();
            if (!DataStream.NextTag()) { AddEmptyParagraph(paragraphs); return; }

            TDrawingParagraphProperties PProps = TDrawingParagraphProperties.Empty;
            TDrawingTextProperties EPProps = TDrawingTextProperties.Empty;

            List<TDrawingTextRun> Runs = new List<TDrawingTextRun>();
            while (!DataStream.AtEndElement(StartElement))
            {
                switch (DataStream.RecordName())
                {
                    case "pPr": PProps = ReadPPr(); break;
                    case "r": Runs.Add(ReadR()); break;
                    case "br": Runs.Add(ReadBr()); break;
                    case "fld": ReadFld(); break;
                    case "endParaRPr": EPProps = ReadRPr();
                        break;
                    default: DataStream.GetXml(); break;
                }
            }

            paragraphs.Add(new TDrawingTextParagraph(Runs.ToArray(), PProps, EPProps));
        }

        private static void AddEmptyParagraph(List<TDrawingTextParagraph> paragraphs)
        {
            paragraphs.Add(new TDrawingTextParagraph((string)null, TDrawingParagraphProperties.Empty, TDrawingTextProperties.Empty));
        }

        private TDrawingParagraphProperties ReadPPr()
        {
            TDrawingParagraphProperties def = TDrawingParagraphProperties.Empty;
            TDrawingParagraphProperties Result =
                new TDrawingParagraphProperties(
                    DataStream.GetAttributeAsInt("marL", def.MarL),
                    DataStream.GetAttributeAsInt("marR", def.MarR),
                    DataStream.GetAttributeAsInt("lvl", def.Lvl),
                    DataStream.GetAttributeAsDrawingCoord("indent", def.Indent),
                    GetDrawingAlign(DataStream.GetAttribute("algn"), def.Algn),
                    DataStream.GetAttributeAsDrawingCoord("defTabSz", def.DefTabSz),
                    DataStream.GetAttributeAsBool("rtl", def.Rtl),
                    DataStream.GetAttributeAsBool("eaLnBrk", def.EaLnBrk),
                    GetDrawingFontAlign(DataStream.GetAttribute("fontAlgn"), def.FontAlgn),
                    DataStream.GetAttributeAsBool("latinLnBrk", def.LatinLnBrk),
                    DataStream.GetAttributeAsBool("hangingPunct", def.HangingPunct)

                );

            DataStream.FinishTagAndIgnoreChildren();
            return Result;
        }

        private TDrawingAlignment GetDrawingAlign(string p, TDrawingAlignment def)
        {
            switch (p)
            {
                case "l": return TDrawingAlignment.Left;
                case "ctr": return TDrawingAlignment.Center;
                case "r": return TDrawingAlignment.Right;
                case "just": return TDrawingAlignment.Justified;
                case "justLow": return TDrawingAlignment.JustLow;
                case "dist": return TDrawingAlignment.Distributed;
                case "thaiDist": return TDrawingAlignment.ThaiDist;
            }
            return def;
        }

        private TDrawingFontAlign GetDrawingFontAlign(string p, TDrawingFontAlign def)
        {
            switch (p)
            {
                case "auto": return TDrawingFontAlign.Automatic;
                case "t": return TDrawingFontAlign.Top;
                case "ctr": return TDrawingFontAlign.Center;
                case "base": return TDrawingFontAlign.BaseLine;
                case "b": return TDrawingFontAlign.Bottom;
            }
            return def;
        }

        private void ReadFld()
        {
            DataStream.GetXml(); //Excel doesn't seem to allow flds.
        }

        private TDrawingTextRun ReadR()
        {
            return ReadR(null);
        }

        private TDrawingTextRun ReadR(string DefaultText)
        {
            TDrawingTextProperties Props = TDrawingTextProperties.Empty;
            string Text = DefaultText;
            if (DataStream.IsSimpleTag) { DataStream.NextTag(); return new TDrawingTextRun(Text, Props); }
            string StartElement = DataStream.RecordName();
            if (!DataStream.NextTag()) return new TDrawingTextRun(Text, Props);

            while (!DataStream.AtEndElement(StartElement))
            {
                switch (DataStream.RecordName())
                {
                    case "rPr": Props = ReadRPr(); break;
                    case "t": Text = ReadT(); break;
 
                    default: DataStream.GetXml(); break;
                }
            }

            return new TDrawingTextRun(Text, Props);
        }

        private TDrawingTextRun ReadBr()
        {
            return ReadR(((char)10).ToString());
        }

        private TDrawingTextProperties ReadRPr()
        {
            TMutableDrawingTextProperties Props = TMutableDrawingTextProperties.Empty;
            TDrawingTextAttributes defAtt = TDrawingTextAttributes.Empty;

            Props.FAttributes =
                new TDrawingTextAttributes(
                        DataStream.GetAttributeAsBool("kumimoji", defAtt.Kumimoji),
                        DataStream.GetAttribute("lang", defAtt.Lang),
                        DataStream.GetAttribute("altLang", defAtt.AltLang),
                        DataStream.GetAttributeAsInt("sz", defAtt.Size),
                        DataStream.GetAttributeAsBool("b", defAtt.Bold),
                        DataStream.GetAttributeAsBool("i", defAtt.Italic),
                        GetUnderline(DataStream.GetAttribute("u"), defAtt.Underline),
                        GetStrike(DataStream.GetAttribute("strike"), defAtt.Strike),
                        DataStream.GetAttributeAsInt("kern", defAtt.Kern),
                        GetCap(DataStream.GetAttribute("cap"), defAtt.Capitalization),
                        DataStream.GetAttributeAsDrawingCoord("spc", defAtt.Spacing),
                        DataStream.GetAttributeAsBool("normalizeH", defAtt.NormalizeH),
                        DataStream.GetAttributeAsPercent("baseline", defAtt.Baseline),
                        DataStream.GetAttributeAsBool("noProof", defAtt.NoProof),
                        DataStream.GetAttributeAsBool("dirty", defAtt.Dirty),
                        DataStream.GetAttributeAsBool("err", defAtt.Err),
                        DataStream.GetAttributeAsBool("smtClean", defAtt.SmartTagClean),
                        DataStream.GetAttributeAsInt("smtId", defAtt.SmartTagId),
                        DataStream.GetAttribute("bmk", defAtt.BookmarkLinkTarget)
                 );

            if (DataStream.IsSimpleTag) { DataStream.NextTag(); return Props.GetProps() ; }
            string StartElement = DataStream.RecordName();
            if (!DataStream.NextTag()) return Props.GetProps();

            while (!DataStream.AtEndElement(StartElement))
            {
                TFillStyle tmpFill = null;
                if (ReadEG_Fill(out tmpFill))
                {
                    Props.FFill = tmpFill;
                }
                else
                {
                    switch (DataStream.RecordName())
                    {
                        case "ln": Props.FLine = ReadLineStyle(); break;

                        case "effectLst":
                        case "effectDag": Props.FEffects = ReadEG_EffectProperties(); break;

                        case "highlight": Props.FHighlight = GetDrawingColor(); break;

                        case "uLnTx":
                        case "uLn": 
                            Props.FUnderline = ReadEG_TextUnderlineLine(Props.FUnderline); break;

                        case "uFillTx":
                        case "uFill": 
                            Props.FUnderline = ReadEG_TextUnderlineFill(Props.FUnderline); break;
                        
                        case "latin": Props.FLatin = ReadThemeTextFont(); break;
                        case "ea": Props.FEastAsian = ReadThemeTextFont(); break;
                        case "cs": Props.FComplexScript = ReadThemeTextFont(); break;
                        case "sym": Props.FSymbol = ReadThemeTextFont(); break;
                        case "hlinkClick": Props.FHyperlinkClick = ReadHlink(); break;
                        case "hlinkMouseOver": Props.FHyperlinkMouseOver = ReadHlink(); break;
                        case "rtl": Props.FRightToLeft = ReadRtl(); break;

                        default: DataStream.GetXml(); break;
                    }
                }
            }

            return Props.GetProps();
        }

        private TDrawingUnderline ReadEG_TextUnderlineLine(TDrawingUnderline OldUnderline)
        {
            string ul = DataStream.GetXml();
            if (OldUnderline != null) return new TDrawingUnderline(ul, OldUnderline.xmlFill); else return new TDrawingUnderline(ul, null);
        }

        private TDrawingUnderline ReadEG_TextUnderlineFill(TDrawingUnderline OldUnderline)
        {
            string uf = DataStream.GetXml();
            if (OldUnderline != null) return new TDrawingUnderline(OldUnderline.xmlLine, uf); else return new TDrawingUnderline(null, uf);
        }

        private bool ReadRtl()
        {
            bool Result = DataStream.GetAttributeAsBool("val", false);
            DataStream.FinishTag();
            return Result;

        }

        private TDrawingUnderlineStyle GetUnderline(string p, TDrawingUnderlineStyle def)
        {
            switch (p)
            {
                case "none": return TDrawingUnderlineStyle.None;
                case "words": return TDrawingUnderlineStyle.Words;
                case "sng": return TDrawingUnderlineStyle.Single;
                case "dbl": return TDrawingUnderlineStyle.Double;
                case "heavy": return TDrawingUnderlineStyle.Heavy;
                case "dotted": return TDrawingUnderlineStyle.Dotted;
                case "dottedHeavy": return TDrawingUnderlineStyle.DottedHeavy;
                case "dash": return TDrawingUnderlineStyle.Dash;
                case "dashHeavy": return TDrawingUnderlineStyle.DashHeavy;
                case "dashLong": return TDrawingUnderlineStyle.DashLong;
                case "dashLongHeavy": return TDrawingUnderlineStyle.DashLongHeavy;
                case "dotDash": return TDrawingUnderlineStyle.DotDash;
                case "dotDashHeavy": return TDrawingUnderlineStyle.DotDashHeavy;
                case "dotDotDash": return TDrawingUnderlineStyle.DotDotDash;
                case "dotDotDashHeavy": return TDrawingUnderlineStyle.DotDotDashHeavy;
                case "wavy": return TDrawingUnderlineStyle.Wavy;
                case "wavyHeavy": return TDrawingUnderlineStyle.WavyHeavy;
                case "wavyDbl": return TDrawingUnderlineStyle.WavyDouble;
            }
            return def;
        }

        private TDrawingTextStrike GetStrike(string p, TDrawingTextStrike def)
        {
            switch (p)
            {
                case "noStrike": return TDrawingTextStrike.None;
                case "sngStrike": return TDrawingTextStrike.Single;
                case "dblStrike": return TDrawingTextStrike.Double;
            }
            return def;
        }

        private TDrawingTextCapitalization GetCap(string p, TDrawingTextCapitalization def)
        {
            switch (p)
            {
                case "none": return TDrawingTextCapitalization.None;
                case "small": return TDrawingTextCapitalization.Small;
                case "all": return TDrawingTextCapitalization.All;
            }
            return def;
        }

        private string ReadT()
        {
            string Result = DataStream.ReadValueAsString();
            return Result;
        }


        #endregion

        #region Theme
        internal

        void ReadTheme(TThemeRecord ThemeRecord)
        {
            DataStream.SelectTheme();
            while (DataStream.NextTag())
            {
                switch (DataStream.RecordName())
                {
                    case "theme":
                        ReadActualTheme(ThemeRecord);
                        break;

                    default: //shouldn't happen
                        DataStream.GetXml();
                        break;
                }
            }
        }

        private void ReadActualTheme(TThemeRecord ThemeRecord)
        {
            TTheme Theme = ThemeRecord.Theme;
            Theme.Name = DataStream.GetAttribute("name");
            if (DataStream.IsSimpleTag) { DataStream.NextTag(); return; }
            string StartElement = DataStream.RecordName();
            if (!DataStream.NextTag()) return;
            while (!DataStream.AtEndElement(StartElement))
            {
                switch (DataStream.RecordName())
                {
                    case "themeElements": ReadThemeElements(ThemeRecord, Theme.Elements); break;
                    //case "objectDefaults": ReadObjectDefaults(Theme); break;
                    //case "extraClrSchemeLst": ReadExtraClrSchemeLst(Theme); break;
                    //case "custClrLst": ReadCustClrLst(Theme); break;

                    default:
                        ThemeRecord.AddFutureStorage(new TFutureStorageRecord(DataStream.GetXml()));
                        break;
                }
            }
        }

        private void ReadThemeElements(TThemeRecord ThemeRecord, TThemeElements Elements)
        {
            if (DataStream.IsSimpleTag) { DataStream.NextTag(); return; }
            string StartElement = DataStream.RecordName();
            if (!DataStream.NextTag()) return;

            while (!DataStream.AtEndElement(StartElement))
            {
                switch (DataStream.RecordName())
                {
                    case "clrScheme": ReadClrScheme(ThemeRecord, Elements.ColorScheme); break;
                    case "fontScheme": ReadFontScheme(ThemeRecord, Elements.FontScheme); break;
                    case "fmtScheme": ReadFmtScheme(ThemeRecord, Elements.FormatScheme); break;

                    default:
                        ThemeRecord.AddElementsFutureStorage(new TFutureStorageRecord(DataStream.GetXml()));
                        break;
                }
            }
        }

        #region Colors
        private void ReadClrScheme(TThemeRecord ThemeRecord, TThemeColorScheme ColorScheme)
        {
            ColorScheme.Name = DataStream.GetAttribute("name");
            if (DataStream.IsSimpleTag) { DataStream.NextTag(); return; }
            string StartElement = DataStream.RecordName();
            if (!DataStream.NextTag()) return;

            while (!DataStream.AtEndElement(StartElement))
            {
                switch (DataStream.RecordName())
                {
                    case "dk1": ReadThemedColor(ColorScheme, TThemeColor.Foreground1); break;
                    case "lt1": ReadThemedColor(ColorScheme, TThemeColor.Background1); break;
                    case "dk2": ReadThemedColor(ColorScheme, TThemeColor.Foreground2); break;
                    case "lt2": ReadThemedColor(ColorScheme, TThemeColor.Background2); break;
                    case "accent1": ReadThemedColor(ColorScheme, TThemeColor.Accent1); break;
                    case "accent2": ReadThemedColor(ColorScheme, TThemeColor.Accent2); break;
                    case "accent3": ReadThemedColor(ColorScheme, TThemeColor.Accent3); break;
                    case "accent4": ReadThemedColor(ColorScheme, TThemeColor.Accent4); break;
                    case "accent5": ReadThemedColor(ColorScheme, TThemeColor.Accent5); break;
                    case "accent6": ReadThemedColor(ColorScheme, TThemeColor.Accent6); break;
                    case "hlink": ReadThemedColor(ColorScheme, TThemeColor.HyperLink); break;
                    case "folHlink": ReadThemedColor(ColorScheme, TThemeColor.FollowedHyperLink); break;

                    default:
                        ThemeRecord.AddColorFutureStorage(new TFutureStorageRecord(DataStream.GetXml()));
                        break;
                }
            }
        }

        private void ReadThemedColor(TThemeColorScheme ColorScheme, TThemeColor ThemeColor)
        {
            TDrawingColor ColorDef = GetDrawingColor();
            ColorScheme[ThemeColor] = ColorDef;
        }

        private TDrawingColor GetDrawingColor()
        {
            TDrawingColor ColorDef = ColorUtil.Empty;
            if (DataStream.IsSimpleTag) { DataStream.NextTag(); return ColorDef; }
            string StartElement = DataStream.RecordName();
            if (!DataStream.NextTag()) return ColorDef;

            while (!DataStream.AtEndElement(StartElement))
            {
                switch (DataStream.RecordName())
                {
                    case "scrgbClr": ColorDef = ReadScrgbClr(); break;
                    case "srgbClr": ColorDef = ReadSrgbClr(); break;
                    case "hslClr": ColorDef = ReadHslClr(); break;
                    case "sysClr": ColorDef = ReadSysClr(); break;
                    case "schemeClr": ColorDef = ReadSchemeClr(); break;
                    case "prstClr": ColorDef = ReadPrstClr(); break;
                }

                ColorDef = TDrawingColor.AddTransform(ColorDef, ReadColorTransform()); //this will finish the tag.
            }
            return ColorDef;
        }

        private TDrawingColor ReadScrgbClr()
        {
            double cR = DataStream.GetAttributeAsPercent("r", 0);
            double cG = DataStream.GetAttributeAsPercent("g", 0);
            double cB = DataStream.GetAttributeAsPercent("b", 0);
            return TDrawingColor.FromScRgb(new TScRGBColor(cR, cG, cB));
        }

        private TDrawingColor ReadSrgbClr()
        {
            unchecked
            {
                long RGB = DataStream.GetAttributeAsHex("val", -1);
                unchecked
                {
                    if (RGB >= 0) return Color.FromArgb((int)RGB);
                }
            }

            return ColorUtil.Empty;
        }

        private TDrawingColor ReadHslClr()
        {
            double hue = DataStream.GetAttributeAsAngle("hue", 0);
            double sat = DataStream.GetAttributeAsPercent("sat", 0);
            double lum = DataStream.GetAttributeAsPercent("lum", 0);
            return TDrawingColor.FromHSL(new THSLColor(hue, sat, lum));

        }

        private TSystemColor ReadSystemColor()
        {
            #region A lot of system colors
            switch (DataStream.GetAttribute("val"))
            {
                case "scrollBar": return TSystemColor.ScrollBar;
                case "background": return TSystemColor.Background;
                case "activeCaption": return TSystemColor.ActiveCaption;
                case "inactiveCaption": return TSystemColor.InactiveCaption;
                case "menu": return TSystemColor.Menu;
                case "window": return TSystemColor.Window;
                case "windowFrame": return TSystemColor.WindowFrame;
                case "menuText": return TSystemColor.MenuText;
                case "windowText": return TSystemColor.WindowText;
                case "captionText": return TSystemColor.CaptionText;
                case "activeBorder": return TSystemColor.ActiveBorder;
                case "inactiveBorder": return TSystemColor.InactiveBorder;
                case "appWorkspace": return TSystemColor.AppWorkspace;
                case "highlight": return TSystemColor.Highlight;
                case "highlightText": return TSystemColor.HighlightText;
                case "btnFace": return TSystemColor.BtnFace;
                case "btnShadow": return TSystemColor.BtnShadow;
                case "grayText": return TSystemColor.GrayText;
                case "btnText": return TSystemColor.BtnText;
                case "inactiveCaptionText": return TSystemColor.InactiveCaptionText;
                case "btnHighlight": return TSystemColor.BtnHighlight;
                case "3dDkShadow": return TSystemColor.DkShadow3d;
                case "3dLight": return TSystemColor.Light3d;
                case "infoText": return TSystemColor.InfoText;
                case "infoBk": return TSystemColor.InfoBk;
                case "hotLight": return TSystemColor.HotLight;
                case "gradientActiveCaption": return TSystemColor.GradientActiveCaption;
                case "gradientInactiveCaption": return TSystemColor.GradientInactiveCaption;
                case "menuHighlight": return TSystemColor.MenuHighlight;
                case "menuBar": return TSystemColor.MenuBar;
            }
            #endregion

            return TSystemColor.None;
        }

        private TDrawingColor ReadSysClr()
        {
            return TDrawingColor.FromSystem(ReadSystemColor());
        }

        private TThemeColor ReadThemeColor()
        {
            switch (DataStream.GetAttribute("val"))
            {
                case "bg1": return TThemeColor.Background1;
                case "tx1": return TThemeColor.Foreground1;
                case "bg2": return TThemeColor.Background2;
                case "tx2": return TThemeColor.Foreground2;
                case "accent1": return TThemeColor.Accent1;
                case "accent2": return TThemeColor.Accent2;
                case "accent3": return TThemeColor.Accent3;
                case "accent4": return TThemeColor.Accent4;
                case "accent5": return TThemeColor.Accent5;
                case "accent6": return TThemeColor.Accent6;
                case "hlink": return TThemeColor.HyperLink;
                case "folHlink": return TThemeColor.FollowedHyperLink;
                case "phClr": return TThemeColor.None;
                case "dk1": return TThemeColor.Foreground1;
                case "lt1": return TThemeColor.Background1;
                case "dk2": return TThemeColor.Foreground2;
                case "lt2": return TThemeColor.Background2;
            }

            return TThemeColor.None;
        }

        private TDrawingColor ReadSchemeClr()
        {
            return TDrawingColor.FromTheme(ReadThemeColor());
        }

        private TPresetColor ReadPresetColor()
        {
            #region A lot of Preset colors
            switch (DataStream.GetAttribute("val"))
            {
                case "aliceBlue": return TPresetColor.AliceBlue;
                case "antiqueWhite": return TPresetColor.AntiqueWhite;
                case "aqua": return TPresetColor.Aqua;
                case "aquamarine": return TPresetColor.Aquamarine;
                case "azure": return TPresetColor.Azure;
                case "beige": return TPresetColor.Beige;
                case "bisque": return TPresetColor.Bisque;
                case "black": return TPresetColor.Black;
                case "blanchedAlmond": return TPresetColor.BlanchedAlmond;
                case "blue": return TPresetColor.Blue;
                case "blueViolet": return TPresetColor.BlueViolet;
                case "brown": return TPresetColor.Brown;
                case "burlyWood": return TPresetColor.BurlyWood;
                case "cadetBlue": return TPresetColor.CadetBlue;
                case "chartreuse": return TPresetColor.Chartreuse;
                case "chocolate": return TPresetColor.Chocolate;
                case "coral": return TPresetColor.Coral;
                case "cornflowerBlue": return TPresetColor.CornflowerBlue;
                case "cornsilk": return TPresetColor.Cornsilk;
                case "crimson": return TPresetColor.Crimson;
                case "cyan": return TPresetColor.Cyan;
                case "darkBlue": return TPresetColor.DkBlue;
                case "darkCyan": return TPresetColor.DkCyan;
                case "darkGoldenrod": return TPresetColor.DkGoldenrod;
                case "darkGray": return TPresetColor.DkGray;
                case "darkGrey": return TPresetColor.DkGray;
                case "darkGreen": return TPresetColor.DkGreen;
                case "darkKhaki": return TPresetColor.DkKhaki;
                case "darkMagenta": return TPresetColor.DkMagenta;
                case "darkOliveGreen": return TPresetColor.DkOliveGreen;
                case "darkOrange": return TPresetColor.DkOrange;
                case "darkOrchid": return TPresetColor.DkOrchid;
                case "darkRed": return TPresetColor.DkRed;
                case "darkSalmon": return TPresetColor.DkSalmon;
                case "darkSeaGreen": return TPresetColor.DkSeaGreen;
                case "darkSlateBlue": return TPresetColor.DkSlateBlue;
                case "darkSlateGray": return TPresetColor.DkSlateGray;
                case "darkSlateGrey": return TPresetColor.DkSlateGray;
                case "darkTurquoise": return TPresetColor.DkTurquoise;
                case "darkViolet": return TPresetColor.DkViolet;
                case "dkBlue": return TPresetColor.DkBlue;
                case "dkCyan": return TPresetColor.DkCyan;
                case "dkGoldenrod": return TPresetColor.DkGoldenrod;
                case "dkGray": return TPresetColor.DkGray;
                case "dkGrey": return TPresetColor.DkGray;
                case "dkGreen": return TPresetColor.DkGreen;
                case "dkKhaki": return TPresetColor.DkKhaki;
                case "dkMagenta": return TPresetColor.DkMagenta;
                case "dkOliveGreen": return TPresetColor.DkOliveGreen;
                case "dkOrange": return TPresetColor.DkOrange;
                case "dkOrchid": return TPresetColor.DkOrchid;
                case "dkRed": return TPresetColor.DkRed;
                case "dkSalmon": return TPresetColor.DkSalmon;
                case "dkSeaGreen": return TPresetColor.DkSeaGreen;
                case "dkSlateBlue": return TPresetColor.DkSlateBlue;
                case "dkSlateGray": return TPresetColor.DkSlateGray;
                case "dkSlateGrey": return TPresetColor.DkSlateGray;
                case "dkTurquoise": return TPresetColor.DkTurquoise;
                case "dkViolet": return TPresetColor.DkViolet;
                case "deepPink": return TPresetColor.DeepPink;
                case "deepSkyBlue": return TPresetColor.DeepSkyBlue;
                case "dimGray": return TPresetColor.DimGray;
                case "dimGrey": return TPresetColor.DimGray;
                case "dodgerBlue": return TPresetColor.DodgerBlue;
                case "firebrick": return TPresetColor.Firebrick;
                case "floralWhite": return TPresetColor.FloralWhite;
                case "forestGreen": return TPresetColor.ForestGreen;
                case "fuchsia": return TPresetColor.Fuchsia;
                case "gainsboro": return TPresetColor.Gainsboro;
                case "ghostWhite": return TPresetColor.GhostWhite;
                case "gold": return TPresetColor.Gold;
                case "goldenrod": return TPresetColor.Goldenrod;
                case "gray": return TPresetColor.Gray;
                case "grey": return TPresetColor.Gray;
                case "green": return TPresetColor.Green;
                case "greenYellow": return TPresetColor.GreenYellow;
                case "honeydew": return TPresetColor.Honeydew;
                case "hotPink": return TPresetColor.HotPink;
                case "indianRed": return TPresetColor.IndianRed;
                case "indigo": return TPresetColor.Indigo;
                case "ivory": return TPresetColor.Ivory;
                case "khaki": return TPresetColor.Khaki;
                case "lavender": return TPresetColor.Lavender;
                case "lavenderBlush": return TPresetColor.LavenderBlush;
                case "lawnGreen": return TPresetColor.LawnGreen;
                case "lemonChiffon": return TPresetColor.LemonChiffon;
                case "lightBlue": return TPresetColor.LtBlue;
                case "lightCoral": return TPresetColor.LtCoral;
                case "lightCyan": return TPresetColor.LtCyan;
                case "lightGoldenrodYellow": return TPresetColor.LtGoldenrodYellow;
                case "lightGray": return TPresetColor.LtGray;
                case "lightGrey": return TPresetColor.LtGray;
                case "lightGreen": return TPresetColor.LtGreen;
                case "lightPink": return TPresetColor.LtPink;
                case "lightSalmon": return TPresetColor.LtSalmon;
                case "lightSeaGreen": return TPresetColor.LtSeaGreen;
                case "lightSkyBlue": return TPresetColor.LtSkyBlue;
                case "lightSlateGray": return TPresetColor.LtSlateGray;
                case "lightSlateGrey": return TPresetColor.LtSlateGray;
                case "lightSteelBlue": return TPresetColor.LtSteelBlue;
                case "lightYellow": return TPresetColor.LtYellow;
                case "ltBlue": return TPresetColor.LtBlue;
                case "ltCoral": return TPresetColor.LtCoral;
                case "ltCyan": return TPresetColor.LtCyan;
                case "ltGoldenrodYellow": return TPresetColor.LtGoldenrodYellow;
                case "ltGray": return TPresetColor.LtGray;
                case "ltGrey": return TPresetColor.LtGray;
                case "ltGreen": return TPresetColor.LtGreen;
                case "ltPink": return TPresetColor.LtPink;
                case "ltSalmon": return TPresetColor.LtSalmon;
                case "ltSeaGreen": return TPresetColor.LtSeaGreen;
                case "ltSkyBlue": return TPresetColor.LtSkyBlue;
                case "ltSlateGray": return TPresetColor.LtSlateGray;
                case "ltSlateGrey": return TPresetColor.LtSlateGray;
                case "ltSteelBlue": return TPresetColor.LtSteelBlue;
                case "ltYellow": return TPresetColor.LtYellow;
                case "lime": return TPresetColor.Lime;
                case "limeGreen": return TPresetColor.LimeGreen;
                case "linen": return TPresetColor.Linen;
                case "magenta": return TPresetColor.Magenta;
                case "maroon": return TPresetColor.Maroon;
                case "medAquamarine": return TPresetColor.MedAquamarine;
                case "medBlue": return TPresetColor.MedBlue;
                case "medOrchid": return TPresetColor.MedOrchid;
                case "medPurple": return TPresetColor.MedPurple;
                case "medSeaGreen": return TPresetColor.MedSeaGreen;
                case "medSlateBlue": return TPresetColor.MedSlateBlue;
                case "medSpringGreen": return TPresetColor.MedSpringGreen;
                case "medTurquoise": return TPresetColor.MedTurquoise;
                case "medVioletRed": return TPresetColor.MedVioletRed;
                case "mediumAquamarine": return TPresetColor.MedAquamarine;
                case "mediumBlue": return TPresetColor.MedBlue;
                case "mediumOrchid": return TPresetColor.MedOrchid;
                case "mediumPurple": return TPresetColor.MedPurple;
                case "mediumSeaGreen": return TPresetColor.MedSeaGreen;
                case "mediumSlateBlue": return TPresetColor.MedSlateBlue;
                case "mediumSpringGreen": return TPresetColor.MedSpringGreen;
                case "mediumTurquoise": return TPresetColor.MedTurquoise;
                case "mediumVioletRed": return TPresetColor.MedVioletRed;
                case "midnightBlue": return TPresetColor.MidnightBlue;
                case "mintCream": return TPresetColor.MintCream;
                case "mistyRose": return TPresetColor.MistyRose;
                case "moccasin": return TPresetColor.Moccasin;
                case "navajoWhite": return TPresetColor.NavajoWhite;
                case "navy": return TPresetColor.Navy;
                case "oldLace": return TPresetColor.OldLace;
                case "olive": return TPresetColor.Olive;
                case "oliveDrab": return TPresetColor.OliveDrab;
                case "orange": return TPresetColor.Orange;
                case "orangeRed": return TPresetColor.OrangeRed;
                case "orchid": return TPresetColor.Orchid;
                case "paleGoldenrod": return TPresetColor.PaleGoldenrod;
                case "paleGreen": return TPresetColor.PaleGreen;
                case "paleTurquoise": return TPresetColor.PaleTurquoise;
                case "paleVioletRed": return TPresetColor.PaleVioletRed;
                case "papayaWhip": return TPresetColor.PapayaWhip;
                case "peachPuff": return TPresetColor.PeachPuff;
                case "peru": return TPresetColor.Peru;
                case "pink": return TPresetColor.Pink;
                case "plum": return TPresetColor.Plum;
                case "powderBlue": return TPresetColor.PowderBlue;
                case "purple": return TPresetColor.Purple;
                case "red": return TPresetColor.Red;
                case "rosyBrown": return TPresetColor.RosyBrown;
                case "royalBlue": return TPresetColor.RoyalBlue;
                case "saddleBrown": return TPresetColor.SaddleBrown;
                case "salmon": return TPresetColor.Salmon;
                case "sandyBrown": return TPresetColor.SandyBrown;
                case "seaGreen": return TPresetColor.SeaGreen;
                case "seaShell": return TPresetColor.SeaShell;
                case "sienna": return TPresetColor.Sienna;
                case "silver": return TPresetColor.Silver;
                case "skyBlue": return TPresetColor.SkyBlue;
                case "slateBlue": return TPresetColor.SlateBlue;
                case "slateGray": return TPresetColor.SlateGray;
                case "slateGrey": return TPresetColor.SlateGray;
                case "snow": return TPresetColor.Snow;
                case "springGreen": return TPresetColor.SpringGreen;
                case "steelBlue": return TPresetColor.SteelBlue;
                case "tan": return TPresetColor.Tan;
                case "teal": return TPresetColor.Teal;
                case "thistle": return TPresetColor.Thistle;
                case "tomato": return TPresetColor.Tomato;
                case "turquoise": return TPresetColor.Turquoise;
                case "violet": return TPresetColor.Violet;
                case "wheat": return TPresetColor.Wheat;
                case "white": return TPresetColor.White;
                case "whiteSmoke": return TPresetColor.WhiteSmoke;
                case "yellow": return TPresetColor.Yellow;
                case "yellowGreen": return TPresetColor.YellowGreen;
            }
            #endregion

            return TPresetColor.None;
        }

        private TDrawingColor ReadPrstClr()
        {
            return TDrawingColor.FromPreset(ReadPresetColor());
        }

        private TColorTransform[] ReadColorTransform()
        {
            TColorTransform[] Result = new TColorTransform[0];
            if (DataStream.IsSimpleTag) { DataStream.NextTag(); return Result; }
            string StartElement = DataStream.RecordName();
            if (!DataStream.NextTag()) return Result;

            List<TColorTransform> TransformList = new List<TColorTransform>();
            while (!DataStream.AtEndElement(StartElement))
            {
                bool TagFinished = false;
                switch (DataStream.RecordName())
                {
                    case "tint": TransformList.Add(new TColorTransform(TColorTransformType.Tint, DataStream.GetAttributeAsPercent("val", 0))); break;
                    case "shade": TransformList.Add(new TColorTransform(TColorTransformType.Shade, DataStream.GetAttributeAsPercent("val", 0))); break;
                    case "comp": TransformList.Add(new TColorTransform(TColorTransformType.Complement, 0)); break;
                    case "inv": TransformList.Add(new TColorTransform(TColorTransformType.Inverse, 0)); break;
                    case "gray": TransformList.Add(new TColorTransform(TColorTransformType.Gray, 0)); break;
                    case "alpha": TransformList.Add(new TColorTransform(TColorTransformType.Alpha, DataStream.GetAttributeAsPercent("val", 0))); break;
                    case "alphaOff": TransformList.Add(new TColorTransform(TColorTransformType.AlphaOff, DataStream.GetAttributeAsPercent("val", 0))); break;
                    case "alphaMod": TransformList.Add(new TColorTransform(TColorTransformType.AlphaMod, DataStream.GetAttributeAsPercent("val", 0))); break;
                    case "hue": TransformList.Add(new TColorTransform(TColorTransformType.Hue, DataStream.GetAttributeAsAngle("val", 0))); break;
                    case "hueOff": TransformList.Add(new TColorTransform(TColorTransformType.HueOff, DataStream.GetAttributeAsAngle("val", 0))); break;
                    case "hueMod": TransformList.Add(new TColorTransform(TColorTransformType.HueMod, DataStream.GetAttributeAsPercent("val", 0))); break;
                    case "sat": TransformList.Add(new TColorTransform(TColorTransformType.Sat, DataStream.GetAttributeAsPercent("val", 0))); break;
                    case "satOff": TransformList.Add(new TColorTransform(TColorTransformType.SatOff, DataStream.GetAttributeAsPercent("val", 0))); break;
                    case "satMod": TransformList.Add(new TColorTransform(TColorTransformType.SatMod, DataStream.GetAttributeAsPercent("val", 0))); break;
                    case "lum": TransformList.Add(new TColorTransform(TColorTransformType.Lum, DataStream.GetAttributeAsPercent("val", 0))); break;
                    case "lumOff": TransformList.Add(new TColorTransform(TColorTransformType.LumOff, DataStream.GetAttributeAsPercent("val", 0))); break;
                    case "lumMod": TransformList.Add(new TColorTransform(TColorTransformType.LumMod, DataStream.GetAttributeAsPercent("val", 0))); break;
                    case "red": TransformList.Add(new TColorTransform(TColorTransformType.Red, DataStream.GetAttributeAsPercent("val", 0))); break;
                    case "redOff": TransformList.Add(new TColorTransform(TColorTransformType.RedOff, DataStream.GetAttributeAsPercent("val", 0))); break;
                    case "redMod": TransformList.Add(new TColorTransform(TColorTransformType.RedMod, DataStream.GetAttributeAsPercent("val", 0))); break;
                    case "green": TransformList.Add(new TColorTransform(TColorTransformType.Green, DataStream.GetAttributeAsPercent("val", 0))); break;
                    case "greenOff": TransformList.Add(new TColorTransform(TColorTransformType.GreenOff, DataStream.GetAttributeAsPercent("val", 0))); break;
                    case "greenMod": TransformList.Add(new TColorTransform(TColorTransformType.GreenMod, DataStream.GetAttributeAsPercent("val", 0))); break;
                    case "blue": TransformList.Add(new TColorTransform(TColorTransformType.Blue, DataStream.GetAttributeAsPercent("val", 0))); break;
                    case "blueOff": TransformList.Add(new TColorTransform(TColorTransformType.BlueOff, DataStream.GetAttributeAsPercent("val", 0))); break;
                    case "blueMod": TransformList.Add(new TColorTransform(TColorTransformType.BlueMod, DataStream.GetAttributeAsPercent("val", 0))); break;
                    case "gamma": TransformList.Add(new TColorTransform(TColorTransformType.Gamma, 0)); break;
                    case "invGamma": TransformList.Add(new TColorTransform(TColorTransformType.InvGamma, 0)); break;
                    default:
                        DataStream.GetXml();
                        TagFinished = true;
                        break;
                }
                if (!TagFinished) DataStream.FinishTag();
            }

            return TransformList.ToArray();
        }
        #endregion

        #region Fonts
        private void ReadFontScheme(TThemeRecord ThemeRecord, TThemeFontScheme FontScheme)
        {
            FontScheme.Name = DataStream.GetAttribute("name");
            if (DataStream.IsSimpleTag) { DataStream.NextTag(); return; }
            string StartElement = DataStream.RecordName();
            if (!DataStream.NextTag()) return;

            while (!DataStream.AtEndElement(StartElement))
            {
                switch (DataStream.RecordName())
                {
                    case "majorFont": FontScheme.MajorFont = ReadThemedFont(ThemeRecord); break;
                    case "minorFont": FontScheme.MinorFont = ReadThemedFont(ThemeRecord); break;
                    default:
                        ThemeRecord.AddFontFutureStorage(new TFutureStorageRecord(DataStream.GetXml()));
                        break;
                }
            }
        }

        private TThemeFont ReadThemedFont(TThemeRecord ThemeRecord)
        {
            if (DataStream.IsSimpleTag) { DataStream.NextTag(); return null; }
            string StartElement = DataStream.RecordName();
            if (!DataStream.NextTag()) return null;

            TThemeFont Result = new TThemeFont();
            while (!DataStream.AtEndElement(StartElement))
            {
                switch (DataStream.RecordName())
                {
                    case "latin": Result.Latin = ReadThemeTextFont(); break;
                    case "ea": Result.EastAsian = ReadThemeTextFont(); break;
                    case "cs": Result.ComplexScript = ReadThemeTextFont(); break;
                    case "font": ReadFont(Result); break;

                    default:
                        ThemeRecord.AddMajorMinorFontFutureStorage(Result, new TFutureStorageRecord(DataStream.GetXml()));
                        break;
                }
            }

            return Result;
        }

        private TThemeTextFont ReadThemeTextFont()
        {
            TThemeTextFont Result = new TThemeTextFont(
                 DataStream.GetAttribute("typeface"),
                 DataStream.GetAttribute("panose"),
                 (TPitchFamily)DataStream.GetAttributeAsInt("pitchFamily", 0),
                 (TFontCharSet)DataStream.GetAttributeAsInt("charset", 1));
            DataStream.FinishTag();

            return Result;
        }

        private void ReadFont(TThemeFont ThemeFont)
        {
            ThemeFont.AddFont(DataStream.GetAttribute("script"), DataStream.GetAttribute("typeface"));
            DataStream.FinishTag();
        }
        #endregion

        #region Format
        private void ReadFmtScheme(TThemeRecord ThemeRecord, TThemeFormatScheme FormatScheme)
        {
            FormatScheme.Name = DataStream.GetAttribute("name");
            if (DataStream.IsSimpleTag) { DataStream.NextTag(); return; }
            string StartElement = DataStream.RecordName();
            if (!DataStream.NextTag()) return;

            while (!DataStream.AtEndElement(StartElement))
            {
                switch (DataStream.RecordName())
                {
                    case "fillStyleLst": ReadFillStyleLst(FormatScheme.FillStyleList); break;
                    case "lnStyleLst": ReadLnStyleLst(FormatScheme.LineStyleList); break;
                    case "effectStyleLst": ReadEffectStyleLst(FormatScheme.EffectStyleList); break;
                    case "bgFillStyleLst": ReadFillStyleLst(FormatScheme.BkFillStyleList); break;

                    default:
                        DataStream.GetXml();
                        break;
                }
            }
        }

        private void ReadFillStyleLst(TFillStyleList FillStyleList)
        {
            if (DataStream.IsSimpleTag) { DataStream.NextTag(); return; }
            string StartElement = DataStream.RecordName();
            if (!DataStream.NextTag()) return;

            while (!DataStream.AtEndElement(StartElement))
            {
                TFillStyle Fill;
                if (ReadEG_Fill(out Fill))
                {
                    if (Fill != null) FillStyleList.Add(Fill);
                }
                else
                {
                    DataStream.GetXml();
                }
            }
        }

        private bool ReadEG_Fill(out TFillStyle Fill)
        {
            Fill = null;
            switch (DataStream.RecordName())
            {
                case "noFill": Fill = new TNoFill(); DataStream.FinishTag(); break;
                case "solidFill": Fill = ReadSolidFill(); break;
                case "gradFill": Fill = ReadGradFill(); break;
                case "blipFill": Fill = ReadBlipFill(); break;
                case "pattFill": Fill = ReadPattFill(); break;
                case "grpFill": Fill = new TGroupFill(); DataStream.FinishTag(); break;
                default:
                    return false;
            }
            return true; 
        }

        private TFillStyle ReadGradFill()
        {
            TFlipMode FlipMode = GetFlip(DataStream.GetAttribute("flip"));
            bool Rotate = DataStream.GetAttributeAsBool("rotWithShape", true);
            TDrawingGradientStop[] GradientStops = null;
            TDrawingGradientDef GradientDef = null;

            TDrawingRelativeRect? TileRect = null;
            if (DataStream.IsSimpleTag)
            {
                DataStream.NextTag();
                return new TGradientFill(TileRect, Rotate, FlipMode, GradientStops, GradientDef);
            }
            string StartElement = DataStream.RecordName();
            if (!DataStream.NextTag()) return null;

            while (!DataStream.AtEndElement(StartElement))
            {
                switch (DataStream.RecordName())
                {
                    case "gsLst": GradientStops = ReadGradientStops(); break;
                    case "lin": GradientDef = ReadLinearGrad(); break;
                    case "path": GradientDef = ReadPathGrad(); break;
                    case "tileRect": TileRect = ReadDrawingRelativeRect(); break;
                    default:
                        DataStream.GetXml();
                        break;
                }
            }

            return new TGradientFill(TileRect, Rotate, FlipMode, GradientStops, GradientDef);
        }

        private TDrawingGradientStop[] ReadGradientStops()
        {
            List<TDrawingGradientStop> GradientStops = new List<TDrawingGradientStop>();

            if (DataStream.IsSimpleTag)
            {
                XlsMessages.ThrowException(XlsErr.ErrExcelInvalid);
            }
            string StartElement = DataStream.RecordName();
            if (!DataStream.NextTag()) XlsMessages.ThrowException(XlsErr.ErrExcelInvalid);

            while (!DataStream.AtEndElement(StartElement))
            {
                switch (DataStream.RecordName())
                {
                    case "gs": GradientStops.Add(ReadOneGradientStop()); break;
                    default:
                        DataStream.GetXml();
                        break;
                }
            }

            return GradientStops.ToArray();

        }

        private TDrawingGradientStop ReadOneGradientStop()
        {
            double Pos = DataStream.GetAttributeAsPercent("pos", 0);
            return new TDrawingGradientStop(Pos, GetDrawingColor());
        }

        private TDrawingLinearGradient ReadLinearGrad()
        {
            TDrawingLinearGradient Result = new TDrawingLinearGradient(DataStream.GetAttributeAsAngle("ang", 0), DataStream.GetAttributeAsBool("scaled", false));
            DataStream.FinishTag();
            return Result;
        }

        private TDrawingPathGradient ReadPathGrad()
        {
            TPathShadeType Path = ReadPathShadeType(DataStream.GetAttribute("path"));
            TDrawingRelativeRect? FillToRect = null;

            if (DataStream.IsSimpleTag)
            {
                DataStream.NextTag();
                return new TDrawingPathGradient(FillToRect, Path);
            }
            string StartElement = DataStream.RecordName();
            if (!DataStream.NextTag()) return new TDrawingPathGradient(FillToRect, Path);

            while (!DataStream.AtEndElement(StartElement))
            {
                switch (DataStream.RecordName())
                {
                    case "fillToRect": FillToRect = ReadDrawingRelativeRect(); break;
                    default:
                        DataStream.GetXml();
                        break;
                }
            }

            return new TDrawingPathGradient(FillToRect, Path);
        }

        private TPathShadeType ReadPathShadeType(string p)
        {
            switch (p)
            {
                case "shape": return TPathShadeType.Shape;
                case "circle": return TPathShadeType.Circle;
                case "rect": return TPathShadeType.Rect;

                default: return TPathShadeType.Shape;
            }
        }

        private void ReadBlipFill(TFillStyleList FillStyleList)
        {
            int Dpi = DataStream.GetAttributeAsInt("dpi", 0);
            bool Rotate = DataStream.GetAttributeAsBool("rotWithShape", true);

            TBlip Blip = null;
            TDrawingRelativeRect? SrcRect = null;
            TBlipFillMode FillMode = null;

            if (DataStream.IsSimpleTag)
            {
                FillStyleList.Add(new TBlipFill(Dpi, Rotate, Blip, SrcRect, FillMode));
                DataStream.NextTag(); return;
            }
            string StartElement = DataStream.RecordName();
            if (!DataStream.NextTag()) return;

            while (!DataStream.AtEndElement(StartElement))
            {
                switch (DataStream.RecordName())
                {
                    case "blip": Blip = ReadBlip(); break;
                    case "srcRect": SrcRect = ReadDrawingRelativeRect(); break;
                    case "tile": FillMode = ReadTile(); break;
                    case "stretch": FillMode = ReadStretch(); break;
                    default:
                        DataStream.GetXml();
                        break;
                }
            }

            FillStyleList.Add(new TBlipFill(Dpi, Rotate, Blip, SrcRect, FillMode));
        }

        private TBlip ReadBlip()
        {
            TBlipCompression Compression = GetBlipCompression(DataStream.GetAttribute("cstate"));
            string ContentType;
            string ImageFileName;
            byte[] PictureData = DataStream.GetRelationshipData("embed", 0, 0, out ContentType, out ImageFileName); 
            ImageFileName = null; //the filename we want is not stored in xlsx drawings.

            if (DataStream.IsSimpleTag)
            {
                DataStream.NextTag();
                return new TBlip(Compression, PictureData, ImageFileName, ContentType, null);
            }

            List<string> Transforms = null;
            string StartElement = DataStream.RecordName();
            if (!DataStream.NextTag()) return new TBlip(Compression, PictureData, ImageFileName, ContentType, null);

            while (!DataStream.AtEndElement(StartElement))
            {
                string t = DataStream.GetXml();
                if (t != null)
                {
                    if (Transforms == null) Transforms = new List<string>();
                    Transforms.Add(t);
                }
            }

            return new TBlip(Compression, PictureData, ImageFileName, ContentType, Transforms.ToArray());
        }

        private TBlipCompression GetBlipCompression(string bc)
        {
            switch (bc)
            {
                case "email": return TBlipCompression.Email;
                case "screen": return TBlipCompression.Screen;
                case "print": return TBlipCompression.Print;
                case "hqprint": return TBlipCompression.HQPrint;

                default: return TBlipCompression.None;
            }
        }

        private TDrawingRelativeRect ReadDrawingRelativeRect()
        {
            TDrawingRelativeRect Result = new TDrawingRelativeRect(
                DataStream.GetAttributeAsPercent("l", 0),
                DataStream.GetAttributeAsPercent("t", 0),
                DataStream.GetAttributeAsPercent("r", 0),
                DataStream.GetAttributeAsPercent("b", 0)
                );
            DataStream.FinishTag();
            return Result;
        }

        private TBlipFillMode ReadTile()
        {
            TBlipFillTile Result = new TBlipFillTile(
                GetDrawingRectAlign(DataStream.GetAttribute("algn")),
                GetFlip(DataStream.GetAttribute("flip")),
                DataStream.GetAttributeAsDrawingCoord("tx", new TDrawingCoordinate(0)),
                DataStream.GetAttributeAsDrawingCoord("ty", new TDrawingCoordinate(0)),
                DataStream.GetAttributeAsPercent("sx", 1),
                DataStream.GetAttributeAsPercent("sy", 1));

            DataStream.FinishTag();
            return Result;
        }

        private TDrawingRectAlign GetDrawingRectAlign(string p)
        {
            switch (p)
            {
                case "tl": return TDrawingRectAlign.TopLeft;
                case "t": return TDrawingRectAlign.Top;
                case "tr": return TDrawingRectAlign.TopRight;
                case "l": return TDrawingRectAlign.Left;
                case "ctr": return TDrawingRectAlign.Center;
                case "r": return TDrawingRectAlign.Right;
                case "bl": return TDrawingRectAlign.BottomLeft;
                case "b": return TDrawingRectAlign.Bottom;
                case "br": return TDrawingRectAlign.BottomRight;

                default:
                    return TDrawingRectAlign.TopLeft;
            }
        }

        private TFlipMode GetFlip(string p)
        {
            switch (p)
            {
                case "none": return TFlipMode.None;
                case "x": return TFlipMode.X;
                case "y": return TFlipMode.Y;
                case "xy": return TFlipMode.XY;
            }

            return TFlipMode.None;
        }

        private TBlipFillMode ReadStretch()
        {
            TBlipFillStretch Result = null;
            if (DataStream.IsSimpleTag)
            {
                DataStream.NextTag(); return null;
            }
            string StartElement = DataStream.RecordName();
            if (!DataStream.NextTag()) return null;

            while (!DataStream.AtEndElement(StartElement))
            {
                switch (DataStream.RecordName())
                {
                    case "fillRect": Result = new TBlipFillStretch(ReadDrawingRelativeRect()); break;
                    default:
                        DataStream.GetXml();
                        break;
                }
            }

            return Result;
        }


        private TFillStyle ReadPattFill()
        {
            TDrawingPattern pat = GetDrawingPattern(DataStream.GetAttribute("prst"));
            TDrawingColor FgColor = Color.Black;
            TDrawingColor BgColor = Color.White;

            if (DataStream.IsSimpleTag)
            {
                DataStream.NextTag();
                return new TPatternFill(FgColor, BgColor, pat);
            }
            string StartElement = DataStream.RecordName();
            if (!DataStream.NextTag())
            {
                return new TPatternFill(FgColor, BgColor, pat);
            }

            while (!DataStream.AtEndElement(StartElement))
            {
                switch (DataStream.RecordName())
                {
                    case "bgClr": BgColor = GetDrawingColor(); break;
                    case "fgClr": FgColor = GetDrawingColor(); break;
                    default:
                        DataStream.GetXml();
                        break;
                }
            }
            return new TPatternFill(FgColor, BgColor, pat);
        }

        private TDrawingPattern GetDrawingPattern(string p)
        {
            switch (p)
            {
                case "pct5": return TDrawingPattern.pct5;
                case "pct10": return TDrawingPattern.pct10;
                case "pct20": return TDrawingPattern.pct20;
                case "pct25": return TDrawingPattern.pct25;
                case "pct30": return TDrawingPattern.pct30;
                case "pct40": return TDrawingPattern.pct40;
                case "pct50": return TDrawingPattern.pct50;
                case "pct60": return TDrawingPattern.pct60;
                case "pct70": return TDrawingPattern.pct70;
                case "pct75": return TDrawingPattern.pct75;
                case "pct80": return TDrawingPattern.pct80;
                case "pct90": return TDrawingPattern.pct90;
                case "horz": return TDrawingPattern.horz;
                case "vert": return TDrawingPattern.vert;
                case "ltHorz": return TDrawingPattern.ltHorz;
                case "ltVert": return TDrawingPattern.ltVert;
                case "dkHorz": return TDrawingPattern.dkHorz;
                case "dkVert": return TDrawingPattern.dkVert;
                case "narHorz": return TDrawingPattern.narHorz;
                case "narVert": return TDrawingPattern.narVert;
                case "dashHorz": return TDrawingPattern.dashHorz;
                case "dashVert": return TDrawingPattern.dashVert;
                case "cross": return TDrawingPattern.cross;
                case "dnDiag": return TDrawingPattern.dnDiag;
                case "upDiag": return TDrawingPattern.upDiag;
                case "ltDnDiag": return TDrawingPattern.ltDnDiag;
                case "ltUpDiag": return TDrawingPattern.ltUpDiag;
                case "dkDnDiag": return TDrawingPattern.dkDnDiag;
                case "dkUpDiag": return TDrawingPattern.dkUpDiag;
                case "wdDnDiag": return TDrawingPattern.wdDnDiag;
                case "wdUpDiag": return TDrawingPattern.wdUpDiag;
                case "dashDnDiag": return TDrawingPattern.dashDnDiag;
                case "dashUpDiag": return TDrawingPattern.dashUpDiag;
                case "diagCross": return TDrawingPattern.diagCross;
                case "smCheck": return TDrawingPattern.smCheck;
                case "lgCheck": return TDrawingPattern.lgCheck;
                case "smGrid": return TDrawingPattern.smGrid;
                case "lgGrid": return TDrawingPattern.lgGrid;
                case "dotGrid": return TDrawingPattern.dotGrid;
                case "smConfetti": return TDrawingPattern.smConfetti;
                case "lgConfetti": return TDrawingPattern.lgConfetti;
                case "horzBrick": return TDrawingPattern.horzBrick;
                case "diagBrick": return TDrawingPattern.diagBrick;
                case "solidDmnd": return TDrawingPattern.solidDmnd;
                case "openDmnd": return TDrawingPattern.openDmnd;
                case "dotDmnd": return TDrawingPattern.dotDmnd;
                case "plaid": return TDrawingPattern.plaid;
                case "sphere": return TDrawingPattern.sphere;
                case "weave": return TDrawingPattern.weave;
                case "divot": return TDrawingPattern.divot;
                case "shingle": return TDrawingPattern.shingle;
                case "wave": return TDrawingPattern.wave;
                case "trellis": return TDrawingPattern.trellis;
                case "zigZag": return TDrawingPattern.zigZag;
                default:
                    return TDrawingPattern.cross;
            }
        }

        private TFillStyle ReadSolidFill()
        {
            return new TSolidFill(GetDrawingColor());
        }

        private void ReadLnStyleLst(TLineStyleList LineStyleList)
        {
            if (DataStream.IsSimpleTag) { DataStream.NextTag(); return; }
            string StartElement = DataStream.RecordName();
            if (!DataStream.NextTag()) return;

            while (!DataStream.AtEndElement(StartElement))
            {
                switch (DataStream.RecordName())
                {
                    case "ln":
                        TLineStyle LineStyle = ReadLineStyle();
                        if (LineStyle != null) LineStyleList.Add(LineStyle);
                        break;

                    default:
                        DataStream.GetXml();
                        break;
                }
            }
        }

        private TLineStyle ReadLineStyle()
        {
            int w = DataStream.GetAttributeAsInt("w", TLineStyle.DefaultWidth);
            TPenAlignment? pa = GetPenAlignment(DataStream.GetAttribute("algn"));
            TLineCap? lc = GetLineCap(DataStream.GetAttribute("cap"));
            TCompoundLineType? clt = GetCompoundLineType(DataStream.GetAttribute("cmpd"));
            TFillStyle Fill = null;
            List<string> extra = new List<string>();
            TLineDashing? ds = null;
            TLineJoin? lj = null;
            TLineArrow? HeadArrow = null;
            TLineArrow? TailArrow = null;

            if (DataStream.IsSimpleTag) { DataStream.NextTag(); return new TLineStyle(Fill, w, pa, lc, clt, ds, lj, HeadArrow, TailArrow, extra.ToArray()); }
            string StartElement = DataStream.RecordName();
            if (!DataStream.NextTag()) return new TLineStyle(Fill, w, pa, lc, clt, ds, lj, HeadArrow, TailArrow, extra.ToArray());

            while (!DataStream.AtEndElement(StartElement))
            {
                TFillStyle tmpFill = null;
                if (ReadEG_Fill(out tmpFill))
                {
                    Fill = tmpFill;
                }
                else
                {
                    switch (DataStream.RecordName())
                    {
                        case "prstDash": ds = GetLineDashing(DataStream.GetAttribute("val")); DataStream.FinishTag(); break;
                        case "bevel": lj = TLineJoin.Bevel; DataStream.FinishTag(); break;
                        case "miter": lj = TLineJoin.Miter; DataStream.FinishTag(); break;
                        case "round": lj = TLineJoin.Round; DataStream.FinishTag(); break;

                        case "tailEnd": TailArrow = GetLineEnd(); DataStream.FinishTag(); break;
                        case "headEnd": HeadArrow = GetLineEnd(); DataStream.FinishTag(); break;
                        default:
                            extra.Add(DataStream.GetXml());
                            break;
                    }
                }
            }
            return new TLineStyle(Fill, w, pa, lc, clt, ds, lj, HeadArrow, TailArrow, extra.ToArray());
        }

        private TLineArrow GetLineEnd()
        {
            return new TLineArrow(GetLineEndType(DataStream.GetAttribute("type")),
                GetLineArrowLen(DataStream.GetAttribute("len")), GetLineArrowWidth(DataStream.GetAttribute("w"))
                );
        }

        private TArrowWidth GetLineArrowWidth(string p)
        {
            switch (p)
            {
                case "lg": return TArrowWidth.Large;
                case "sm": return TArrowWidth.Small;
                case "med":
                default:
                    return TArrowWidth.Medium;
            }
        }
        private TArrowLen GetLineArrowLen(string p)
        {
            switch (p)
            {
                case "lg": return TArrowLen.Large;
                case "sm": return TArrowLen.Small;
                case "med":
                default:
                    return TArrowLen.Medium;
            }
        }

        private TArrowStyle GetLineEndType(string p)
        {
            switch (p)
            {
                case "none": return TArrowStyle.None;
                case "triangle": return TArrowStyle.Normal;
                case "stealth": return TArrowStyle.Stealth;
                case "diamond": return TArrowStyle.Diamond;
                case "oval": return TArrowStyle.Oval;
                case "arrow": return TArrowStyle.Open;
            }

            return TArrowStyle.None;
        }

        private TLineDashing GetLineDashing(string p)
        {
            switch (p)
            {
                case "solid": return TLineDashing.Solid;
                case "dot": return TLineDashing.DotGEL;
                case "dash": return TLineDashing.DashGEL;
                case "lgDash": return TLineDashing.LongDashGEL;
                case "dashDot": return TLineDashing.DashDotGEL;
                case "lgDashDot": return TLineDashing.LongDashDotGEL;
                case "lgDashDotDot": return TLineDashing.LongDashDotDotGEL;
                case "sysDash": return TLineDashing.DashSys;
                case "sysDot": return TLineDashing.DotSys;
                case "sysDashDot": return TLineDashing.DashDotSys;
                case "sysDashDotDot": return TLineDashing.DashDotDotSys;
            }
            return TLineDashing.Solid;
        }

        private TPenAlignment? GetPenAlignment(string p)
        {
            switch (p)
            {
                case "ctr": return TPenAlignment.Center;
                case "in": return TPenAlignment.Inset;
                default: return null;
            }
        }

        private TLineCap? GetLineCap(string p)
        {
            switch (p)
            {
                case "rnd": return TLineCap.Round;
                case "sq": return TLineCap.Square; 
                case "flat": return TLineCap.Flat;
                default:
                    return null;
            }
        }

        private TCompoundLineType? GetCompoundLineType(string p)
        {
            switch (p)
            {
                case "sng": return TCompoundLineType.Single;
                case "dbl": return TCompoundLineType.Double;
                case "thickThin": return TCompoundLineType.ThickThin;
                case "thinThick": return TCompoundLineType.ThinThick;
                case "tri": return TCompoundLineType.Triple;
                default:
                    return null;
            }
        }

        private void ReadEffectStyleLst(TEffectStyleList EffectStyleList)
        {
            EffectStyleList.Xml = DataStream.GetXml();
        }

        #endregion

        #endregion

    }
}
