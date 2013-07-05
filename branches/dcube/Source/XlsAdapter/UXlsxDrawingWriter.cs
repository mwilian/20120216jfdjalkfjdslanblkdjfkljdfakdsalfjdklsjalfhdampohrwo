using System;
using System.Collections.Generic;
using System.Text;
using FlexCel.Core;
using System.Globalization;
using System.Reflection;
using System.IO.Packaging;
using System.IO;

namespace FlexCel.XlsAdapter
{
    class TXlsxDrawingWriter
    {
        #region Variables
        private TOpenXmlWriter DataStream;
        private ExcelFile xls;
        private TWorkbook Workbook;

        private int CurrentMediaRelId;
        public int GetCurrentMediaRelId { get { return CurrentMediaRelId; } }

        private Dictionary<string, string> ExistingImages;
        #endregion

        internal TXlsxDrawingWriter(TOpenXmlWriter aDataStream, ExcelFile axls, TWorkbook aWorkbook)
        {
            DataStream = aDataStream;
            xls = axls;
            Workbook = aWorkbook;
            ExistingImages = new Dictionary<string, string>();
        }

        internal TWorkbookGlobals Globals
        {
            get
            {
                return Workbook.Globals;
            }
        }

        #region Drawing
        internal void WriteDrawing(int relId, TSheet Sheet, int SheetId, Uri SheetUri)
        {
            DataStream.CreatePart(new Uri(TOpenXmlManager.DrawingBaseURI + "drawing" + SheetId.ToString(CultureInfo.InvariantCulture) + ".xml", UriKind.Relative),
                TOpenXmlManager.DrawingContentType);

            DataStream.CreateRelationshipFromUri(SheetUri, TOpenXmlWriter.DrawingRelationshipType, relId);

            DataStream.WriteStartDocument("wsDr", "xdr", TOpenXmlManager.SpreadsheetDrawingNamespace);
            DataStream.DefaultNamespacePrefix = "xdr";
            DataStream.WriteAtt("xmlns", "a", null, TOpenXmlManager.DrawingNamespace);
            WriteActualDrawing(Sheet);
            DataStream.WriteEndDocument();
            DataStream.DefaultNamespacePrefix = null;
        }

        private void WriteActualDrawing(TSheet Sheet)
        {
            for (int i = 0; i < Sheet.Drawing.ObjectCount; i++)
			{
                WriteTwoCellAnchor(Sheet, i);
			}
        }

        internal static bool DrawingMustBeSaved(TObjectType ObjType, TEscherOPTRecord opt)
        {
            switch (ObjType)
            {
                case TObjectType.Group:
                    return true;

                case TObjectType.Picture:
                    return opt.HasBlip(); //DDE pictures don't have blips and are saved as legacy

                case TObjectType.Line:
                case TObjectType.Rectangle:
                case TObjectType.Oval:
                case TObjectType.Arc:
                case TObjectType.Text:
                case TObjectType.Polygon:
                case TObjectType.EditBox:
                case TObjectType.MicrosoftOfficeDrawing:
                    return true;
            }
            return false;
        }


        private void WriteTwoCellAnchor(TSheet Sheet, int DrawingId)
        {
            TClientAnchor Anchor = Sheet.Drawing.GetObjectAnchor(DrawingId).Inc(); //Need to inc it so dx can be calculated correctly
            TEscherOPTRecord opt = Sheet.Drawing.GetOPT(DrawingId);
            TMsObj msobj = opt.GetObj();

            if (!DrawingMustBeSaved(msobj.ObjType, opt)) return;
            DataStream.WriteStartElement("twoCellAnchor");

            DataStream.WriteAtt("editAs", GetFlxAnchor(Anchor.AnchorType));

            IRowColSize rc = new RowColSize(xls.HeightCorrection, xls.WidthCorrection, Sheet);
            WriteMarker("from", Anchor.Row1, Anchor.Col1, TDrawingCoordinate.FromPixels(Anchor.Dy1Pix(rc)), TDrawingCoordinate.FromPixels(Anchor.Dx1Pix(rc)));
            WriteMarker("to", Anchor.Row2, Anchor.Col2, TDrawingCoordinate.FromPixels(Anchor.Dy2Pix(rc)), TDrawingCoordinate.FromPixels(Anchor.Dx2Pix(rc)));

            WriteEG_ObjectChoices(Sheet, opt, msobj);
            WriteClientData(opt, msobj);
            DataStream.WriteEndElement();
        }

        private string GetObjName(TObjectType ObjType, TShapeType ShapeType)
        {
            switch (ObjType)
            {
                case TObjectType.MicrosoftOfficeDrawing:
                    return TCompactFramework.EnumGetName(typeof(TShapeType), ShapeType);
                default:
                    return TCompactFramework.EnumGetName(typeof(TObjectType), ObjType);
            }
        }

        private static TBlipFill GetBlipFill(TEscherOPTRecord opt)
        {
            byte[] BlipData;
            TXlsImgType imageType = TXlsImgType.Unknown;
            using (MemoryStream ms = new MemoryStream())
            {
                opt.GetImageFromStream(ms, ref imageType);
                BlipData = ms.ToArray();
            }

            return new TBlipFill(0, true, new TBlip(TBlipCompression.None, BlipData, opt.FileName, GetContentType(imageType)),
                GetSrcRect(opt.CropArea), new TBlipFillStretch(new TDrawingRelativeRect()));
        }

        private static TDrawingRelativeRect? GetSrcRect(TCropArea CropArea)
        {
            return new TDrawingRelativeRect(GetCrop(CropArea.CropFromLeft), GetCrop(CropArea.CropFromTop), GetCrop(CropArea.CropFromRight), GetCrop(CropArea.CropFromBottom));
        }

        private static double GetCrop(int p)
        {
            return p / 65536.0;
        }

        public static string GetContentType(TXlsImgType imageType)
        {
            switch (imageType)
            {
                case TXlsImgType.Gif: return "image/gif";
                case TXlsImgType.Png: return "image/png";
                case TXlsImgType.Tiff: return "image/tiff";
                case TXlsImgType.Jpeg: return "image/jpeg";
                case TXlsImgType.Pict: return "image/pict";
                case TXlsImgType.Wmf: return "image/x-wmf";
                case TXlsImgType.Emf: return "image/x-emf";
                case TXlsImgType.Bmp: return "image/bmp";
            }
            XlsMessages.ThrowException(XlsErr.ErrExcelInvalid);
            return "image";
        }

        private void WriteMarker(string TagName, int Row, int Col, TDrawingCoordinate RowOffs, TDrawingCoordinate ColOffs)
        {
            DataStream.WriteStartElement(TagName);
            DataStream.WriteElement("col", Col - 1); //this one is 0-based
            DataStream.WriteElement("colOff", TOpenXmlWriter.ConvertFromDrawingCoord(ColOffs));
            DataStream.WriteElement("row", Row - 1); //0-based
            DataStream.WriteElement("rowOff", TOpenXmlWriter.ConvertFromDrawingCoord(RowOffs));
            DataStream.WriteEndElement();
        }

        private string GetFlxAnchor(TFlxAnchorType anchorType)
        {
            switch (anchorType)
            {
                case TFlxAnchorType.DontMoveAndDontResize: return "absolute";
                case TFlxAnchorType.MoveAndDontResize: return "oneCell";

                default: return null;
            }
        }

        private void WriteClientData(TEscherOPTRecord opt, TMsObj msobj)
        {
            DataStream.WriteStartElement("clientData", false);
            DataStream.WriteAtt("fLocksWithSheet", msobj.IsLocked, true);
            DataStream.WriteAtt("fPrintsWithSheet", msobj.IsPrintable, true);
            DataStream.WriteEndElement();
        }

        private void WriteEG_ObjectChoices(TSheet Sheet, TEscherOPTRecord opt, TMsObj msobj)
        {
            if (!DrawingMustBeSaved(msobj.ObjType, opt)) return;

            switch (msobj.ObjType)
            {
                case TObjectType.Group:
                    WriteGroup(Sheet, opt, msobj);
                    break;
                case TObjectType.Picture:
                    WritePic(Sheet, opt, msobj);
                    break;

                case TObjectType.Line:
                case TObjectType.Rectangle:
                case TObjectType.Oval:
                case TObjectType.Arc:
                case TObjectType.Text:
                case TObjectType.Polygon:
                case TObjectType.EditBox:
                case TObjectType.MicrosoftOfficeDrawing:
                    WriteShape(Sheet, opt, msobj);
                    break;
            }
        }

        private void WriteShape(TSheet Sheet, TEscherOPTRecord opt, TMsObj msobj)
        {
            DataStream.WriteStartElement("sp");
            DataStream.WriteAtt("macro", msobj.GetFmlaMacroXlsx(Sheet.Cells.CellList));
            DataStream.WriteAtt("fPublished", msobj.IsPublished, false);

            WriteNvSpPr(Sheet, opt, msobj);
            WriteSpPr(Sheet, opt);
            WriteStyle(opt, msobj);
            WriteTxBody(opt, msobj);

            DataStream.WriteEndElement();
        }

        private void WriteNvSpPr(TSheet Sheet, TEscherOPTRecord opt, TMsObj msobj)
        {
            DataStream.WriteStartElement("nvSpPr");
            WriteCNvPr(Sheet, opt, msobj);
            WriteCNvSpPr(opt, msobj);
            DataStream.WriteEndElement();
        }

        private void WriteCNvSpPr(TEscherOPTRecord opt, TMsObj msobj)
        {
            DataStream.WriteStartElement("cNvSpPr", false);
            DataStream.WriteEndElement();
        }

        private void WriteStyle(TEscherOPTRecord opt, TMsObj msobj)
        {
            TShapeLine shline = opt.GetRawShapeLine(); //for the theme we don't need to read xls records, since those don't use themes.
            TShapeFill shfill = opt.GetRawShapeFill();
            TShapeFont shfont = opt.GetRawShapeFont();
            TShapeEffects sheffects = opt.GetRawShapeEffects();

            if (shline == null || shfill == null || shfont == null || sheffects == null) return; //"style" is not required, but if we write it, we need to write all subrecords.

            DataStream.WriteStartElement("style");
            string SaveDefNamespace = DataStream.DefaultNamespacePrefix;
            DataStream.DefaultNamespacePrefix = "a";
            try
            {
                WriteLnRef(shline);
                WriteFillRef(shfill);
                WriteEffectRef(sheffects);
                WriteFontRef(shfont);
                DataStream.WriteEndElement();
            }
            finally
            {
                DataStream.DefaultNamespacePrefix = SaveDefNamespace;
            }
        }

        private void WriteLnRef(TShapeLine shline)
        {
            DataStream.WriteStartElement("lnRef", false);
            if (shline != null)
            {
                DataStream.WriteAtt("idx", shline.GetIdx());
                if (shline.ThemeColor.HasValue) WriteColorDef(shline.ThemeColor.Value);
            }
            DataStream.WriteEndElement();
        }

        private void WriteFillRef(TShapeFill shfill)
        {
            DataStream.WriteStartElement("fillRef", false);
            if (shfill != null)
            {
                DataStream.WriteAtt("idx", shfill.GetIdx());
                if (shfill.ThemeColor.HasValue) WriteColorDef(shfill.ThemeColor.Value);
            }

            DataStream.WriteEndElement();
        }

        private void WriteEffectRef(TShapeEffects shapeeffects)
        {
            DataStream.WriteStartElement("effectRef", false);
            DataStream.WriteAtt("idx", (int)shapeeffects.ThemeStyle);
            WriteColorDef(shapeeffects.ThemeColor);
            DataStream.WriteEndElement();
        }

        private void WriteFontRef(TShapeFont shapefont)
        {
            DataStream.WriteStartElement("fontRef", false);
            DataStream.WriteAtt("idx", GetFontIdx(shapefont.ThemeScheme));
            WriteColorDef(shapefont.ThemeColor);

            DataStream.WriteEndElement();
        }

        private string GetFontIdx(TFontScheme fs)
        {
            switch (fs)
            {
                case TFontScheme.Minor:
                    return "minor";

                case TFontScheme.Major:
                    return "major";

                default:
                    return "none";
            }
        }

        private void WriteTxBody(TEscherOPTRecord opt, TMsObj msobj)
        {
            TDrawingRichString rtext = opt.GetRawTextExt();
            if (rtext == null)
            {
                TTXO txo = opt.GetTXO();
                if (txo == null) return;
                TRichString RichText = txo.GetText();
                rtext = TDrawingRichString.FromRichString(RichText, xls);
            }

            if (rtext == null || string.IsNullOrEmpty(rtext.Value)) return;

            DataStream.WriteStartElement("txBody", false);
            string SaveDefNamespace = DataStream.DefaultNamespacePrefix;
            DataStream.DefaultNamespacePrefix = "a";
            try
            {
                WriteBodyPr(opt);
                WriteLstStyle(opt.GetRawLstStyle());
                WritePs(rtext);
            }
            finally
            {
                DataStream.DefaultNamespacePrefix = SaveDefNamespace;
            }

            DataStream.WriteEndElement();
        }

        private void WriteLstStyle(string LstStyle)
        {
            if (String.IsNullOrEmpty(LstStyle)) return;
            DataStream.WriteRaw(LstStyle);
        }

        private void WriteBodyPr(TEscherOPTRecord opt)
        {
            TBodyPr BodyPr = opt.GetRawBodyPr();
            if (BodyPr == null || String.IsNullOrEmpty(BodyPr.xml))
            {
                DataStream.WriteStartElement("bodyPr", false);
                int Margin;
                if (opt.UseTextMargins())
                {
                    Margin = opt.GetInt(TShapeOption.dxTextLeft, 91440); DataStream.WriteAtt("lIns", Margin, 91440);
                    Margin = opt.GetInt(TShapeOption.dyTextTop, 45720); DataStream.WriteAtt("tIns", Margin, 45720);
                    Margin = opt.GetInt(TShapeOption.dxTextRight, 91440); DataStream.WriteAtt("rIns", Margin, 91440);
                    Margin = opt.GetInt(TShapeOption.dyTextBottom, 45720); DataStream.WriteAtt("bIns", Margin, 45720);
                }
                else
                {
                    DataStream.WriteAtt("lIns", 18288);
                    DataStream.WriteAtt("tIns", 45720 / 2);
                    DataStream.WriteAtt("rIns", 0);
                    DataStream.WriteAtt("bIns", 0);
                }
                DataStream.WriteEndElement();
                
            }
            else
            {
                DataStream.WriteRaw(BodyPr.xml);
            }
        }

        private void WritePs(TDrawingRichString text)
        {
            for(int i = 0; i < text.ParagraphCount; i++)
            {
                TDrawingTextParagraph p = text.Paragraph(i);
                DataStream.WriteStartElement("p", false);

                WritePpr(p.Properties);

                for (int k = 0; k < p.TextRunCount; k++)
                {
                    TDrawingTextRun r = p.TextRun(k);
                    WriteR(r);
                }

                WriteEpr(p.EndParagraphProperties);

                DataStream.WriteEndElement();
            }
        }

        private void WritePpr(TDrawingParagraphProperties pProps)
        {
            TDrawingParagraphProperties def = TDrawingParagraphProperties.Empty;
            DataStream.WriteStartElement("pPr");
            DataStream.WriteAtt("marL", pProps.MarL, def.MarL);
            DataStream.WriteAtt("marR", pProps.MarR, def.MarR);
            DataStream.WriteAtt("lvl", pProps.Lvl, def.Lvl);
            DataStream.WriteAtt("indent", pProps.Indent.Emu, def.Indent.Emu);
            WriteDrawingAlign("algn", pProps.Algn, def.Algn);
            DataStream.WriteAtt("defTabSz", pProps.DefTabSz.Emu, def.DefTabSz.Emu);
            DataStream.WriteAtt("rtl", pProps.Rtl, def.Rtl);
            DataStream.WriteAtt("eaLnBrk", pProps.EaLnBrk, def.EaLnBrk);
            WriteDrawingFontAlign("fontAlgn", pProps.FontAlgn, def.FontAlgn);
            DataStream.WriteAtt("latinLnBrk", pProps.LatinLnBrk, def.LatinLnBrk);
            DataStream.WriteAtt("hangingPunct", pProps.HangingPunct, def.HangingPunct);
            DataStream.WriteEndElement();
        }

        private void WriteDrawingAlign(string p, TDrawingAlignment val, TDrawingAlignment def)
        {
            if (val == def) return;
            switch (val)
            {
                case TDrawingAlignment.Left: DataStream.WriteAtt(p, "l"); break;
                case TDrawingAlignment.Center: DataStream.WriteAtt(p, "ctr"); break;
                case TDrawingAlignment.Right: DataStream.WriteAtt(p, "r"); break;
                case TDrawingAlignment.Justified: DataStream.WriteAtt(p, "just"); break;
                case TDrawingAlignment.JustLow: DataStream.WriteAtt(p, "justLow"); break;
                case TDrawingAlignment.Distributed: DataStream.WriteAtt(p, "dist"); break;
                case TDrawingAlignment.ThaiDist: DataStream.WriteAtt(p, "thaiDist"); break;
            }
        }
        
        private void WriteDrawingFontAlign(string p, TDrawingFontAlign val, TDrawingFontAlign def)
        {
            if (val == def) return;
            switch (val)
            {
                case TDrawingFontAlign.Automatic: DataStream.WriteAtt(p, "auto"); break;
                case TDrawingFontAlign.Top: DataStream.WriteAtt(p, "t"); break;
                case TDrawingFontAlign.Center: DataStream.WriteAtt(p, "ctr"); break;
                case TDrawingFontAlign.BaseLine: DataStream.WriteAtt(p, "base"); break;
                case TDrawingFontAlign.Bottom: DataStream.WriteAtt(p, "b"); break;
            }
        }

        private void WriteR(TDrawingTextRun r)
        {
            bool IsBreak = r.IsBreak;
            if (IsBreak)
            {
                DataStream.WriteStartElement("br", false);
            }
            else
            {
                DataStream.WriteStartElement("r", false);
            }
            WriteRpr(r.TextProperties);
            if (!IsBreak) DataStream.WriteElement("t", r.Text, false);
            DataStream.WriteEndElement();
        }

        private void WriteRpr(TDrawingTextProperties rProps)
        {
            DataStream.WriteStartElement("rPr");

            WriteRprAtts(rProps);
            WriteRprElements(rProps);

            DataStream.WriteEndElement();
        }

        private void WriteRprElements(TDrawingTextProperties rProps)
        {
            WriteLn(rProps.Line);
            WriteFill(rProps.Fill, true);
            WriteEffects(rProps.Effects);
            if (rProps.Highlight.HasValue)
            {
                DataStream.WriteStartElement("highlight");
                WriteColorDef(rProps.Highlight.Value);
                DataStream.WriteEndElement();
            }

            if (rProps.Underline != null)
            {
                if (rProps.Underline.xmlLine != null) DataStream.WriteRaw(rProps.Underline.xmlLine);
                if (rProps.Underline.xmlFill != null) DataStream.WriteRaw(rProps.Underline.xmlFill);
            }


            if (rProps.Latin.HasValue) WriteLatin(rProps.Latin.Value);
            if (rProps.EastAsian.HasValue) WriteEastAsian(rProps.EastAsian.Value);
            if (rProps.ComplexScript.HasValue) WriteComplexScript(rProps.ComplexScript.Value);
            if (rProps.Symbol.HasValue) WriteSymbol(rProps.Symbol.Value);

            WriteDrawingHLink(rProps.HyperlinkClick);
            WriteDrawingHLink(rProps.HyperlinkMouseOver);
            if (rProps.RightToLeft) WriteRtl();

        }

        private void WriteEffects(TEffectProperties props)
        {
            if (props != null) DataStream.WriteRaw(props.xml);
        }

        private void WriteRtl()
        {
            DataStream.WriteStartElement("rtl");
            DataStream.WriteAtt("val", true);
            DataStream.WriteEndElement();
        }

        private void WriteDrawingHLink(TDrawingHyperlink hlink)
        {
            if (hlink == null) return;
            //hyperinks have relationships, in both the id and hte sound. those must be saved too...
            //DataStream.WriteRaw(hlink.xml);
        }

        private void WriteRprAtts(TDrawingTextProperties rProps)
        {
            TDrawingTextAttributes rAtts = rProps.Attributes;
            TDrawingTextAttributes defAtt = TDrawingTextAttributes.Empty;
            DataStream.WriteAtt("kumimoji", rAtts.Kumimoji, defAtt.Kumimoji);
            DataStream.WriteAtt("lang", rAtts.Lang, defAtt.Lang);
            DataStream.WriteAtt("altLang", rAtts.AltLang, defAtt.AltLang);
            DataStream.WriteAtt("sz", rAtts.Size, defAtt.Size);
            DataStream.WriteAtt("b", rAtts.Bold, defAtt.Bold);
            DataStream.WriteAtt("i", rAtts.Italic, defAtt.Italic);
            WriteUnderline("u", rAtts.Underline, defAtt.Underline);
            WriteStrike("strike", rAtts.Strike, defAtt.Strike);
            DataStream.WriteAtt("kern", rAtts.Kern, defAtt.Kern);
            WriteCap("cap", rAtts.Capitalization, defAtt.Capitalization);
            DataStream.WriteAtt("spc", rAtts.Spacing.Emu, defAtt.Spacing.Emu);
            DataStream.WriteAtt("normalizeH", rAtts.NormalizeH, defAtt.NormalizeH);
            if (rAtts.Baseline != defAtt.Baseline) DataStream.WriteAttPercent("baseline", rAtts.Baseline);
            DataStream.WriteAtt("noProof", rAtts.NoProof, defAtt.NoProof);
            DataStream.WriteAtt("dirty", rAtts.Dirty, defAtt.Dirty);
            DataStream.WriteAtt("err", rAtts.Err, defAtt.Err);
            DataStream.WriteAtt("smtClean", rAtts.SmartTagClean, defAtt.SmartTagClean);
            DataStream.WriteAtt("smtId", rAtts.SmartTagId, defAtt.SmartTagId);
            DataStream.WriteAtt("bmk", rAtts.BookmarkLinkTarget, defAtt.BookmarkLinkTarget);
        }

        private void WriteUnderline(string p, TDrawingUnderlineStyle val, TDrawingUnderlineStyle def)
        {
            if (val == def) return;
            switch (val)
            {
                case TDrawingUnderlineStyle.None: DataStream.WriteAtt(p, "none"); break;
                case TDrawingUnderlineStyle.Words: DataStream.WriteAtt(p, "words"); break;
                case TDrawingUnderlineStyle.Single: DataStream.WriteAtt(p, "sng"); break;
                case TDrawingUnderlineStyle.Double: DataStream.WriteAtt(p, "dbl"); break;
                case TDrawingUnderlineStyle.Heavy: DataStream.WriteAtt(p, "heavy"); break;
                case TDrawingUnderlineStyle.Dotted: DataStream.WriteAtt(p, "dotted"); break;
                case TDrawingUnderlineStyle.DottedHeavy: DataStream.WriteAtt(p, "dottedHeavy"); break;
                case TDrawingUnderlineStyle.Dash: DataStream.WriteAtt(p, "dash"); break;
                case TDrawingUnderlineStyle.DashHeavy: DataStream.WriteAtt(p, "dashHeavy"); break;
                case TDrawingUnderlineStyle.DashLong: DataStream.WriteAtt(p, "dashLong"); break;
                case TDrawingUnderlineStyle.DashLongHeavy: DataStream.WriteAtt(p, "dashLongHeavy"); break;
                case TDrawingUnderlineStyle.DotDash: DataStream.WriteAtt(p, "dotDash"); break;
                case TDrawingUnderlineStyle.DotDashHeavy: DataStream.WriteAtt(p, "dotDashHeavy"); break;
                case TDrawingUnderlineStyle.DotDotDash: DataStream.WriteAtt(p, "dotDotDash"); break;
                case TDrawingUnderlineStyle.DotDotDashHeavy: DataStream.WriteAtt(p, "dotDotDashHeavy"); break;
                case TDrawingUnderlineStyle.Wavy: DataStream.WriteAtt(p, "wavy"); break;
                case TDrawingUnderlineStyle.WavyHeavy: DataStream.WriteAtt(p, "wavyHeavy"); break;
                case TDrawingUnderlineStyle.WavyDouble: DataStream.WriteAtt(p, "wavyDbl"); break;
            }
        }

        private void WriteStrike(string p, TDrawingTextStrike val, TDrawingTextStrike def)
        {
            if (val == def) return;
            switch (val)
            {
                case TDrawingTextStrike.None: DataStream.WriteAtt(p, "noStrike"); break;
                case TDrawingTextStrike.Single: DataStream.WriteAtt(p, "sngStrike"); break;
                case TDrawingTextStrike.Double: DataStream.WriteAtt(p, "dblStrike"); break;
            }
        }

        private void WriteCap(string p, TDrawingTextCapitalization val, TDrawingTextCapitalization def)
        {
            if (val == def) return;
            switch (val)
            {
                case TDrawingTextCapitalization.None: DataStream.WriteAtt(p, "none"); break;
                case TDrawingTextCapitalization.Small: DataStream.WriteAtt(p, "small"); break;
                case TDrawingTextCapitalization.All: DataStream.WriteAtt(p, "all"); break;
            }
        }


        private void WriteEpr(TDrawingTextProperties rProps)
        {
            WriteRpr(rProps);
        }

        private void WriteGroup(TSheet Sheet, TEscherOPTRecord opt, TMsObj msobj)
        {
            DataStream.WriteStartElement("grpSp");
            WriteNvGrpSpPr(Sheet, opt, msobj);
            WriteGrpSpPr(opt, msobj);

            TEscherContainerRecord GrpParent = opt.Parent.Parent;
            for (int i = 1; i < GrpParent.FContainedRecords.Count; i++)
            {
                TEscherContainerRecord GrpNext = GrpParent.ContainedRecords[i] as TEscherContainerRecord;
                if (GrpNext is TEscherSpgrContainerRecord) GrpNext = GrpNext.ContainedRecords[0] as TEscherContainerRecord;
                TEscherOPTRecord NextOpt = GrpNext.FindRec<TEscherOPTRecord>();
                WriteEG_ObjectChoices(Sheet, NextOpt, NextOpt.GetObj());                
            }
            DataStream.WriteEndElement();
        }

        private void WriteGrpSpPr(TEscherOPTRecord opt, TMsObj msobj)
        {
            DataStream.WriteStartElement("grpSpPr");
            TEscherSpgrRecord Spgr = opt.Parent.FindRec<TEscherSpgrRecord>();
            int[] Bounds = Spgr.Bounds;
            WriteXfrm(opt, Bounds);

            DataStream.WriteEndElement();
        }

        private void WriteXfrm(TEscherOPTRecord opt, int[] ChildBounds)
        {
            string SaveDefNamespace = DataStream.DefaultNamespacePrefix;
            DataStream.DefaultNamespacePrefix = "a";
            try
            {
                DataStream.WriteStartElement("xfrm");
                DataStream.WriteAttAsAngle("rot", opt.Rotation, 0);
                TEscherSpContainerRecord SpC = opt.Parent as TEscherSpContainerRecord;
                DataStream.WriteAtt("flipH", (SpC.SP.Flags & 0x40) != 0, false);
                DataStream.WriteAtt("flipV", (SpC.SP.Flags & 0x80) != 0, false);

                int[] Bounds = GetChildAnchor(opt);
                if (Bounds == null) Bounds = ChildBounds;

                if (Bounds != null)
                {
                    WriteOffs("off", Bounds[0], Bounds[1]);
                    WriteExt("ext", Bounds[2] - Bounds[0], Bounds[3] - Bounds[1]);
                    if (ChildBounds != null)
                    {
                        WriteOffs("chOff", ChildBounds[0], ChildBounds[1]);
                        WriteExt("chExt", ChildBounds[2] - ChildBounds[0], ChildBounds[3] - ChildBounds[1]);
                    }
                }
                DataStream.WriteEndElement();
            }
            finally
            {
                DataStream.DefaultNamespacePrefix = SaveDefNamespace;
            }
        }

        private int[] GetChildAnchor(TEscherOPTRecord opt)
        {
            TEscherChildAnchorRecord Anchor = opt.Parent.FindRec<TEscherChildAnchorRecord>();
            if (Anchor != null)
            {
                return new int[] { Anchor.Dx1, Anchor.Dy1, Anchor.Dx2, Anchor.Dy2 };

            }
            return null;
        }

        private void WriteOffs(string name, int dx1, int dy1)
        {
            DataStream.WriteStartElement(name);
            DataStream.WriteAtt("x", dx1);
            DataStream.WriteAtt("y", dy1);
            DataStream.WriteEndElement();
        }

        private void WriteExt(string name, int w, int h)
        {
            DataStream.WriteStartElement(name);
            DataStream.WriteAtt("cx", w);
            DataStream.WriteAtt("cy", h);
            DataStream.WriteEndElement();
        }

        private void WriteNvGrpSpPr(TSheet Sheet, TEscherOPTRecord opt, TMsObj msobj)
        {
            DataStream.WriteStartElement("nvGrpSpPr", false);
            WriteCNvPr(Sheet, opt, msobj);
            WriteCNvGrpSpPr(opt, msobj);
            WriteNvPr(opt, msobj);
            DataStream.WriteEndElement();
        }

        private void WriteNvPr(TEscherOPTRecord opt, TMsObj msobj)
        {
        }

        private void WriteCNvGrpSpPr(TEscherOPTRecord opt, TMsObj msobj)
        {
            DataStream.WriteStartElement("cNvGrpSpPr", false);
            DataStream.WriteEndElement();
        }

        private void WritePic(TSheet Sheet, TEscherOPTRecord opt, TMsObj msobj)
        {

            DataStream.WriteStartElement("pic");
            DataStream.WriteAtt("macro", msobj.GetFmlaMacroXlsx(Sheet.Cells.CellList));
            DataStream.WriteAtt("fPublished", msobj.IsPublished, false);
            WriteNvPicPr(Sheet, opt, msobj);
            if (opt.HasBlip())
            {
                TBlipFill BlipFill = GetBlipFill(opt);
                WriteDrawingBlipFill(BlipFill);
            }
            WriteSpPr(Sheet, opt);
            DataStream.WriteEndElement();
        }

        private void WriteNvPicPr(TSheet Sheet, TEscherOPTRecord opt, TMsObj msobj)
        {
            DataStream.WriteStartElement("nvPicPr", false);
            WriteCNvPr(Sheet, opt, msobj);
            WriteCNvPicPr();
            WriteNvPr();
            DataStream.WriteEndElement();
        }

        private void WriteCNvPr(TSheet Sheet, TEscherOPTRecord opt, TMsObj msobj)
        {
            DataStream.WriteStartElement("cNvPr", false);
            DataStream.WriteAtt("id", opt.ShapeId());

            TShapeType ShapeType = Sheet.Drawing.ShapeType(opt);
            string ShapeName = opt.ShapeName;
            if (string.IsNullOrEmpty(ShapeName)) ShapeName = GetObjName(msobj.ObjType, ShapeType) + " " + opt.ShapeId(); //this value can't be null

            DataStream.WriteAtt("name", ShapeName, false);
            DataStream.WriteAtt("descr", opt.AltText); 
            DataStream.WriteAtt("hidden", !opt.Visible, false);
            DataStream.WriteAtt("title", null);

            WriteDrawingHLink(opt.GetRawHLinkClick());
            WriteDrawingHLink(opt.GetRawHLinkHover());

           DataStream.WriteEndElement();
        }

        private void WriteCNvPicPr()
        {
            DataStream.WriteStartElement("cNvPicPr", false);
            DataStream.WriteEndElement();
        }

        private void WriteNvPr()
        {
            DataStream.WriteStartElement("nvPr");
            DataStream.WriteEndElement();
        }

        private void WriteDrawingBlipFill(TBlipFill BlipFill)
        {
            DataStream.WriteStartElement("blipFill", false);

            string SaveDefNamespace = DataStream.DefaultNamespacePrefix;
            DataStream.DefaultNamespacePrefix = "a";
            try
            {
                WriteBlipFill(BlipFill);
            }
            finally
            {
                DataStream.DefaultNamespacePrefix = SaveDefNamespace;
            }

            DataStream.WriteEndElement();
        }

        private void WriteSpPr(TSheet Sheet, TEscherOPTRecord opt)
        {
            DataStream.WriteStartElement("spPr", false);
            DataStream.WriteAtt("bwMode", GetBWMode(opt.BwMode));

            string SaveDefNamespace = DataStream.DefaultNamespacePrefix;
            DataStream.DefaultNamespacePrefix = "a";
            try
            {
                WriteXfrm(opt, null);
                WriteGeom(opt);
                TShapeFill fs = opt.GetFillColor(xls);
                if (fs != null) WriteFill(fs.FillStyle, fs.HasFill);

                TShapeLine ln = opt.GetLine(xls);
                if (ln != null) WriteLn(ln.LineStyle);
                WriteEffects(opt.GetRawEffectProps());
                //    WriteScene3d(opt);
                //    WriteSp3d(opt);
                //    DataStream.WriteFutureStorage(opt.FutureStorage);
            }
            finally
            {
                DataStream.DefaultNamespacePrefix = SaveDefNamespace;
            }

            DataStream.WriteEndElement();
        }

        private string GetBWMode(TBwMode BwMode)
        {
            switch (BwMode)
            {
                case TBwMode.GrayScale: return "gray";
                case TBwMode.LightGrayScale: return "ltGray";
                case TBwMode.InverseGray: return "invGray";
                case TBwMode.GrayOutline: return "grayWhite";
                case TBwMode.BlackTextLine: return "blackGray";
                case TBwMode.HighContrast: return "blackWhite";
                case TBwMode.Black: return "black";
                case TBwMode.White: return "white";
                case TBwMode.DontShow: return "hidden";
            }

            return null;
        }


        private void WriteGeom(TEscherOPTRecord opt)
        {
            string SaveDefNamespace = DataStream.DefaultNamespacePrefix;
            DataStream.DefaultNamespacePrefix = "a";
            try
            {
                TShapeGeom geom = opt.GetGeom();
                if (geom == null) geom = new TShapeGeom("rect");
                if (geom.Name != null)
                {
                    WritePrstGeom(geom);
                }
                else
                {
                    WriteCustGeom(geom);
                }
            }
            finally
            {
                DataStream.DefaultNamespacePrefix = SaveDefNamespace;
            }
        }

        private void WriteCustGeom(TShapeGeom geom)
        {
            DataStream.WriteStartElement("custGeom", false);
            WriteAvLst(geom);
            WriteGdLst(geom);
            WriteAhLst(geom);
            WriteCxnLst(geom);
            WriteRect(geom);
            WritePathLst(geom);
            DataStream.WriteEndElement();
        }

        private void WritePathLst(TShapeGeom geom)
        {
            DataStream.WriteStartElement("pathLst", false);
            foreach (TShapePath path in geom.PathList)
            {
                WritePath(path);
            }
            DataStream.WriteEndElement();            
        }

        private void WritePath(TShapePath path)
        {
            DataStream.WriteStartElement("path");
            DataStream.WriteAtt("w", path.Width, 0);
            DataStream.WriteAtt("h", path.Height, 0);
            DataStream.WriteAtt("fill", GetPathFill(path.PathFill));
            DataStream.WriteAtt("stroke", path.PathStroke, true);
            DataStream.WriteAtt("extrusionOk", path.ExtrusionOk, true);

            foreach (TShapeAction Action in path.Actions)
            {
                switch (Action.ActionType)
                {
                    case TShapeActionType.Close:
                        DataStream.WriteStartElement("close", false);
                        DataStream.WriteEndElement();
                        break;

                    case TShapeActionType.MoveTo:
                        DataStream.WriteStartElement("moveTo", false);
                        WritePt(((TShapeActionMoveTo)Action).Target);
                        DataStream.WriteEndElement();
                        break;

                    case TShapeActionType.LineTo:
                        DataStream.WriteStartElement("lnTo", false);
                        WritePt(((TShapeActionLineTo)Action).Target);
                        DataStream.WriteEndElement();
                        break;

                    case TShapeActionType.ArcTo:
                        DataStream.WriteStartElement("arcTo", false);
                        TShapeActionArcTo ArcTo = ((TShapeActionArcTo)Action);
                        DataStream.WriteAtt("wR", ArcTo.WidthRadius.NameOrValue());
                        DataStream.WriteAtt("hR", ArcTo.HeightRadius.NameOrValue());
                        DataStream.WriteAtt("stAng", ArcTo.StartAngle.NameOrValue());
                        DataStream.WriteAtt("swAng", ArcTo.SwingAngle.NameOrValue());
                        DataStream.WriteEndElement();
                        break;

                    case TShapeActionType.CubicBezierTo:
                        DataStream.WriteStartElement("cubicBezTo", false);
                        WriteBezPoints(Action);
                        DataStream.WriteEndElement();
                        break;

                    case TShapeActionType.QuadBezierTo:
                        DataStream.WriteStartElement("quadBezTo", false);
                        WriteBezPoints(Action);
                        DataStream.WriteEndElement();
                        break;
                }
            }

            DataStream.WriteEndElement();
        }

        private void WriteBezPoints(TShapeAction Action)
        {
            TShapeActionBezierTo BezTo = ((TShapeActionBezierTo)Action);
            foreach (TShapePoint point in BezTo.Target)
            {
                WritePt(point);
            }
        }

        private string GetPathFill(TPathFillMode p)
        {
            switch (p)
            {
                case TPathFillMode.None: return "none";
                case TPathFillMode.Norm: return "norm";
                case TPathFillMode.Lighten: return "lighten";
                case TPathFillMode.LightenLess: return "lightenLess";
                case TPathFillMode.Darken: return "darken";
                case TPathFillMode.DarkenLess: return "darkenLess";
            }
            return "norm";
        }

        private void WriteRect(TShapeGeom geom)
        {
            if (geom.TextRect == null) return;
            DataStream.WriteStartElement("rect");
            DataStream.WriteAtt("l", geom.TextRect.Left.NameOrValue());
            DataStream.WriteAtt("t", geom.TextRect.Top.NameOrValue());
            DataStream.WriteAtt("r", geom.TextRect.Right.NameOrValue());
            DataStream.WriteAtt("b", geom.TextRect.Bottom.NameOrValue());
            DataStream.WriteEndElement();

        }

        private void WriteGdLst(TShapeGeom geom)
        {
            DataStream.WriteStartElement("gdLst", false);
            WriteGuides(geom.GdList);
            DataStream.WriteEndElement();
        }

        private void WriteAhLst(TShapeGeom geom)
        {
            DataStream.WriteStartElement("ahLst", false);
            foreach (TShapeAdjustHandle ah in geom.AhList)
            {
                WriteAdjustHandle(ah);
            }
            DataStream.WriteEndElement();
        }

        private void WriteAdjustHandle(TShapeAdjustHandle ah)
        {
            TShapeAdjustHandleXY ahXY = ah as TShapeAdjustHandleXY;
            if (ahXY != null)
            {
                DataStream.WriteStartElement("ahXY");
                if (ahXY.GdRefX != null)
                {
                    DataStream.WriteAtt("gdRefX", ahXY.GdRefX.Name);
                    DataStream.WriteAtt("minX", ahXY.MinX.NameOrValue());
                    DataStream.WriteAtt("maxX", ahXY.MaxX.NameOrValue());
                }

                if (ahXY.GdRefY != null)
                {
                    DataStream.WriteAtt("gdRefY", ahXY.GdRefY.Name);
                    DataStream.WriteAtt("minY", ahXY.MinY.NameOrValue());
                    DataStream.WriteAtt("maxY", ahXY.MaxY.NameOrValue());
                }
            }
            else
            {
                TShapeAdjustHandlePolar ahp = (TShapeAdjustHandlePolar)ah;
                DataStream.WriteStartElement("ahPolar");
                if (ahp.GdRefR != null)
                {
                    DataStream.WriteAtt("gdRefR", ahp.GdRefR.Name);
                    DataStream.WriteAtt("minR", ahp.MinR.NameOrValue());
                    DataStream.WriteAtt("maxR", ahp.MaxR.NameOrValue());
                }

                if (ahp.GdRefAng != null)
                {
                    DataStream.WriteAtt("gdRefAng", ahp.GdRefAng.Name);
                    DataStream.WriteAtt("minAng", ahp.MinAng.NameOrValue());
                    DataStream.WriteAtt("maxAng", ahp.MaxAng.NameOrValue());
                }

            }

            WritePos(ah.Location);

            DataStream.WriteEndElement();
        }

        private void WritePos(TShapePoint Location)
        {
            WritePos(Location, "pos");
        }

        private void WritePt(TShapePoint Location)
        {
            WritePos(Location, "pt");
        }

        private void WritePos(TShapePoint Location, string TagName)
        {
            DataStream.WriteStartElement(TagName);
            DataStream.WriteAtt("x", Location.x.NameOrValue());
            DataStream.WriteAtt("y", Location.y.NameOrValue());
            DataStream.WriteEndElement();
        }

        private void WriteCxnLst(TShapeGeom geom)
        {
            DataStream.WriteStartElement("cxnLst", false);
            foreach (TShapeConnection cxn in geom.ConnList)
            {
                WriteCxn(cxn);
            }
            DataStream.WriteEndElement();
        }

        private void WriteCxn(TShapeConnection cxn)
        {
            DataStream.WriteStartElement("cxn");
            DataStream.WriteAtt("ang", cxn.Angle.NameOrValue());
            WritePos(cxn.Position);
            DataStream.WriteEndElement();
        }

        private void WritePrstGeom(TShapeGeom geom)
        {
            DataStream.WriteStartElement("prstGeom", false);
            DataStream.WriteAtt("prst", geom.Name);
            WriteAvLst(geom);
            DataStream.WriteEndElement();
        }

        private void WriteAvLst(TShapeGeom geom)
        {
            DataStream.WriteStartElement("avLst", false);
            WriteGuides(geom.AvList);
            DataStream.WriteEndElement();
        }

        private void WriteGuides(TShapeGuideList guides)
        {
            foreach (TShapeGuide av in guides)
            {
                if (av.Name != null) WriteGuide(av);
            }
        }

        private void WriteGuide(TShapeGuide av)
        {
            DataStream.WriteStartElement("gd");
            DataStream.WriteAtt("name", av.Name);
            DataStream.WriteAtt("fmla", av.Fmla.XlsxString());
            DataStream.WriteEndElement();
        }

        #endregion


        #region Theme
        internal void WriteThemeManager()
        {
            DataStream.CreatePart(TOpenXmlWriter.ThemeManagerURI, TOpenXmlWriter.ThemeManagerContentType);
            DataStream.CreateRelationshipFromUri(null, TOpenXmlWriter.documentRelationshipType, TOpenXmlManager.RelIdThemeManager);

            DataStream.WriteStartDocument("themeManager", null, TOpenXmlManager.DrawingNamespace);
            DataStream.WriteEndDocument();
        }

        internal void WriteTheme(bool Standalone)
        {
            DataStream.CreatePart(TOpenXmlWriter.ThemeURI, TOpenXmlWriter.ThemeContentType);
            Uri RootUri = Standalone ? TOpenXmlWriter.ThemeManagerURI : TOpenXmlWriter.WorkbookURI;
            DataStream.CreateRelationshipFromUri(RootUri, TOpenXmlWriter.themeRelationshipType, Globals.SheetCount + TOpenXmlManager.RelIdThemes);

            DataStream.WriteStartDocument("theme", "a", TOpenXmlManager.DrawingNamespace);
            string SaveDefNamespace = DataStream.DefaultNamespacePrefix;
            DataStream.DefaultNamespacePrefix = "a";
            try
            {
                DataStream.WriteAtt("name", Globals.Theme.Name);
                WriteActualTheme();
                DataStream.WriteEndDocument();
            }
            finally
            {
                DataStream.DefaultNamespacePrefix = SaveDefNamespace;
            }

        }

        private void WriteActualTheme()
        {
            WriteThemeElements();
            //WriteObjectDefaults();
            //WriteExtraClrSchemeLst();
            //WriteCustClrLst();
            WriteTheme_ExtLst();
        }

        private void WriteThemeElements()
        {
            DataStream.WriteStartElement("themeElements");
            WriteClrScheme();
            WriteFontScheme();
            WriteFmtScheme();
            WriteThemeElements_ExtLst();
            DataStream.WriteEndElement();
        }

        #region Color Scheme
        private void WriteClrScheme()
        {
            DataStream.WriteStartElement("clrScheme");
            DataStream.WriteAtt("name", Globals.Theme.Elements.ColorScheme.Name);

            WriteColorDef(TThemeColor.Foreground1, "dk1");
            WriteColorDef(TThemeColor.Background1, "lt1");
            WriteColorDef(TThemeColor.Foreground2, "dk2");
            WriteColorDef(TThemeColor.Background2, "lt2");
            WriteColorDef(TThemeColor.Accent1, "accent1");
            WriteColorDef(TThemeColor.Accent2, "accent2");
            WriteColorDef(TThemeColor.Accent3, "accent3");
            WriteColorDef(TThemeColor.Accent4, "accent4");
            WriteColorDef(TThemeColor.Accent5, "accent5");
            WriteColorDef(TThemeColor.Accent6, "accent6");

            WriteColorDef(TThemeColor.HyperLink, "hlink");
            WriteColorDef(TThemeColor.FollowedHyperLink, "folHlink");

            DataStream.WriteFutureStorage(Globals.ThemeRecord.ColorFutureStorage);

            DataStream.WriteEndElement();
        }

        private void WriteColorDef(TThemeColor tc, string ColorName)
        {
            TDrawingColor cdef = Globals.Theme.Elements.ColorScheme[tc];
            DataStream.WriteStartElement(ColorName);

            WriteColorDef(cdef);

            DataStream.WriteEndElement();

        }

        private void WriteColorDef(TDrawingColor cdef)
        {
            switch (cdef.ColorType)
            {
                case TDrawingColorType.HSL:
                    WriteHslClr(cdef.HSL);
                    break;

                case TDrawingColorType.Preset:
                    WritePrstClr(cdef.Preset);
                    break;

                case TDrawingColorType.RGB:
                    WriteSrgbClr(cdef.RGB);
                    break;

                case TDrawingColorType.scRGB:
                    WriteScrgbClr(cdef.ScRGB);
                    break;

                case TDrawingColorType.System:
                    WriteSysClr(cdef.System);
                    break;

                case TDrawingColorType.Theme:
                    WriteSchemeClr(cdef.Theme);
                    break;

                default:
                    XlsMessages.ThrowException(XlsErr.ErrInternal);
                    break;
            }
            
            WriteColorTransform(cdef.GetColorTransform());
            DataStream.WriteEndElement();
        }

        private void WriteColorTransform(TColorTransform[] Transform)
        {
            if (Transform == null) return;
            foreach (TColorTransform ct in Transform)
            {
                bool NeedsValue = true;
                switch (ct.ColorTransformType)
                {
                    case TColorTransformType.Tint: DataStream.WriteStartElement("tint");break;
                    case TColorTransformType.Shade: DataStream.WriteStartElement("shade");break;
                    case TColorTransformType.Inverse: DataStream.WriteStartElement("inv", false); NeedsValue = false; break;
                    case TColorTransformType.Complement: DataStream.WriteStartElement("comp", false); NeedsValue = false; break;
                    case TColorTransformType.Gray: DataStream.WriteStartElement("gray", false); NeedsValue = false; break;
                    case TColorTransformType.Alpha: DataStream.WriteStartElement("alpha");break;
                    case TColorTransformType.AlphaOff: DataStream.WriteStartElement("alphaOff");break;
                    case TColorTransformType.AlphaMod: DataStream.WriteStartElement("alphaMod"); break;
                    case TColorTransformType.Hue: DataStream.WriteStartElement("hue"); DataStream.WriteAttAsAngle("val", ct.Value); NeedsValue = false; break;
                    case TColorTransformType.HueOff: DataStream.WriteStartElement("hueOff"); DataStream.WriteAttAsAngle("val", ct.Value); NeedsValue = false; break;
                    case TColorTransformType.HueMod: DataStream.WriteStartElement("hueMod");break;
                    case TColorTransformType.Sat: DataStream.WriteStartElement("sat");break;
                    case TColorTransformType.SatOff: DataStream.WriteStartElement("satOff");break;
                    case TColorTransformType.SatMod: DataStream.WriteStartElement("satMod");break;
                    case TColorTransformType.Lum: DataStream.WriteStartElement("lum");break;
                    case TColorTransformType.LumOff: DataStream.WriteStartElement("lumOff");break;
                    case TColorTransformType.LumMod: DataStream.WriteStartElement("lumMod");break;
                    case TColorTransformType.Red: DataStream.WriteStartElement("red");break;
                    case TColorTransformType.RedOff: DataStream.WriteStartElement("redOff");break;
                    case TColorTransformType.RedMod: DataStream.WriteStartElement("redMod");break;
                    case TColorTransformType.Green: DataStream.WriteStartElement("green");break;
                    case TColorTransformType.GreenOff: DataStream.WriteStartElement("greenOff");break;
                    case TColorTransformType.GreenMod: DataStream.WriteStartElement("greenMod");break;
                    case TColorTransformType.Blue: DataStream.WriteStartElement("blue");break;
                    case TColorTransformType.BlueOff: DataStream.WriteStartElement("blueOff");break;
                    case TColorTransformType.BlueMod: DataStream.WriteStartElement("blueMod");break;
                    case TColorTransformType.Gamma: DataStream.WriteStartElement("gamma", false); NeedsValue = false; break;
                    case TColorTransformType.InvGamma: DataStream.WriteStartElement("invGamma", false); NeedsValue = false; break;

                    default:
                        FlxMessages.ThrowException(FlxErr.ErrInternal);
                        break;
                }

                if (NeedsValue) DataStream.WriteAttPercent("val", ct.Value);
                DataStream.WriteEndElement();
            }
        }

        private void WriteScrgbClr(TScRGBColor aColor)
        {
            DataStream.WriteStartElement("scrgbClr");
            DataStream.WriteAttPercent("r", aColor.ScR);
            DataStream.WriteAttPercent("g", aColor.ScG);
            DataStream.WriteAttPercent("b", aColor.ScB);
        }

        private void WriteSrgbClr(long aColor)
        {
            DataStream.WriteStartElement("srgbClr");
            DataStream.WriteAttHex("val", aColor & ~0xFF000000, 6);
        }

        private void WriteHslClr(THSLColor aColor)
        {
            DataStream.WriteStartElement("hslClr");
            DataStream.WriteAttAsAngle("hue", aColor.Hue);
            DataStream.WriteAttPercent("sat", aColor.Sat);
            DataStream.WriteAttPercent("lum", aColor.Lum);
        }

        private void WriteSysClr(TSystemColor aColor)
        {
            DataStream.WriteStartElement("sysClr");
            DataStream.WriteAtt("val", GetSysColor(aColor));
            DataStream.WriteAttHex("lastClr", TDrawingColor.GetSystemColor(aColor).ToArgb() & ~0xFF000000, 6);
        }

        private string GetSysColor(TSystemColor aColor)
        {
            switch (aColor)
            {
                case TSystemColor.ScrollBar: return "scrollBar";
                case TSystemColor.Background: return "background";
                case TSystemColor.ActiveCaption: return "activeCaption";
                case TSystemColor.InactiveCaption: return "inactiveCaption";
                case TSystemColor.Menu: return "menu";
                case TSystemColor.Window: return "window";
                case TSystemColor.WindowFrame: return "windowFrame";
                case TSystemColor.MenuText: return "menuText";
                case TSystemColor.WindowText: return "windowText";
                case TSystemColor.CaptionText: return "captionText";
                case TSystemColor.ActiveBorder: return "activeBorder";
                case TSystemColor.InactiveBorder: return "inactiveBorder";
                case TSystemColor.AppWorkspace: return "appWorkspace";
                case TSystemColor.Highlight: return "highlight";
                case TSystemColor.HighlightText: return "highlightText";
                case TSystemColor.BtnFace: return "btnFace";
                case TSystemColor.BtnShadow: return "btnShadow";
                case TSystemColor.GrayText: return "grayText";
                case TSystemColor.BtnText: return "btnText";
                case TSystemColor.InactiveCaptionText: return "inactiveCaptionText";
                case TSystemColor.BtnHighlight: return "btnHighlight";
                case TSystemColor.DkShadow3d: return "3dDkShadow";
                case TSystemColor.Light3d: return "3dLight";
                case TSystemColor.InfoText: return "infoText";
                case TSystemColor.InfoBk: return "infoBk";
                case TSystemColor.HotLight: return "hotLight";
                case TSystemColor.GradientActiveCaption: return "gradientActiveCaption";
                case TSystemColor.GradientInactiveCaption: return "gradientInactiveCaption";
                case TSystemColor.MenuHighlight: return "menuHighlight";
                case TSystemColor.MenuBar: return "menuBar";
            }
            return String.Empty;
        }

        private void WriteSchemeClr(TThemeColor aColor)
        {
            DataStream.WriteStartElement("schemeClr");
            DataStream.WriteAtt("val", GetThemeColor(aColor));
        }

        private string GetThemeColor(TThemeColor aColor)
        {
            switch (aColor)
            {
                case TThemeColor.Accent1: return "accent1";
                case TThemeColor.Accent2: return "accent2";
                case TThemeColor.Accent3: return "accent3";
                case TThemeColor.Accent4: return "accent4";
                case TThemeColor.Accent5: return "accent5";
                case TThemeColor.Accent6: return "accent6";
                case TThemeColor.HyperLink: return "hlink";
                case TThemeColor.FollowedHyperLink: return "folHlink";
                case TThemeColor.None: return "phClr";
                case TThemeColor.Foreground1: return "dk1";
                case TThemeColor.Background1: return "lt1";
                case TThemeColor.Foreground2: return "dk2";
                case TThemeColor.Background2: return "lt2";
            }
            return String.Empty;
        }

        private void WritePrstClr(TPresetColor aColor)
        {
            DataStream.WriteStartElement("prstClr");
            DataStream.WriteAtt("val", GetPresetColor(aColor));
        }

        private static string GetPresetColor(TPresetColor aColor)
        {
            #region A lot of Preset colors
            switch (aColor)
            {
                case TPresetColor.AliceBlue: return "aliceBlue";
                case TPresetColor.AntiqueWhite: return "antiqueWhite";
                case TPresetColor.Aqua: return "aqua";
                case TPresetColor.Aquamarine: return "aquamarine";
                case TPresetColor.Azure: return "azure";
                case TPresetColor.Beige: return "beige";
                case TPresetColor.Bisque: return "bisque";
                case TPresetColor.Black: return "black";
                case TPresetColor.BlanchedAlmond: return "blanchedAlmond";
                case TPresetColor.Blue: return "blue";
                case TPresetColor.BlueViolet: return "blueViolet";
                case TPresetColor.Brown: return "brown";
                case TPresetColor.BurlyWood: return "burlyWood";
                case TPresetColor.CadetBlue: return "cadetBlue";
                case TPresetColor.Chartreuse: return "chartreuse";
                case TPresetColor.Chocolate: return "chocolate";
                case TPresetColor.Coral: return "coral";
                case TPresetColor.CornflowerBlue: return "cornflowerBlue";
                case TPresetColor.Cornsilk: return "cornsilk";
                case TPresetColor.Crimson: return "crimson";
                case TPresetColor.Cyan: return "cyan";
                case TPresetColor.DkBlue: return "dkBlue";
                case TPresetColor.DkCyan: return "dkCyan";
                case TPresetColor.DkGoldenrod: return "dkGoldenrod";
                case TPresetColor.DkGray: return "dkGray";
                case TPresetColor.DkGreen: return "dkGreen";
                case TPresetColor.DkKhaki: return "dkKhaki";
                case TPresetColor.DkMagenta: return "dkMagenta";
                case TPresetColor.DkOliveGreen: return "dkOliveGreen";
                case TPresetColor.DkOrange: return "dkOrange";
                case TPresetColor.DkOrchid: return "dkOrchid";
                case TPresetColor.DkRed: return "dkRed";
                case TPresetColor.DkSalmon: return "dkSalmon";
                case TPresetColor.DkSeaGreen: return "dkSeaGreen";
                case TPresetColor.DkSlateBlue: return "dkSlateBlue";
                case TPresetColor.DkSlateGray: return "dkSlateGray";
                case TPresetColor.DkTurquoise: return "dkTurquoise";
                case TPresetColor.DkViolet: return "dkViolet";
                case TPresetColor.DeepPink: return "deepPink";
                case TPresetColor.DeepSkyBlue: return "deepSkyBlue";
                case TPresetColor.DimGray: return "dimGray";
                case TPresetColor.DodgerBlue: return "dodgerBlue";
                case TPresetColor.Firebrick: return "firebrick";
                case TPresetColor.FloralWhite: return "floralWhite";
                case TPresetColor.ForestGreen: return "forestGreen";
                case TPresetColor.Fuchsia: return "fuchsia";
                case TPresetColor.Gainsboro: return "gainsboro";
                case TPresetColor.GhostWhite: return "ghostWhite";
                case TPresetColor.Gold: return "gold";
                case TPresetColor.Goldenrod: return "goldenrod";
                case TPresetColor.Gray: return "gray";
                case TPresetColor.Green: return "green";
                case TPresetColor.GreenYellow: return "greenYellow";
                case TPresetColor.Honeydew: return "honeydew";
                case TPresetColor.HotPink: return "hotPink";
                case TPresetColor.IndianRed: return "indianRed";
                case TPresetColor.Indigo: return "indigo";
                case TPresetColor.Ivory: return "ivory";
                case TPresetColor.Khaki: return "khaki";
                case TPresetColor.Lavender: return "lavender";
                case TPresetColor.LavenderBlush: return "lavenderBlush";
                case TPresetColor.LawnGreen: return "lawnGreen";
                case TPresetColor.LemonChiffon: return "lemonChiffon";
                case TPresetColor.LtBlue: return "ltBlue";
                case TPresetColor.LtCoral: return "ltCoral";
                case TPresetColor.LtCyan: return "ltCyan";
                case TPresetColor.LtGoldenrodYellow: return "ltGoldenrodYellow";
                case TPresetColor.LtGray: return "ltGray";
                case TPresetColor.LtGreen: return "ltGreen";
                case TPresetColor.LtPink: return "ltPink";
                case TPresetColor.LtSalmon: return "ltSalmon";
                case TPresetColor.LtSeaGreen: return "ltSeaGreen";
                case TPresetColor.LtSkyBlue: return "ltSkyBlue";
                case TPresetColor.LtSlateGray: return "ltSlateGray";
                case TPresetColor.LtSteelBlue: return "ltSteelBlue";
                case TPresetColor.LtYellow: return "ltYellow";
                case TPresetColor.Lime: return "lime";
                case TPresetColor.LimeGreen: return "limeGreen";
                case TPresetColor.Linen: return "linen";
                case TPresetColor.Magenta: return "magenta";
                case TPresetColor.Maroon: return "maroon";
                case TPresetColor.MedAquamarine: return "medAquamarine";
                case TPresetColor.MedBlue: return "medBlue";
                case TPresetColor.MedOrchid: return "medOrchid";
                case TPresetColor.MedPurple: return "medPurple";
                case TPresetColor.MedSeaGreen: return "medSeaGreen";
                case TPresetColor.MedSlateBlue: return "medSlateBlue";
                case TPresetColor.MedSpringGreen: return "medSpringGreen";
                case TPresetColor.MedTurquoise: return "medTurquoise";
                case TPresetColor.MedVioletRed: return "medVioletRed";
                case TPresetColor.MidnightBlue: return "midnightBlue";
                case TPresetColor.MintCream: return "mintCream";
                case TPresetColor.MistyRose: return "mistyRose";
                case TPresetColor.Moccasin: return "moccasin";
                case TPresetColor.NavajoWhite: return "navajoWhite";
                case TPresetColor.Navy: return "navy";
                case TPresetColor.OldLace: return "oldLace";
                case TPresetColor.Olive: return "olive";
                case TPresetColor.OliveDrab: return "oliveDrab";
                case TPresetColor.Orange: return "orange";
                case TPresetColor.OrangeRed: return "orangeRed";
                case TPresetColor.Orchid: return "orchid";
                case TPresetColor.PaleGoldenrod: return "paleGoldenrod";
                case TPresetColor.PaleGreen: return "paleGreen";
                case TPresetColor.PaleTurquoise: return "paleTurquoise";
                case TPresetColor.PaleVioletRed: return "paleVioletRed";
                case TPresetColor.PapayaWhip: return "papayaWhip";
                case TPresetColor.PeachPuff: return "peachPuff";
                case TPresetColor.Peru: return "peru";
                case TPresetColor.Pink: return "pink";
                case TPresetColor.Plum: return "plum";
                case TPresetColor.PowderBlue: return "powderBlue";
                case TPresetColor.Purple: return "purple";
                case TPresetColor.Red: return "red";
                case TPresetColor.RosyBrown: return "rosyBrown";
                case TPresetColor.RoyalBlue: return "royalBlue";
                case TPresetColor.SaddleBrown: return "saddleBrown";
                case TPresetColor.Salmon: return "salmon";
                case TPresetColor.SandyBrown: return "sandyBrown";
                case TPresetColor.SeaGreen: return "seaGreen";
                case TPresetColor.SeaShell: return "seaShell";
                case TPresetColor.Sienna: return "sienna";
                case TPresetColor.Silver: return "silver";
                case TPresetColor.SkyBlue: return "skyBlue";
                case TPresetColor.SlateBlue: return "slateBlue";
                case TPresetColor.SlateGray: return "slateGray";
                case TPresetColor.Snow: return "snow";
                case TPresetColor.SpringGreen: return "springGreen";
                case TPresetColor.SteelBlue: return "steelBlue";
                case TPresetColor.Tan: return "tan";
                case TPresetColor.Teal: return "teal";
                case TPresetColor.Thistle: return "thistle";
                case TPresetColor.Tomato: return "tomato";
                case TPresetColor.Turquoise: return "turquoise";
                case TPresetColor.Violet: return "violet";
                case TPresetColor.Wheat: return "wheat";
                case TPresetColor.White: return "white";
                case TPresetColor.WhiteSmoke: return "whiteSmoke";
                case TPresetColor.Yellow: return "yellow";
                case TPresetColor.YellowGreen: return "yellowGreen";
            }
            #endregion

            return String.Empty;
        }
        #endregion

        #region Font Scheme
        private void WriteFontScheme()
        {
            DataStream.WriteStartElement("fontScheme");
            DataStream.WriteAtt("name", Globals.Theme.Elements.ColorScheme.Name);

            WriteFontDef("majorFont", Globals.Theme.Elements.FontScheme.MajorFont);
            WriteFontDef("minorFont", Globals.Theme.Elements.FontScheme.MinorFont);

            DataStream.WriteFutureStorage(Globals.ThemeRecord.FontFutureStorage);

            DataStream.WriteEndElement();
        }

        private void WriteFontDef(string tagName, TThemeFont ThemeFont)
        {
            DataStream.WriteStartElement(tagName);
            WriteLatin(ThemeFont.Latin);
            WriteEastAsian(ThemeFont.EastAsian);
            WriteComplexScript(ThemeFont.ComplexScript);
            WriteFont(ThemeFont);
            
            WriteFontDef_ExtLst(ThemeFont);

            DataStream.WriteEndElement();
        }

        private void WriteLatin(TThemeTextFont ThemeFont)
        {
            DataStream.WriteStartElement("latin", false);
            WriteThemeTextFont(ThemeFont);
            DataStream.WriteEndElement();
        }

        private void WriteEastAsian(TThemeTextFont ThemeFont)
        {
            DataStream.WriteStartElement("ea", false);
            WriteThemeTextFont(ThemeFont);
            DataStream.WriteEndElement();
        }

        private void WriteComplexScript(TThemeTextFont ThemeFont)
        {
            DataStream.WriteStartElement("cs", false);
            WriteThemeTextFont(ThemeFont);
            DataStream.WriteEndElement();
        }

        private void WriteSymbol(TThemeTextFont ThemeFont)
        {
            DataStream.WriteStartElement("sym", false);
            WriteThemeTextFont(ThemeFont);
            DataStream.WriteEndElement();
        }

        private void WriteThemeTextFont(TThemeTextFont ThemeTextFont)
        {
            DataStream.WriteAtt("typeface", ThemeTextFont.Typeface, false);
            DataStream.WriteAtt("panose", ThemeTextFont.Panose);
            if (ThemeTextFont.Pitch != TPitchFamily.DEFAULT_PITCH__UNKNOWN_FONT_FAMILY) DataStream.WriteAtt("pitchFamily", (int)ThemeTextFont.Pitch);
            if (ThemeTextFont.CharSet != TFontCharSet.Default) DataStream.WriteAtt("charset", (int)ThemeTextFont.CharSet);
        }

        private void WriteFont(TThemeFont ThemeFont)
        {
            if (ThemeFont == null) return;
            string[] scripts = ThemeFont.GetFontScripts();
            if (scripts == null || scripts.Length == 0) return;

            foreach (string script in scripts)
            {
                DataStream.WriteStartElement("font");
                DataStream.WriteAtt("script", script);
                DataStream.WriteAtt("typeface", ThemeFont.GetFont(script));
                DataStream.WriteEndElement();
            }
        }

        private void WriteFontDef_ExtLst(TThemeFont ThemeFont)
        {
            if (ThemeFont == Globals.Theme.Elements.FontScheme.MajorFont)
            {
                DataStream.WriteFutureStorage(Globals.ThemeRecord.MajorFontFutureStorage);
                return;
            }

            if (ThemeFont == Globals.Theme.Elements.FontScheme.MinorFont)
            {
                DataStream.WriteFutureStorage(Globals.ThemeRecord.MinorFontFutureStorage);
                return;
            }

            XlsMessages.ThrowException(XlsErr.ErrInternal);
        }
        #endregion

        #region Format Scheme
        private void WriteFmtScheme()
        {
            DataStream.WriteStartElement("fmtScheme");
            DataStream.WriteAtt("name", Globals.Theme.Elements.FormatScheme.Name);
            WriteFillStyleLst();
            WriteLnStyleLst();
            WriteEffectStyleLst();
            WriteBgFillStyleLst();
            DataStream.WriteEndElement();
        }

        private void WriteFillStyleLst()
        {
            DataStream.WriteStartElement("fillStyleLst", false);
            WriteGenericFillStyle(Globals.Theme.Elements.FormatScheme.FillStyleList, false);
            DataStream.WriteEndElement();
        }

        private void WriteGenericFillStyle(TFillStyleList FillStyleList, bool IsBkFill)
        {
            int aCount = Math.Max(3, FillStyleList.Count);

            for (int i = 0; i < aCount; i++)
            {
                TFillStyle fs = i < FillStyleList.Count ? FillStyleList[(TFormattingType)i] : null;

                if (fs == null)
                {
                    if (IsBkFill) fs = TFillStyleList.GetDefaultBkFillStyle(i); else fs = TFillStyleList.GetDefaultFillStyle(i);
                }

                WriteFill(fs, true);
            }

        }

        private void WriteFill(TFillStyle fs, bool HasFill)
        {
            if (fs == null) return;
            TFillStyleType FillStyleType = fs.FillStyleType;
            if (!HasFill) FillStyleType = TFillStyleType.NoFill;
            switch (FillStyleType)
            {
                case TFillStyleType.NoFill:
                    DataStream.WriteStartElement("noFill", false);
                    break;

                case TFillStyleType.Solid:
                    DataStream.WriteStartElement("solidFill", false);
                    WriteColorDef(((TSolidFill)fs).Color);
                    break;

                case TFillStyleType.Gradient:
                    DataStream.WriteStartElement("gradFill", false);
                    WriteGradientFill((TGradientFill)fs);
                    break;

                case TFillStyleType.Blip:
                    DataStream.WriteStartElement("blipFill", false);
                    WriteBlipFill((TBlipFill)fs);

                    break;

                case TFillStyleType.Pattern:
                    DataStream.WriteStartElement("pattFill", false);
                    WritePatternFill((TPatternFill)fs);
                    break;

                case TFillStyleType.Group:
                    DataStream.WriteStartElement("grpFill", false);
                    break;

                default: FlxMessages.ThrowException(FlxErr.ErrInternal);
                    break;

            }

            DataStream.WriteEndElement();
        }

        private void WritePatternFill(TPatternFill pf)
        {
            DataStream.WriteAtt("prst", GetDrawingPattern(pf.Pattern));
            DataStream.WriteStartElement("fgClr");
            WriteColorDef(pf.FgColor);
            DataStream.WriteEndElement();
            DataStream.WriteStartElement("bgClr");
            WriteColorDef(pf.BgColor);
            DataStream.WriteEndElement();
        }

        private string GetDrawingPattern(TDrawingPattern pattern)
        {
            switch (pattern)
            {
                case TDrawingPattern.pct5: return "pct5";
                case TDrawingPattern.pct10: return "pct10";
                case TDrawingPattern.pct20: return "pct20";
                case TDrawingPattern.pct25: return "pct25";
                case TDrawingPattern.pct30: return "pct30";
                case TDrawingPattern.pct40: return "pct40";
                case TDrawingPattern.pct50: return "pct50";
                case TDrawingPattern.pct60: return "pct60";
                case TDrawingPattern.pct70: return "pct70";
                case TDrawingPattern.pct75: return "pct75";
                case TDrawingPattern.pct80: return "pct80";
                case TDrawingPattern.pct90: return "pct90";
                case TDrawingPattern.horz: return "horz";
                case TDrawingPattern.vert: return "vert";
                case TDrawingPattern.ltHorz: return "ltHorz";
                case TDrawingPattern.ltVert: return "ltVert";
                case TDrawingPattern.dkHorz: return "dkHorz";
                case TDrawingPattern.dkVert: return "dkVert";
                case TDrawingPattern.narHorz: return "narHorz";
                case TDrawingPattern.narVert: return "narVert";
                case TDrawingPattern.dashHorz: return "dashHorz";
                case TDrawingPattern.dashVert: return "dashVert";
                case TDrawingPattern.cross: return "cross";
                case TDrawingPattern.dnDiag: return "dnDiag";
                case TDrawingPattern.upDiag: return "upDiag";
                case TDrawingPattern.ltDnDiag: return "ltDnDiag";
                case TDrawingPattern.ltUpDiag: return "ltUpDiag";
                case TDrawingPattern.dkDnDiag: return "dkDnDiag";
                case TDrawingPattern.dkUpDiag: return "dkUpDiag";
                case TDrawingPattern.wdDnDiag: return "wdDnDiag";
                case TDrawingPattern.wdUpDiag: return "wdUpDiag";
                case TDrawingPattern.dashDnDiag: return "dashDnDiag";
                case TDrawingPattern.dashUpDiag: return "dashUpDiag";
                case TDrawingPattern.diagCross: return "diagCross";
                case TDrawingPattern.smCheck: return "smCheck";
                case TDrawingPattern.lgCheck: return "lgCheck";
                case TDrawingPattern.smGrid: return "smGrid";
                case TDrawingPattern.lgGrid: return "lgGrid";
                case TDrawingPattern.dotGrid: return "dotGrid";
                case TDrawingPattern.smConfetti: return "smConfetti";
                case TDrawingPattern.lgConfetti: return "lgConfetti";
                case TDrawingPattern.horzBrick: return "horzBrick";
                case TDrawingPattern.diagBrick: return "diagBrick";
                case TDrawingPattern.solidDmnd: return "solidDmnd";
                case TDrawingPattern.openDmnd: return "openDmnd";
                case TDrawingPattern.dotDmnd: return "dotDmnd";
                case TDrawingPattern.plaid: return "plaid";
                case TDrawingPattern.sphere: return "sphere";
                case TDrawingPattern.weave: return "weave";
                case TDrawingPattern.divot: return "divot";
                case TDrawingPattern.shingle: return "shingle";
                case TDrawingPattern.wave: return "wave";
                case TDrawingPattern.trellis: return "trellis";
                case TDrawingPattern.zigZag: return "zigZag";
                default:
                    return "cross";
            }
        }

        private void WriteGradientFill(TGradientFill gf)
        {
            if (gf.Flip != TFlipMode.None) DataStream.WriteAtt("flip", GetFlip(gf.Flip));
            DataStream.WriteAtt("rotWithShape", gf.RotateWithShape);

            WriteGsLst(gf);

            TDrawingLinearGradient lin = gf.GradientDef as TDrawingLinearGradient;
            if (lin != null)
            {
                WriteLinGrad(lin);
            }
            else
            {
                TDrawingPathGradient pat = gf.GradientDef as TDrawingPathGradient;
                if (pat != null)
                {
                    WritePathGrad(pat);
                }
            }

            if (gf.TileRect.HasValue) WriteRelativeRect("tileRect", gf.TileRect.Value);
        }

        private void WriteGsLst(TGradientFill gf)
        {
            DataStream.WriteStartElement("gsLst");
            foreach (TDrawingGradientStop stop in gf.GradientStops)
            {
                DataStream.WriteStartElement("gs");
                DataStream.WriteAttPercent("pos", stop.Position);
                WriteColorDef(stop.Color);

                DataStream.WriteEndElement();
            }

            DataStream.WriteEndElement();
        }

        private void WriteLinGrad(TDrawingLinearGradient lin)
        {
            DataStream.WriteStartElement("lin");
            DataStream.WriteAttAsAngle("ang", lin.Angle);
            DataStream.WriteAtt("scaled", lin.Scaled);
            DataStream.WriteEndElement();
        }

        private void WritePathGrad(TDrawingPathGradient pat)
        {
            DataStream.WriteStartElement("path");
            DataStream.WriteAtt("path", GetPathShadeType(pat.Path));
            if (pat.FillToRect.HasValue) WriteRelativeRect("fillToRect", pat.FillToRect.Value);
            DataStream.WriteEndElement();            
        }

        private string GetPathShadeType(TPathShadeType ShadeType)
        {
            switch (ShadeType)
            {
                case TPathShadeType.Shape: return "shape";
                case TPathShadeType.Circle: return "circle";
                case TPathShadeType.Rect: return "rect";
                default: return "shape";
            }
        }

        private void WriteBlipFill(TBlipFill bf)
        {
            if (bf.Dpi != 0) DataStream.WriteAtt("dpi", bf.Dpi);
            DataStream.WriteAtt("rotWithShape", bf.RotateWithShape);

            if (bf.Blip != null) WriteBlip(bf.Blip);
            if (bf.SourceRect.HasValue) WriteSrcRect(bf.SourceRect.Value);

            if (bf.FillMode != null)
            {
                TBlipFillTile bft = bf.FillMode as TBlipFillTile;
                if (bft != null)
                {
                    WriteTile(bft);
                }
                else
                {
                    TBlipFillStretch bfs = bf.FillMode as TBlipFillStretch;
                    if (bfs != null)
                    {
                        WriteStretch(bfs);
                    }
                }

            }

        }

        private void WriteBlip(TBlip blip)
        {
            DataStream.WriteStartElement("blip");
            DataStream.WriteAtt("cstate", GetCompressionState(blip.CompressionState));
            if (blip.PictureData != null) WriteBlipData(DataStream, blip.ImageFileName, blip.ContentType, blip.PictureData, true);

            if (blip.Transforms != null)
            {
                foreach (string s in blip.Transforms)
                {
                    DataStream.WriteRaw(s);
                }
            }
            DataStream.WriteEndElement();
        }

        public void WriteBlipData(TOpenXmlWriter DataStream, string ImageName, string ContentType, byte[] PicData, bool WriteRel)
        {
            CurrentMediaRelId++;
            
            //if (string.IsNullOrEmpty(ImageName) || ExistingImages.ContainsKey(ImageName))
            ImageName = "image" + (ExistingImages.Count + 1).ToString(CultureInfo.InvariantCulture) + GetDefaultExt(ContentType);            
            ExistingImages.Add(ImageName, ImageName);
            
            if (WriteRel)
            {
                DataStream.WriteRelationship("embed", CurrentMediaRelId);
            }
            Uri PicUri = TOpenXmlManager.ResolvePartUri(DataStream.CurrentUri, new Uri("../media/" + Path.GetFileName(ImageName), UriKind.Relative));
            DataStream.WritePart(PicUri, ContentType, PicData);
            DataStream.CreateRelationshipToUri(PicUri, TargetMode.Internal, TOpenXmlManager.imageRelationshipType,
                TOpenXmlWriter.FlexCelRid + CurrentMediaRelId.ToString(CultureInfo.InvariantCulture));
        }

        private static string GetDefaultExt(string ContentType)
        {
            switch (ContentType)
            {
                case "image/gif": return ".gif";
                case "image/png": return ".png";
                case "image/tiff": return ".tiff";
                case "image/jpeg": return ".jpeg";
                case "image/pict": return ".pic";
                case "image/x-wmf": return ".wmf";
                case "image/x-emf": return ".emf";
                case "image/bmp": return ".bmp";
            }
            return String.Empty;
        } 

        private string GetCompressionState(TBlipCompression bc)
        {
            switch (bc)
            {
                case TBlipCompression.Email: return "email";
                case TBlipCompression.Screen: return "screen";
                case TBlipCompression.Print: return "print";
                case TBlipCompression.HQPrint: return "hqprint";
            }
            return null;
        }

        private void WriteRelativeRect(string TagName, TDrawingRelativeRect r)
        {
            DataStream.WriteStartElement(TagName, false);
            if (r.Left != 0) DataStream.WriteAttPercent("l", r.Left);
            if (r.Top != 0) DataStream.WriteAttPercent("t", r.Top);
            if (r.Right != 0) DataStream.WriteAttPercent("r", r.Right);
            if (r.Bottom != 0) DataStream.WriteAttPercent("b", r.Bottom);

            DataStream.WriteEndElement();
        }

        private void WriteSrcRect(TDrawingRelativeRect r)
        {
            WriteRelativeRect("srcRect", r);
        }

        private void WriteTile(TBlipFillTile bft)
        {
            DataStream.WriteStartElement("tile");
            DataStream.WriteAtt("tx", bft.Tx.Emu);
            DataStream.WriteAtt("ty", bft.Ty.Emu);
            DataStream.WriteAttPercent("sx", bft.ScaleX);
            DataStream.WriteAttPercent("sy", bft.ScaleY);
            DataStream.WriteAtt("flip", GetFlip(bft.Flip));
            DataStream.WriteAtt("algn", GetAlign(bft.Align));
            DataStream.WriteEndElement();
        }

        private string GetAlign(TDrawingRectAlign Align)
        {
            switch (Align)
            {
                case TDrawingRectAlign.TopLeft: return "tl";
                case TDrawingRectAlign.Top: return "t";
                case TDrawingRectAlign.TopRight: return "tr";
                case TDrawingRectAlign.Left: return "l";
                case TDrawingRectAlign.Center: return "ctr";
                case TDrawingRectAlign.Right: return "r";
                case TDrawingRectAlign.BottomLeft: return "bl";
                case TDrawingRectAlign.Bottom: return "b";
                case TDrawingRectAlign.BottomRight: return "br";
                default:
                    return "tl";
            }
        }

        private string GetFlip(TFlipMode Flip)
        {
            switch (Flip)
            {
                case TFlipMode.None: return "none";
                case TFlipMode.X: return "x";
                case TFlipMode.Y: return "y";
                case TFlipMode.XY: return "xy";
                default:
                    return "none";
            }
        }

        private void WriteStretch(TBlipFillStretch bfs)
        {
            DataStream.WriteStartElement("stretch");
            WriteRelativeRect("fillRect", bfs.FillRect);
            DataStream.WriteEndElement();
        }

        private void WriteLnStyleLst()
        {
            DataStream.WriteStartElement("lnStyleLst", false);
            int RealCount = Globals.Theme.Elements.FormatScheme.LineStyleList.Count;
            int Count = Math.Max(3, RealCount);
            for (int i = 0; i < Count; i++)
            {
                TLineStyle LineStyle = i < RealCount ? Globals.Theme.Elements.FormatScheme.LineStyleList[(TFormattingType)i] : TLineStyleList.GetDefaultLineStyle(i);
                WriteLn(LineStyle);
            }
            DataStream.WriteEndElement();
        }

        private void WriteLn(TLineStyle LineStyle)
        {
            if (LineStyle == null) return;
            DataStream.WriteStartElement("ln");
            if (LineStyle.Width.HasValue) DataStream.WriteAtt("w", LineStyle.Width.Value);
            if (LineStyle.LineCap.HasValue) DataStream.WriteAtt("cap", GetLineCap(LineStyle.LineCap.Value));
            if (LineStyle.CompoundLineType.HasValue) DataStream.WriteAtt("cmpd", GetCompoundLineStyle(LineStyle.CompoundLineType.Value));
            if (LineStyle.PenAlign.HasValue) DataStream.WriteAtt("algn", GetPenAlign(LineStyle.PenAlign.Value));

            WriteFill(LineStyle.Fill, true);

            if (LineStyle.Dashing.HasValue)
            {
                DataStream.WriteStartElement("prstDash");
                DataStream.WriteAtt("val", GetDashing(LineStyle.Dashing));
                DataStream.WriteEndElement();
            }

            if (LineStyle.Join.HasValue)
            {
                switch (LineStyle.Join.Value)
                {
                    case TLineJoin.Bevel:
                        DataStream.WriteStartElement("bevel", false);
                        DataStream.WriteEndElement();
                        break;

                    case TLineJoin.Miter:
                        DataStream.WriteStartElement("miter", false);
                        DataStream.WriteEndElement();
                        break;

                    case TLineJoin.Round:
                        DataStream.WriteStartElement("round", false);
                        DataStream.WriteEndElement();
                        break;
                }
            }

            if (LineStyle.HeadArrow.HasValue) WriteArrow("headEnd", LineStyle.HeadArrow.Value);
            if (LineStyle.TailArrow.HasValue) WriteArrow("tailEnd", LineStyle.TailArrow.Value);

            if (LineStyle.FExtra != null)
            {
                foreach (string s in LineStyle.FExtra)
                {
                    DataStream.WriteRaw(s);
                }
            }
            DataStream.WriteEndElement();
        }

        private void WriteArrow(string elementName, TLineArrow arrow)
        {
            if (arrow.Style == TArrowStyle.None) return;
            DataStream.WriteStartElement(elementName);
            DataStream.WriteAtt("type", GetArrowType(arrow.Style));
            DataStream.WriteAtt("w", GetArrowWidth(arrow.Width));
            DataStream.WriteAtt("len", GetArrowLen(arrow.Len));
            DataStream.WriteEndElement();
        }

        private string GetArrowType(TArrowStyle aArrowStyle)
        {
            switch (aArrowStyle)
            {
                case TArrowStyle.Normal: return "triangle";
                case TArrowStyle.Stealth: return "stealth";
                case TArrowStyle.Diamond: return "diamond";
                case TArrowStyle.Oval: return "oval";
                case TArrowStyle.Open: return "arrow";
                
                case TArrowStyle.None: 
                default:
                return "none";

            }
        }

        private string GetArrowWidth(TArrowWidth aArrowWidth)
        {
            switch (aArrowWidth)
            {
                case TArrowWidth.Large: return "lg";
                case TArrowWidth.Small: return "sm";
                
                case TArrowWidth.Medium:
                default: return "med";
            }
        }
        private string GetArrowLen(TArrowLen aArrowLen)
        {
            switch (aArrowLen)
            {
                case TArrowLen.Large: return "lg";
                case TArrowLen.Small: return "sm";
                
                case TArrowLen.Medium:
                default: return "med";
            }
        }

        private string GetDashing(TLineDashing? ld)
        {
            if (ld.HasValue)
            {
                switch (ld.Value)
                {
                    case TLineDashing.Solid: return "solid";
                    case TLineDashing.DotGEL: return "dot";
                    case TLineDashing.DashGEL: return "dash";
                    case TLineDashing.LongDashGEL: return "lgDash";
                    case TLineDashing.DashDotGEL: return "dashDot";
                    case TLineDashing.LongDashDotGEL: return "lgDashDot";
                    case TLineDashing.LongDashDotDotGEL: return "lgDashDotDot";
                    case TLineDashing.DashSys: return "sysDash";
                    case TLineDashing.DotSys: return "sysDot";
                    case TLineDashing.DashDotSys: return "sysDashDot";
                    case TLineDashing.DashDotDotSys: return "sysDashDotDot";
                }
            }
            return "solid";
        }

        private string GetLineCap(TLineCap lc)
        {
            switch (lc)
            {
                case TLineCap.Round: return "rnd";
                case TLineCap.Square: return "sq";
                case TLineCap.Flat: return "flat";
                default:
                    return "flat";
            }
        }

        private string GetCompoundLineStyle(TCompoundLineType clt)
        {
            switch (clt)
            {
                case TCompoundLineType.Single: return "sng";
                case TCompoundLineType.Double: return "dbl";
                case TCompoundLineType.ThickThin: return "thickThin";
                case TCompoundLineType.ThinThick: return "thinThick";
                case TCompoundLineType.Triple: return "tri";
                default:
                    return "sng";
            }
        }

        private string GetPenAlign(TPenAlignment pa)
        {
            switch (pa)
            {
                case TPenAlignment.Center: return "ctr";
                case TPenAlignment.Inset: return "in";
                default:
                    return "ctr";
            }
        }

        private void WriteEffectStyleLst()
        {
            //DataStream.WriteStartElement("effectStyleLst", false);
            string s = Globals.Theme.Elements.FormatScheme.EffectStyleList.Xml;
            if (s != null) DataStream.WriteRaw(s);
            else DataStream.WriteRaw( //Not yet parsed.
@"			<a:effectStyleLst>
				<a:effectStyle>
					<a:effectLst>
						<a:outerShdw blurRad=""40000"" dist=""20000"" dir=""5400000"" rotWithShape=""0"">
							<a:srgbClr val=""000000"">
								<a:alpha val=""38000"" />
							</a:srgbClr>
						</a:outerShdw>
					</a:effectLst>
				</a:effectStyle>
				<a:effectStyle>
					<a:effectLst>
						<a:outerShdw blurRad=""40000"" dist=""23000"" dir=""5400000"" rotWithShape=""0"">
							<a:srgbClr val=""000000"">
								<a:alpha val=""35000"" />
							</a:srgbClr>
						</a:outerShdw>
					</a:effectLst>
				</a:effectStyle>
				<a:effectStyle>
					<a:effectLst>
						<a:outerShdw blurRad=""40000"" dist=""23000"" dir=""5400000"" rotWithShape=""0"">
							<a:srgbClr val=""000000"">
								<a:alpha val=""35000"" />
							</a:srgbClr>
						</a:outerShdw>
					</a:effectLst>
					<a:scene3d>
						<a:camera prst=""orthographicFront"">
							<a:rot lat=""0"" lon=""0"" rev=""0"" />
						</a:camera>
						<a:lightRig rig=""threePt"" dir=""t"">
							<a:rot lat=""0"" lon=""0"" rev=""1200000"" />
						</a:lightRig>
					</a:scene3d>
					<a:sp3d>
						<a:bevelT w=""63500"" h=""25400"" />
					</a:sp3d>
				</a:effectStyle>
			</a:effectStyleLst>
"
                );
            //DataStream.WriteEndElement();
        }

        private void WriteBgFillStyleLst()
        {
            DataStream.WriteStartElement("bgFillStyleLst", false);
            WriteGenericFillStyle(Globals.Theme.Elements.FormatScheme.BkFillStyleList, true);
            DataStream.WriteEndElement();
        }
        #endregion

        private void WriteThemeElements_ExtLst()
        {
            DataStream.WriteFutureStorage(Globals.ThemeRecord.ElementsFutureStorage);
        }

        private void WriteObjectDefaults()
        {
            DataStream.WriteStartElement("objectDefaults");
            DataStream.WriteEndElement();
        }

        private void WriteExtraClrSchemeLst()
        {
            DataStream.WriteStartElement("extraClrSchemeLst");
            DataStream.WriteEndElement();
        }

        private void WriteCustClrLst()
        {
            DataStream.WriteStartElement("custClrLst");
            DataStream.WriteEndElement();
        }

        private void WriteTheme_ExtLst()
        {
            DataStream.WriteFutureStorage(Globals.ThemeRecord.FutureStorage);
        }
        #endregion
    }
}
