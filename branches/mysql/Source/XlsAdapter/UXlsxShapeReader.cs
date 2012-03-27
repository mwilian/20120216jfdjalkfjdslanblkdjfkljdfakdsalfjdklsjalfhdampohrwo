using System;
using System.Collections.Generic;
using FlexCel.Core;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.IO.Compression;
using System.Globalization;

namespace FlexCel.XlsAdapter
{
    static class TShapePresets
    {
        internal static Dictionary<string, TShapeGeom> ShapeList = ReadShapes();

        private static Dictionary<string, TShapeGeom> ReadShapes()
        {
            Dictionary<string, TShapeGeom> Result = new Dictionary<string, TShapeGeom>();
            using (Stream PresetStream = GetPresetStream())
            {
                using (TOpenXmlReader DataStream = TOpenXmlReader.CreateFromSimpleStream(PresetStream))
                {
                    DataStream.DefaultNamespace = "";
                    DataStream.NextTag();
                    Debug.Assert(DataStream.RecordName() == "presetShapeDefinitons", DataStream.RecordName());

                    string StartElement = DataStream.RecordName();
                    if (!DataStream.NextTag()) FlxMessages.ThrowException(FlxErr.ErrInternal);

                    TXlsxShapeReader ShapeReader = new TXlsxShapeReader(DataStream);
                    while (!DataStream.AtEndElement(StartElement))
                    {
                        Result[DataStream.RecordName()] = ShapeReader.ReadShapeDef(DataStream.RecordName());
                    }
                }
            }

            return Result;
        }

        private static Stream GetPresetStream()
        {
            Stream Gziped = Assembly.GetExecutingAssembly().GetManifestResourceStream("FlexCel.XlsAdapter.PresetShapes.xml.gz");
            return new GZipStream(Gziped, CompressionMode.Decompress, false);
        }


        internal static bool GetBiff8Adjust(string AdjName, out TShapeOption so)
        {
            so = TShapeOption.None;
            if (!AdjName.StartsWith("adj", StringComparison.InvariantCultureIgnoreCase)) return false;
            if (AdjName.Length == 3)
            {
                so = TShapeOption.adjustValue;
                return true;
            }

            string AdjNum = AdjName.Substring(3);
            int n;
            if (!int.TryParse(AdjNum, NumberStyles.Integer, CultureInfo.InvariantCulture, out n)) return false;
            if (n < 1 || n > 10) return false;

            so = TShapeOption.adjustValue + n - 1;
            return true;
        }

        internal static int ConvertAdjustFromBiff8(int value, string sp)
        {
            return Convert.ToInt32(value / .216);
        }

        internal static long ConvertAdjustToBiff8(TShapeType st, double vald, out long def)
        {
            def = 10800;
            long Result = Convert.ToInt64(Math.Round(vald * .216));
            if (Result > 21600) Result = 21600;

            switch (st)
            {
                case TShapeType.RoundRectangle: def = 0xd60; break;

                case TShapeType.Parallelogram:
                case TShapeType.Trapezoid:
                case TShapeType.Trapezoid2007:
                case TShapeType.Hexagon:
                case TShapeType.Plus:
                case TShapeType.Can:
                case TShapeType.FlowChartMagneticDisk:
                    def = 21600 / 4; break;

                case TShapeType.FlowChartInputOutput:
                case TShapeType.FlowChartManualOperation:
                case TShapeType.FlowChartPreparation:
                    def = 21600 / 5; break;

                case TShapeType.Octagon:
                    def = 21600 * 2 / 7; break;

                case TShapeType.SmileyFace:
                    def = 17400;
                    Result += 16515;
                    break;
            }

            return Result;
        }
        
    }

    class TXlsxShapeReader
    {
        private TOpenXmlReader DataStream;

        public TXlsxShapeReader(TOpenXmlReader aDataStream)
        {
            DataStream = aDataStream;
        }

        internal TShapeGeom ReadShapeDef(string GeomName)
        {
            TShapeGeom Result = new TShapeGeom(GeomName);

            string DefaultNamespace = DataStream.DefaultNamespace;
            DataStream.DefaultNamespace = "http://schemas.openxmlformats.org/drawingml/2006/main";
            try
            {
                string StartElement = DataStream.RecordName();
                if (DataStream.IsSimpleTag) { DataStream.NextTag(); return Result; }
                if (!DataStream.NextTag()) return Result;

                while (!DataStream.AtEndElement(StartElement))
                {
                    switch (DataStream.RecordName())
                    {
                        case "avLst": ReadGdLst(Result.AvList, Result); break;
                        case "gdLst": ReadGdLst(Result.GdList, Result); break;
                        case "ahLst": ReadAhLst(Result.AhList, Result); break;
                        case "cxnLst": ReadCxnLst(Result.ConnList, Result); break;
                        case "rect": Result.TextRect = ReadRect(Result); break;
                        case "pathLst": ReadPathLst(Result); break;

                        default: DataStream.GetXml(); break;
                    }
                }
            }
            finally
            {
                DataStream.DefaultNamespace = DefaultNamespace;
            }

            return Result;
        }

        private void ReadGdLst(TShapeGuideList GuideList, TShapeGeom ShapeGeom)
        {
            if (DataStream.IsSimpleTag) { DataStream.NextTag(); return; }
            string StartElement = DataStream.RecordName();
            if (!DataStream.NextTag()) return;

            while (!DataStream.AtEndElement(StartElement))
            {
                switch (DataStream.RecordName())
                {
                    case "gd":
                        GuideList.Add(ReadGuide(ShapeGeom));
                        DataStream.FinishTag();
                        break;

                    default: DataStream.GetXml(); break;
                }
            }
            
        }

        private TShapeGuide ReadGuide(TShapeGeom ShapeGeom)
        {
            return new TShapeGuide(DataStream.GetAttribute("name"), GetShapeFmla(DataStream.GetAttribute("fmla").Trim(), ShapeGeom));
        }

        private TShapeFormula GetShapeFmla(string fmla, TShapeGeom ShapeGeom)
        {
            string[] Args = fmla.Split((char[])null, StringSplitOptions.RemoveEmptyEntries);
            switch (Args[0])
            {
                case "*/":
                    return Get3Args(Args, new TShapeMulDiv(), ShapeGeom);
                case "+-":
                    return Get3Args(Args, new TShapeAddSub(), ShapeGeom);
                case "+/":
                    return Get3Args(Args, new TShapeAddDiv(), ShapeGeom);
                case "?:":
                    return Get3Args(Args, new TShapeIfElse(), ShapeGeom);
                case "abs":
                    return Get1Arg(Args, new TShapeAbs(), ShapeGeom);
                case "at2":
                    return Get2Args(Args, new TShapeArcTan(), ShapeGeom);
                case "cat2":
                    return Get3Args(Args, new TShapeCosArcTan(), ShapeGeom);
                case "cos":
                    return Get2Args(Args, new TShapeCos(), ShapeGeom);
                case "max":
                    return Get2Args(Args, new TShapeMax(), ShapeGeom);
                case "min":
                    return Get2Args(Args, new TShapeMin(), ShapeGeom);
                case "mod":
                    return Get3Args(Args, new TShapeMod(), ShapeGeom);
                case "pin":
                    return Get3Args(Args, new TShapePin(), ShapeGeom);
                case "sat2":
                    return Get3Args(Args, new TShapeSinArcTan(), ShapeGeom);
                case "sin":
                    return Get2Args(Args, new TShapeSin(), ShapeGeom);
                case "sqrt":
                    return Get1Arg(Args, new TShapeSqrt(), ShapeGeom);
                case "tan":
                    return Get2Args(Args, new TShapeTan(), ShapeGeom);
                case "val":

                    return GetGuide(Args[1], ShapeGeom).Fmla;
            }

            return new TShapeUndefFormula(fmla);
        }

        private TShapeFormula Get1Arg(string[] Args, T1ArgShapeFormula aShape, TShapeGeom ShapeGeom)
        {
            aShape.x = GetGuide(Args[1], ShapeGeom);
            return aShape;
        }

        private TShapeFormula Get2Args(string[] Args, T2ArgShapeFormula aShape, TShapeGeom ShapeGeom)
        {
            Get1Arg(Args, aShape, ShapeGeom);
            aShape.y = GetGuide(Args[2], ShapeGeom);
            return aShape;
        }

        private TShapeFormula Get3Args(string[] Args, T3ArgShapeFormula aShape, TShapeGeom ShapeGeom)
        {
            Get2Args(Args, aShape, ShapeGeom);
            aShape.z = GetGuide(Args[3], ShapeGeom);
            return aShape;
        }

        private void ReadAhLst(TShapeAdjustHandleList ShapeAdjustHandleList, TShapeGeom ShapeGeom)
        {
            if (DataStream.IsSimpleTag) { DataStream.NextTag(); return; }
            string StartElement = DataStream.RecordName();
            if (!DataStream.NextTag()) return;

            while (!DataStream.AtEndElement(StartElement))
            {
                switch (DataStream.RecordName())
                {
                    case "ahXY":
                       ShapeAdjustHandleList.Add(ReadAhXY(ShapeGeom));
                        break;

                    case "ahPolar":
                       ShapeAdjustHandleList.Add(ReadAhPolar(ShapeGeom));
                        break;
                    default: DataStream.GetXml(); break;
                }
            }
        }

        private TShapeAdjustHandle ReadAhXY(TShapeGeom ShapeGeom)
        {
            TShapeAdjustHandleXY ah = new TShapeAdjustHandleXY();
            ah.GdRefX = GetGuideFromAttAllowNull("gdRefX", ShapeGeom);
            ah.MinX = GetGuideFromAttAllowNull("minX", ShapeGeom);
            ah.MaxX = GetGuideFromAttAllowNull("maxX", ShapeGeom);
            ah.GdRefY = GetGuideFromAttAllowNull("gdRefY", ShapeGeom);
            ah.MinY = GetGuideFromAttAllowNull("minY", ShapeGeom);
            ah.MaxY = GetGuideFromAttAllowNull("maxY", ShapeGeom);
            ah.Location = ReadPoint(ShapeGeom, "pos");
            return ah;
        }

        private TShapeAdjustHandle ReadAhPolar(TShapeGeom ShapeGeom)
        {
            TShapeAdjustHandlePolar ah = new TShapeAdjustHandlePolar();
            ah.GdRefR = GetGuideFromAttAllowNull("gdRefR", ShapeGeom);
            ah.MinR = GetGuideFromAttAllowNull("minR", ShapeGeom);
            ah.MaxR = GetGuideFromAttAllowNull("maxR", ShapeGeom);
             ah.GdRefAng = GetGuideFromAttAllowNull("gdRefAng", ShapeGeom);
            ah.MinAng = GetGuideFromAttAllowNull("minAng", ShapeGeom);
            ah.MaxAng = GetGuideFromAttAllowNull("maxAng", ShapeGeom);
            ah.Location = ReadPoint(ShapeGeom, "pos");
            return ah;

        }

        private void ReadCxnLst(TShapeConnectionList ShapeConnectionList, TShapeGeom ShapeGeom)
        {
            if (DataStream.IsSimpleTag) { DataStream.NextTag(); return; }
            string StartElement = DataStream.RecordName();
            if (!DataStream.NextTag()) return;

            while (!DataStream.AtEndElement(StartElement))
            {
                switch (DataStream.RecordName())
                {
                    case "cxn":
                        ShapeConnectionList.Add(ReadCxn(ShapeGeom));
                        break;
                    default: DataStream.GetXml(); break;
                }
            }
        }
        private TShapeConnection ReadCxn(TShapeGeom ShapeGeom)
        {
            TShapeConnection cxn = new TShapeConnection();
            cxn.Angle = GetGuideFromAtt("ang", ShapeGeom);
            cxn.Position = ReadPoint(ShapeGeom, "pos");
            return cxn;
        }

        private TShapeTextRect ReadRect(TShapeGeom ShapeGeom)
        {
            TShapeTextRect ShapeTextRect = new TShapeTextRect();
            ShapeTextRect.Left = GetGuideFromAtt("l", ShapeGeom);
            ShapeTextRect.Right = GetGuideFromAtt("r", ShapeGeom);
            ShapeTextRect.Top = GetGuideFromAtt("t", ShapeGeom);
            ShapeTextRect.Bottom = GetGuideFromAtt("b", ShapeGeom);
            DataStream.FinishTag();
            return ShapeTextRect;
        }

        private void ReadPathLst(TShapeGeom ShapeGeom)
        {
            if (DataStream.IsSimpleTag) { DataStream.NextTag(); return; }
            string StartElement = DataStream.RecordName();
            if (!DataStream.NextTag()) return;

            while (!DataStream.AtEndElement(StartElement))
            {
                switch (DataStream.RecordName())
                {
                    case "path":
                        ShapeGeom.PathList.Add(ReadPath(ShapeGeom));
                        break;

                    default: DataStream.GetXml(); break;

                }
            }
        }

        private TShapePath ReadPath(TShapeGeom ShapeGeom)
        {
            TShapePath Result = new TShapePath();
            if (DataStream.IsSimpleTag) { DataStream.NextTag(); return Result; }

            Result.Width = DataStream.GetAttributeAsInt("w", 0);
            Result.Height = DataStream.GetAttributeAsInt("h", 0);
            Result.PathFill = GetPathFill(DataStream.GetAttribute("fill"));
            Result.PathStroke = DataStream.GetAttributeAsBool("stroke", true);
            Result.ExtrusionOk = DataStream.GetAttributeAsBool("extrusionOk", true);

            string StartElement = DataStream.RecordName();
            if (!DataStream.NextTag()) return Result;

            while (!DataStream.AtEndElement(StartElement))
            {
                switch (DataStream.RecordName())
                {
                    case "close":
                        Result.Actions.Add(new TShapeActionClose());
                        DataStream.FinishTag();
                        break;

                    case "moveTo":
                        Result.Actions.Add(new TShapeActionMoveTo(ReadPt(ShapeGeom)));
                        break;

                    case "lnTo":
                        Result.Actions.Add(new TShapeActionLineTo(ReadPt(ShapeGeom)));
                        break;

                    case "arcTo":
                        TShapeActionArcTo ArcTo = new TShapeActionArcTo();
                        ArcTo.WidthRadius = GetGuideFromAtt("wR", ShapeGeom);
                        ArcTo.HeightRadius = GetGuideFromAtt("hR", ShapeGeom);
                        ArcTo.StartAngle = GetGuideFromAtt("stAng", ShapeGeom);
                        ArcTo.SwingAngle = GetGuideFromAtt("swAng", ShapeGeom);
                        Result.Actions.Add(ArcTo);
                        DataStream.FinishTag();
                        break;

                    case "quadBezTo":
                        Result.Actions.Add(new TShapeActionQuadBezierTo(ReadPts(ShapeGeom, "pt").ToArray()));
                        break;

                    case "cubicBezTo":
                        Result.Actions.Add(new TShapeActionCubicBezierTo(ReadPts(ShapeGeom, "pt").ToArray()));
                        break;

                    default: DataStream.GetXml(); break;

                }
            }

            return Result;
        }

        private TShapeGuide GetGuideFromAtt(string att, TShapeGeom ShapeGeom)
        {
            string s = DataStream.GetAttribute(att);
            return GetGuide(s, ShapeGeom);
        }

        private TShapeGuide GetGuideFromAttAllowNull(string att, TShapeGeom ShapeGeom)
        {
            string s = DataStream.GetAttribute(att);
            if (s == null) return null;
            return GetGuide(s, ShapeGeom);
        }

        private static TShapeGuide GetGuide(string s, TShapeGeom ShapeGeom)
        {
            TShapeGuide Result;
            if (ShapeGeom.FindGuide(s, out Result)) return Result;
            return new TShapeGuide(null, new TShapeVal(Convert.ToInt32(s)));
        }

        private TShapePoint ReadPt(TShapeGeom ShapeGeom)
        {
            return ReadPoint(ShapeGeom, "pt");
        }

        private TShapePoint ReadPoint(TShapeGeom ShapeGeom, string TagName)
        {
            List<TShapePoint> Pts = ReadPts(ShapeGeom, TagName);
            if (Pts.Count != 1) FlxMessages.ThrowException(FlxErr.ErrInternal);
            return Pts[0];
        }

        private List<TShapePoint> ReadPts(TShapeGeom ShapeGeom, string TagName)
        {
            List<TShapePoint> Result = new List<TShapePoint>();
            if (DataStream.IsSimpleTag) { DataStream.NextTag(); return Result; }
            string StartElement = DataStream.RecordName();
            if (!DataStream.NextTag()) return Result;

            while (!DataStream.AtEndElement(StartElement))
            {
                if (DataStream.RecordName() == TagName)
                {
                    Result.Add(new TShapePoint(GetGuideFromAtt("x", ShapeGeom), GetGuideFromAtt("y", ShapeGeom)));
                    DataStream.FinishTag();
                }
                else DataStream.GetXml(); 
            }
            return Result;
        }

        private TPathFillMode GetPathFill(string p)
        {
            switch (p)
            {
                case "none": return TPathFillMode.None;
                case "norm": return TPathFillMode.Norm;
                case "lighten": return TPathFillMode.Lighten;
                case "lightenLess": return TPathFillMode.LightenLess;
                case "darken": return TPathFillMode.Darken;
                case "darkenLess": return TPathFillMode.DarkenLess;
            }
            return TPathFillMode.Norm;
        }
    }
}
