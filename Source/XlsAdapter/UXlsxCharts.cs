using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using FlexCel.Core;
using System.Xml;
using System.Globalization;

namespace FlexCel.XlsAdapter
{
    internal class TXlsxChartReader
    {
        #region Variables
        private ExcelFile xls;
        private TOpenXmlReader DataStream;
        #endregion

        #region Constructors
        internal TXlsxChartReader(TOpenXmlReader aDataStream, ExcelFile axls)
        {
            DataStream = aDataStream;
            xls = axls;
        }
        #endregion

        internal void ReadChart(string relId, TFlxChart Chart)
        {
            DataStream.SelectFromCurrentPartAndPush(relId, TOpenXmlManager.ChartNamespace, false);

            while (DataStream.NextTag())
            {
                switch (DataStream.RecordName())
                {
                    case "chartSpace":
                        ReadChartSpace(Chart);
                        break;

                    default: //shouldn't happen
                        DataStream.GetXml();
                        break;
                }
            }

            DataStream.PopPart();
        }


        public void ReadChartSpace(TFlxChart Chart)
        {
            if (DataStream.IsSimpleTag) { DataStream.NextTag(); return; }
            string StartElement = DataStream.RecordName();
            if (!DataStream.NextTag()) return;
            while (!DataStream.AtEndElement(StartElement))
            {
                switch (DataStream.RecordName())
                {
/*
                    case "date1904": ReadDate1904(Chart); break;
                    case "lang": ReadLang(Chart); break;
                    case "roundedCorners": ReadRoundedCorners(Chart); break;
                    case "style": ReadStyle(Chart); break;
                    case "clrMapOvr": ReadClrMapOvr(Chart); break;
                    case "pivotSource": ReadPivotSource(Chart); break;
                    case "protection": ReadProtection(Chart); break;*/
                    case "chart": ReadChart(Chart); break;
                   /* case "spPr": ReadSpPr(Chart); break;
                    case "txPr": ReadTxPr(Chart); break;
                    case "externalData": ReadExternalData(Chart); break;
                    case "printSettings": ReadPrintSettings(Chart); break;
                    case "userShapes": ReadUserShapes(Chart); break;
                    case "extLst": ReadExtLst(Chart); break;*/
                    default: DataStream.GetXml();
                        break;
                }
            }
        }


        internal void ReadChart(TFlxChart Chart)
        {
            if (DataStream.IsSimpleTag) { DataStream.NextTag(); return; }
            string StartElement = DataStream.RecordName();
            if (!DataStream.NextTag()) return;
            while (!DataStream.AtEndElement(StartElement))
            {
                switch (DataStream.RecordName())
                {
                 /*   case "title": ReadTitle(Chart); break;
                    case "autoTitleDeleted": ReadAutoTitleDeleted(Chart); break;
                    case "pivotFmts": ReadPivotFmts(Chart); break;
                    case "view3D": ReadView3D(Chart); break;
                    case "floor": ReadFloor(Chart); break;
                    case "sideWall": ReadSideWall(Chart); break;
                    case "backWall": ReadBackWall(Chart); break;*/
                    case "plotArea": ReadPlotArea(Chart); break;
                   /* case "legend": ReadLegend(Chart); break;
                    case "plotVisOnly": ReadPlotVisOnly(Chart); break;
                    case "dispBlanksAs": ReadDispBlanksAs(Chart); break;
                    case "showDLblsOverMax": ReadShowDLblsOverMax(Chart); break;
                    case "extLst": ReadExtLst(Chart); break;*/
                    default: DataStream.GetXml();
                        break;
                }
            }
        }


        public void ReadPlotArea(TFlxChart Chart)
        {
            if (DataStream.IsSimpleTag) { DataStream.NextTag(); return; }
            string StartElement = DataStream.RecordName();
            if (!DataStream.NextTag()) return;
            while (!DataStream.AtEndElement(StartElement))
            {
                switch (DataStream.RecordName())
                {
                        /*
                    case "layout": ReadLayout(Chart); break;

                    //choice
                    case "areaChart": ReadAreaChart(Chart); break;
                    case "area3DChart": ReadArea3DChart(Chart); break;
                    case "lineChart": ReadLineChart(Chart); break;
                    case "line3DChart": ReadLine3DChart(Chart); break;
                    case "stockChart": ReadStockChart(Chart); break;
                    case "radarChart": ReadRadarChart(Chart); break;
                    case "scatterChart": ReadScatterChart(Chart); break;
                    case "pieChart": ReadPieChart(Chart); break;
                    case "pie3DChart": ReadPie3DChart(Chart); break;
                    case "doughnutChart": ReadDoughnutChart(Chart); break;
                    case "barChart": ReadBarChart(Chart); break;
                    case "bar3DChart": ReadBar3DChart(Chart); break;
                    case "ofPieChart": ReadOfPieChart(Chart); break;
                    case "surfaceChart": ReadSurfaceChart(Chart); break;
                    case "surface3DChart": ReadSurface3DChart(Chart); break;
                    case "bubbleChart": ReadBubbleChart(Chart); break;

                    //end choice

                    case "valAx": ReadValAx(Chart); break;
                    case "catAx": ReadCatAx(Chart); break;
                    case "dateAx": ReadDateAx(Chart); break;
                    case "serAx": ReadSerAx(Chart); break;
                    case "spPr": ReadSpPr(Chart); break;
                    case "extLst": ReadExtLst(Chart); break;*/
                    default: DataStream.GetXml();
                        break;
                }


            }
        }
    }
}
   
   
