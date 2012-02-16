using System;
using FlexCel.Core;
using System.Collections.Generic;


namespace FlexCel.XlsAdapter
{
	/// <summary>
	/// Implements an ExcelChart interface.
	/// </summary>
	public class XlsChart: ExcelChart
	{
		TFlxChart CurrentChart;
		XlsFile Workbook;
		TSheet CurrentSheet;
		TCellList CellList; //we might later add or remove sheets, we want to keep this one fixed to the sheet. This is why we do not save the ActiveSheet.

		TChartFrameOptions FBackground;
		TFlxChartFont FDefaultFont;
		TFlxChartFont FDefaultLabelFont;
		TFlxChartFont FDefaultAxisFont;
		double FVerticalFontScaling;
		double FHorizontalFontScaling;		
		TPlotEmptyCells FPlotEmptyCells;

		TDataLabel[] FDataLabels;

		/// <summary>
		/// You cannot create instances of this class. It must be returned with a call to <see cref="XlsFile.GetChart(int, string)"/>
		/// </summary>
		internal XlsChart(XlsFile aWorkbook, TFlxChart aCurrentChart): base()
		{
			Workbook = aWorkbook;
			CurrentChart = aCurrentChart;
			CurrentSheet = aWorkbook.InternalWorkbook.Sheets[aWorkbook.ActiveSheet - 1];
			if (aWorkbook.InternalWorkbook.IsWorkSheet(aWorkbook.ActiveSheet - 1))
			{
				CellList = aWorkbook.InternalWorkbook.WorkSheets(aWorkbook.ActiveSheet - 1).Cells.CellList;
			}
			else
			{
				CellList = new TCellList(aWorkbook.InternalWorkbook.Globals, null, null); //for chart sheets.
			}

			FHorizontalFontScaling = 1;
			FVerticalFontScaling = 1;
			FPlotEmptyCells = TPlotEmptyCells.NotPlotted;
            if (CurrentChart != null && CurrentChart.Chart.GetChartCache != null)
            {
                List<TDataLabel> ArrDataLabels = new List<TDataLabel>();
                TChartRecordList Children = CurrentChart.Chart.GetChartCache.Children;
                for (int i = 0; i < Children.Count; i++)
                {
                    TxChartBaseRecord R = Children[i] as TxChartBaseRecord;
                    if (R != null)
                    {
                        switch ((xlr)R.Id)
                        {
                            case xlr.ChartPlotgrowth:
                                TChartPlotGrowthRecord PlotGrowth = R as TChartPlotGrowthRecord;
                                FHorizontalFontScaling = PlotGrowth.XScaling / 65536.0;
                                FVerticalFontScaling = PlotGrowth.YScaling / 65536.0;
                                break;

                            case xlr.ChartFrame:
                                TChartFrameRecord BackgroundFrame = R as TChartFrameRecord;
                                if (BackgroundFrame != null)
                                {
                                    FBackground = BackgroundFrame.GetFrameOptions();
                                }
                                break;

                            case xlr.ChartDefaulttext:
                                TChartDefaultTextRecord DT = R as TChartDefaultTextRecord;
                                switch (DT.AppliesTo)
                                {
                                    case 2:  //default text for all text in the chart.
                                        {
                                            if (i + 1 < Children.Count)
                                            {
                                                TChartTextRecord TR = Children[i + 1] as TChartTextRecord;
                                                if (TR != null)
                                                {
                                                    TChartFontXRecord FontX = (TChartFontXRecord)TR.FindRec<TChartFontXRecord>();
                                                    if (FontX != null) FDefaultFont = FontX.GetFont(Workbook.InternalWorkbook.Globals, Math.Min(FVerticalFontScaling, FHorizontalFontScaling));
                                                }
                                            }
                                            break;

                                        }
                                    case 0:
                                        {
                                            if (i + 1 < Children.Count)
                                            {
                                                TChartTextRecord TR = Children[i + 1] as TChartTextRecord;
                                                if (TR != null)
                                                {
                                                    TChartFontXRecord FontX = (TChartFontXRecord)TR.FindRec<TChartFontXRecord>();
                                                    if (FontX != null) FDefaultLabelFont = FontX.GetFont(Workbook.InternalWorkbook.Globals, Math.Min(FVerticalFontScaling, FHorizontalFontScaling));
                                                }
                                            }
                                            break;
                                        }
                                    case 3: //Not documented, but it is the default font for axis.
                                        {
                                            if (i + 1 < Children.Count)
                                            {
                                                TChartTextRecord TR = Children[i + 1] as TChartTextRecord;
                                                if (TR != null)
                                                {
                                                    TChartFontXRecord FontX = (TChartFontXRecord)TR.FindRec<TChartFontXRecord>();
                                                    if (FontX != null) FDefaultAxisFont = FontX.GetFont(Workbook.InternalWorkbook.Globals, Math.Min(FVerticalFontScaling, FHorizontalFontScaling));
                                                }
                                            }
                                            break;
                                        }

                                }
                                break;

                            case xlr.ChartText:
                                {
                                    if (i - 1 < 0 || !(Children[i - 1] is TChartDefaultTextRecord))
                                    {
                                        TChartTextRecord TR = R as TChartTextRecord;
                                        int SheetIndex = Workbook.InternalWorkbook.Sheets.IndexOf(CurrentSheet) + 1;
                                        if (SheetIndex >= 1)
                                        {
                                            ArrDataLabels.Add(TR.GetDataLabel(Workbook, CellList, SheetIndex, true, true, Math.Min(FVerticalFontScaling, FHorizontalFontScaling)));
                                        }
                                    }
                                    break;
                                }

                            case xlr.ChartShtprops:
                                TChartShtPropsRecord SP = R as TChartShtPropsRecord;
                                FPlotEmptyCells = (TPlotEmptyCells)SP.PlotEmptyCells;
                                break;
                        }
                    }
                }

                FDataLabels = ArrDataLabels.ToArray();
            }

			if (FDefaultFont == null) //FDefaultFont = Workbook.GetFont(15); //Excel does not use this one.
			{
				FDefaultFont = new TFlxChartFont();
				FDefaultFont.Font.Size20 = 200;
				FDefaultFont.Font.Name = "Arial";
				FDefaultFont.Font.Style = TFlxFontStyles.None;
				FDefaultFont.Scale = 1;
			}

			if (FDefaultLabelFont == null) 
			{
				FDefaultLabelFont = FDefaultFont;
			}

			if (FDefaultAxisFont == null) 
			{
				FDefaultAxisFont = new TFlxChartFont(); 
				FDefaultAxisFont.Font = Workbook.GetFormat(FlxConsts.DefaultFormatId).Font;  //We do not fall out to FDefaultFont, but rather to the default font in the workbook.
				FDefaultAxisFont.Scale = 1;
			}

		}

		#region Utilities
		private void CheckConnected()
		{
			if (CurrentChart==null) FlxMessages.ThrowException(FlxErr.ErrNotConnected);
		}

		private static void CheckRange(int val, int lowest, int highest, FlxParam paramName)
		{
            XlsFile.CheckRange(val, lowest, highest, paramName);
		}

        private static void CheckRangeObjPath(string ObjPath, int val, int lowest, int highest, FlxParam paramName)
        {
            XlsFile.CheckRangeObjPath(ObjPath, val, lowest, highest, paramName);
        }

		#endregion

		#region Background
		/// <summary>
		/// Options for the background of the full chart. If this member is null, the options for the Autoshape will be used.
		/// </summary>
		public override TChartFrameOptions Background
		{
			get
			{
				return FBackground;
			}
		}

		/// <summary>
		/// Returns the default font for all text in the chart that does not have a font defined.
		/// </summary>
		public override TFlxChartFont DefaultFont
		{
			get
			{
				return FDefaultFont;
			}
		}

		/// <summary>
		/// Returns the default font for all labels in the chart that do not have a font defined.
		/// </summary>
		public override TFlxChartFont DefaultLabelFont
		{
			get
			{
				return FDefaultLabelFont;
			}
		}

		/// <summary>
		/// Returns the default font for the Axis in the chart that do not have a font defined.
		/// </summary>
		public override TFlxChartFont DefaultAxisFont
		{
			get
			{
				return FDefaultAxisFont;
			}
		}

 
		#endregion

		#region ChartOptions
		///<inheritdoc />
		public override TChartOptions[] ChartOptions
		{
			get
			{
				if (CurrentChart == null || CurrentChart.Chart.GetChartCache == null) return new TChartOptions[]{new TUnknownChartOptions()};

				List<TChartOptions> Result = new List<TChartOptions>();
				for (int i = 0; i < CurrentChart.Chart.GetChartCache.Children.Count; i++)
				{
					TChartAxisParentRecord AP = CurrentChart.Chart.GetChartCache.Children[i] as TChartAxisParentRecord;

					if (AP == null) continue;
					for (int k = 0; k < AP.Children.Count; k++)
					{
						TChartChartFormatRecord CF = AP.Children[k] as TChartChartFormatRecord;

						if (CF == null)
						{
							continue;
						}

						TChartFormatBaseRecord FB = (TChartFormatBaseRecord)CF.FindRec<TChartFormatBaseRecord>();

						if (FB == null)
						{
							Result.Add(new TUnknownChartOptions());
							continue;
						}

                        TChartDataFormatRecord DF = null;
						TChartDropBarRecord UpBar = null;
						TChartDropBarRecord DownBar = null;
						ChartLineOptions DropLineFormat = null;
						ChartLineOptions HiLoLineFormat = null;
						ChartLineOptions SeriesLineFormat = null;
						TDataLabel SeriesDefaultLabel = null;
                        TChartAttachedLabelRecord AttachedLabel = null;

						for (int z = 0; z < CF.Children.Count; z++)
						{
							TxChartBaseRecord R = CF.Children[z] as TxChartBaseRecord;
							if (R == null) continue;
					
							switch ((xlr)R.Id)
							{
								case xlr.ChartDataformat: DF = (TChartDataFormatRecord)R; break;
								case xlr.ChartDropbar: if (UpBar == null) UpBar = (TChartDropBarRecord)R; 
													   else if (DownBar == null) DownBar = (TChartDropBarRecord)R;
									break;
								case xlr.ChartChartline: 
								{
									TChartChartLineRecord CCL = (TChartChartLineRecord)R;
									if (CCL.HasDropLines)
									{
										if (z + 1 < CF.Children.Count)
										{
											TChartLineFormatRecord FR = CF.Children[z+1] as TChartLineFormatRecord;
											if (FR != null) DropLineFormat = FR.GetLineFormat();
										}
									}
									if (CCL.HasHiLoLines)
									{
										if (z + 1 < CF.Children.Count)
										{
											TChartLineFormatRecord FR = CF.Children[z+1] as TChartLineFormatRecord;
											if (FR != null) HiLoLineFormat = FR.GetLineFormat();
										}
									}
									if (CCL.HasSeriesLines)
									{
										if (z + 1 < CF.Children.Count)
										{
											TChartLineFormatRecord FR = CF.Children[z+1] as TChartLineFormatRecord;
											if (FR != null) SeriesLineFormat = FR.GetLineFormat();
										}
									}
									break;
								}

								case xlr.ChartDefaulttext:
								{
									TChartDefaultTextRecord DT = (TChartDefaultTextRecord)R;
									if (z + 1 <CF.Children.Count && DT.AppliesTo != 2)
									{
										TChartTextRecord TR = CF.Children[z+1] as TChartTextRecord;
										if (TR != null)
										{
											int SheetIndex = Workbook.InternalWorkbook.Sheets.IndexOf(CurrentSheet) + 1;
											if (SheetIndex >= 1)
											{
												SeriesDefaultLabel = TR.GetDataLabel(Workbook, CellList, SheetIndex, false, true, Math.Min(FVerticalFontScaling, FHorizontalFontScaling));
											}
										}
									}
									break;
								}

                                case xlr.ChartAttachedlabel:
                                AttachedLabel = (TChartAttachedLabelRecord)R;
                                break;
							}
						}

                        if (AttachedLabel == null)
                        {
                            if (DF != null)
                            {
                                AttachedLabel = (TChartAttachedLabelRecord)DF.FindRec<TChartAttachedLabelRecord>();
                            }
                        }

                        if (AttachedLabel == null)
                        {
                            SeriesDefaultLabel = null;
                        }

						Result.Add(GetOneChart(AP, CF, FB, DF, UpBar, DownBar, DropLineFormat, HiLoLineFormat, SeriesLineFormat, SeriesDefaultLabel));
					}
				}

				if (Result.Count == 0) return new TChartOptions[]{new TUnknownChartOptions()};
				Result.Sort();
				return Result.ToArray();
			}
		}

		private static TChartOptions GetOneChart(TChartAxisParentRecord AP, TChartChartFormatRecord CF, TChartFormatBaseRecord FB, TChartDataFormatRecord DF, 
			TChartDropBarRecord UpBar, TChartDropBarRecord DownBar, 
			ChartLineOptions DropLineFormat, ChartLineOptions HiLoLineFormat, ChartLineOptions SeriesLineFormat, TDataLabel SeriesDefaultLabel)
		{
			if (FB == null) return new TUnknownChartOptions();

			TChartAreaRecord Area = (FB as TChartAreaRecord);
			if (Area != null)
			{
				int f = Area.Flags;
				TStackedMode StackedMode = TStackedMode.None;
				if ((f & 0x1) != 0)
				{
					if ((f & 0x2) != 0) StackedMode = TStackedMode.Stacked100; else StackedMode = TStackedMode.Stacked;
				}

				return new TAreaChartOptions(AP.Index, CF.ChangeColorsOnEachSeries, CF.ZOrder, StackedMode, (f & 0x4) != 0, TChartSeriesRecord.GetSeriesOptions(DF), TChartSeriesRecord.GetPlotArea(AP), 
					GetDropBarOptions(UpBar, DownBar, DropLineFormat, HiLoLineFormat), SeriesDefaultLabel);
			}

			TChartBarRecord Bar = (FB as TChartBarRecord);
			if (Bar != null)
			{
				int f = Bar.Flags;
				TStackedMode StackedMode = TStackedMode.None;
				if ((f & 0x2) != 0) 
				{
					if ((f & 0x4) != 0) StackedMode = TStackedMode.Stacked100; else StackedMode = TStackedMode.Stacked;
				}
				return new TBarChartOptions(AP.Index, CF.ChangeColorsOnEachSeries, CF.ZOrder, Bar.BarOverlap / 100f, Bar.CategoriesOverlap / 100f, 
					(f & 0x1) != 0, StackedMode, (f & 0x8) != 0, TChartSeriesRecord.GetSeriesOptions(DF), TChartSeriesRecord.GetPlotArea(AP), 
					SeriesDefaultLabel, SeriesLineFormat) ;
			}

			TChartLineRecord Line = (FB as TChartLineRecord);
			if (Line != null)
			{
				int f = Line.Flags;
				TStackedMode StackedMode = TStackedMode.None;
				if ((f & 0x1) != 0)
				{
					if ((f & 0x2) != 0) StackedMode = TStackedMode.Stacked100; else StackedMode = TStackedMode.Stacked;
				}

				return new TLineChartOptions(AP.Index, CF.ChangeColorsOnEachSeries, CF.ZOrder, StackedMode, (f & 0x4) != 0, TChartSeriesRecord.GetSeriesOptions(DF), TChartSeriesRecord.GetPlotArea(AP), 
					GetDropBarOptions(UpBar, DownBar, DropLineFormat, HiLoLineFormat), SeriesDefaultLabel);
			}

			TChartPieRecord Pie = (FB as TChartPieRecord);
			if (Pie != null)
			{
				int f = Pie.Flags;
				bool HasLeaderLines = (f & 0x2) != 0;

				ChartLineOptions LeaderLineFormat = null;
				if (HasLeaderLines)
				{
					TChartLineFormatRecord CLF = (TChartLineFormatRecord)CF.FindRec<TChartLineFormatRecord>();
					if (CLF != null)
					{
						LeaderLineFormat = CLF.GetLineFormat();
					}
				}
				return new TPieChartOptions(AP.Index, CF.ChangeColorsOnEachSeries, CF.ZOrder, Pie.FirstSliceAngle, Pie.DonutRadius, (f & 0x1) != 0, HasLeaderLines, LeaderLineFormat, TChartSeriesRecord.GetSeriesOptions(DF), TChartSeriesRecord.GetPlotArea(AP), SeriesDefaultLabel);
			}

			TChartRadarRecord Radar = (FB as TChartRadarRecord);
			if (Radar != null)
			{
			}

			TChartScatterRecord Scatter = (FB as TChartScatterRecord);
			if (Scatter != null)
			{
				int f = Scatter.Flags;
				TBubbleSizeType BubbleSize = TBubbleSizeType.BubbleSizeIsArea;
				if (Scatter.BubbleSize == 2)
				{
					BubbleSize = TBubbleSizeType.BubbleSizeIsWidth;
				}

				return new TScatterChartOptions(AP.Index, CF.ChangeColorsOnEachSeries, CF.ZOrder, TChartSeriesRecord.GetSeriesOptions(DF), 
					TChartSeriesRecord.GetPlotArea(AP), Scatter.PercentOfLargestBubble, BubbleSize, (f & 0x1) != 0, (f & 0x2) != 0, (f & 0x3) != 0, SeriesDefaultLabel);
			}

			TChartSurfaceRecord Surface = (FB as TChartSurfaceRecord);
			if (Surface != null)
			{
			}

			return new TUnknownChartOptions();
		
		}

		private static TChartDropBars GetDropBarOptions(TChartDropBarRecord UpDropBar, TChartDropBarRecord DownDropBar, 
			ChartLineOptions DropLineFormat, ChartLineOptions HiLoLineFormat)
		{
			TChartOneDropBar UpBar = null;
			if (UpDropBar != null)
			{
				UpBar = new TChartOneDropBar(UpDropBar.GetFrame(), UpDropBar.GapWidth);
			}

			TChartOneDropBar DownBar = null;
			if (DownDropBar != null)
			{
				DownBar = new TChartOneDropBar(DownDropBar.GetFrame(),DownDropBar.GapWidth);
			}
			return new TChartDropBars(DropLineFormat, HiLoLineFormat, UpBar, DownBar);
		}

		

		///<inheritdoc />
		public override TPlotEmptyCells PlotEmptyCells
		{
			get
			{
				return FPlotEmptyCells;
			}
			set
			{
				FPlotEmptyCells = value;
			}
		}


		#endregion

		#region Series
		///<inheritdoc />
		public override int AddSeries(ChartSeries value)
		{
			CheckConnected();
			int SheetIndex = Workbook.InternalWorkbook.Sheets.IndexOf(CurrentSheet) + 1;
			CurrentChart.Chart.AddSeries(TChartSeriesRecord.CreateFromData(Workbook, value, SheetIndex, CurrentChart.Chart.Cache));
			return SeriesCount - 1;
		}

		///<inheritdoc />
		public override void DeleteSeries(int index)
		{
			CheckConnected();
			CheckRange(index, 1, SeriesCount, FlxParam.SeriesIndex);
			CurrentChart.Chart.DeleteSeries(index - 1);
            CurrentChart.SeriesData.DeleteSeries(index);
		}

		///<inheritdoc />
		public override ChartSeries GetSeries(int index, bool getDefinitions, bool getValues, bool getOptions)
		{
			CheckConnected();
			CheckRange(index, 1, SeriesCount, FlxParam.SeriesIndex);

			int SheetIndex = Workbook.InternalWorkbook.Sheets.IndexOf(CurrentSheet) + 1;
			if (SheetIndex < 1)
			{
				return new ChartSeries(string.Empty, string.Empty, string.Empty, TFlxFormulaErrorValue.ErrRef, new object[]{TFlxFormulaErrorValue.ErrRef}, new object[]{TFlxFormulaErrorValue.ErrRef}, null, null, 0, 0, 0);
			}
			return CurrentChart.Chart.Cache.Series[index - 1].GetValuesAndDefinition(Workbook, CellList, SheetIndex, getDefinitions, getValues, getOptions, Math.Min(FVerticalFontScaling, FHorizontalFontScaling));
		}

		///<inheritdoc />
		public override void SetSeries(int index, ChartSeries value)
		{
			CheckConnected();
			CheckRange(index, 1, SeriesCount, FlxParam.SeriesIndex);
			CurrentChart.Chart.Cache.Series[index - 1].SetDefinition(Workbook, value, CurrentChart.Chart.Cache);
		}

		///<inheritdoc />
		public override int SeriesCount
		{
			get
			{
				return CurrentChart.Chart.Cache.Series.Count;
			}
		}
		#endregion

		#region Axis
		///<inheritdoc />
		public override TChartAxis[] GetChartAxis()
		{
			return CurrentChart.Chart.GetChartAxis(Workbook, CurrentSheet, CellList, Math.Min(FVerticalFontScaling, FHorizontalFontScaling));
		}

		#endregion

		#region Legend
		///<inheritdoc />
		public override TChartLegend GetChartLegend()
		{
			return CurrentChart.Chart.GetChartLegend(Math.Min(FVerticalFontScaling, FHorizontalFontScaling));
		}

		#endregion

		#region Labels
		///<inheritdoc />
		public override TDataLabel[] GetDataLabels()
		{
			return FDataLabels;
		}

        ///<inheritdoc />
        public override void SetDataLabels(TDataLabel[] labels)
        {
            FDataLabels = labels;
        }

		#endregion

		#region Objects
		///<inheritdoc />
		public override int ObjectCount
		{
			get
			{
				CheckConnected();
				return CurrentChart.Drawing.ObjectCount;
			}
		}

		///<inheritdoc />
		public override TShapeProperties GetObjectProperties(int objectIndex, bool GetShapeOptions)
		{
			CheckConnected();
			CheckRange(objectIndex, 1, ObjectCount, FlxParam.ObjectIndex);
			return CurrentChart.Drawing.GetObjectProperties(objectIndex - 1, GetShapeOptions);
		}

		///<inheritdoc />
		public override void SetObjectText(int objectIndex, string objectPath, TRichString text)
		{
			CheckConnected();
			CheckRangeObjPath(objectPath, objectIndex, 1, ObjectCount, FlxParam.ObjectIndex);
			CurrentChart.Drawing.SetObjectText(objectIndex -1, objectPath, text, Workbook, null);
		}

		///<inheritdoc />
		public override void DeleteObject(int objectIndex)
		{
			CheckConnected();
			CheckRange(objectIndex, 1, ObjectCount, FlxParam.ObjectIndex);
			CurrentChart.Drawing.DeleteObject(objectIndex-1);
		}


		#endregion
	}
}
