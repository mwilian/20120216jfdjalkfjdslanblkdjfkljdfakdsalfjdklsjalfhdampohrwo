using System;
using System.Collections;
using System.Collections.Generic;

#if (MONOTOUCH)
    using Color = MonoTouch.UIKit.UIColor;
#endif

#if (WPF)
using System.Windows.Media;
using Rectangle = System.Windows.Rect;
#else
using System.Drawing;
#endif


namespace FlexCel.Core
{
	#region Enums
	/// <summary>
	/// Chart style.
	/// </summary>
	public enum TChartType
	{
		/// <summary>
		/// FlexCel cannot determine the chart type.
		/// </summary>
		Unknown,

		/// <summary>
		/// Area chart.
		/// </summary>
		Area,

		/// <summary>
		/// Bar chart.
		/// </summary>
		Bar,

		/// <summary>
		/// Line chart.
		/// </summary>
		Line,

		/// <summary>
		/// Pie chart.
		/// </summary>
		Pie,

		/// <summary>
		/// Radar chart.
		/// </summary>
		Radar,

		/// <summary>
		/// Scatter chart.
		/// </summary>
		Scatter,

		/// <summary>
		/// Surface chart.
		/// </summary>
		Surface
        
	}

	/// <summary>
	/// Way the series stack one to the other. This does not apply to all chart types.
	/// </summary>
	public enum TStackedMode
	{
		/// <summary>
		/// Data is not Stacked.
		/// </summary>
		None,

		/// <summary>
		/// Data is Stacked. (for example in stacked bar chart)
		/// </summary>
		Stacked,

		/// <summary>
		/// Data is stacked and normalized to 100%. (This is a 100% stacked bar/column chart)
		/// </summary>
		Stacked100
	}


	/// <summary>
	/// Defines how empty cells will be plotted in the chart.
	/// </summary>
	public enum TPlotEmptyCells
	{
		/// <summary>
		/// The cell will not be plotted. There will be a gap in the chart.
		/// </summary>
		NotPlotted = 0,

		/// <summary>
		/// The cell will be plotted with value = 0.
		/// </summary>
		Zero = 1,

		/// <summary>
		/// The cell will be plotted with value interpolated between its nearest points.
		/// </summary>
		Interpolated = 2
	}

	/// <summary>
	/// Line styles for a chart object.
	/// </summary>
	public enum TChartLineStyle
	{
		/// <summary>
		/// Sloid line.
		/// </summary>
		Solid = 0,

		/// <summary>
		/// Dashed line.
		/// </summary>
		Dash = 1,

		/// <summary>
		/// Dotted line.
		/// </summary>
		Dot = 2,

		/// <summary>
		/// Dash-dot line.
		/// </summary>
		DashDot = 3,

		/// <summary>
		/// Dash dot dot line.
		/// </summary>
		DashDotDot = 4,

		/// <summary>
		/// No line.
		/// </summary>
		None = 5,

		/// <summary>
		/// Dark gray line.
		/// </summary>
		DarkGray = 6,

		/// <summary>
		/// Medium gray line.
		/// </summary>
		MediumGray = 7,

		/// <summary>
		/// Light gray line.
		/// </summary>
		LightGray = 8
	}

	/// <summary>
	/// Kind of marker.
	/// </summary>
	public enum TChartMarkerType
	{
		/// <summary>
		/// No marker.
		/// </summary>
		None = 0,

		/// <summary>
		/// Square.
		/// </summary>
		Square = 1,
 
		/// <summary>
		/// Diamond.
		/// </summary>
		Diamond = 2,

		/// <summary>
		/// Up triangle.
		/// </summary>
		Triangle = 3,

		/// <summary>
		/// X sign.
		/// </summary>
		X = 4,

		/// <summary>
		/// Star.
		/// </summary>
		Star = 5,

		/// <summary>
		/// Dow jones symbol. (small horizontal line)
		/// </summary>
		DowJones = 6,

		/// <summary>
		/// Standard deviation symbol. (big horizontal line)
		/// </summary>
		StandardDeviation = 7,

		/// <summary>
		/// Circle.
		/// </summary>
		Circle = 8,

		/// <summary>
		/// Plus sign.
		/// </summary>
		Plus = 9
	}

	/// <summary>
	/// Different widths for chart lines.
	/// </summary>
	public enum TChartLineWeight
	{
		/// <summary>
		/// Minimum width.
		/// </summary>
		Hair = -1,

		/// <summary>
		/// Single width.
		/// </summary>
		Narrow = 0,
        
		/// <summary>
		/// Double width.
		/// </summary>
		Medium = 1,
	    
		/// <summary>
		/// Triple width.
		/// </summary>
		Wide = 2
	}

	#endregion

	#region Series 
	/// <summary>
	/// The definiton for a series, and the values of it.
	/// </summary>
	public class ChartSeries
	{
		private string FTitleDefinition;
		private string FDataDefinition;
		private string FCategoriesDefinition;
		private object FTitleValue;
		private object[] FDataValues;
		private object[] FCategoriesValues;
		
		private string[] FDataFormats;
		private string[] FCategoriesFormats;

		private int FChartOptionsIndex;
		private int FSeriesIndex;
		private int FSeriesNumber;

		private TSeriesOptionsList FOptions;
		private TLegendOptionsList FLegendOptions;

		/// <summary>
		/// Creates a new empty instance.
		/// </summary>
		public ChartSeries()
		{
			FOptions = new TSeriesOptionsList();
			FLegendOptions = new TLegendOptionsList();
		}

		/// <summary>
		/// Creates a new instance and fills it with default values.
		/// </summary>
		/// <param name="aTitleDefinition">Title of the series.</param>
		/// <param name="aDataDefinition">Formula defining the Data for the series.</param>
		/// <param name="aCategoriesDefinition">Formula defining the Categories for the series.</param>
		/// <param name="aTitleValue">Evaluated value of the series title.</param>
		/// <param name="aCategoriesValues">Actual values for the series.</param>
		/// <param name="aDataValues">Actual values for the categories.</param>
		/// <param name="aChartOptionsIndex">Index to the ChartOptions object that applies to this series.</param>
		/// <param name="aSeriesIndex">Index of this series on the file.</param>
		/// <param name="aSeriesNumber">Series number as shown on the Legend box. This might be different from the SeriesIndex if the order of the series is changed.</param>
		/// <param name="aDataFormats">See <see cref="DataFormats"/></param>
		/// <param name="aCategoriesFormats">See <see cref="CategoriesFormats"/></param>
		public ChartSeries(string aTitleDefinition, string aDataDefinition, string aCategoriesDefinition, object aTitleValue, object[] aDataValues, object[] aCategoriesValues, 
			string[] aDataFormats, string[] aCategoriesFormats, int aChartOptionsIndex, int aSeriesIndex, int aSeriesNumber): this()
		{
			FTitleDefinition = aTitleDefinition;
			FDataDefinition = aDataDefinition;
			FCategoriesDefinition = aCategoriesDefinition;
			FTitleValue = aTitleValue;
			FDataValues = aDataValues;
			FCategoriesValues = aCategoriesValues;
			FChartOptionsIndex = aChartOptionsIndex;
			FSeriesIndex = aSeriesIndex;
			FSeriesNumber = aSeriesNumber;
			FCategoriesFormats = aCategoriesFormats;
			FDataFormats = aDataFormats;
		}

		/// <summary>
        /// Formula or text defining the Series caption. Start with an &quot;=&quot; sign to enter a formula.
		/// </summary>
		public string TitleDefinition {get {return FTitleDefinition;} set{FTitleDefinition = value;}}

		/// <summary>
        /// Formula or values that define the series. For example, &quot;=A1:A5&quot; or {1,2,3}. Start with an &quot;=&quot; sign to enter a formula.
		/// </summary>
		public string DataDefinition {get {return FDataDefinition;} set {FDataDefinition = value;}}

		/// <summary>
        /// Formula or text defining the Series Categories (normally the x Axis). Start with an &quot;=&quot; sign to enter a formula.
		/// </summary>
		public string CategoriesDefinition {get {return FCategoriesDefinition;} set{FCategoriesDefinition = value;}}


		/// <summary>
		/// Evaluated text of the title.
		/// </summary>
		public object TitleValue {get {return FTitleValue;} set{FTitleValue = value;}}
		
		/// <summary>
		/// Actual values for the series.
		/// </summary>
		public object[] DataValues {get {return FDataValues;} set {FDataValues = value;}}

		/// <summary>
		/// Actual values for the Series Categories (normally the x Axis).
		/// </summary>
		public object[] CategoriesValues {get {return FCategoriesValues;} set{FCategoriesValues = value;}}

		/// <summary>
		/// Format on the cell where the data is. This format should be applied to the data if the data format for the axis or the label is null.
		/// </summary>
		public string[] DataFormats {get {return FDataFormats;} set{FDataFormats = value;}}

		/// <summary>
		/// Format on the cell where the data is. This format should be applied to the data if the data format for the axis or the label is null.
		/// </summary>
		public string[] CategoriesFormats {get {return FCategoriesFormats;} set{FCategoriesFormats = value;}}

		/// <summary>
		/// Index to the ChartOptions object that applies to this series.
		/// </summary>
		public int ChartOptionsIndex {get {return FChartOptionsIndex;} set{FChartOptionsIndex = value;}}

		/// <summary>
		/// Index of this series on the file.
		/// </summary>
		public int SeriesIndex {get {return FSeriesIndex;} set{FSeriesIndex = value;}}

		/// <summary>
		/// Series number as shown on the Legend box. This might be different from the <see cref="SeriesIndex"/> if the order of the series is changed.
		/// </summary>
		public int SeriesNumber {get {return FSeriesNumber;} set{FSeriesNumber = value;}}

		/// <summary>
		/// Options for this series and their data points. -1 means the whole series, and n is options for the n-point.
		/// </summary>
		public TSeriesOptionsList Options {get {return FOptions;}}

		/// <summary>
		/// Options for the legend entry associated with this series. (when legend is showing series), or
		/// with a point on the series. (when legend is showing all the entries on series[0], for example on pie charts)
		/// </summary>
		public TLegendOptionsList LegendOptions {get {return FLegendOptions;} set{FLegendOptions = value;}}

	}


	#region Series Options
	/// <summary>
	/// Options for the whole series or for a data point inside it.
	/// </summary>
	public class ChartSeriesOptions: ICloneable
	{
		private int FPointNumber;
		private ChartSeriesFillOptions FFillOptions;
		private ChartSeriesLineOptions FLineOptions;
		private ChartSeriesPieOptions FPieOptions;
		private ChartSeriesMarkerOptions FMarkerOptions;
		private ChartSeriesMiscOptions FMiscOptions;
		private TShapeOptionList FExtraOptions;

		/// <summary>
		/// Creates a new instance of ChartSeriesOptions. Objects will be cloned, so you can change their values later and they will not
		/// change the value on this class.
		/// </summary>
		/// <param name="aPointNumber">Point number where this options apply. -1 means that the options apply for the whole series.</param>
		/// <param name="aFillOptions">Fill options for the series or point.</param>
		/// <param name="aLineOptions">Line options for the series or point.</param>
		/// <param name="aPieOptions">If the chart type is pie, options for the pie. If it is not a pie chart, this value has no meaning.</param>
		/// <param name="aMarkerOptions">Color and shape of the markers on Line and Scatter charts.</param>
		/// <param name="aMiscOptions">Misc Options.</param>
		public ChartSeriesOptions(int aPointNumber, ChartSeriesFillOptions aFillOptions, ChartSeriesLineOptions aLineOptions, ChartSeriesPieOptions aPieOptions, ChartSeriesMarkerOptions aMarkerOptions, ChartSeriesMiscOptions aMiscOptions)
		{
			FPointNumber = aPointNumber;
			if (aFillOptions != null) FFillOptions = (ChartSeriesFillOptions)aFillOptions.Clone();
			if (aLineOptions != null) FLineOptions =  (ChartSeriesLineOptions)aLineOptions.Clone();
			if (aPieOptions != null) FPieOptions = (ChartSeriesPieOptions)aPieOptions.Clone();
			if (aMarkerOptions != null) FMarkerOptions = (ChartSeriesMarkerOptions)aMarkerOptions.Clone();
			if (aMiscOptions != null) FMiscOptions = (ChartSeriesMiscOptions)aMiscOptions.Clone();
		}

		/// <summary>
		/// Point number where this options apply. -1 means that the options apply for the whole series.
		/// </summary>
		public int PointNumber {get {return FPointNumber;} set{FPointNumber = value;}}

		/// <summary>
		/// Fill options for the series or point.
		/// </summary>
		public ChartSeriesFillOptions FillOptions {get {return FFillOptions;} set{FFillOptions = value;}}

		/// <summary>
		/// Line options for the series or point.
		/// </summary>
		public ChartSeriesLineOptions LineOptions {get {return FLineOptions;} set{FLineOptions = value;}}

		/// <summary>
		/// If the chart type is pie, options for the pie. If it is not a pie chart, this value has no meaning.
		/// </summary>
		public ChartSeriesPieOptions PieOptions {get {return FPieOptions;} set{FPieOptions = value;}}

		/// <summary>
		/// Color and shape of the markers on Line and Scatter charts.
		/// </summary>
		public ChartSeriesMarkerOptions MarkerOptions {get {return FMarkerOptions;} set{FMarkerOptions = value;}}

		/// <summary>
		/// Other options not related to the specific parts.
		/// </summary>
		public ChartSeriesMiscOptions MiscOptions {get {return FMiscOptions;} set{FMiscOptions = value;}}

		/// <summary>
		/// Extra options used to specify gradients, textures, etc. If this is no null, the color values are not used.
		/// </summary>
		public TShapeOptionList ExtraOptions {get {return FExtraOptions;} set{FExtraOptions = value;}}


		#region ICloneable Members

		/// <summary>
		/// Returns a deep copy of this object.
		/// </summary>
		/// <returns></returns>
		public object Clone()
		{
			ChartSeriesOptions Result = (ChartSeriesOptions) MemberwiseClone();
			if (FFillOptions != null)   Result.FFillOptions = (ChartSeriesFillOptions)FFillOptions.Clone();
			if (FLineOptions != null)   Result.FLineOptions =  (ChartSeriesLineOptions)FLineOptions.Clone();
			if (FPieOptions != null)    Result.FPieOptions = (ChartSeriesPieOptions)FPieOptions.Clone();
			if (FMarkerOptions != null) Result.FMarkerOptions = (ChartSeriesMarkerOptions)FMarkerOptions.Clone();
			if (FMiscOptions != null)   Result.FMiscOptions = (ChartSeriesMiscOptions)FMiscOptions.Clone();
			if (FExtraOptions != null)  Result.ExtraOptions = (TShapeOptionList)FExtraOptions.Clone();
			return Result;
		}

		#endregion
	}


	/// <summary>
	/// Fill style used in a pattern inside chart elements.
	/// </summary>
	public enum TChartPatternStyle
	{
		/// <summary>
		/// Transparent pattern.
		/// </summary>
		None = 0,

		/// <summary>
		/// Pattern defined by the context.
		/// </summary>
		Automatic = 1
	}

	/// <summary>
	/// Fill options for a chart element.
	/// </summary>
	public class ChartFillOptions: ICloneable
	{
		private Color FFgColor;
		private Color FBgColor;
		private TChartPatternStyle FPattern;
		
		/// <summary>
		/// Creates a new instance of this object.
		/// </summary>
		/// <param name="aFgColor">Foreground color for the pattern.</param>
		/// <param name="aBgColor">Background color for the pattern.</param>
		/// <param name="aPattern">Pattern style.</param>
		public ChartFillOptions(Color aFgColor, Color aBgColor, TChartPatternStyle aPattern)
		{
			FFgColor = aFgColor;
			FBgColor = aBgColor;
			FPattern = aPattern;
		}

		/// <summary>
		/// Foreground color for the pattern.
		/// </summary>
		public Color FgColor {get {return FFgColor;} set{FFgColor = value;}}

		/// <summary>
		/// Background color for the pattern.
		/// </summary>
		public Color BgColor {get {return FBgColor;} set{FBgColor = value;}}

		/// <summary>
		/// Pattern style.
		/// </summary>
		public TChartPatternStyle Pattern {get {return FPattern;} set{FPattern = value;}}

		#region ICloneable Members

		/// <summary>
		/// Creates a deep copy of this object.
		/// </summary>
		/// <returns></returns>
		public virtual object Clone()
		{
			return MemberwiseClone();
		}

		#endregion
	}

	/// <summary>
	/// Fill options for a series or a point inside a series.
	/// </summary>
	public class ChartSeriesFillOptions: ChartFillOptions, ICloneable
	{
		private bool FAutomaticColors;
		private bool FInvertNegativeValues;

		/// <summary>
		/// Creates a new instance of this object.
		/// </summary>
		/// <param name="aFgColor">Foreground color for the pattern.</param>
		/// <param name="aBgColor">Background color for the pattern.</param>
		/// <param name="aPattern">Pattern style.</param>
		/// <param name="aAutomaticColors">When true, fill colors are assigned automatically.</param>
		/// <param name="aInvertNegativeValues">When true and values of the series are negative, foreground and background colors are reversed.</param>
		public ChartSeriesFillOptions(Color aFgColor, Color aBgColor, TChartPatternStyle aPattern, bool aAutomaticColors, bool aInvertNegativeValues): base(aFgColor, aBgColor, aPattern)
		{
			FAutomaticColors = aAutomaticColors;
			FInvertNegativeValues = aInvertNegativeValues;
		}

		/// <summary>
		/// When true, fill colors are assigned automatically.
		/// </summary>
		public bool AutomaticColors {get {return FAutomaticColors;} set{FAutomaticColors = value;}}

		/// <summary>
		/// When true and values of the series are negative, foreground and background colors are reversed.
		/// </summary>
		public bool InvertNegativeValues {get {return FInvertNegativeValues;} set{FInvertNegativeValues = value;}}

		#region ICloneable Members
		/// <summary>
		/// Returns a deep copy of this object.
		/// </summary>
		/// <returns></returns>
		public override object Clone()
		{
			return MemberwiseClone();
		}

		#endregion
	}

	/// <summary>
	/// Line style options for an element of a chart.
	/// </summary>
	public class ChartLineOptions: ICloneable
	{
		private Color FLineColor;
		private TChartLineStyle FStyle;
		private TChartLineWeight FLineWeight;

		/// <summary>
		/// Creates a new ChartSeriesLineOptions instance.
		/// </summary>
		/// <param name="aLineColor">Line color.</param>
		/// <param name="aStyle">Line style.</param>
		/// <param name="aLineWeight">Line weight.</param>
		public ChartLineOptions(Color aLineColor, TChartLineStyle aStyle, TChartLineWeight aLineWeight)
		{
			FLineColor = aLineColor;
			FStyle = aStyle;
			FLineWeight = aLineWeight;
		}

		/// <summary>
		/// Line color.
		/// </summary>
		public Color LineColor {get {return FLineColor;} set{FLineColor = value;}}

		/// <summary>
		/// Line style.
		/// </summary>
		public TChartLineStyle Style {get {return FStyle;} set{FStyle = value;}}

		/// <summary>
		/// Line weight.
		/// </summary>
		public TChartLineWeight LineWeight {get {return FLineWeight;} set{FLineWeight = value;}}

		#region ICloneable Members
		/// <summary>
		/// Returns a deep copy of this object.
		/// </summary>
		/// <returns></returns>
		public virtual object Clone()
		{
			return MemberwiseClone();
		}

		#endregion

	}


	/// <summary>
	/// Line style options for a series or a point inside a series.
	/// </summary>
	public class ChartSeriesLineOptions: ChartLineOptions, ICloneable
	{
		private bool FAutomaticColors;

		/// <summary>
		/// Creates a new ChartSeriesLineOptions instance.
		/// </summary>
		/// <param name="aLineColor">Line color.</param>
		/// <param name="aStyle">Line style.</param>
		/// <param name="aLineWeight">Line weight.</param>
		/// <param name="aAutomaticColors">When true, line colors are assigned automatically.</param>
		public ChartSeriesLineOptions(Color aLineColor, TChartLineStyle aStyle, TChartLineWeight aLineWeight, bool aAutomaticColors): base (aLineColor, aStyle, aLineWeight)
		{
			FAutomaticColors = aAutomaticColors;
		}

		/// <summary>
		/// When true, line colors are assigned automatically.
		/// </summary>
		public bool AutomaticColors {get {return FAutomaticColors;} set{FAutomaticColors = value;}}

		#region ICloneable Members
		/// <summary>
		/// Returns a deep copy of this object.
		/// </summary>
		/// <returns></returns>
		public override object Clone()
		{
			return MemberwiseClone();
		}

		#endregion

	}

	/// <summary>
	/// Pie options for a series or a slice of the series when the chart is a pie chart.
	/// </summary>
	public class ChartSeriesPieOptions: ICloneable
	{
		private int FSliceDistance;

		/// <summary>
		/// Creates a new ChartSeriesPieOptions instance.
		/// </summary>
		/// <param name="aSliceDistance">Distance of the pie slice from the center on percent of the pie diemeter.</param>
		public ChartSeriesPieOptions(int aSliceDistance)
		{
			FSliceDistance = aSliceDistance;
		}

		/// <summary>
		/// Distance of the pie slice from the center on percent of the pie diammeter.
		/// </summary>
		public int SliceDistance {get {return FSliceDistance;} set{FSliceDistance = value;}}

		#region ICloneable Members
		/// <summary>
		/// Returns a deep copy of this object.
		/// </summary>
		/// <returns></returns>
		public object Clone()
		{
			return MemberwiseClone();
		}

		#endregion

	}

	/// <summary>
	/// Marker options for the whole series or a point in the series when the chart is line or scatter.
	/// </summary>
	public class ChartSeriesMarkerOptions: ICloneable
	{
		private Color FFgColor;
		private Color FBgColor;
		private TChartMarkerType FMarkerType;
		private bool FAutomaticColors;
		private bool FNoBackground;
		private bool FNoForeground;
		private int FMarkerSize;

		/// <summary>
		/// Creates a new ChartSeriesMarkerOptions instance.
		/// </summary>
		/// <param name="aFgColor">Color of the marker lines.</param>
		/// <param name="aBgColor">Color of the marker fill.</param>
		/// <param name="aMarkerType">Type of marker.</param>
		/// <param name="aAutomaticColors">When true, marker colors are assigned automatically.</param>
		/// <param name="aNoBackground">When true, the marker has no fill.</param>
		/// <param name="aNoForeground">When true, the marker has no lines.</param>
		/// <param name="aMarkerSize">Marker size.</param>
		public ChartSeriesMarkerOptions(Color aFgColor, Color aBgColor, TChartMarkerType aMarkerType, bool aAutomaticColors, bool aNoBackground, bool aNoForeground, int aMarkerSize)
		{
			FFgColor =aFgColor;
			FBgColor = aBgColor;
			FMarkerType = aMarkerType;
			FAutomaticColors = aAutomaticColors;
			FNoBackground = aNoBackground;
			FNoForeground = aNoForeground;
			FMarkerSize = aMarkerSize;
		}


		/// <summary>
		/// Color of the marker lines.
		/// </summary>
		public Color FgColor {get {return FFgColor;} set{FFgColor = value;}}

		/// <summary>
		/// Color of the marker fill.
		/// </summary>
		public Color BgColor {get {return FBgColor;} set{FBgColor = value;}}

		/// <summary>
		/// Type of marker.
		/// </summary>
		public TChartMarkerType MarkerType {get {return FMarkerType;} set{FMarkerType = value;}}

		/// <summary>
		/// When true, marker colors are assigned automatically.
		/// </summary>
		public bool AutomaticColors {get {return FAutomaticColors;} set{FAutomaticColors = value;}}

		/// <summary>
		/// When true, the marker has no fill.
		/// </summary>
		public bool NoBackground {get {return FNoBackground;} set{FNoBackground = value;}}

		/// <summary>
		/// When true, the marker has no lines.
		/// </summary>
		public bool NoForeground {get {return FNoForeground;} set{FNoForeground = value;}}

		/// <summary>
		/// Size of the marker.
		/// </summary>
		public int MarkerSize {get {return FMarkerSize;} set{FMarkerSize = value;}}

		#region ICloneable Members
		/// <summary>
		/// Returns a deep copy of this object.
		/// </summary>
		/// <returns></returns>
		public object Clone()
		{
			return MemberwiseClone();
		}

		#endregion
	}

	/// <summary>
	/// Misc options for the whole series or a point in the series that do no enter in any other category.
	/// </summary>
	public class ChartSeriesMiscOptions: ICloneable
	{
		private bool FSmoothedLines;
		private bool FBubbles3D;
		private bool FHasShadow;

		/// <summary>
		/// Creates a new ChartSeriesMiscOptions instance.
		/// </summary>
		/// <param name="aBubbles3D">Draw bubbles with 3D effects.</param>
		/// <param name="aHasShadow">Series has shadow.</param>
		/// <param name="aSmoothedLines">Lines should be smoothed (line and scatter charts).</param>
		public ChartSeriesMiscOptions(bool aSmoothedLines, bool aBubbles3D, bool aHasShadow)
		{
			FSmoothedLines = aSmoothedLines;
			FBubbles3D = aBubbles3D;
			FHasShadow = aHasShadow;
		}

		/// <summary>
		/// Lines should be smoothed (line and scatter charts).
		/// </summary>
		public bool SmoothedLines {get {return FSmoothedLines;} set{FSmoothedLines = value;}}

		/// <summary>
		/// Draw bubbles with 3D effects.
		/// </summary>
		public bool Bubbles3D {get {return FBubbles3D;} set{FBubbles3D = value;}}
		
		/// <summary>
		/// Series has shadow.
		/// </summary>
		public bool HasShadow {get {return FHasShadow;} set{FHasShadow = value;}}


		#region ICloneable Members
		/// <summary>
		/// Returns a deep copy of this object.
		/// </summary>
		/// <returns></returns>
		public object Clone()
		{
			return MemberwiseClone();
		}

		#endregion
	}
	#endregion

	#region Series Options List
	/// <summary>
	/// A list of options for the whole series and for specific data points inside these series.
	/// </summary>
	public class TSeriesOptionsList: IEnumerable
#if(FRAMEWORK20 && !DELPHIWIN32)
, IEnumerable<int>
#endif
	{
#if(FRAMEWORK20)
        private Dictionary<int, ChartSeriesOptions> FList = new Dictionary<int, ChartSeriesOptions>();
#else
		private Hashtable FList = new Hashtable();
#endif

		/// <summary>
		/// Creates a new TSeriesOptionsList instance.
		/// </summary>
		public TSeriesOptionsList()
		{
		}

		/// <summary>
		/// Gets or sets the value for a data point or for the whole series (when key = -1).
		/// </summary>
		public ChartSeriesOptions this[int key]
		{
			get
			{
#if(FRAMEWORK20)
				ChartSeriesOptions Result = null;
                if (FList.TryGetValue(key, out Result))
                    return Result;
                return null;
#else
				return (ChartSeriesOptions)FList[key];
#endif
			}
			set
			{
				FList[key] = value;
			}
		}


		/// <summary>
		/// Adds a new option to the list. If the option is null, nothing will be done.
		/// </summary>
		/// <param name="Options">Options to add.</param>
		public void Add(ChartSeriesOptions Options)
		{
			if (Options == null) return;
			FList[Options.PointNumber] = Options;
		}

		/// <summary>
		/// Gets all the values of the series.
		/// </summary>
		/// <returns></returns>
		public ChartSeriesOptions[] GetValues()
		{
			ChartSeriesOptions[] Result = new ChartSeriesOptions[FList.Count];
			FList.Values.CopyTo(Result, 0);
			return Result;
		}

		#region IEnumerable Members

		/// <summary>
		/// Gets the enumerator for this class. 
		/// </summary>
		/// <returns></returns>
		public System.Collections.IEnumerator GetEnumerator()
		{
			return FList.GetEnumerator();
		}

		#endregion

        #region IEnumerable<ChartSeriesOptions> Members
#if (FRAMEWORK20 && !DELPHIWIN32)
        IEnumerator<int> IEnumerable<int>.GetEnumerator()
        {
            return FList.Keys.GetEnumerator();
        }
#endif

        #endregion
    }
	#endregion

	#region Legend Options List
	/// <summary>
	/// A list of options for the legend on the whole series and for specific data points inside these series.
	/// </summary>
	public class TLegendOptionsList: IEnumerable, IEnumerable<int>
	{
        private Dictionary<int, TLegendEntryOptions> FList = new Dictionary<int, TLegendEntryOptions>();

		/// <summary>
		/// Creates a new TLegendOptionsList instance.
		/// </summary>
		public TLegendOptionsList()
		{
		}

		/// <summary>
		/// Gets or sets the value for a data point or for the whole series (when key = -1).
		/// </summary>
		public TLegendEntryOptions this[int key]
		{
			get
			{
#if(FRAMEWORK20)
				TLegendEntryOptions Result = null;
                if (FList.TryGetValue(key, out Result))
                    return Result;
                return null;
#else
				return (TLegendEntryOptions)FList[key];
#endif
			}
			set
			{
				FList[key] = value;
			}
		}


		/// <summary>
		/// Adds a new option to the list. If the option is null, nothing will be done.
		/// </summary>
		/// <param name="PointNumber">Point where this value applies. -1 means the whole series.</param>
		/// <param name="Options">Options to add.</param>
		public void Add(int PointNumber, TLegendEntryOptions Options)
		{
			if (Options == null) return;
			FList[PointNumber] = Options;
		}

		/// <summary>
		/// Gets all the values of the series.
		/// </summary>
		/// <returns></returns>
		public TLegendEntryOptions[] GetValues()
		{
			TLegendEntryOptions[] Result = new TLegendEntryOptions[FList.Count];
			FList.Values.CopyTo(Result, 0);
			return Result;
		}

		#region IEnumerable Members

		/// <summary>
		/// Gets the enumerator for this class. 
		/// </summary>
		/// <returns></returns>
		public System.Collections.IEnumerator GetEnumerator()
		{
			return FList.GetEnumerator();
		}

		#endregion

        #region IEnumerable<TLegendEntryOptions> Members
#if (FRAMEWORK20 && !DELPHIWIN32)
        IEnumerator<int> IEnumerable<int>.GetEnumerator()
        {
            return FList.Keys.GetEnumerator();
        }
#endif
        #endregion
    }
	#endregion

	#endregion

	#region ExcelChart
	/// <summary>
	/// Information for a chart inside a sheet or an object.
	/// </summary>
	public abstract class ExcelChart: IEmbeddedObjects
	{	
		#region Background
		/// <summary>
		/// Returns the chart background if there is one defined, or null if there is none.
		/// </summary>
		public abstract TChartFrameOptions Background{get;}

		/// <summary>
		/// Returns the default font for all text in the chart that do not have a font defined.
		/// </summary>
		public abstract TFlxChartFont DefaultFont{get;}

		/// <summary>
		/// Returns the default font for all labels in the chart that do not have a font defined.
		/// </summary>
		public abstract TFlxChartFont DefaultLabelFont{get;}

		/// <summary>
		/// Returns the default font for Axis in the chart that do not have a font defined.
		/// </summary>
		public abstract TFlxChartFont DefaultAxisFont{get;}
		#endregion

		#region Options
		/// <summary>
		/// Returns the type of chart and the options specific for that kind of chart.
		/// There might be more than one ChartOptions returned, since you can mix more than one type of 
		/// chart on a simple chart. (One for each series). You need to look at the series ChartOptionsIndex to 
		/// know to which one it refers.
		/// </summary>
		public abstract TChartOptions[] ChartOptions{get;}

		/// <summary>
		/// Defines how null cells will be plotted on the chart.
		/// </summary>
		public abstract TPlotEmptyCells PlotEmptyCells{get; set;}

		#endregion

		#region Series
		/// <summary>
		/// Returns a series definition.
		/// </summary>
		/// <param name="index">Index of the series you want to return.</param>
		/// <param name="getDefinitions">If false, this method will not return the series formulas, so it will be a little faster.</param>
		/// <param name="getValues">If false, this method will not return the series values, so it will be a little faster and use less memory.</param>
		/// <param name="getOptions">If false, this method will not return the series options.</param>
		/// <returns>series description.</returns>
		public abstract ChartSeries GetSeries(int index, bool getDefinitions, bool getValues, bool getOptions);

		/// <summary>
		/// Sets a Serie value.
		/// </summary>
		/// <param name="index">Index of the serie to set.</param>
		/// <param name="value">Series definition.</param>
		public abstract void SetSeries(int index, ChartSeries value);

		/// <summary>
		/// Adds a series to the chart.
		/// </summary>
		/// <param name="value">Definition of the new series</param>
		/// <returns>Index of the newly added series.</returns>
		public abstract int AddSeries(ChartSeries value);

		/// <summary>
		/// Deletes the series at position index.
		/// </summary>
		/// <param name="index"></param>
		public abstract void DeleteSeries(int index);

		/// <summary>
		/// Returns the count of series on this chart.
		/// </summary>
		public abstract int SeriesCount	{get;}
		#endregion

		#region Axis
		/// <summary>
		/// Returns the axis information for this chart. Note that this might be more than one, if the chart has a secondary axis.
		/// </summary>
		/// <returns></returns>
        public abstract TChartAxis[] GetChartAxis();
		#endregion

		#region Legend
		/// <summary>
		/// Information about the Legend of the chart.
		/// </summary>
		/// <returns></returns>
        public abstract TChartLegend GetChartLegend();
		#endregion

		#region Labels
		/// <summary>
		/// Returns all the labels for the chart and the main title. Note that Axis have their labels defined inside their own definition.
		/// </summary>
		/// <returns>Label values.</returns>
		public abstract TDataLabel[] GetDataLabels();

        /// <summary>
        /// Changes the labels for the chart. You should always get the values with <see cref="GetDataLabels"/>,
        /// modify them, and change them back with this method.
        /// </summary>
        /// <param name="labels">New labels for the chart.</param>
        public abstract void SetDataLabels(TDataLabel[] labels);
		#endregion

		#region Objects
		/// <summary>
		/// The number of objects that are embedded inside this chart.
		/// </summary>
		public abstract int ObjectCount{get;}

		/// <summary>
		/// An object embedded inside a chart.
		/// </summary>
		/// <param name="objectIndex">Index of the object, between 1 and <see cref="ObjectCount"/></param>
		/// <returns>The properties for the embedded object.</returns>
		/// <param name="GetShapeOptions">When true, shape options will be retrieved. As this can be a slow operation,
		/// only specify true when you really need those options.</param>
		public abstract TShapeProperties GetObjectProperties(int objectIndex, bool GetShapeOptions);

		/// <summary>
		/// Changes the text inside an object of the chart.
		/// </summary>
		/// <param name="objectIndex">Index of the object, between 1 and <see cref="ObjectCount"/></param>
		/// <param name="objectPath">Index to the child object you want to change the text.
        /// If it is a simple object, you can use String.Empty here, if not you need to get the ObjectPath from <see cref="GetObjectProperties"/><br></br>
        /// If it is "absolute"(it starts with "\\"), then the path includes the objectIndex, and the objectIndex is
        /// not used. An object path of "\\1\\2\\3" is exactly the same as using objectIndex = 1 and objectPath = "2\\3"</param>
		/// <param name="text">Text you want to use. Use null to delete text from an AutoShape.</param>
		public abstract void SetObjectText(int objectIndex, string objectPath, TRichString text);

		/// <summary>
		/// Deletes the graphic object at objectIndex. Use it with care, there are some graphics objects you
		/// <b>don't</b> want to remove (like comment boxes when you don't delete the associated comment.)
		/// </summary>
		/// <param name="objectIndex">Index of the object (1 based).</param>
		public abstract void DeleteObject(int objectIndex);	


		#endregion

	}
	#endregion

	#region ChartOptions

	/// <summary>
	/// Interface for charts that can be stacked.
	/// </summary>
	public interface IStackedOptions
	{
		/// <summary>
		/// Stacked mode for a series.
		/// </summary>
		TStackedMode StackedMode{get;set;}
	}
	
	/// <summary>
	/// Base class for options specific to the type of chart.
	/// </summary>
	public abstract class TChartOptions: IComparable, IComparable<TChartOptions>
	{
		private TChartType FChartType;
		private bool FChangeColorsOnEachSeries;
		private int FZOrder;
		private int FAxisNumber;
		private ChartSeriesOptions FSeriesOptions;
		private TChartPlotArea FChartPlotArea;
		private TDataLabel FDefaultLabel;

		/// <summary>
		/// Creates a new TChartOptions instance.
		/// </summary>
		/// <param name="aChartType">Type of the chart.</param>
		/// <param name="aAxisNumber">Axis where this chart group belongs, 0 is primary, 1 is secondary.</param>
		/// <param name="aChangeColorsOnEachSeries">If false, all series will be the same color.</param>
		/// <param name="aZOrder">Z-Order of this chart group, with 0 being the bottom. Chart groups with lower z-Order are drawn below the ones with higher ones.</param>
		/// <param name="aSeriesOptions">Global options that apply to all the series on this group. This instance will be copied.</param>
		/// <param name="aPlotArea">Plot area line style and fill settings. This value will be copied.</param>
		/// <param name="aDefaultLabel">Default label properties for labels on this series group.</param>
		protected TChartOptions(TChartType aChartType, bool aChangeColorsOnEachSeries, int aZOrder, int aAxisNumber, ChartSeriesOptions aSeriesOptions, 
			TChartPlotArea aPlotArea, TDataLabel aDefaultLabel)
		{
			FChartType = aChartType;
			FChangeColorsOnEachSeries = aChangeColorsOnEachSeries;
			FZOrder = aZOrder;
			FAxisNumber = aAxisNumber;
			if (aSeriesOptions != null) FSeriesOptions = (ChartSeriesOptions)aSeriesOptions.Clone();
			if (aPlotArea != null) FChartPlotArea = (TChartPlotArea)aPlotArea.Clone();
			if (aDefaultLabel != null) DefaultLabel = (TDataLabel)aDefaultLabel.Clone();
		}

		/// <summary>
		/// Chart Type.
		/// </summary>
		public TChartType ChartType {get {return FChartType;}}

		/// <summary>
		/// If false, all series will be the same color.
		/// </summary>
		public bool ChangeColorsOnEachSeries {get {return FChangeColorsOnEachSeries;} set{FChangeColorsOnEachSeries = value;}}

		/// <summary>
		/// Z-Order of this chart group, with 0 being the bottom. Chart groups with lower z-Order are drawn below the ones with higher ones.
		/// </summary>
		public int ZOrder {get {return FZOrder;} set{FZOrder = value;}}

		/// <summary>
		/// Axis where this chart group belongs, 0 is primary, 1 is secondary.
		/// </summary>
		public int AxisNumber {get {return FAxisNumber;} set{FAxisNumber = value;}}

		/// <summary>
		/// Global options for all the series on this chart group.
		/// </summary>
		public ChartSeriesOptions SeriesOptions {get {return FSeriesOptions;} set{FSeriesOptions = value;}}
		
		/// <summary>
		/// Plot area for this chart group.
		/// </summary>
		public TChartPlotArea PlotArea {get {return FChartPlotArea;} set{FChartPlotArea = value;}}

		/// <summary>
		/// Default label properties for this group of charts.
		/// </summary>
		public TDataLabel DefaultLabel {get {return FDefaultLabel;} set{FDefaultLabel = value;}}

		
		#region IComparable Members

		/// <summary>
		/// Orders the chart options depending on their z-order.
		/// </summary>
		/// <param name="obj">Object to compare to.</param>
		/// <returns>-1, 0 or 1 depending if the series has a bigger z-order than obj.</returns>
		public int CompareTo(TChartOptions obj)
		{
			if (obj == null) return -1;
			return ZOrder.CompareTo(obj.ZOrder);
		}

        /// <summary></summary>
        public int CompareTo(object obj)
        {
            return CompareTo(obj as TChartOptions);
        }

		/// <summary>
		/// Returns true if both objects are equal.
		/// </summary>
		/// <param name="obj"></param>
		/// <returns></returns>
		public override bool Equals(object obj)
		{
			return CompareTo(obj as TChartOptions) == 0;
		}

        /// <summary>
        /// Returns true if both objects are equal. 
        /// Note this is for backwards compatibility, this is a class and not immutable,
        /// so this method should return true if references are different. But that would break old code.
        /// </summary>
        /// <param name="o1"></param>
        /// <param name="o2"></param>
        /// <returns></returns>
        public static bool operator==(TChartOptions o1, TChartOptions o2)
        {
            if ((object)o1 == null) return (object)o2 == null;
            return o1.Equals(o2);
        }

        /// <summary>
        /// Returns true if both objects do not have the same value.
        /// Note this is for backwards compatibility, this is a class and not immutable,
        /// so this method should return true if references are different. But that would break old code.
        /// </summary>
        /// <param name="s1"></param>
        /// <param name="s2"></param>
        /// <returns></returns>
        public static bool operator!=(TChartOptions s1, TChartOptions s2)
        {
            if ((object)s1 == null) return (object)s2 != null;
            return !(s1.Equals(s2));
        }

        /// <summary>
        /// Returns true if o1 is bigger than o2.
        /// </summary>
        /// <param name="o1"></param>
        /// <param name="o2"></param>
        /// <returns></returns>
        public static bool operator >(TChartOptions o1, TChartOptions o2)
        {
            if ((object)o1 == null)
            {
                if ((object)o2 == null) return false;
                return true;
            }
            if ((object)o2 == null) return true;
            return o1.CompareTo(o2) > 0;
        }

        /// <summary>
        /// Returns true if o1 is less than o2.
        /// </summary>
        /// <param name="o1"></param>
        /// <param name="o2"></param>
        /// <returns></returns>
        public static bool operator <(TChartOptions o1, TChartOptions o2)
        {
            if ((object)o1 == null)
            {
                if ((object)o2 == null) return true;
                return true;
            }
            if ((object)o2 == null) return false;
            return o1.CompareTo(o2) < 0;
        }

		/// <summary>
		/// Returns a hashcode for the object.
		/// </summary>
		/// <returns></returns>
		public override int GetHashCode()
		{
			return ZOrder.GetHashCode ();
		}



		#endregion
	}

	/// <summary>
	/// Options for an unknown chart. This class does not have any information except the charttype.
	/// </summary>
	public class TUnknownChartOptions: TChartOptions
	{
		/// <summary>
		/// Creates a new TUnknownChartOptions instance.
		/// </summary>
		public TUnknownChartOptions(): base(TChartType.Unknown, false, -1, 0, null, null, null){}
	}

	/// <summary>
	/// Options for a Bar or Column chart.
	/// </summary>
	public class TBarChartOptions: TChartOptions, IStackedOptions
	{
		private float FBarOverlap;
		private float FCategoriesGap;
		private bool FHorizontal;
		private TStackedMode FStackedMode;
		private ChartLineOptions FSeriesLines;
		private bool FHasShadow;

		/// <summary>
		/// Creates a new TBarChartOptions instance.
		/// </summary>
		/// <param name="aChangeColorsOnEachSeries">If false, all series will be the same color.</param>
		/// <param name="aZOrder">Z-Order of this chart group, with 0 being the bottom. Chart groups with lower z-Order are drawn below the ones with higher ones.</param>
		/// <param name="aBarOverlap">Space between bars in percent of the bar width.</param>
		/// <param name="aCategoriesOverlap">Space between categories in percent of bar width.</param>
		/// <param name="aHorizontal">If true, bars are horizontal and this is a bar chart. If false, bars are vertical and this is a column chart.</param>
		/// <param name="aStackedMode"><see cref="TStackedMode"/> of the chart.</param>
		/// <param name="aHasShadow">True if the bars have shadows.</param>
		/// <param name="aAxisNumber">Axis where this chart group belongs, 0 is primary, 1 is secondary.</param>
		/// <param name="aSeriesOptions">Global options that apply to all the series on this group. This instance will be copied.</param>
		/// <param name="aPlotArea">Plot area fill and line style for this group.</param>
		/// <param name="aDefaultLabel">Default label properties for labels in this group.</param>
		/// <param name="aSeriesLines">See <see cref="SeriesLines"/></param>
		public TBarChartOptions(int aAxisNumber, bool aChangeColorsOnEachSeries, int aZOrder, float aBarOverlap, float aCategoriesOverlap, bool aHorizontal, TStackedMode aStackedMode, bool aHasShadow, ChartSeriesOptions aSeriesOptions, TChartPlotArea aPlotArea, TDataLabel aDefaultLabel, ChartLineOptions aSeriesLines): 
			base(TChartType.Bar, aChangeColorsOnEachSeries, aZOrder, aAxisNumber, aSeriesOptions, aPlotArea, aDefaultLabel)
		{
			FBarOverlap = aBarOverlap;
			FCategoriesGap = aCategoriesOverlap;
			FHorizontal = aHorizontal;
			FStackedMode = aStackedMode;
			FHasShadow = aHasShadow;
			FSeriesLines = aSeriesLines;
		}
		
		/// <summary>
		/// Space between bars in percent of the bar width.
		/// </summary>
		public float BarOverlap {get {return FBarOverlap;} set{FBarOverlap = value;}}

		/// <summary>
		/// Space between categories in percent of bar width.
		/// </summary>
		public float CategoriesGap {get {return FCategoriesGap;} set{FCategoriesGap = value;}}

		/// <summary>
		/// If true, bars are horizontal and this is a bar chart. If false, bars are vertical and this is a column chart.
		/// </summary>
		public bool Horizontal {get {return FHorizontal;} set{FHorizontal = value;}}

		/// <summary>
		/// <see cref="TStackedMode"/> of the chart.
		/// </summary>
		public TStackedMode StackedMode {get {return FStackedMode;} set{FStackedMode = value;}}

		/// <summary>
		/// True if the bars have shadows.
		/// </summary>
		public bool HasShadow {get {return FHasShadow;} set{FHasShadow = value;}}

		/// <summary>
		/// Line style for the Lines between Series if they exist, null otherwise.
		/// </summary>
		public ChartLineOptions SeriesLines {get {return FSeriesLines;} set{FSeriesLines = value;}}

	}

	/// <summary>
	/// Options specific for a Line or Area chart.
	/// </summary>
	public class TAreaLineChartOptions: TChartOptions, IStackedOptions
	{
		private TStackedMode FStackedMode;
		private bool FHasShadow;
		private TChartDropBars FDropBars;

		/// <summary>
		/// Creates a new TAreaLineChartOptions instance.
		/// </summary>
		/// <param name="aStackedMode"><see cref="TStackedMode"/> of the chart.</param>
		/// <param name="aHasShadow">True if the chart lines have shadows.</param>
		/// <param name="aChangeColorsOnEachSeries">If false, all series will be the same color.</param>
		/// <param name="aChartType">Type of chart. Must be Area or line.</param>
		/// <param name="aZOrder">Z-Order of this chart group, with 0 being the bottom. Chart groups with lower z-Order are drawn below the ones with higher ones.</param>
		/// <param name="aAxisNumber">Axis where this chart group belongs, 0 is primary, 1 is secondary.</param>
		/// <param name="aSeriesOptions">Global options that apply to all the series on this group. This instance will be copied.</param>
		/// <param name="aPlotArea">Plot area fill and line style for this group.</param>
		/// <param name="aDropBars">Drop Bars, lines and Hi/lo information. This value will be copied.</param>
		/// <param name="aDefaultLabel">Default label properties for labels in this group.</param>
		public TAreaLineChartOptions(TChartType aChartType, int aAxisNumber, bool aChangeColorsOnEachSeries, int aZOrder, TStackedMode aStackedMode, bool aHasShadow, ChartSeriesOptions aSeriesOptions, TChartPlotArea aPlotArea, TChartDropBars aDropBars, TDataLabel aDefaultLabel): 
			base(aChartType, aChangeColorsOnEachSeries, aZOrder, aAxisNumber, aSeriesOptions, aPlotArea, aDefaultLabel)
		{
			FStackedMode = aStackedMode;
			FHasShadow = aHasShadow;
			if (aDropBars != null) FDropBars = (TChartDropBars)aDropBars.Clone();
		}
		
		/// <summary>
		/// <see cref="TStackedMode"/> of the chart.
		/// </summary>
		public TStackedMode StackedMode {get {return FStackedMode;} set{FStackedMode = value;}}

		/// <summary>
		/// True if the chart lines have shadows.
		/// </summary>
		public bool HasShadow {get {return FHasShadow;} set{FHasShadow = value;}}

		/// <summary>
		/// If the chart group has drop lines, the line information. Null otherwise.
		/// </summary>
		public TChartDropBars DropBars {get {return FDropBars;} set{FDropBars = value;}}
	}

	/// <summary>
	/// Options specific for a Line chart.
	/// </summary>
	public class TLineChartOptions: TAreaLineChartOptions
	{
		/// <summary>
		/// Creates a new TLineChartOptions instance.
		/// </summary>
		/// <param name="aStackedMode"><see cref="TStackedMode"/> of the chart.</param>
		/// <param name="aHasShadow">True if the chart lines have shadows.</param>
		/// <param name="aChangeColorsOnEachSeries">If false, all series will be the same color.</param>
		/// <param name="aZOrder">Z-Order of this chart group, with 0 being the bottom. Chart groups with lower z-Order are drawn below the ones with higher ones.</param>
		/// <param name="aAxisNumber">Axis where this chart group belongs, 0 is primary, 1 is secondary.</param>
		/// <param name="aSeriesOptions">Global options that apply to all the series on this group. This instance will be copied.</param>
		/// <param name="aPlotArea">Plot area fill and line style for this group.</param>
		/// <param name="aDropBars">Drop bars and lines information. This value will be copied.</param>
		/// <param name="aDefaultLabel">Default label properties for labels in this group.</param>
		public TLineChartOptions(int aAxisNumber, bool aChangeColorsOnEachSeries, int aZOrder, TStackedMode aStackedMode, bool aHasShadow, ChartSeriesOptions aSeriesOptions, TChartPlotArea aPlotArea, TChartDropBars aDropBars, TDataLabel aDefaultLabel): 
			base(TChartType.Line, aAxisNumber, aChangeColorsOnEachSeries, aZOrder, aStackedMode, aHasShadow, aSeriesOptions, aPlotArea, aDropBars, aDefaultLabel)
		{
		}
	}

	/// <summary>
	/// Options specific for an Area chart.
	/// </summary>
	public class TAreaChartOptions: TAreaLineChartOptions
	{
		/// <summary>
		/// Creates a new TAreaChartOptions instance.
		/// </summary>
		/// <param name="aStackedMode"><see cref="TStackedMode"/> of the chart.</param>
		/// <param name="aHasShadow">True if the chart lines have shadows.</param>
		/// <param name="aChangeColorsOnEachSeries">If false, all series will be the same color.</param>
		/// <param name="aZOrder">Z-Order of this chart group, with 0 being the bottom. Chart groups with lower z-Order are drawn below the ones with higher ones.</param>
		/// <param name="aAxisNumber">Axis where this chart group belongs, 0 is primary, 1 is secondary.</param>
		/// <param name="aSeriesOptions">Global options that apply to all the series on this group. This instance will be copied.</param>
		/// <param name="aPlotArea">Plot area fill and line style for this group.</param>
		/// <param name="aDropBars">Drop bars and lines information. This value will be copied.</param>
		/// <param name="aDefaultLabel">Default label properties for labels in this group.</param>
		public TAreaChartOptions(int aAxisNumber, bool aChangeColorsOnEachSeries, int aZOrder, TStackedMode aStackedMode, bool aHasShadow, ChartSeriesOptions aSeriesOptions, TChartPlotArea aPlotArea, TChartDropBars aDropBars, TDataLabel aDefaultLabel): 
			base(TChartType.Area, aAxisNumber, aChangeColorsOnEachSeries, aZOrder, aStackedMode, aHasShadow, aSeriesOptions, aPlotArea, aDropBars, aDefaultLabel)
		{
		}
	}

	/// <summary>
	/// Options specific for a Pie chart.
	/// </summary>
	public class TPieChartOptions: TChartOptions
	{
		private int FFirstSliceAngle;
		private int FDonutRadius;
		private bool FHasShadow;
		private bool FLeaderLines;
		private ChartLineOptions FLeaderLineStyle;

		/// <summary>
		/// Creates a new TPieChartOptions instance.
		/// </summary>
		/// <param name="aChangeColorsOnEachSeries">If false, all series will be the same color.</param>
		/// <param name="aZOrder">Z-Order of this chart group, with 0 being the bottom. Chart groups with lower z-Order are drawn below the ones with higher ones.</param>
		/// <param name="aFirstSliceAngle">Angle of the first slice in degrees. It can go from 0 to 359.</param>
		/// <param name="aDonutRadius">Radius of the center of the donut in Percet. 0 Means a Pie without hole.</param>
		/// <param name="aHasShadow">True if the chart lines have shadows.</param>
		/// <param name="aLeaderLines">True if there are lines from the slices to the labels.</param>
		/// <param name="aAxisNumber">Axis where this chart group belongs, 0 is primary, 1 is secondary.</param>
		/// <param name="aSeriesOptions">Global options that apply to all the series on this group. This instance will be copied.</param>
		/// <param name="aPlotArea">Plot area fill and line style for this group.</param>
		/// <param name="aLeaderLineStyle">Line style for the leader lines, only has meaning if <see cref="LeaderLines"/> is true. This value will be copied.</param>
		/// <param name="aDefaultLabel">Default label properties for labels in this group.</param>
		public TPieChartOptions(int aAxisNumber, bool aChangeColorsOnEachSeries, int aZOrder, int aFirstSliceAngle, int aDonutRadius, bool aHasShadow, bool aLeaderLines, ChartLineOptions aLeaderLineStyle, ChartSeriesOptions aSeriesOptions, TChartPlotArea aPlotArea, TDataLabel aDefaultLabel): 
			base(TChartType.Pie, aChangeColorsOnEachSeries, aZOrder, aAxisNumber, aSeriesOptions, aPlotArea, aDefaultLabel)
		{
			FFirstSliceAngle = aFirstSliceAngle;
			FDonutRadius = aDonutRadius;
			FHasShadow = aHasShadow;
			FLeaderLines = aLeaderLines;
			if (aLeaderLineStyle != null) FLeaderLineStyle = (ChartLineOptions) aLeaderLineStyle.Clone();
		}
		
		/// <summary>
		/// Angle of the first slice in degrees. It can go from 0 to 359.
		/// </summary>
		public int FirstSliceAngle {get {return FFirstSliceAngle;} set{FFirstSliceAngle = value % 360;}}

		/// <summary>
		/// Radius of the center of the donut in Percet. 0 Means a Pie without hole.
		/// </summary>
		public int DonutRadius {get {return FDonutRadius;} set{FDonutRadius = value;}}

		/// <summary>
		/// True if the chart lines have shadows.
		/// </summary>
		public bool HasShadow {get {return FHasShadow;} set{FHasShadow = value;}}
	
		/// <summary>
		/// True if there are lines from the slices to the labels.
		/// </summary>
		public bool LeaderLines {get {return FLeaderLines;} set{FLeaderLines = value;}}

		/// <summary>
		/// Line style for the leader lines, only has smeaning if <see cref="LeaderLines"/> is true.
		/// </summary>
		public ChartLineOptions LeaderLineStyle {get {return FLeaderLineStyle;} set{FLeaderLineStyle = value;}}

	}

	/// <summary>
	/// What the bubble size means.
	/// </summary>
	public enum TBubbleSizeType
	{
        /// <summary>
        /// Not used.
        /// </summary>
        None = 0,

		/// <summary>
		/// Bubble size represents the area.
		/// </summary>
		BubbleSizeIsArea = 1,

		/// <summary>
		/// Bubble size represents the width.
		/// </summary>
		BubbleSizeIsWidth = 2
	}

	/// <summary>
	/// Options specific for a Scatter/Bubble chart.
	/// </summary>
	public class TScatterChartOptions: TChartOptions
	{
		private int FBubblePercentRatio;
		private TBubbleSizeType FBubbleSizeType;
		private bool FIsBubbleChart;
		private bool FShowNegativeBubbles;
		private bool FBubblesHaveShadow;

		/// <summary>
		/// Creates a new TScatterChartOptions instance.
		/// </summary>
		/// <param name="aAxisNumber">See <see cref="TChartOptions.AxisNumber"/></param>
		/// <param name="aChangeColorsOnEachSeries">See <see cref="TChartOptions.ChangeColorsOnEachSeries"/></param>
		/// <param name="aZOrder">See <see cref="TChartOptions.ZOrder"/></param>
		/// <param name="aSeriesOptions">See <see cref="TChartOptions.SeriesOptions"/></param>
		/// <param name="aPlotArea">See <see cref="TChartOptions.PlotArea"/></param>
		/// <param name="aBubblePercentRatio">See <see cref="BubblePercentRatio"/></param>
		/// <param name="aBubbleSizeType">See <see cref="BubbleSizeType"/></param>
		/// <param name="aIsBubbleChart">See <see cref="IsBubbleChart"/></param>
		/// <param name="aShowNegativeBubbles">See <see cref="ShowNegativeBubbles"/></param>
		/// <param name="aBubblesHaveShadow">See <see cref="BubblesHaveShadow"/></param>
		/// <param name="aDefaultLabel">Default label properties for labels in this group.</param>
		public TScatterChartOptions(int aAxisNumber, bool aChangeColorsOnEachSeries, int aZOrder, ChartSeriesOptions aSeriesOptions, TChartPlotArea aPlotArea,
			int aBubblePercentRatio, TBubbleSizeType aBubbleSizeType, bool aIsBubbleChart, bool aShowNegativeBubbles, bool aBubblesHaveShadow, TDataLabel aDefaultLabel): 
			base(TChartType.Scatter, aChangeColorsOnEachSeries, aZOrder, aAxisNumber, aSeriesOptions, aPlotArea, aDefaultLabel)
		{
			FBubblePercentRatio = aBubblePercentRatio;
			FBubbleSizeType = aBubbleSizeType;
			FIsBubbleChart = aIsBubbleChart;
			FShowNegativeBubbles = aShowNegativeBubbles;
			FBubblesHaveShadow = aBubblesHaveShadow;
		}
		
		/// <summary>
		/// Percent of largest bubble compared to chart in general.
		/// </summary>
		public int BubblePercentRatio {get {return FBubblePercentRatio;} set{FBubblePercentRatio = value;}}

		/// <summary>
		/// What the bubble size represents.
		/// </summary>
		public TBubbleSizeType BubbleSizeType {get {return FBubbleSizeType;} set{FBubbleSizeType = value;}}

		/// <summary>
		/// True if this is a bubble chart, false if it is a scatter chart.
		/// </summary>
		public bool IsBubbleChart {get {return FIsBubbleChart;} set{FIsBubbleChart = value;}}

		/// <summary>
		/// True if negative bubbles should be shwown.
		/// </summary>
		public bool ShowNegativeBubbles {get {return FShowNegativeBubbles;} set{FShowNegativeBubbles = value;}}

		/// <summary>
		/// True if the bubbles have shadows.
		/// </summary>
		public bool BubblesHaveShadow {get {return FBubblesHaveShadow;} set{FBubblesHaveShadow = value;}}
	}

	#endregion

	#region TFrameOptions
	/// <summary>
	/// Description of a box, its coordinates, fill style and line style.
	/// </summary>
	public class TChartFrameOptions: ICloneable
	{
		#region Privates
		private ChartLineOptions FLineOptions;
		private ChartFillOptions FFillOptions;
	
		private TShapeOptionList FExtraOptions;

		#endregion

		/// <summary>
		/// Creates a new TChartFrameOptions instance.
		/// </summary>
		/// <param name="aLineOptions">See <see cref="LineOptions"/></param>
		/// <param name="aFillOptions">See <see cref="FillOptions"/></param>
		/// <param name="aExtraOptions">See <see cref="ExtraOptions"/></param>
		public TChartFrameOptions(ChartLineOptions aLineOptions, ChartFillOptions aFillOptions, TShapeOptionList aExtraOptions)
		{
			FLineOptions = aLineOptions;
			FFillOptions = aFillOptions;
			FExtraOptions = aExtraOptions;
		}

		#region Properties
		/// <summary>
		/// Fill options for the frame.
		/// </summary>
		public ChartFillOptions FillOptions {get {return FFillOptions;} set{FFillOptions = value;}}

		/// <summary>
		/// Line options for the frame.
		/// </summary>
		public ChartLineOptions LineOptions {get {return FLineOptions;} set{FLineOptions = value;}}

		/// <summary>
		/// Extra options used to specify gradients, textures, etc. If this is no null, the color values are not used.
		/// </summary>
		public TShapeOptionList ExtraOptions {get {return FExtraOptions;} set{FExtraOptions = value;}}

		#endregion

		#region ICloneable Members

		/// <summary>
		/// Returns a deep copy of this object.
		/// </summary>
		/// <returns></returns>
		public object Clone()
		{
			ChartLineOptions  NewLineOptions = FLineOptions==null? null: (ChartLineOptions) FLineOptions.Clone();
			ChartFillOptions  NewFillOptions = FFillOptions==null? null: (ChartFillOptions) FFillOptions.Clone();
			TShapeOptionList  NewExtraOptions = FExtraOptions==null? null: (TShapeOptionList) FExtraOptions.Clone();
			return new TChartFrameOptions(NewLineOptions, NewFillOptions, NewExtraOptions);
		}

		#endregion
	}
	#endregion

	#region TextOptions
    /// <summary>
    /// How coordinates are stored for a label when it is manually positioned.
    /// </summary>
    public enum TChartLabelPositionMode
    {
        /// <summary>
        /// Relative position to the chart, in points. 
        /// </summary>
        MDFX = 0x0000,

        /// <summary>
        /// Absolute width and height in points; can only be applied to the mdBotRt field of Pos
        /// </summary>

        MDABS = 0x0001,

        /// <summary>
        /// Owner of Pos determines how to interpret the position data. 
        /// </summary>
        MDPARENT = 0x0002,

        /// <summary>
        /// Offset to default position, in 1/1000 th  of the plot area size. 
        /// </summary>
        MDKTH = 0x0003,

        /// <summary>
        /// Relative position to the chart, in SPRC. 
        /// </summary>
        MDCHART = 0x0005
    }

    /// <summary>
    /// Defines the position of a label in the chart when it is manually positioned.
    /// </summary>
    public class TChartLabelPosition: ICloneable
    {
        private TChartLabelPositionMode FTopLeftMode;
        private TChartLabelPositionMode FBottomRightMode;
        private int FX1;
        private int FY1;
        private int FX2;
        private int FY2;

        /// <summary>
        /// Creates a new TChartLabelPosition class.
        /// </summary>
        /// <param name="aTopLeftMode">See <see cref="TopLeftMode" /></param>
        /// <param name="aBottomRightMode">See <see cref="BottomRightMode" /></param>
        /// <param name="aX1">See <see cref="X1" /></param>
        /// <param name="aY1">See <see cref="Y1" /></param>
        /// <param name="aX2">See <see cref="X2" /></param>
        /// <param name="aY2">See <see cref="Y2" /></param>
        public TChartLabelPosition(TChartLabelPositionMode aTopLeftMode, TChartLabelPositionMode aBottomRightMode, int aX1, int aY1, int aX2, int aY2)
        {
            TopLeftMode = aTopLeftMode;
            BottomRightMode = aBottomRightMode;
            X1 = aX1;
            Y1 = aY1;
            X2 = aX2;
            Y2 = aY2;
        }

        /// <summary>
        /// This value specifies how the top left coordinates (X1, Y1) are stored and what they mean.
        /// </summary>
        public TChartLabelPositionMode TopLeftMode { get { return FTopLeftMode; } set { FTopLeftMode = value; } }

        /// <summary>
        /// This value specifies how the bottom right coordinates (X2, Y2) are stored and what they mean.
        /// </summary>
        public TChartLabelPositionMode BottomRightMode { get { return FBottomRightMode; } set { FBottomRightMode = value; } }

        /// <summary>
        /// Left coordinate. The meaning of this property depends in <see cref="TopLeftMode"/>
        /// </summary>
        public int X1 { get { return FX1; } set { FX1 = value; } }

        /// <summary>
        /// Top coordinate. The meaning of this property depends in <see cref="TopLeftMode"/>
        /// </summary>
        public int Y1 { get { return FY1; } set { FY1 = value; } }

        /// <summary>
        /// Right coordinate. The meaning of this property depends in <see cref="BottomRightMode"/>
        /// </summary>
        public int X2 { get { return FX2; } set { FX2 = value; } }

        /// <summary>
        /// Bottom coordinate. The meaning of this property depends in <see cref="BottomRightMode"/>
        /// </summary>
        public int Y2 { get { return FY2; } set { FY2 = value; } }


        #region ICloneable Members

        /// <summary>
        /// Returns a deep clone of this object.
        /// </summary>
        /// <returns></returns>
        public object Clone()
        {
            return MemberwiseClone();
        }

        #endregion
    }

	/// <summary>
	/// Options for text inside the chart.
	/// </summary>
	public class TChartTextOptions: ICloneable
	{
		private long FX;
		private long FY;
		private long FWidth;
		private long FHeight;

		private TFlxChartFont FFont;
		private THFlxAlignment FHAlign;
		private TVFlxAlignment FVAlign;
		private TBackgroundMode FBackgoundMode;
		private ChartFillOptions FTextColor;

		private byte FRotation;


		/// <summary>
		/// Font style for the text.
		/// </summary>
		public TFlxChartFont Font {get {return FFont;} set{FFont = value;}}

		/// <summary>
		/// Color of the text.
		/// </summary>
		public ChartFillOptions TextColor {get {return FTextColor;} set{FTextColor = value;}}

		/// <summary>
		/// Background mode, transparent or opaque.
		/// </summary>
		public TBackgroundMode BackgoundMode {get {return FBackgoundMode;} set{FBackgoundMode = value;}}

		/// <summary>
		/// Horizontal alginment for the text.
		/// </summary>
		public THFlxAlignment HAlign {get {return FHAlign;} set{FHAlign = value;}}

		/// <summary>
		/// Vertical alignment for the text.
		/// </summary>
		public TVFlxAlignment VAlign {get {return FVAlign;} set{FVAlign = value;}}

		/// <summary>
		/// Text Rotation on degrees. 
		/// 0 - 90 is up, 
		/// 91 - 180 is down, 
		/// 255 is vertical.
		/// </summary>
		public byte Rotation {get {return FRotation;} set{FRotation = value;}}

		/// <summary>
		/// X coordinate on 1/4000 units of chart area.
		/// </summary>
		public long X {get {return FX;} set{FX = value;}}

		/// <summary>
		/// Y coordinate on 1/4000 units of chart area.
		/// </summary>
		public long Y {get {return FY;} set{FY = value;}}

        /// <summary>
        /// Position of the label. If this value is null, then X and Y properties in this class are used instead. If this value is not null,
        /// then X and Y have no meaning.
        /// </summary>
        public TChartLabelPosition Position;

		/// <summary>
		/// Hieght of the bounding box, on 1/4000 units of chart size.
		/// </summary>
		public long Height {get {return FHeight;} set{FHeight = value;}}
		
		/// <summary>
		/// Width of the bounding box, on 1/4000 units of chart size.
		/// </summary>
		public long Width {get {return FWidth;} set{FWidth = value;}}

		#region ICloneable Members

		/// <summary>
		/// Returns a deep copy of the object.
		/// </summary>
		/// <returns></returns>
		public object Clone()
		{
			TChartTextOptions Result = (TChartTextOptions)MemberwiseClone();
			if (Font != null) Result.Font = (TFlxChartFont)(Font.Clone());
			if (TextColor != null) Result.TextColor = (ChartFillOptions)TextColor.Clone();
            if (Position != null) Result.Position = (TChartLabelPosition)Position.Clone(); 
			return null;
		}

		#endregion
	}
	#endregion

	#region PlotArea
	/// <summary>
	/// Line and fill styles for the Plot Area.
	/// </summary>
	public class TChartPlotArea: ICloneable
	{
		#region Privates
		TChartFrameOptions FChartFrameOptions;
		#endregion

		#region Constructors
		/// <summary>
		/// Creates a new PlotArea instance.
		/// </summary>
		/// <param name="aChartFrameOptions">See <see cref="ChartFrameOptions"/></param>
		public TChartPlotArea(TChartFrameOptions aChartFrameOptions)
		{
			FChartFrameOptions = aChartFrameOptions;
		}
		#endregion

		#region Properties
		/// <summary>
		/// Line and fill style for this PlotArea.
		/// </summary>
		public TChartFrameOptions ChartFrameOptions {get {return FChartFrameOptions;} set{FChartFrameOptions = value;}}
		#endregion

		#region ICloneable Members

		/// <summary>
		/// Returns a deep copy of this object.
		/// </summary>
		/// <returns></returns>
		public object Clone()
		{
			TChartFrameOptions NewFrameOptions = null;
			if (FChartFrameOptions != null) NewFrameOptions = (TChartFrameOptions) FChartFrameOptions.Clone();

			return new TChartPlotArea(NewFrameOptions);
		}

		#endregion
	}

	#endregion

	#region Axis
	#region AxisType
	
	/// <summary>
	/// Type of axis.
	/// </summary>
	public enum TAxisType
	{
		/// <summary>
		/// Category Axis.
		/// </summary>
		Category = 0,

		/// <summary>
		/// Value Axis.
		/// </summary>
		Value = 1,

		/// <summary>
		/// Series Axis.
		/// </summary>
		Series =2
	}

	#endregion

	#region AxisLineOptions
	/// <summary>
	/// Line options for an Axis.
	/// </summary>
	public class TAxisLineOptions: ICloneable
	{
		private ChartLineOptions FMainAxis;
		private ChartLineOptions FMajorGridLines;
		private ChartLineOptions FMinorGridLines;
		private ChartLineOptions FWallLines;
		private bool FDoNotDrawLabelsIfNotDrawingAxis;

		/// <summary>
		/// Line options for the main axis line.
		/// </summary>
		public ChartLineOptions MainAxis {get {return FMainAxis;} set{FMainAxis = value;}}

		/// <summary>
		/// Line options for the major gridlines along the axis.
		/// </summary>
		public ChartLineOptions MajorGridLines {get {return FMajorGridLines;} set{FMajorGridLines = value;}}

		/// <summary>
		/// Line options for the minor gridlines along the axis.
		/// </summary>
		public ChartLineOptions MinorGridLines {get {return FMinorGridLines;} set{FMinorGridLines = value;}}

		/// <summary>
		/// Line options for the walls if this axis is <see cref="TAxisType.Category"/> or <see cref="TAxisType.Series"/> or 
		/// line options for the floor otherwise.
		/// </summary>
		public ChartLineOptions WallLines {get {return FWallLines;} set{FWallLines = value;}}

		/// <summary>
		/// If true and the line format is none, the axis labels will not be drawn.
		/// </summary>
		public bool DoNotDrawLabelsIfNotDrawingAxis {get {return FDoNotDrawLabelsIfNotDrawingAxis;} set{FDoNotDrawLabelsIfNotDrawingAxis = value;}}

		#region ICloneable Members

		/// <summary>
		/// Returns a deep copy of this object.
		/// </summary>
		/// <returns></returns>
		public object Clone()
		{
			TAxisLineOptions Result = new TAxisLineOptions();
			if (FMainAxis != null) Result.FMainAxis = (ChartLineOptions)FMainAxis.Clone();
			if (FMajorGridLines != null) Result.FMajorGridLines = (ChartLineOptions)FMajorGridLines.Clone();
			if (FMinorGridLines != null) Result.FMinorGridLines = (ChartLineOptions)FMinorGridLines.Clone();
			if (FWallLines != null) Result.FWallLines = (ChartLineOptions)FWallLines.Clone();
			Result.DoNotDrawLabelsIfNotDrawingAxis = DoNotDrawLabelsIfNotDrawingAxis;
			return Result;
		}

		#endregion
	}
	#endregion

	#region AxisTickOptions
	/// <summary>
	/// Ticks for an axis.
	/// </summary>
	public enum TTickType
	{
		/// <summary>
		/// No Ticks.
		/// </summary>
		None = 0,

		/// <summary>
		/// Inside axis line.
		/// </summary>
		Inside = 1,

		/// <summary>
		/// Outside axis line.
		/// </summary>
		Outside = 2,

		/// <summary>
		/// Crosses the axis line.
		/// </summary>
		Cross = 3
	}

	/// <summary>
	/// Position of the labels on the axis.
	/// </summary>
	public enum TAxisLabelPosition
	{
		/// <summary>
		/// This axis has no labels.
		/// </summary>
		None = 0,

		/// <summary>
		/// Labels go at the left of the chart, or to the bottom if axis is vertical.
		/// </summary>
		LowEnd = 1,

		/// <summary>
		/// Labels go at the right of the chart, or to the top if axis is vertical.
		/// </summary>
		HighEnd = 2,

		/// <summary>
		/// Labels go next the axis, not to the chart.
		/// </summary>
		NextToAxis = 3
	}

	/// <summary>
	/// How to draw backgrounds of text.
	/// </summary>
	public enum TBackgroundMode
	{
		/// <summary>
		/// This is equivalent to transparent.
		/// </summary>
		Automatic = 0,

		/// <summary>
		/// Text will be drawn transpaerntly.
		/// </summary>
		Transparent = 1,

		/// <summary>
		/// Text will be drawn on a box.
		/// </summary>
		Opaque = 2
	}

	/// <summary>
	/// Properties for the ticks and labels of an axis.
	/// </summary>
	public class TAxisTickOptions: ICloneable
	{
		private TTickType FMajorTickType;
		private TTickType FMinorTickType;
		private TAxisLabelPosition FLabelPosition;
		private TBackgroundMode FBackgroundMode;
		private Color FLabelColor;
		private int FRotation;

		/// <summary>
		/// Creates a new TAxisTickOptions instance.
		/// </summary>
		/// <param name="aMajorTickType">See <see cref="MajorTickType"/></param>
		/// <param name="aMinorTickType">See <see cref="MinorTickType"/></param>
		/// <param name="aLabelPosition">See <see cref="LabelPosition"/></param>
		/// <param name="aBackgroundMode">See <see cref="BackgroundMode"/></param>
		/// <param name="aLabelColor">See <see cref="LabelColor"/></param>
		/// <param name="aRotation">See <see cref="Rotation"/></param>
		public TAxisTickOptions(TTickType aMajorTickType, TTickType aMinorTickType, TAxisLabelPosition aLabelPosition, TBackgroundMode aBackgroundMode, Color aLabelColor, int aRotation)
		{
			FMajorTickType = aMajorTickType;
			FMinorTickType = aMinorTickType;
			FLabelPosition = aLabelPosition;
			FBackgroundMode = aBackgroundMode;
			FLabelColor = aLabelColor;
			FRotation = aRotation;
		}

		/// <summary>
		/// Majot ticks type.
		/// </summary>
		public TTickType MinorTickType {get {return FMinorTickType;} set{FMinorTickType = value;}}

		/// <summary>
		/// Minor ticks type.
		/// </summary>
		public TTickType MajorTickType {get {return FMajorTickType;} set{FMajorTickType = value;}}

		/// <summary>
		/// Position of the label relative to the axis.
		/// </summary>
		public TAxisLabelPosition LabelPosition {get {return FLabelPosition;} set{FLabelPosition = value;}}

		/// <summary>
		/// how the background of text will be rendered.
		/// </summary>
		public TBackgroundMode BackgroundMode {get {return FBackgroundMode;} set{FBackgroundMode = value;}}

		/// <summary>
		/// Color of labels in this axis.
		/// </summary>
		public Color LabelColor {get {return FLabelColor;} set{FLabelColor = value;}}

		/// <summary>
		/// Text Rotation on degrees. 
		/// 0 - 90 is up, 
		/// 91 - 180 is down, 
		/// 255 is vertical.
		/// </summary>
		public int Rotation {get {return FRotation;} set{FRotation = value;}}


		#region ICloneable Members

		/// <summary>
		/// Returns a deep copy of this object.
		/// </summary>
		/// <returns></returns>
		public object Clone()
		{
			return MemberwiseClone();
		}

		#endregion
	}
	#endregion

	#region AxisRangeOptions
	/// <summary>
	/// Properties for the ranges of an axis.
	/// </summary>
	public class TAxisRangeOptions: ICloneable
	{
		private int FLabelFrequency;
		private int FTickFrequency;
		private bool FValueAxisBetweenCategories;
		private bool FValueAxisAtMaxCategory;
		private bool FReverseCategories;

		/// <summary>
		/// Creates a new TAxisRangeOptions instance.
		/// </summary>
		/// <param name="aLabelFrequency">See <see cref="LabelFrequency"/></param>
		/// <param name="aTickFrequency">See <see cref="TickFrequency"/></param>
		/// <param name="aValueAxisBetweenCategories">See <see cref="ValueAxisBetweenCategories"/></param>
		/// <param name="aValueAxisAtMaxCategory">See <see cref="ValueAxisAtMaxCategory"/></param>
		/// <param name="aReverseCategories">See <see cref="ReverseCategories"/></param>
		public TAxisRangeOptions(int aLabelFrequency, int aTickFrequency, bool aValueAxisBetweenCategories, bool aValueAxisAtMaxCategory, bool aReverseCategories)
		{
			FLabelFrequency = aLabelFrequency;
			FTickFrequency = aTickFrequency;
            FValueAxisBetweenCategories = aValueAxisBetweenCategories;
			FValueAxisAtMaxCategory = aValueAxisAtMaxCategory;
			FReverseCategories = aReverseCategories;
		}

		/// <summary>
		/// Frequency at what the labels on categories are displayed. 1 means display all labels, 2 display one label and skip one, and so on.
		/// </summary>
		public int LabelFrequency {get {return FLabelFrequency;} set{FLabelFrequency = value;}}

		/// <summary>
		/// Frequency at what the ticks on categories are displayed. 1 means display all ticks, 2 display one tick and skip one, and so on.
		/// </summary>
		public int TickFrequency {get {return FTickFrequency;} set{FTickFrequency = value;}}

        /// <summary>
        /// Specifies if the Y Axis crosses between categories or in the middle of one. Normally a Column Chart 
        /// cross in the middle, and an area chart will cross between.
        /// </summary>
		public bool ValueAxisBetweenCategories {get {return FValueAxisBetweenCategories;} set{FValueAxisBetweenCategories = value;}}

		/// <summary>
		/// True if the Y axis is at the left.
		/// </summary>
		public bool ValueAxisAtMaxCategory {get {return FValueAxisAtMaxCategory;} set{FValueAxisAtMaxCategory = value;}}

		/// <summary>
		/// True if categories should be printed in reverse order.
		/// </summary>
		public bool ReverseCategories {get {return FReverseCategories;} set{FReverseCategories = value;}}


		#region ICloneable Members

		/// <summary>
		/// Returns a deep copy of the object.
		/// </summary>
		/// <returns></returns>
		public object Clone()
		{
			return MemberwiseClone();
		}

		#endregion
	}
	#endregion

	#region Base Axis
	/// <summary>
	/// Common ancestor for all Axis types.
	/// </summary>
	public abstract class TBaseAxis
	{
		#region Privates
		
		private TFlxChartFont FFont;
		private TAxisTickOptions FTickOptions;
		private string FNumberFormat;
		private TAxisLineOptions FAxisLineOptions;
		private TAxisRangeOptions FRangeOptions;

		private TDataLabel FCaption;
		#endregion

		#region Constructor
		/// <summary>
		/// Creates a new TBAseAxis instance.
		/// </summary>
		/// <param name="aFont">See <see cref="Font"/></param>
		/// <param name="aNumberFormat">See <see cref="NumberFormat"/></param>
		/// <param name="aAxisLineOptions">See <see cref="AxisLineOptions"/></param>
		/// <param name="aTickOptions">See <see cref="TickOptions"/></param>
		/// <param name="aRangeOptions">See <see cref="RangeOptions"/></param>
		/// <param name="aCaption">See <see cref="Caption"/></param>
		protected TBaseAxis(TFlxChartFont aFont, string aNumberFormat, TAxisLineOptions aAxisLineOptions, TAxisTickOptions aTickOptions, TAxisRangeOptions aRangeOptions, TDataLabel aCaption)
		{
			if (aFont == null) FFont = null; else FFont = (TFlxChartFont)aFont.Clone();
			FNumberFormat = aNumberFormat;
			if (aAxisLineOptions != null) FAxisLineOptions = (TAxisLineOptions) aAxisLineOptions.Clone();
			if (aTickOptions != null) FTickOptions = (TAxisTickOptions) aTickOptions.Clone();
			if (aRangeOptions != null) FRangeOptions = (TAxisRangeOptions) aRangeOptions.Clone();
			if (aCaption != null) FCaption = (TDataLabel) aCaption.Clone();
		}

		#endregion

		#region Properties
		/// <summary>
		/// Font used on this axis.
		/// </summary>
		public TFlxChartFont Font {get {return FFont;}}

		/// <summary>
		/// Format for the numbers on this axis.
		/// </summary>
		public string NumberFormat {get {return FNumberFormat;} set{FNumberFormat = value;}}

		/// <summary>
		/// Linestyles for the different lines of this axis.
		/// </summary>
		public TAxisLineOptions AxisLineOptions {get {return FAxisLineOptions;} set{FAxisLineOptions = value;}}

		/// <summary>
		/// Options for the ticks and the font used on the labels.
		/// </summary>
		public TAxisTickOptions TickOptions {get {return FTickOptions;} set{FTickOptions = value;}}

		/// <summary>
		/// Options for the range of this axis.
		/// </summary>
		public TAxisRangeOptions RangeOptions {get {return FRangeOptions;} set{FRangeOptions = value;}}

		/// <summary>
		/// Axis Caption.
		/// </summary>
		public TDataLabel Caption {get {return FCaption;} set{FCaption = value;}}

		#endregion
	}
	#endregion

	#region Category Axis
	/// <summary>
	/// Information about an Axis of categories. (normally the x axis)
	/// </summary>
	public class TCategoryAxis: TBaseAxis
	{
		private int FMin;
		private int FMax;
		private int FMajorValue;
		private int FMajorUnit;
		private int FMinorValue;
		private int FMinorUnit;

		private int FBaseUnit;

		private int FCrossValue;

		private TCategoryAxisOptions FAxisOptions;

		/// <summary>
		/// Constructs a new TCategoryAxisOption instance with all values set to automatic.
		/// </summary>
		public TCategoryAxis(): base(null, String.Empty, null, null, null, null)
		{
			FAxisOptions = TCategoryAxisOptions.AutoMin | 
				TCategoryAxisOptions.AutoMax |
				TCategoryAxisOptions.AutoMajor |
				TCategoryAxisOptions.AutoMinor |
				//This is not a date axis
				TCategoryAxisOptions.AutoBase |
				TCategoryAxisOptions.AutoCrossDate |
				TCategoryAxisOptions.AutoDate;

		}

		/// <summary>
		/// Constructs a new TCategoryAxisOptions instance.
		/// </summary>
		/// <param name="aAxisOptions">See <see cref="AxisOptions"/></param>
		/// <param name="aBaseUnit">See <see cref="BaseUnit"/></param>
		/// <param name="aCrossValue">See <see cref="CrossValue"/></param>
		/// <param name="aMajorUnit">See <see cref="MajorUnit"/></param>
		/// <param name="aMajorValue">See <see cref="MajorValue"/></param>
		/// <param name="aMax">See <see cref="Max"/></param>
		/// <param name="aMin">See <see cref="Min"/></param>
		/// <param name="aMinorUnit">See <see cref="MinorUnit"/></param>
		/// <param name="aMinorValue">See <see cref="MinorValue"/></param>
		/// <param name="aFont">See <see cref="Font"/></param>
		/// <param name="aNumberFormat">See <see cref="TBaseAxis.NumberFormat"/></param>
		/// <param name="aAxisLineOptions">See <see cref="TBaseAxis.AxisLineOptions"/>. This parameter will be cloned.</param>
		/// <param name="aTickOptions">See <see cref="TBaseAxis.TickOptions"/> This parameter will be cloned.</param>
		/// <param name="aRangeOptions">See <see cref="TBaseAxis.RangeOptions"/> This parameter will be cloned.</param>
		/// <param name="aCaption">See <see cref="TBaseAxis.Caption"/> This parameter will be cloned.</param>
		public TCategoryAxis(int aMin, int aMax, int aMajorValue, int aMajorUnit, int aMinorValue, int aMinorUnit, int aBaseUnit, int aCrossValue, TCategoryAxisOptions aAxisOptions, TFlxChartFont aFont, string aNumberFormat, TAxisLineOptions aAxisLineOptions, TAxisTickOptions aTickOptions, TAxisRangeOptions aRangeOptions, TDataLabel aCaption):
			base(aFont, aNumberFormat, aAxisLineOptions, aTickOptions, aRangeOptions, aCaption)
		{
			FMin			= aMin;
			FMax			= aMax;
			FMajorValue		= aMajorValue;
			FMajorUnit		= aMajorUnit;
			FMinorValue		= aMinorValue;
			FMinorUnit		= aMinorUnit;
				
			FBaseUnit		= aBaseUnit;
				
			FCrossValue		= aCrossValue;

			FAxisOptions	= aAxisOptions;

		}


		/// <summary>
		/// Minimum value for the axis, when not set to automatic in <see cref="AxisOptions"/>.
		/// </summary>
		public int Min {get {return FMin;} set{FMin = value;}}

		/// <summary>
		/// Maximum value for the axis, when not set to automatic in <see cref="AxisOptions"/>.
		/// </summary>
		public int Max {get {return FMax;} set{FMax = value;}}
		
		/// <summary>
		/// Value for the major unit, when not set to automatic in <see cref="AxisOptions"/>.
		/// </summary>
		public int MajorValue {get {return FMajorValue;} set{FMajorValue = value;}}

		/// <summary>
		/// Units for the major unit, when not set to automatic in <see cref="AxisOptions"/>.
		/// </summary>
		public int MajorUnit {get {return FMajorUnit;} set{FMajorUnit = value;}}

		/// <summary>
		/// Value for the minor unit, when not set to automatic in <see cref="AxisOptions"/>.
		/// </summary>
		public int MinorValue {get {return FMinorValue;} set{FMinorValue = value;}}
		
		/// <summary>
		/// Units for the minor unit, when not set to automatic in <see cref="AxisOptions"/>.
		/// </summary>
		public int MinorUnit {get {return FMinorUnit;} set{FMinorUnit = value;}}

		/// <summary>
		/// Base units for the axis, when not set to automatic in <see cref="AxisOptions"/>.
		/// </summary>
		public int BaseUnit {get {return FBaseUnit;} set{FBaseUnit = value;}}

		/// <summary>
		/// Value where the other Axis will cross this one, when not set to automatic in <see cref="AxisOptions"/>.
		/// </summary>
		public int CrossValue {get {return FCrossValue;} set{FCrossValue = value;}}

		/// <summary>
		/// Enumerates which of the other options contain valid values or are automatic.
		/// </summary>
		public TCategoryAxisOptions AxisOptions {get {return FAxisOptions;} set{FAxisOptions = value;}}
	}

	/// <summary>
	/// Options for a Category Axis.
	/// </summary>
	[Flags]
	public enum TCategoryAxisOptions
	{
		/// <summary>
		/// All magnitudes will be manual.
		/// </summary>
		None = 0x00,

		/// <summary>
		/// Use automatic minimum.
		/// </summary>
		AutoMin = 0x01,

		/// <summary>
		/// Use automatic maximum.
		/// </summary>
		AutoMax = 0x02,

		/// <summary>
		/// Use automatic major units.
		/// </summary>
		AutoMajor = 0x04,

		/// <summary>
		/// Use automatic minor units.
		/// </summary>
		AutoMinor = 0x08,

		/// <summary>
		/// This is a date Axis.
		/// </summary>
		DateAxis = 0x10,

		/// <summary>
		/// Use automatic base unit.
		/// </summary>
		AutoBase = 0x20,

		/// <summary>
		/// Use automatic date crossing point.
		/// </summary>
		AutoCrossDate = 0x40,

		/// <summary>
		/// Use automatic date settings.
		/// </summary>
		AutoDate = 0x80
	}


	#endregion

	#region Value Axis
	/// <summary>
	/// Information about an Axis of values. (normally the y axis)
	/// </summary>
	public class TValueAxis: TBaseAxis
	{
		private double FMin;
		private double FMax;
		private double FMajor;
		private double FMinor;
		private double FCrossValue;

		private TValueAxisOptions FAxisOptions;

		/// <summary>
		/// Constructs a new TValueAxisOptions instance with all values set to automatic.
		/// </summary>
		public TValueAxis(): base(null, String.Empty, null, null, null, null)
		{
			FAxisOptions = TValueAxisOptions.AutoMin | 
				TValueAxisOptions.AutoMax |
				TValueAxisOptions.AutoMajor |
				TValueAxisOptions.AutoMinor |
				TValueAxisOptions.AutoCross;
			
		}

		/// <summary>
		/// Constructs a new TValueAxisOptions instance.
		/// </summary>
		/// <param name="aMin">See <see cref="Min"/></param>
		/// <param name="aMax">See <see cref="Max"/></param>
		/// <param name="aMajor">See <see cref="Major"/></param>
		/// <param name="aMinor">See <see cref="Minor"/></param>
		/// <param name="aCrossValue">See <see cref="CrossValue"/></param>
		/// <param name="aAxisOptions">See <see cref="AxisOptions"/></param>
		/// <param name="aFont">See <see cref="Font"/>. This parameter will be cloned.</param>
		/// <param name="aNumberFormat">See <see cref="TBaseAxis.NumberFormat"/></param>
		/// <param name="aAxisLineOptions">See <see cref="TBaseAxis.AxisLineOptions"/>. This parameter will be cloned.</param>
		/// <param name="aTickOptions">See <see cref="TBaseAxis.TickOptions"/>. This parameter will be cloned.</param>
		/// <param name="aRangeOptions">See <see cref="TBaseAxis.RangeOptions"/> This parameter will be cloned.</param>
		/// <param name="aCaption">See <see cref="TBaseAxis.Caption"/> This parameter will be cloned.</param>
		public TValueAxis(double aMin, double aMax, double aMajor, double aMinor, double aCrossValue, 
			TValueAxisOptions aAxisOptions, TFlxChartFont aFont, string aNumberFormat, TAxisLineOptions aAxisLineOptions, TAxisTickOptions aTickOptions, TAxisRangeOptions aRangeOptions, TDataLabel aCaption): 
			base(aFont, aNumberFormat, aAxisLineOptions, aTickOptions, aRangeOptions, aCaption)
		{
			FMin			= aMin;
			FMax			= aMax;
			FMajor			= aMajor;
			FMinor			= aMinor;
				
			FCrossValue		= aCrossValue;

			FAxisOptions	= aAxisOptions;
		}

		/// <summary>
		/// Minimum value for the axis, when not set to automatic in <see cref="AxisOptions"/>.
		/// </summary>
		public double Min {get {return FMin;} set{FMin = value;}}

		/// <summary>
		/// Maximum value for the axis, when not set to automatic in <see cref="AxisOptions"/>.
		/// </summary>
		public double Max {get {return FMax;} set{FMax = value;}}
		
		/// <summary>
		/// Value for the major unit, when not set to automatic in <see cref="AxisOptions"/>.
		/// </summary>
		public double Major {get {return FMajor;} set{FMajor = value;}}

		/// <summary>
		/// Value for the minor unit, when not set to automatic in <see cref="AxisOptions"/>.
		/// </summary>
		public double Minor {get {return FMinor;} set{FMinor = value;}}
		
		/// <summary>
		/// Value where the other Axis will cross this one, when not set to automatic in <see cref="AxisOptions"/>.
		/// </summary>
		public double CrossValue {get {return FCrossValue;} set{FCrossValue = value;}}

		/// <summary>
		/// Enumerates which of the other options contain valid values or are automatic.
		/// </summary>
		public TValueAxisOptions AxisOptions {get {return FAxisOptions;} set{FAxisOptions = value;}}

	}

	/// <summary>
	/// Options for a Value Axis.
	/// </summary>
	[Flags]
	public enum TValueAxisOptions
	{
		/// <summary>
		/// All magnitudes will be manual.
		/// </summary>
		None = 0x00,

		/// <summary>
		/// Use automatic minimum.
		/// </summary>
		AutoMin = 0x01,

		/// <summary>
		/// Use automatic maximum.
		/// </summary>
		AutoMax = 0x02,

		/// <summary>
		/// Use automatic major units.
		/// </summary>
		AutoMajor = 0x04,

		/// <summary>
		/// Use automatic minor units.
		/// </summary>
		AutoMinor = 0x08,

		/// <summary>
		/// Use automatic crossing point.
		/// </summary>
		AutoCross = 0x10,

		/// <summary>
		/// Use logarithmic scale.
		/// </summary>
		LogScale = 0x20,

		/// <summary>
		/// Values will be in reverse order.
		/// </summary>
		Reverse = 0x40,

		/// <summary>
		/// The category Axis will cross at the maximum value.
		/// </summary>
		MaxCross = 0x80
	}

	#endregion

	#region Chart Axis

	/// <summary>
	/// A class encapsulating the information of an axis
	/// </summary>
	public class TChartAxis
	{
		private int FIndex;
		private Rectangle FAxisLocation;

		private TBaseAxis FCategoryAxis;
		private TValueAxis FValueAxis;

		/// <summary>
		/// Creates a new ChartAxis instance.
		/// </summary>
		/// <param name="aIndex">Axis Index. 0 means primary, 1 secundary.</param>
		/// <param name="aAxisLocation">Location of the axis, in 1/4000 units of the current chart dimensions.</param>
		/// <param name="aCategoryAxis">Data of the Category Axis. This might be a CategoryAxis for line or bar charts, or a ValuesAxis for a Scatter Chart.</param>
		/// <param name="aValueAxis">Data of the value Axis.</param>
		public TChartAxis(int aIndex, Rectangle aAxisLocation, TBaseAxis aCategoryAxis, TValueAxis aValueAxis)
		{
			FIndex = aIndex;
			FAxisLocation = aAxisLocation;

			if (aCategoryAxis == null) FCategoryAxis = new TCategoryAxis(); else FCategoryAxis = aCategoryAxis;
			if (aValueAxis == null) FValueAxis = new TValueAxis(); else FValueAxis = aValueAxis;

		}

		/// <summary>
		/// Axis Index. 0 means primary, 1 secundary.
		/// </summary>
		public int Index {get {return FIndex;} set{FIndex = value;}}

		/// <summary>
		/// Returns the Axis coordinates on 1/4000 parts of the chart area.
		/// </summary>
		public Rectangle AxisLocation {get {return FAxisLocation;} set{FAxisLocation = value;}}

		/// <summary>
		/// Returns information about the Category Axis (X-Axis on a non rtated chart). Note that this might be a <see cref="TValueAxis"/> axis for scatter charts, or a <see cref="TCategoryAxis"/> for line or bar charts.
		/// </summary>
		public TBaseAxis CategoryAxis {get {return FCategoryAxis;}}

		/// <summary>
		/// Returns information about the Value Axis (Y-Axis on a non rotated chart).
		/// </summary>
		public TValueAxis ValueAxis {get {return FValueAxis;} set{FValueAxis = value;}}

	}

	#endregion

	#endregion

	#region Legend

	/// <summary>
	/// Position of hte legend inside the chart.
	/// </summary>
	public enum TChartLegendPos
	{
		/// <summary>
		/// At bottom.
		/// </summary>
		Bottom = 0,

		/// <summary>
		/// At the corner.
		/// </summary>
		Corner = 1,

		/// <summary>
		/// At the top.
		/// </summary>
		Top = 2,

		/// <summary>
		/// At the right.
		/// </summary>
		Right = 3,

		/// <summary>
		/// At the left.
		/// </summary>
		Left = 4,

		/// <summary>
		/// The legend is not docked inside the chart area.
		/// </summary>
		NotDocked = 7
	}

	/// <summary>
	/// Description of the chart's legend box.
	/// </summary>
	public class TChartLegend: ICloneable
	{
		#region Privates
		private TChartFrameOptions FFrame;
		private long FX;
		private long FY;
		private long FWidth;
		private long FHeight;
		private TChartLegendPos FPlacement;
		private TChartTextOptions FTextOptions;
		#endregion

		/// <summary>
		/// Creates a new TChartLegend instance.
		/// </summary>
		/// <param name="aX">See <see cref="X"/></param>
		/// <param name="aY">See <see cref="Y"/></param>
		/// <param name="aWidth">See <see cref="Width"/></param>
		/// <param name="aHeight">See <see cref="Height"/></param>
		/// <param name="aFrame">See <see cref="Frame"/></param>
		/// <param name="aPlacement">See <see cref="Placement"/></param>
		/// <param name="aTextOptions">See <see cref="TextOptions"/></param>
		public TChartLegend(long aX, long aY, long aWidth, long aHeight, TChartLegendPos aPlacement, TChartFrameOptions aFrame, TChartTextOptions aTextOptions)
		{
			FX = aX;
			FY = aY;
			FWidth = aWidth;
			FHeight = aHeight;
			FPlacement = aPlacement;
			FFrame = aFrame;
			FTextOptions = aTextOptions;
		}

		#region Properties

		/// <summary>
		/// X coordinate on 1/4000 units of chart area.
		/// </summary>
		public long X {get {return FX;} set{FX = value;}}

		/// <summary>
		/// Y coordinate on 1/4000 units of chart area.
		/// </summary>
		public long Y {get {return FY;} set{FY = value;}}

		/// <summary>
		/// Height of the bounding box, on 1/4000 units of chart size.
		/// </summary>
		public long Height {get {return FHeight;} set{FHeight = value;}}
		
		/// <summary>
		/// Width of the bounding box, on 1/4000 units of chart size.
		/// </summary>
		public long Width {get {return FWidth;} set{FWidth = value;}}

		/// <summary>
		/// Placement of the legend inside of the chart.
		/// </summary>
		public TChartLegendPos Placement {get {return FPlacement;} set{FPlacement = value;}}

		/// <summary>
		/// Line and fill style options for the frame.
		/// </summary>
		public TChartFrameOptions Frame {get {return FFrame;} set{FFrame = value;}}

		/// <summary>
		/// Global font options for the legend labels.
		/// </summary>
		public TChartTextOptions TextOptions {get {return FTextOptions;} set{FTextOptions = value;}}


		#endregion

		#region ICloneable Members

		/// <summary>
		/// Returns a deep copy of this object.
		/// </summary>
		/// <returns></returns>
		public object Clone()
		{
			TChartLegend Result = (TChartLegend) MemberwiseClone();
			if (Frame != null) Result.Frame = (TChartFrameOptions)FFrame.Clone();
			return Result;
		}

		#endregion
	}
	#endregion

	#region Legend Options
	/// <summary>
	/// Description of one particular entry on the Legend box.
	/// </summary>
	public class TLegendEntryOptions: ICloneable
	{
		#region Privates
		private bool FEntryDeleted;
		private TChartTextOptions FTextFormat;
		#endregion

		/// <summary>
		/// Creates a new TLegendEntryOptions instance.
		/// </summary>
		/// <param name="aEntryDeleted">See <see cref="EntryDeleted"/></param>
		/// <param name="aTextFormat">See <see cref="TextFormat"/></param>
		public TLegendEntryOptions(bool aEntryDeleted, TChartTextOptions aTextFormat)
		{
			FEntryDeleted = aEntryDeleted;
			FTextFormat = aTextFormat;
		}

		#region Properties
		/// <summary>
		/// If true, this series has been deleted from the legend box, and it should not be displayed.
		/// </summary>
		public bool EntryDeleted {get {return FEntryDeleted;} set{FEntryDeleted = value;}}

		/// <summary>
		/// Font to use on this particular entry. If null, the default font for the legend box should be used.
		/// </summary>
		public TChartTextOptions TextFormat {get {return FTextFormat;} set{FTextFormat = value;}}

		#endregion

		#region ICloneable Members

		/// <summary>
		/// Returns a deep copy of this object.
		/// </summary>
		/// <returns></returns>
		public object Clone()
		{
			TLegendEntryOptions Result = (TLegendEntryOptions) MemberwiseClone();
			if (TextFormat != null) Result.TextFormat = (TChartTextOptions) TextFormat.Clone();
			return Result;
		}

		#endregion
	}
	#endregion

	#region DropBars
	/// <summary>
	/// Information about a Drop Bar.
	/// </summary>
	public class TChartDropBars: ICloneable
	{
		#region Privates
		private TChartOneDropBar FUpBar;
		private TChartOneDropBar FDownBar;
		private ChartLineOptions FDropLines;
		private ChartLineOptions FHiLoLines;
		#endregion

		/// <summary>
		/// Creates a new TChartDropBars instance.
		/// </summary>
		/// <param name="aDropLines">See <see cref="DropLines"/></param>
		/// <param name="aUpBar">See <see cref="UpBar"/></param>
		/// <param name="aDownBar">See <see cref="DownBar"/></param>
		/// <param name="aHiLoLines">See <see cref="HiLoLines"/></param>
		public TChartDropBars(ChartLineOptions aDropLines, ChartLineOptions aHiLoLines, 
			TChartOneDropBar aUpBar, TChartOneDropBar aDownBar)
		{
			FDropLines = aDropLines;
			FHiLoLines = aHiLoLines;
			FUpBar = aUpBar;
			FDownBar = aDownBar;
		}

		/// <summary>
		/// Line style for the drop lines if they exist, null otherwise.
		/// </summary>
		public ChartLineOptions DropLines {get {return FDropLines;} set{FDropLines = value;}}

		/// <summary>
		/// Lie srtyle for the High-Low lines if they exist, null otherwise.
		/// </summary>
		public ChartLineOptions HiLoLines {get {return FHiLoLines;} set{FHiLoLines = value;}}

		/// <summary>
		/// The data for the up drop bar.
		/// </summary>
		public TChartOneDropBar UpBar {get {return FUpBar;} set{FUpBar = value;}}

		/// <summary>
		/// The data for the down drop bar.
		/// </summary>
		public TChartOneDropBar DownBar {get {return FDownBar;} set{FDownBar = value;}}

		#region ICloneable Members

		/// <summary>
		/// Returns a deep copy of this object.
		/// </summary>
		/// <returns></returns>
		public object Clone()
		{
			TChartDropBars Result = (TChartDropBars) MemberwiseClone();
			if (FUpBar != null) Result.FUpBar = (TChartOneDropBar)FUpBar.Clone();
			if (FDownBar != null) Result.FDownBar = (TChartOneDropBar)FDownBar.Clone();
			return Result;
		}

		#endregion
	}

	/// <summary>
	/// Information about one specific drop bar.
	/// </summary>
	public class TChartOneDropBar: ICloneable
	{
		#region Privates
		private TChartFrameOptions FFrame;
		private int FGapWidth;
		#endregion

		/// <summary>
		/// Creates a new instance.
		/// </summary>
		/// <param name="aFrame">See <see cref="Frame"/></param>
		/// <param name="aGapWidth">See <see cref="GapWidth"/></param>
		public TChartOneDropBar(TChartFrameOptions aFrame, int aGapWidth)
		{
			FFrame = aFrame;
			FGapWidth = aGapWidth;
		}

		#region Properties
		/// <summary>
		/// Properties for the drop bar.
		/// </summary>
		public TChartFrameOptions Frame {get {return FFrame;} set{FFrame = value;}}

		/// <summary>
		/// Gap width on percent.(0 to 100)
		/// </summary>
		public int GapWidth {get {return FGapWidth;} set{FGapWidth = value;}}

		#endregion

		#region ICloneable Members

		/// <summary>
		/// Returns a deep copy of this object.
		/// </summary>
		/// <returns></returns>
		public object Clone()
		{
			TChartOneDropBar Result = (TChartOneDropBar) MemberwiseClone();
			if (Frame != null) Result.Frame = (TChartFrameOptions)FFrame.Clone();
			return Result;
		}

		#endregion
	}
	
	#endregion

	#region Data Labels

	/// <summary>
	/// Specifies to which object the label is linked to.
	/// </summary>
	public enum TLinkOption
	{
        /// <summary>
        /// Not used. It means the label is not linked.
        /// </summary>
        None = 0,

		/// <summary>
		/// This is the chart title.
		/// </summary>
		ChartTitle = 1,

		/// <summary>
		/// This is the caption for the Y-Axis.
		/// </summary>
		YAxisTitle = 2,

		/// <summary>
		/// This is the caption for the X-Axis.
		/// </summary>
		XAxisTitle = 3,

		/// <summary>
		/// This is a Data Label, linked to a series or a point in the series.
		/// </summary>
		DataLabel = 4,

		/// <summary>
		/// This is the caption for the Z-Axis.
		/// </summary>
		ZAxisTitle = 7
	}

	/// <summary>
	/// Defines what information a DataLabel should show.
	/// </summary>
	public enum TLabelDataValue
	{
		/// <summary>
		/// This label has a manually entered value.
		/// </summary>
		Manual,

		/// <summary>
		/// This label shows the values and/or categories of the series.
		/// </summary>
		SeriesInfo,
	}

	/// <summary>
	/// Defines where a label is displayed on a chart.
	/// </summary>
	public enum TDataLabelPosition
	{
		/// <summary>
		/// Label is placed on the default position. This is AutoPie for Pie charts, Right for Line Charts, Outside for BarCharts and Center for Stacked Bar charts.
		/// </summary>
		Automatic = 0,

		/// <summary>
		/// Label is outside the data. Applies to Bar, pie charts.
		/// </summary>
		Outside = 1,
	
		/// <summary>
		/// Label is inside the data. Applies to Bar, pie charts.
		/// </summary>
		Inside = 2,
	
		/// <summary>
		/// Label is centered on the data. Applies to Bar, line, pie charts.
		/// </summary>
		Center = 3,
	
		/// <summary>
		/// Label is placed on the Axis. Applies to Bar charts.
		/// </summary>
		Axis = 4,
 
		/// <summary>
		/// Label is placed Above the data. Applies to Line charts.
		/// </summary>
		Above = 5,
 
		/// <summary>
		/// Label is placed Below the data. Applies to Line charts.
		/// </summary>
		Below = 6,
 
		/// <summary>
		/// Label is placed Left of the data. Applies to Line charts.
		/// </summary>
		Left = 7,
 
		/// <summary>
		/// Label is placed Right of the data. Applies to Line charts.
		/// </summary>
		Right = 8,
 
		/// <summary>
		/// Label is placed automatically on the chart. This applies to Pie charts, and it is equivalent to TDataLabelPosition.Automatic.
		/// </summary>
		AutoPie = 9,

		/// <summary>
		/// Label has been manually moved and it is placed at an arbitrary place.
		/// </summary>
		Any = 10
	}

	/// <summary>
	/// Options for a data label.
	/// </summary>
	public class TDataLabelOptions: ICloneable
	{
		bool FAutoColor;
		bool FShowLegendKey;
		bool FShowSeriesName;
		bool FShowCategories;
		bool FShowValues;
		bool FShowPercents;
		bool FShowBubbles;
        string FSeparator;
		TLabelDataValue FDataType;
		bool FDeleted;
		TDataLabelPosition FPosition;

		/// <summary>
		/// True if this label will use automatic coloring, false if the color is user defined.
		/// </summary>
		public bool AutoColor {get {return FAutoColor;} set{FAutoColor = value;}}

		/// <summary>
		/// If true, the legend key will be shown along with the label. 
		/// </summary>
		public bool ShowLegendKey {get {return FShowLegendKey;} set{FShowLegendKey = value;}}

		/// <summary>
		/// If true and this label <see cref="DataType"/> is SeriesInfo, this label will display the Series name.
		/// </summary>
		public bool ShowSeriesName {get {return FShowSeriesName;} set{FShowSeriesName = value;}}

		/// <summary>
		/// If true and this label <see cref="DataType"/> is SeriesInfo, this label will display the Categories.
		/// </summary>
		public bool ShowCategories {get {return FShowCategories;} set{FShowCategories = value;}}

		/// <summary>
		/// If true and this label <see cref="DataType"/> is SeriesInfo, this label will display the actual value of the data.
		/// </summary>
		public bool ShowValues {get {return FShowValues;} set{FShowValues = value;}}

		/// <summary>
		/// If true and this label <see cref="DataType"/> is SeriesInfo, this label will display the percentage of the total data. This value only applies to PIE charts.
		/// </summary>
		public bool ShowPercents {get {return FShowPercents;} set{FShowPercents = value;}}

		/// <summary>
		/// If true and this label <see cref="DataType"/> is SeriesInfo, this label will display the percentage bubble size. This value only applies to BUBBLE charts.
		/// </summary>
		public bool ShowBubbles {get {return FShowBubbles;} set{FShowBubbles = value;}}

        /// <summary>
        /// The separator that will be used to separate labels when they contain more than one value. (For example, if the labes contains both the value and
        /// the category, they will be separated by this string).
        /// </summary>
        public string Separator{ get { return FSeparator; } set { FSeparator = value; } }

		/// <summary>
		/// Defines which information this label displays.
		/// </summary>
		public TLabelDataValue DataType {get {return FDataType;} set{FDataType = value;}}

		/// <summary>
		/// If true, this label has been manually deleted by the user and should not be displayed.
		/// </summary>
		public bool Deleted {get {return FDeleted;} set{FDeleted = value;}}

		/// <summary>
		/// Where the label is placed.
		/// </summary>
		public TDataLabelPosition Position {get {return FPosition;} set{FPosition = value;}}

		#region ICloneable Members

		/// <summary>
		/// Returns a deep copy of the object.
		/// </summary>
		/// <returns></returns>
		public object Clone()
		{
			return MemberwiseClone();
		}

		#endregion
	}

	/// <summary>
	/// Represents one data label on the chart.
	/// </summary>
	public class TDataLabel: ICloneable
	{
		private TChartTextOptions FTextOptions;
		private TChartFrameOptions FFrame;
		private object[] FLabelValues;
		private string FLabelDefinition;
		private TLinkOption FLinkedTo;
		private int FSeriesIndex;
		private int FDataPointIndex;
		private TDataLabelOptions FLabelOptions;
		private string FNumberFormat;


		/// <summary>
		/// Formatting options for this label.
		/// </summary>
		public TChartTextOptions TextOptions {get {return FTextOptions;} set{FTextOptions = value;}}

		/// <summary>
		/// Data options for the label.
		/// </summary>
		public TDataLabelOptions LabelOptions {get {return FLabelOptions;} set{FLabelOptions = value;}}


		/// <summary>
		/// Background for the label, if there is one. Null otherwise.
		/// </summary>
		public TChartFrameOptions Frame {get {return FFrame;} set{FFrame = value;}}

		/// <summary>
		/// A list with the actual values for the label, evaluated from the formula at <see cref="LabelDefinition"/>
		/// IMPORTANT NOTE: The values here only are valid if <see cref="LabelOptions"/> indicates a manual DataType.
		/// </summary>
		public object[] LabelValues {get {return FLabelValues;} set{FLabelValues = value;}}

		/// <summary>
		/// The formula defining the values on this label. You can access the actual values of the labels with <see cref="LabelValues"/>.
		/// </summary>
		public string LabelDefinition {get {return FLabelDefinition;} set{FLabelDefinition = value;}}

		/// <summary>
		/// Defines to which object this label is linked.
		/// </summary>
		public TLinkOption LinkedTo {get {return FLinkedTo;} set{FLinkedTo = value;}}

		/// <summary>
		/// Series number for the series this label displays. This value only has meaning if <see cref="LinkedTo"/> is TLinkOptions.DataLabel
		/// </summary>
		public int SeriesIndex {get {return FSeriesIndex;} set{FSeriesIndex = value;}}

		/// <summary>
		/// Point index of of the point this label displays.  This value only has meaning if <see cref="LinkedTo"/> is TLinkOptions.DataLabel
		/// and the label is for only one point. When the label is for the whole series, DataPointIndex is -1.
		/// </summary>
		public int DataPointIndex {get {return FDataPointIndex;} set{FDataPointIndex = value;}}

		/// <summary>
		/// Numeric format for this label.
		/// </summary>
		public string NumberFormat {get {return FNumberFormat;} set{FNumberFormat = value;}}


		#region ICloneable Members

		/// <summary>
		/// Returns a deep copy of the object.
		/// </summary>
		/// <returns></returns>
		public object Clone()
		{
			TDataLabel Result = (TDataLabel) MemberwiseClone();

			if (TextOptions != null) Result.TextOptions = (TChartTextOptions) TextOptions.Clone();
			if (LabelOptions != null) Result.LabelOptions = (TDataLabelOptions) LabelOptions.Clone();
			if (Frame != null) Result.Frame = (TChartFrameOptions) Frame.Clone();
			if (LabelValues != null) Result.LabelValues = (object[]) LabelValues.Clone();
			return Result;
		}

		#endregion
	}

	#endregion

	#region TFlxChartFont
	/// <summary>
	/// A TFlxFont with Scaling factor. Scaling factor might be different than 1 if
	/// the chart has Autosize Fonts. To get the real value of the font, you need to multiply by the factor.
	/// </summary>
	public class TFlxChartFont:ICloneable
	{
		private double FScale;
		private TFlxFont FFont;

        /// <summary>
        /// Constructs a new instance.
        /// </summary>
		public TFlxChartFont()
		{
			FFont = new TFlxFont();
		}

		/// <summary>
		/// Scale for the font. Multiply by this value to get the real size in points.
		/// </summary>
		public double Scale {get {return FScale;} set{FScale = value;}}

		/// <summary>
		/// Actual information for the font.
		/// </summary>
		public TFlxFont Font {get {return FFont;} set{FFont = value;}}

		#region ICloneable Members

        /// <summary>
        /// Return a deep copy of the copy.
        /// </summary>
        /// <returns></returns>
		public object Clone()
		{
			TFlxChartFont Result = new TFlxChartFont();
			Result.Scale = Scale;
			if (FFont != null) Result.FFont = (TFlxFont) FFont.Clone();
			return Result;
		}

		#endregion
	}
	#endregion

}
