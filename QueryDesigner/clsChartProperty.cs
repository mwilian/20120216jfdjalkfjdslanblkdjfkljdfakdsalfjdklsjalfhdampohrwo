using System;
using System.Collections.Generic;

using System.Text;
using System.ComponentModel;
using System.Drawing;
using System.Text.RegularExpressions;

namespace dCube
{
    /// <summary>
    // Customer class to be displayed in the property grid
    /// </summary>
    ///
    [DefaultPropertyAttribute("_Name")]
    public class clsChartProperty
    {
        #region Designer
        string _chartname;
        [TypeConverter(typeof(ChartList)), CategoryAttribute("01 - Designer"), DescriptionAttribute("Chart _Name")]
        public string ChartName
        {
            get { return _chartname; }
            set { _chartname = value; }
        }
        string _sheetChart = "";
        [CategoryAttribute("01 - Designer"), DescriptionAttribute("Excel Sheet include data for chart")]
        public string SheetChart
        {
            get { return _sheetChart; }
            set { _sheetChart = value; }
        }
        string _dataRange = "";
        [CategoryAttribute("01 - Designer"), DescriptionAttribute("Excel _Name Range include data for chart")]
        public string DataRange
        {
            get { return _dataRange; }
            set { _dataRange = value; }
        }
        string _captionRange;
        [CategoryAttribute("01 - Designer"), DescriptionAttribute("Excel _Name Range include Caption for chart")]
        public string CaptionRange
        {
            get { return _captionRange; }
            set { _captionRange = value; }
        }
        string _subCaptionRange;
        [CategoryAttribute("01 - Designer"), DescriptionAttribute("Excel _Name Range include SubCaption for chart")]
        public string SubCaptionRange
        {
            get { return _subCaptionRange; }
            set { _subCaptionRange = value; }
        }
        #endregion Designer

        #region Functional
        Boolean _animation = false;
        Double _palette;
        Boolean _showLabels = true;
        String _labelDisplay;
        Boolean _rotateLabels;
        Boolean _slantLabels;
        Double _labelStep;
        Double _staggerLines;
        Boolean _showValues = true;
        Boolean _rotateValues;
        Boolean _placeValuesInside;
        Boolean _showYAxisValues = true;
        Boolean _showLimits;
        Boolean _showDivLineValues;
        Double _yAxisValuesStep;
        Boolean _showShadow = true;
        Boolean _adjustDiv;
        Boolean _rotateYAxisName;
        Double _yAxisNameWidth;
        String _clickURL;
        Boolean _defaultAnimation = true;
        Double _yAxisMinValue;
        Double _yAxisMaxValue;
        Boolean _setAdaptiveYMin;
        /*
        Boolean _showAboutMenuItem = false;
        string _aboutMenuItemLabel = "About Ta Vi Co. LTD";        
        string _aboutMenuItemLink = "tavicosoft.com";

        [CategoryAttribute("02 - Functional"), DescriptionAttribute("Show About")]
        public string AboutMenuItemLabel
        {
            get { return _aboutMenuItemLabel; }
            set { _aboutMenuItemLabel = value; }
        }
        [CategoryAttribute("02 - Functional"), DescriptionAttribute("Show About")]
        public string AboutMenuItemLink
        {
            get { return _aboutMenuItemLink; }
            set { _aboutMenuItemLink = value; }
        }        
        [CategoryAttribute("02 - Functional"), DescriptionAttribute("Show About")]
        public Boolean ShowAboutMenuItem
        {
            get { return _showAboutMenuItem; }
            set { _showAboutMenuItem = value; }
        }*/
        [CategoryAttribute("02 - Functional"), DescriptionAttribute("This attribute lets you set the configuration whether the chart should appear in an animated fashion. If you do not want to animate any part of the chart, set this as 0.")]
        public Boolean Animation { get { return _animation; } set { _animation = value; } }
        [CategoryAttribute("02 - Functional"), DescriptionAttribute("Charts introduces the concept of Color Palettes. Each chart has 5 pre-defined color palettes which you can choose from. Each palette renders the chart in a different color theme. Valid values are 1-5.")]
        public Double Palette { get { return _palette; } set { _palette = value; } }
        [CategoryAttribute("02 - Functional"), DescriptionAttribute("It sets the configuration whether the x-axis labels will be displayed or not.")]
        public Boolean ShowLabels { get { return _showLabels; } set { _showLabels = value; } }
        [TypeConverter(typeof(DisplayList)), CategoryAttribute("02 - Functional"), DescriptionAttribute("Using this attribute, you can control how your data labels (x-axis labels) would appear on the chart. There are 4 options: WRAP, STAGGER, ROTATE or NONE. WRAP wraps the label text if it's longer than the allotted area. ROTATE rotates the label in vertical or slanted position. STAGGER divides the labels into multiple lines.")]
        public String LabelDisplay { get { return _labelDisplay; } set { _labelDisplay = value; } }
        [CategoryAttribute("02 - Functional"), DescriptionAttribute("This attribute lets you set whether the data labels would show up as rotated labels on the chart.")]
        public Boolean RotateLabels { get { return _rotateLabels; } set { _rotateLabels = value; } }
        [CategoryAttribute("02 - Functional"), DescriptionAttribute("If you've opted to show rotated labels on chart, this attribute lets you set the configuration whether the labels would show as slanted labels or fully vertical ones.")]
        public Boolean SlantLabels { get { return _slantLabels; } set { _slantLabels = value; } }
        [CategoryAttribute("02 - Functional"), DescriptionAttribute("By default, all the labels are displayed on the chart. However, if you've a set of streaming data (like name of months or days of week), you can opt to hide every n-th label for better clarity. This attributes just lets you do so. It allows to skip every n(th) X-axis label.")]
        public Double LabelStep { get { return _labelStep; } set { _labelStep = value; } }
        [CategoryAttribute("02 - Functional"), DescriptionAttribute("If you've opted for STAGGER mode as labelDisplay, using this attribute you can control how many lines to stagger the label to. By default, all labels are displayed in a single line.")]
        public Double StaggerLines { get { return _staggerLines; } set { _staggerLines = value; } }
        [CategoryAttribute("02 - Functional"), DescriptionAttribute("Sets the configuration whether data values would be displayed along with the data plot on chart.")]
        public Boolean ShowValues { get { return _showValues; } set { _showValues = value; } }
        [CategoryAttribute("02 - Functional"), DescriptionAttribute("If you've opted to show data values, you can rotate them using this attribute.")]
        public Boolean RotateValues { get { return _rotateValues; } set { _rotateValues = value; } }
        [CategoryAttribute("02 - Functional"), DescriptionAttribute("If you've opted to show data values, you can show them inside the columns using this attribute. By default, the data values show outside the column.")]
        public Boolean PlaceValuesInside { get { return _placeValuesInside; } set { _placeValuesInside = value; } }
        [CategoryAttribute("02 - Functional"), DescriptionAttribute("Charts y-axis is divided into vertical sections using div (divisional) lines. Each div line assumes a value based on its position. Using this attribute you can set whether to show those div line (y-axis) values or not.")]
        public Boolean ShowYAxisValues { get { return _showYAxisValues; } set { _showYAxisValues = value; } }
        [CategoryAttribute("02 - Functional"), DescriptionAttribute("Whether to show chart limit values? showYAxisValues is the single new attribute in which over-rides this value.")]
        public Boolean ShowLimits { get { return _showLimits; } set { _showLimits = value; } }
        [CategoryAttribute("02 - Functional"), DescriptionAttribute("Whether to show div line values? showYAxisValues is the single new attribute in which over-rides this value.")]
        public Boolean ShowDivLineValues { get { return _showDivLineValues; } set { _showDivLineValues = value; } }
        [CategoryAttribute("02 - Functional"), DescriptionAttribute("By default, all div lines show their values. However, you can opt to skip every x(th) div line value using this attribute.")]
        public Double YAxisValuesStep { get { return _yAxisValuesStep; } set { _yAxisValuesStep = value; } }
        [CategoryAttribute("02 - Functional"), DescriptionAttribute("Whether to show column shadows?")]
        public Boolean ShowShadow { get { return _showShadow; } set { _showShadow = value; } }
        [CategoryAttribute("02 - Functional"), DescriptionAttribute("Charts automatically tries to adjust divisional lines and limit values based on the data provided. However, if you want to set your explicit lower and upper limit values and number of divisional lines, first set this attribute to false. That would disable automatic adjustment of divisional lines.")]
        public Boolean AdjustDiv { get { return _adjustDiv; } set { _adjustDiv = value; } }
        [CategoryAttribute("02 - Functional"), DescriptionAttribute("If you do not wish to rotate y-axis name, set this as 0. It specifically comes to use when you've special characters (UTF8) in your y-axis name that do not show up in rotated mode.")]
        public Boolean RotateYAxisName { get { return _rotateYAxisName; } set { _rotateYAxisName = value; } }
        [CategoryAttribute("02 - Functional"), DescriptionAttribute("If you opt to not rotate y-axis name, you can choose a maximum width that will be applied to y-axis name.")]
        public Double YAxisNameWidth { get { return _yAxisNameWidth; } set { _yAxisNameWidth = value; } }
        [CategoryAttribute("02 - Functional"), DescriptionAttribute("The entire chart can now act as a hotspot. Use this URL to define the hotspot link for the chart. The link can be specified in Charts Link Format.")]
        public String ClickURL { get { return _clickURL; } set { _clickURL = value; } }
        [CategoryAttribute("02 - Functional"), DescriptionAttribute("By default, each chart animates some of its elements. If you wish to switch off the default animation patterns, you can set this attribute to 0. It can be particularly useful when you want to define your own animation patterns using STYLE feature.")]
        public Boolean DefaultAnimation { get { return _defaultAnimation; } set { _defaultAnimation = value; } }
        [CategoryAttribute("02 - Functional"), DescriptionAttribute("This attribute helps you explicitly set the lower limit of the chart. If you don't specify this value, it is automatically calculated by Charts based on the data provided by you.")]
        public Double YAxisMinValue { get { return _yAxisMinValue; } set { _yAxisMinValue = value; } }
        [CategoryAttribute("02 - Functional"), DescriptionAttribute("This attribute helps you explicitly set the upper limit of the chart. If you don't specify this value, it is automatically calculated by Charts based on the data provided by you.")]
        public Double YAxisMaxValue { get { return _yAxisMaxValue; } set { _yAxisMaxValue = value; } }
        [CategoryAttribute("02 - Functional"), DescriptionAttribute("This attribute lets you set whether the y-axis lower limit would be 0 (in case of all positive values on chart) or should the y-axis lower limit adapt itself to a different figure based on values provided to the chart.")]
        public Boolean SetAdaptiveYMin { get { return _setAdaptiveYMin; } set { _setAdaptiveYMin = value; } }


        #endregion Functional

        #region Chart Titles and Axis Names
        String _caption;
        String _subCaption;
        String _xAxisName;
        String _yAxisName;

        [CategoryAttribute("03 - Chart Titles and Axis Names"), DescriptionAttribute("Caption of the chart.")]
        public String Caption { get { return _caption; } set { _caption = value; } }
        [CategoryAttribute("03 - Chart Titles and Axis Names"), DescriptionAttribute("Sub-caption of the chart.")]
        public String SubCaption { get { return _subCaption; } set { _subCaption = value; } }
        [CategoryAttribute("03 - Chart Titles and Axis Names"), DescriptionAttribute("X-Axis Title of the Chart.")]
        public String XAxisName { get { return _xAxisName; } set { _xAxisName = value; } }
        [CategoryAttribute("03 - Chart Titles and Axis Names"), DescriptionAttribute("Y-Axis Title of the chart.")]
        public String YAxisName { get { return _yAxisName; } set { _yAxisName = value; } }

        #endregion Chart Titles and Axis Names

        #region Chart Cosmetics
        Color _bgColor;
        Double _bgAlpha;
        Double _bgRatio;
        Double _bgAngle;
        String _bgSWF;
        Double _bgSWFAlpha;
        Color _canvasBgColor;
        Double _canvasBgAlpha;
        Double _canvasBgRatio;
        Double _canvasBgAngle;
        Color _canvasBorderColor;
        Double _canvasBorderThickness;
        Double _canvasBorderAlpha;
        Boolean _showBorder;
        Color _borderColor;
        Double _borderThickness;
        Double _borderAlpha;

        [CategoryAttribute("04 - Chart Cosmetics"), DescriptionAttribute("This attribute sets the background color for the chart. You can set any hex color code as the value of this attribute. To specify a gradient as background color, separate the hex color codes of each color in the gradient using comma. Example: FF5904,FFFFFF. Remember to remove # and any spaces in between. See the gradient specification page for more details.")]
        public Color BgColor { get { return _bgColor; } set { _bgColor = value; } }
        [CategoryAttribute("04 - Chart Cosmetics"), DescriptionAttribute("Sets the alpha (transparency) for the background. If you've opted for gradient background, you need to set a list of alpha(s) separated by comma. See the gradient specification page for more details.")]
        public Double BgAlpha { get { return _bgAlpha; } set { _bgAlpha = value; } }
        [CategoryAttribute("04 - Chart Cosmetics"), DescriptionAttribute("If you've opted for a gradient background, this attribute lets you set the ratio of each color constituent. See the gradient specification page for more details.")]
        public Double BgRatio { get { return _bgRatio; } set { _bgRatio = value; } }
        [CategoryAttribute("04 - Chart Cosmetics"), DescriptionAttribute("Angle of the background color, in case of a gradient. See the gradient specification page for more details.")]
        public Double BgAngle { get { return _bgAngle; } set { _bgAngle = value; } }
        [CategoryAttribute("04 - Chart Cosmetics"), DescriptionAttribute("To place any Flash movie as background of the chart, enter the (path and) name of the background SWF. It should be in the same domain as the chart.")]
        public String BgSWF { get { return _bgSWF; } set { _bgSWF = value; } }
        [CategoryAttribute("04 - Chart Cosmetics"), DescriptionAttribute("Helps you specify alpha for the loaded background SWF.")]
        public Double BgSWFAlpha { get { return _bgSWFAlpha; } set { _bgSWFAlpha = value; } }
        [CategoryAttribute("04 - Chart Cosmetics"), DescriptionAttribute("Sets Canvas background color. For Gradient effect, enter colors separated by comma.")]
        public Color CanvasBgColor { get { return _canvasBgColor; } set { _canvasBgColor = value; } }
        [CategoryAttribute("04 - Chart Cosmetics"), DescriptionAttribute("Sets alpha for Canvas Background. For gradient, enter alpha list separated by commas.")]
        public Double CanvasBgAlpha { get { return _canvasBgAlpha; } set { _canvasBgAlpha = value; } }
        [CategoryAttribute("04 - Chart Cosmetics"), DescriptionAttribute("Helps you specify canvas background ratio for gradients.")]
        public Double CanvasBgRatio { get { return _canvasBgRatio; } set { _canvasBgRatio = value; } }
        [CategoryAttribute("04 - Chart Cosmetics"), DescriptionAttribute("Helps you specify canvas background angle in case of gradient.")]
        public Double CanvasBgAngle { get { return _canvasBgAngle; } set { _canvasBgAngle = value; } }
        [CategoryAttribute("04 - Chart Cosmetics"), DescriptionAttribute("Lets you specify canvas border color.")]
        public Color CanvasBorderColor { get { return _canvasBorderColor; } set { _canvasBorderColor = value; } }
        [CategoryAttribute("04 - Chart Cosmetics"), DescriptionAttribute("Lets you specify canvas border thickness.")]
        public Double CanvasBorderThickness { get { return _canvasBorderThickness; } set { _canvasBorderThickness = value; } }
        [CategoryAttribute("04 - Chart Cosmetics"), DescriptionAttribute("Lets you control transparency of canvas border.")]
        public Double CanvasBorderAlpha { get { return _canvasBorderAlpha; } set { _canvasBorderAlpha = value; } }
        [CategoryAttribute("04 - Chart Cosmetics"), DescriptionAttribute("Whether to show a border around the chart or not?")]
        public Boolean ShowBorder { get { return _showBorder; } set { _showBorder = value; } }
        [CategoryAttribute("04 - Chart Cosmetics"), DescriptionAttribute("Border color of the chart.")]
        public Color BorderColor { get { return _borderColor; } set { _borderColor = value; } }
        [CategoryAttribute("04 - Chart Cosmetics"), DescriptionAttribute("Border thickness of the chart.")]
        public Double BorderThickness { get { return _borderThickness; } set { _borderThickness = value; } }
        [CategoryAttribute("04 - Chart Cosmetics"), DescriptionAttribute("Border alpha of the chart.")]
        public Double BorderAlpha { get { return _borderAlpha; } set { _borderAlpha = value; } }

        #endregion Chart Cosmetics

        #region Data Plot Cosmetics
        Boolean _useRoundEdges = true;
        Boolean _showPlotBorder = true;
        Color _plotBorderColor;
        Double _plotBorderThickness;
        Double _plotBorderAlpha;
        Boolean _plotBorderDashed;
        Double _plotBorderDashLen;
        Double _plotBorderDashGap;
        Double _plotFillAngle;
        Double _plotFillRatio;
        Double _plotFillAlpha;
        Color _plotGradientColor;

        [CategoryAttribute("05 - Data Plot Cosmetics"), DescriptionAttribute("If you wish to plot columns with round edges and fill them with a glass effect gradient, set this attribute to 1. \nThe following functionalities wouldn't work when this attribute is set to 1: \nshowShadow attribute doesn't work any more. If you want to remove shadow from columns, you'll have to over-ride the shadow with a new shadow style (applied to DATAPLOT) with alpha as 0.\nPlot fill properties like gradient color, angle etc. wouldn't work any more, as the colors for gradient are now calculated by the chart itself.\nPlot border properties also do not work in this mode. Also, you cannot render the border as dash in this mode.")]
        public Boolean UseRoundEdges { get { return _useRoundEdges; } set { _useRoundEdges = value; } }
        [CategoryAttribute("05 - Data Plot Cosmetics"), DescriptionAttribute("Whether the column, area, pie etc. border would show up.")]
        public Boolean ShowPlotBorder { get { return _showPlotBorder; } set { _showPlotBorder = value; } }
        [CategoryAttribute("05 - Data Plot Cosmetics"), DescriptionAttribute("Color for column, area, pie border")]
        public Color PlotBorderColor { get { return _plotBorderColor; } set { _plotBorderColor = value; } }
        [CategoryAttribute("05 - Data Plot Cosmetics"), DescriptionAttribute("Thickness for column, area, pie border")]
        public Double PlotBorderThickness { get { return _plotBorderThickness; } set { _plotBorderThickness = value; } }
        [CategoryAttribute("05 - Data Plot Cosmetics"), DescriptionAttribute("Alpha for column, area, pie border")]
        public Double PlotBorderAlpha { get { return _plotBorderAlpha; } set { _plotBorderAlpha = value; } }
        [CategoryAttribute("05 - Data Plot Cosmetics"), DescriptionAttribute("Whether the plot border should appear as dashed?")]
        public Boolean PlotBorderDashed { get { return _plotBorderDashed; } set { _plotBorderDashed = value; } }
        [CategoryAttribute("05 - Data Plot Cosmetics"), DescriptionAttribute("If plot border is to appear as dash, this attribute lets you control the length of each dash.")]
        public Double PlotBorderDashLen { get { return _plotBorderDashLen; } set { _plotBorderDashLen = value; } }
        [CategoryAttribute("05 - Data Plot Cosmetics"), DescriptionAttribute("If plot border is to appear as dash, this attribute lets you control the length of each gap between two dash.")]
        public Double PlotBorderDashGap { get { return _plotBorderDashGap; } set { _plotBorderDashGap = value; } }
        [CategoryAttribute("05 - Data Plot Cosmetics"), DescriptionAttribute("If you've opted to fill the plot (column, area etc.) as gradient, this attribute lets you set the fill angle for gradient.")]
        public Double PlotFillAngle { get { return _plotFillAngle; } set { _plotFillAngle = value; } }
        [CategoryAttribute("05 - Data Plot Cosmetics"), DescriptionAttribute("If you've opted to fill the plot (column, area etc.) as gradient, this attribute lets you set the ratio for gradient.")]
        public Double PlotFillRatio { get { return _plotFillRatio; } set { _plotFillRatio = value; } }
        [CategoryAttribute("05 - Data Plot Cosmetics"), DescriptionAttribute("If you've opted to fill the plot (column, area etc.) as gradient, this attribute lets you set the fill alpha for gradient.")]
        public Double PlotFillAlpha { get { return _plotFillAlpha; } set { _plotFillAlpha = value; } }
        [CategoryAttribute("05 - Data Plot Cosmetics"), DescriptionAttribute("You can globally add a gradient color to the entire plot of chart by specifying the second color as this attribute. For example, if you've specified individual colors for your columns and now you want a gradient that ends in white. So, specify FFFFFF (white) as this color and the chart will now draw plots as gradient.")]
        public Color PlotGradientColor { get { return _plotGradientColor; } set { _plotGradientColor = value; } }

        #endregion Data Plot Cosmetics

        #region Tool-tip
        Boolean _showToolTip = true;
        Color _toolTipBgColor;
        Color _toolTipBorderColor;
        String _toolTipSepChar;

        [CategoryAttribute("07 - Tool-tip"), DescriptionAttribute("Whether to show tool tip on chart?")]
        public Boolean ShowToolTip { get { return _showToolTip; } set { _showToolTip = value; } }
        [CategoryAttribute("07 - Tool-tip"), DescriptionAttribute("Background Color for tool tip.")]
        public Color ToolTipBgColor { get { return _toolTipBgColor; } set { _toolTipBgColor = value; } }
        [CategoryAttribute("07 - Tool-tip"), DescriptionAttribute("Border Color for tool tip.")]
        public Color ToolTipBorderColor { get { return _toolTipBorderColor; } set { _toolTipBorderColor = value; } }
        [CategoryAttribute("07 - Tool-tip"), DescriptionAttribute("The character specified as the value of this attribute separates the name and value displayed in tool tip.")]
        public String ToolTipSepChar { get { return _toolTipSepChar; } set { _toolTipSepChar = value; } }

        #endregion Tool-tip

        #region Chart Padding & Margins
        Double _captionPadding;
        Double _xAxisNamePadding;
        Double _yAxisNamePadding;
        Double _yAxisValuesPadding;
        Double _labelPadding;
        Double _valuePadding;
        Double _plotSpacePercent;
        Double _chartLeftMargin;
        Double _chartRightMargin;
        Double _chartTopMargin;
        Double _chartBottomMargin;

        [CategoryAttribute("08 - Chart Padding & Margins"), DescriptionAttribute("This attribute lets you control the space (in pixels) between the sub-caption and top of the chart canvas. If the sub-caption is not defined, it controls the space between caption and top of chart canvas. If neither caption, nor sub-caption is defined, this padding does not come into play.")]
        public Double CaptionPadding { get { return _captionPadding; } set { _captionPadding = value; } }
        [CategoryAttribute("08 - Chart Padding & Margins"), DescriptionAttribute("Using this, you can set the distance between the top end of x-axis title and the bottom end of data labels (or canvas, if data labels are not to be shown).")]
        public Double XAxisNamePadding { get { return _xAxisNamePadding; } set { _xAxisNamePadding = value; } }
        [CategoryAttribute("08 - Chart Padding & Margins"), DescriptionAttribute("Using this, you can set the distance between the right end of y-axis title and the start of y-axis values (or canvas, if the y-axis values are not to be shown).")]
        public Double YAxisNamePadding { get { return _yAxisNamePadding; } set { _yAxisNamePadding = value; } }
        [CategoryAttribute("08 - Chart Padding & Margins"), DescriptionAttribute("This attribute helps you set the horizontal space between the canvas left edge and the y-axis values or trend line values (on left/right side). This is particularly useful, when you want more space between your canvas and y-axis values.")]
        public Double YAxisValuesPadding { get { return _yAxisValuesPadding; } set { _yAxisValuesPadding = value; } }
        [CategoryAttribute("08 - Chart Padding & Margins"), DescriptionAttribute("This attribute sets the vertical space between the labels and canvas bottom edge. If you want more space between the canvas and the x-axis labels, you can use this attribute to control it.")]
        public Double LabelPadding { get { return _labelPadding; } set { _labelPadding = value; } }
        [CategoryAttribute("08 - Chart Padding & Margins"), DescriptionAttribute("It sets the vertical space between the end of columns and start of value textboxes. This basically helps you control the space you want between your columns/anchors and the value textboxes.")]
        public Double ValuePadding { get { return _valuePadding; } set { _valuePadding = value; } }
        [CategoryAttribute("08 - Chart Padding & Margins"), DescriptionAttribute("On a column chart, there is spacing defined between two columns. By default, the spacing is set to 20% of canvas width. If you intend to increase or decrease the spacing between columns, you can do so using this attribute. For example, if you wanted all columns to stick to each other without any space in between, you can set plotSpacePercent to 0. Similarly, if you want very thin columns, you can set plotSpacePercent to its max value of 80.")]
        public Double PlotSpacePercent { get { return _plotSpacePercent; } set { _plotSpacePercent = value; } }
        [CategoryAttribute("08 - Chart Padding & Margins"), DescriptionAttribute("Amount of empty space that you want to put on the left side of your chart. Nothing is rendered in this space.")]
        public Double ChartLeftMargin { get { return _chartLeftMargin; } set { _chartLeftMargin = value; } }
        [CategoryAttribute("08 - Chart Padding & Margins"), DescriptionAttribute("Amount of empty space that you want to put on the right side of your chart. Nothing is rendered in this space.")]
        public Double ChartRightMargin { get { return _chartRightMargin; } set { _chartRightMargin = value; } }
        [CategoryAttribute("08 - Chart Padding & Margins"), DescriptionAttribute("Amount of empty space that you want to put on the top of your chart. Nothing is rendered in this space.")]
        public Double ChartTopMargin { get { return _chartTopMargin; } set { _chartTopMargin = value; } }
        [CategoryAttribute("08 - Chart Padding & Margins"), DescriptionAttribute("Amount of empty space that you want to put on the bottom of your chart. Nothing is rendered in this space.")]
        public Double ChartBottomMargin { get { return _chartBottomMargin; } set { _chartBottomMargin = value; } }

        #endregion Chart Padding & Margins

        #region Number Formatting
        Boolean _formatNumber = true;
        Boolean _formatNumberScale;
        String _defaultNumberScale;
        String _numberScaleUnit;
        String _numberScaleValue;
        String _numberPrefix;
        String _numberSuffix;
        String _decimalSeparator;
        String _thousandSeparator;
        String _inDecimalSeparator;
        String _inThousandSeparator;
        Double _decimals;
        Boolean _forceDecimals;
        Double _yAxisValueDecimals;

        [CategoryAttribute("09 - Number Formatting"), DescriptionAttribute("This configuration determines whether the numbers displayed on the chart will be formatted using commas, e.g., 40,000 if formatNumber='1' and 40000 if formatNumber='0 '")]
        public Boolean FormatNumber { get { return _formatNumber; } set { _formatNumber = value; } }
        [CategoryAttribute("09 - Number Formatting"), DescriptionAttribute("Configuration whether to add K (thousands) and M (millions) to a number after truncating and rounding it - e.g., if formatNumberScale is set to 1, 1043 would become 1.04K (with decimals set to 2 places). Same with numbers in millions - a M will added at the end. For more details, please see Advanced Number Formatting section.")]
        public Boolean FormatNumberScale { get { return _formatNumberScale; } set { _formatNumberScale = value; } }
        [CategoryAttribute("09 - Number Formatting"), DescriptionAttribute("The default unit of the numbers that you're providing to the chart. For more details, please see Advanced Number Formatting section.")]
        public String DefaultNumberScale { get { return _defaultNumberScale; } set { _defaultNumberScale = value; } }
        [CategoryAttribute("09 - Number Formatting"), DescriptionAttribute("Unit of each block of the scale. For more details, please see Advanced Number Formatting section.")]
        public String NumberScaleUnit { get { return _numberScaleUnit; } set { _numberScaleUnit = value; } }
        [CategoryAttribute("09 - Number Formatting"), DescriptionAttribute("Range of the various blocks that constitute the scale. For more details, please see Advanced Number Formatting section.")]
        public String NumberScaleValue { get { return _numberScaleValue; } set { _numberScaleValue = value; } }
        [CategoryAttribute("09 - Number Formatting"), DescriptionAttribute("Using this attribute, you could add prefix to all the numbers visible on the graph. For example, to represent all dollars figure on the chart, you could specify this attribute to ' $' to show like $40000, $50000. For more details, please see Advanced Number Formatting section.")]
        public String NumberPrefix { get { return _numberPrefix; } set { _numberPrefix = value; } }
        [CategoryAttribute("09 - Number Formatting"), DescriptionAttribute("Using this attribute, you could add suffix to all the numbers visible on the graph. For example, to represent all figure quantified as per annum on the chart, you could specify this attribute to ' /a' to show like 40000/a, 50000/a. For more details, please see Advanced Number Formatting section.")]
        public String NumberSuffix { get { return _numberSuffix; } set { _numberSuffix = value; } }
        [CategoryAttribute("09 - Number Formatting"), DescriptionAttribute("This option helps you specify the character to be used as the decimal separator in a number. For more details, please see Advanced Number Formatting section.")]
        public String DecimalSeparator { get { return _decimalSeparator; } set { _decimalSeparator = value; } }
        [CategoryAttribute("09 - Number Formatting"), DescriptionAttribute("This option helps you specify the character to be used as the thousands separator in a number. For more details, please see Advanced Number Formatting section.")]
        public String ThousandSeparator { get { return _thousandSeparator; } set { _thousandSeparator = value; } }
        [CategoryAttribute("09 - Number Formatting"), DescriptionAttribute("In some countries, commas are used as decimal separators and dots as thousand separators. In XML, if you specify such values, it will give a error while converting to number. So, Charts accepts the input decimal and thousand separator from user, so that it can covert it accordingly into the required format. This attribute lets you input the decimal separator. For more details, please see Advanced Number Formatting section.")]
        public String InDecimalSeparator { get { return _inDecimalSeparator; } set { _inDecimalSeparator = value; } }
        [CategoryAttribute("09 - Number Formatting"), DescriptionAttribute("In some countries, commas are used as decimal separators and dots as thousand separators. In XML, if you specify such values, it will give a error while converting to number. So, Charts accepts the input decimal and thousand separator from user, so that it can covert it accordingly into the required format. This attribute lets you input the thousand separator. For more details, please see Advanced Number Formatting section.")]
        public String InThousandSeparator { get { return _inThousandSeparator; } set { _inThousandSeparator = value; } }
        [CategoryAttribute("09 - Number Formatting"), DescriptionAttribute("Number of decimal places to which all numbers on the chart would be rounded to.")]
        public Double Decimals { get { return _decimals; } set { _decimals = value; } }
        [CategoryAttribute("09 - Number Formatting"), DescriptionAttribute("Whether to add 0 padding at the end of decimal numbers? For example, if you set decimals as 2 and a number is 23.4. If forceDecimals is set to 1, Charts will convert the number to 23.40 (note the extra 0 at the end)")]
        public Boolean ForceDecimals { get { return _forceDecimals; } set { _forceDecimals = value; } }
        [CategoryAttribute("09 - Number Formatting"), DescriptionAttribute("If you've opted to not adjust div lines, you can specify the div line values decimal precision using this attribute.")]
        public Double YAxisValueDecimals { get { return _yAxisValueDecimals; } set { _yAxisValueDecimals = value; } }

        #endregion Number Formatting

        #region Font Properties
        Font _baseFont;
        Double _baseFontSize;
        Color _baseFontColor;
        String _outCnvBaseFont;
        Double _outCnvBaseFontSize;
        Color _outCnvBaseFontColor;

        [CategoryAttribute("10 - Font Properties"), DescriptionAttribute("This attribute lets you set the font face (family) of all the text (data labels, values etc.) on chart. If you specify outCnvBaseFont attribute also, then this attribute controls only the font face of text within the chart canvas bounds.")]
        public Font BaseFont { get { return _baseFont; } set { _baseFont = value; } }
        [CategoryAttribute("10 - Font Properties"), DescriptionAttribute("This attribute sets the base font size of the chart i.e., all the values and the names in the chart which lie on the canvas will be displayed using the font size provided here.")]
        public Double BaseFontSize { get { return _baseFontSize; } set { _baseFontSize = value; } }
        [CategoryAttribute("10 - Font Properties"), DescriptionAttribute("This attribute sets the base font color of the chart i.e., all the values and the names in the chart which lie on the canvas will be displayed using the font color provided here.")]
        public Color BaseFontColor { get { return _baseFontColor; } set { _baseFontColor = value; } }
        [CategoryAttribute("10 - Font Properties"), DescriptionAttribute("This attribute sets the base font family of the chart font which lies outside the canvas i.e., all the values and the names in the chart which lie outside the canvas will be displayed using the font name provided here.")]
        public String OutCnvBaseFont { get { return _outCnvBaseFont; } set { _outCnvBaseFont = value; } }
        [CategoryAttribute("10 - Font Properties"), DescriptionAttribute("This attribute sets the base font size of the chart i.e., all the values and the names in the chart which lie outside the canvas will be displayed using the font size provided here.")]
        public Double OutCnvBaseFontSize { get { return _outCnvBaseFontSize; } set { _outCnvBaseFontSize = value; } }
        [CategoryAttribute("10 - Font Properties"), DescriptionAttribute("This attribute sets the base font color of the chart i.e., all the values and the names in the chart which lie outside the canvas will be displayed using the font color provided here.")]
        public Color OutCnvBaseFontColor { get { return _outCnvBaseFontColor; } set { _outCnvBaseFontColor = value; } }

        #endregion Font Properties


        public clsChartProperty()
        {
        }
        public void ReadProperty(string property)
        {
            MatchCollection mathcoll = Regex.Matches(property, @"[a-zA-Z]+='.*?' *");
            foreach (Match x in mathcoll)
            {
                string tmp = x.Value.Trim();
                if (tmp.Contains("chartname")) _chartname = Convert.ToString(tmp.Replace("chartname=", "").Replace("'", ""));
                else if (tmp.Contains("sheetChart")) _sheetChart = Convert.ToString(tmp.Replace("sheetChart=", "").Replace("'", ""));
                else if (tmp.Contains("dataRange")) _dataRange = Convert.ToString(tmp.Replace("dataRange=", "").Replace("'", ""));
                else if (tmp.Contains("captionRange")) _captionRange = Convert.ToString(tmp.Replace("captionRange=", "").Replace("'", ""));
                else if (tmp.Contains("subCaptionRange")) _subCaptionRange = Convert.ToString(tmp.Replace("subCaptionRange=", "").Replace("'", ""));
                else if (tmp.Contains("animation")) _animation = tmp.Replace("animation=", "").Replace("'", "") == "1" ? true : false;
                else if (tmp.Contains("palette")) _palette = Convert.ToDouble(tmp.Replace("palette=", "").Replace("'", ""));
                else if (tmp.Contains("showLabels")) _showLabels = tmp.Replace("showLabels=", "").Replace("'", "") == "1" ? true : false;
                else if (tmp.Contains("labelDisplay")) _labelDisplay = Convert.ToString(tmp.Replace("labelDisplay=", "").Replace("'", ""));
                else if (tmp.Contains("rotateLabels")) _rotateLabels = tmp.Replace("rotateLabels=", "").Replace("'", "") == "1" ? true : false;
                else if (tmp.Contains("slantLabels")) _slantLabels = tmp.Replace("slantLabels=", "").Replace("'", "") == "1" ? true : false;
                else if (tmp.Contains("labelStep")) _labelStep = Convert.ToDouble(tmp.Replace("labelStep=", "").Replace("'", ""));
                else if (tmp.Contains("staggerLines")) _staggerLines = Convert.ToDouble(tmp.Replace("staggerLines=", "").Replace("'", ""));
                else if (tmp.Contains("showValues")) _showValues = tmp.Replace("showValues=", "").Replace("'", "") == "1" ? true : false;
                else if (tmp.Contains("rotateValues")) _rotateValues = tmp.Replace("rotateValues=", "").Replace("'", "") == "1" ? true : false;
                else if (tmp.Contains("placeValuesInside")) _placeValuesInside = tmp.Replace("placeValuesInside=", "").Replace("'", "") == "1" ? true : false;
                else if (tmp.Contains("showYAxisValues")) _showYAxisValues = tmp.Replace("showYAxisValues=", "").Replace("'", "") == "1" ? true : false;
                else if (tmp.Contains("showLimits")) _showLimits = tmp.Replace("showLimits=", "").Replace("'", "") == "1" ? true : false;
                else if (tmp.Contains("showDivLineValues")) _showDivLineValues = tmp.Replace("showDivLineValues=", "").Replace("'", "") == "1" ? true : false;
                else if (tmp.Contains("yAxisValuesStep")) _yAxisValuesStep = Convert.ToDouble(tmp.Replace("yAxisValuesStep=", "").Replace("'", ""));
                else if (tmp.Contains("showShadow")) _showShadow = tmp.Replace("showShadow=", "").Replace("'", "") == "1" ? true : false;
                else if (tmp.Contains("adjustDiv")) _adjustDiv = tmp.Replace("adjustDiv=", "").Replace("'", "") == "1" ? true : false;
                else if (tmp.Contains("rotateYAxisName")) _rotateYAxisName = tmp.Replace("rotateYAxisName=", "").Replace("'", "") == "1" ? true : false;
                else if (tmp.Contains("yAxisNameWidth")) _yAxisNameWidth = Convert.ToDouble(tmp.Replace("yAxisNameWidth=", "").Replace("'", ""));
                else if (tmp.Contains("clickURL")) _clickURL = Convert.ToString(tmp.Replace("clickURL=", "").Replace("'", ""));
                else if (tmp.Contains("defaultAnimation")) _defaultAnimation = tmp.Replace("defaultAnimation=", "").Replace("'", "") == "1" ? true : false;
                else if (tmp.Contains("yAxisMinValue")) _yAxisMinValue = Convert.ToDouble(tmp.Replace("yAxisMinValue=", "").Replace("'", ""));
                else if (tmp.Contains("yAxisMaxValue")) _yAxisMaxValue = Convert.ToDouble(tmp.Replace("yAxisMaxValue=", "").Replace("'", ""));
                else if (tmp.Contains("setAdaptiveYMin")) _setAdaptiveYMin = tmp.Replace("setAdaptiveYMin=", "").Replace("'", "") == "1" ? true : false;
                    /*
                else if (tmp.Contains("showAboutMenuItem")) _showAboutMenuItem = tmp.Replace("showAboutMenuItem=", "").Replace("'", "") == "1" ? true : false;
                else if (tmp.Contains("aboutMenuItemLabel")) _aboutMenuItemLabel = tmp.Replace("aboutMenuItemLabel=", "").Replace("'", "");
                else if (tmp.Contains("aboutMenuItemLink")) _aboutMenuItemLink = tmp.Replace("aboutMenuItemLink=", "").Replace("'", "");
                    */
                else if (tmp.Contains("caption")) _caption = Convert.ToString(tmp.Replace("caption=", "").Replace("'", ""));
                else if (tmp.Contains("subCaption")) _subCaption = Convert.ToString(tmp.Replace("subCaption=", "").Replace("'", ""));
                else if (tmp.Contains("xAxisName")) _xAxisName = Convert.ToString(tmp.Replace("xAxisName=", "").Replace("'", ""));
                else if (tmp.Contains("yAxisName")) _yAxisName = Convert.ToString(tmp.Replace("yAxisName=", "").Replace("'", ""));

                else if (tmp.Contains("bgColor")) _bgColor = System.Drawing.ColorTranslator.FromHtml("#" + (tmp.Replace("bgColor=", "").Replace("'", "")));
                else if (tmp.Contains("bgAlpha")) _bgAlpha = Convert.ToDouble(tmp.Replace("bgAlpha=", "").Replace("'", ""));
                else if (tmp.Contains("bgRatio")) _bgRatio = Convert.ToDouble(tmp.Replace("bgRatio=", "").Replace("'", ""));
                else if (tmp.Contains("bgAngle")) _bgAngle = Convert.ToDouble(tmp.Replace("bgAngle=", "").Replace("'", ""));
                else if (tmp.Contains("bgSWF")) _bgSWF = Convert.ToString(tmp.Replace("bgSWF=", "").Replace("'", ""));
                else if (tmp.Contains("bgSWFAlpha")) _bgSWFAlpha = Convert.ToDouble(tmp.Replace("bgSWFAlpha=", "").Replace("'", ""));
                else if (tmp.Contains("canvasBgColor")) _canvasBgColor = System.Drawing.ColorTranslator.FromHtml("#" + (tmp.Replace("canvasBgColor=", "").Replace("'", "")));
                else if (tmp.Contains("canvasBgAlpha")) _canvasBgAlpha = Convert.ToDouble(tmp.Replace("canvasBgAlpha=", "").Replace("'", ""));
                else if (tmp.Contains("canvasBgRatio")) _canvasBgRatio = Convert.ToDouble(tmp.Replace("canvasBgRatio=", "").Replace("'", ""));
                else if (tmp.Contains("canvasBgAngle")) _canvasBgAngle = Convert.ToDouble(tmp.Replace("canvasBgAngle=", "").Replace("'", ""));
                else if (tmp.Contains("canvasBorderColor")) _canvasBorderColor = System.Drawing.ColorTranslator.FromHtml("#" + (tmp.Replace("canvasBorderColor=", "").Replace("'", "")));
                else if (tmp.Contains("canvasBorderThickness")) _canvasBorderThickness = Convert.ToDouble(tmp.Replace("canvasBorderThickness=", "").Replace("'", ""));
                else if (tmp.Contains("canvasBorderAlpha")) _canvasBorderAlpha = Convert.ToDouble(tmp.Replace("canvasBorderAlpha=", "").Replace("'", ""));
                else if (tmp.Contains("showBorder")) _showBorder = tmp.Replace("showBorder=", "").Replace("'", "") == "1" ? true : false;
                else if (tmp.Contains("borderColor")) _borderColor = System.Drawing.ColorTranslator.FromHtml("#" + (tmp.Replace("borderColor=", "").Replace("'", "")));
                else if (tmp.Contains("borderThickness")) _borderThickness = Convert.ToDouble(tmp.Replace("borderThickness=", "").Replace("'", ""));
                else if (tmp.Contains("borderAlpha")) _borderAlpha = Convert.ToDouble(tmp.Replace("borderAlpha=", "").Replace("'", ""));

                else if (tmp.Contains("useRoundEdges")) _useRoundEdges = tmp.Replace("useRoundEdges=", "").Replace("'", "") == "1" ? true : false;
                else if (tmp.Contains("showPlotBorder")) _showPlotBorder = tmp.Replace("showPlotBorder=", "").Replace("'", "") == "1" ? true : false;
                else if (tmp.Contains("plotBorderColor")) _plotBorderColor = System.Drawing.ColorTranslator.FromHtml("#" + (tmp.Replace("plotBorderColor=", "").Replace("'", "")));
                else if (tmp.Contains("plotBorderThickness")) _plotBorderThickness = Convert.ToDouble(tmp.Replace("plotBorderThickness=", "").Replace("'", ""));
                else if (tmp.Contains("plotBorderAlpha")) _plotBorderAlpha = Convert.ToDouble(tmp.Replace("plotBorderAlpha=", "").Replace("'", ""));
                else if (tmp.Contains("plotBorderDashed")) _plotBorderDashed = tmp.Replace("plotBorderDashed=", "").Replace("'", "") == "1" ? true : false;
                else if (tmp.Contains("plotBorderDashLen")) _plotBorderDashLen = Convert.ToDouble(tmp.Replace("plotBorderDashLen=", "").Replace("'", ""));
                else if (tmp.Contains("plotBorderDashGap")) _plotBorderDashGap = Convert.ToDouble(tmp.Replace("plotBorderDashGap=", "").Replace("'", ""));
                else if (tmp.Contains("plotFillAngle")) _plotFillAngle = Convert.ToDouble(tmp.Replace("plotFillAngle=", "").Replace("'", ""));
                else if (tmp.Contains("plotFillRatio")) _plotFillRatio = Convert.ToDouble(tmp.Replace("plotFillRatio=", "").Replace("'", ""));
                else if (tmp.Contains("plotFillAlpha")) _plotFillAlpha = Convert.ToDouble(tmp.Replace("plotFillAlpha=", "").Replace("'", ""));
                else if (tmp.Contains("plotGradientColor")) _plotGradientColor = System.Drawing.ColorTranslator.FromHtml("#" + (tmp.Replace("plotGradientColor=", "").Replace("'", "")));

                else if (tmp.Contains("showToolTip")) _showToolTip = tmp.Replace("showToolTip=", "").Replace("'", "") == "1" ? true : false;
                else if (tmp.Contains("toolTipBgColor")) _toolTipBgColor = System.Drawing.ColorTranslator.FromHtml("#" + (tmp.Replace("toolTipBgColor=", "").Replace("'", "")));
                else if (tmp.Contains("toolTipBorderColor")) _toolTipBorderColor = System.Drawing.ColorTranslator.FromHtml("#" + (tmp.Replace("toolTipBorderColor=", "").Replace("'", "")));
                else if (tmp.Contains("toolTipSepChar")) _toolTipSepChar = Convert.ToString(tmp.Replace("toolTipSepChar=", "").Replace("'", ""));

                else if (tmp.Contains("captionPadding")) _captionPadding = Convert.ToDouble(tmp.Replace("captionPadding=", "").Replace("'", ""));
                else if (tmp.Contains("xAxisNamePadding")) _xAxisNamePadding = Convert.ToDouble(tmp.Replace("xAxisNamePadding=", "").Replace("'", ""));
                else if (tmp.Contains("yAxisNamePadding")) _yAxisNamePadding = Convert.ToDouble(tmp.Replace("yAxisNamePadding=", "").Replace("'", ""));
                else if (tmp.Contains("yAxisValuesPadding")) _yAxisValuesPadding = Convert.ToDouble(tmp.Replace("yAxisValuesPadding=", "").Replace("'", ""));
                else if (tmp.Contains("labelPadding")) _labelPadding = Convert.ToDouble(tmp.Replace("labelPadding=", "").Replace("'", ""));
                else if (tmp.Contains("valuePadding")) _valuePadding = Convert.ToDouble(tmp.Replace("valuePadding=", "").Replace("'", ""));
                else if (tmp.Contains("plotSpacePercent")) _plotSpacePercent = Convert.ToDouble(tmp.Replace("plotSpacePercent=", "").Replace("'", ""));
                else if (tmp.Contains("chartLeftMargin")) _chartLeftMargin = Convert.ToDouble(tmp.Replace("chartLeftMargin=", "").Replace("'", ""));
                else if (tmp.Contains("chartRightMargin")) _chartRightMargin = Convert.ToDouble(tmp.Replace("chartRightMargin=", "").Replace("'", ""));
                else if (tmp.Contains("chartTopMargin")) _chartTopMargin = Convert.ToDouble(tmp.Replace("chartTopMargin=", "").Replace("'", ""));
                else if (tmp.Contains("chartBottomMargin")) _chartBottomMargin = Convert.ToDouble(tmp.Replace("chartBottomMargin=", "").Replace("'", ""));

                else if (tmp.Contains("formatNumber")) _formatNumber = tmp.Replace("formatNumber=", "").Replace("'", "") == "1" ? true : false;
                else if (tmp.Contains("formatNumberScale")) _formatNumberScale = tmp.Replace("formatNumberScale=", "").Replace("'", "") == "1" ? true : false;
                else if (tmp.Contains("defaultNumberScale")) _defaultNumberScale = Convert.ToString(tmp.Replace("defaultNumberScale=", "").Replace("'", ""));
                else if (tmp.Contains("numberScaleUnit")) _numberScaleUnit = Convert.ToString(tmp.Replace("numberScaleUnit=", "").Replace("'", ""));
                else if (tmp.Contains("numberScaleValue")) _numberScaleValue = Convert.ToString(tmp.Replace("numberScaleValue=", "").Replace("'", ""));
                else if (tmp.Contains("numberPrefix")) _numberPrefix = Convert.ToString(tmp.Replace("numberPrefix=", "").Replace("'", ""));
                else if (tmp.Contains("numberSuffix")) _numberSuffix = Convert.ToString(tmp.Replace("numberSuffix=", "").Replace("'", ""));
                else if (tmp.Contains("decimalSeparator")) _decimalSeparator = Convert.ToString(tmp.Replace("decimalSeparator=", "").Replace("'", ""));
                else if (tmp.Contains("thousandSeparator")) _thousandSeparator = Convert.ToString(tmp.Replace("thousandSeparator=", "").Replace("'", ""));
                else if (tmp.Contains("inDecimalSeparator")) _inDecimalSeparator = Convert.ToString(tmp.Replace("inDecimalSeparator=", "").Replace("'", ""));
                else if (tmp.Contains("inThousandSeparator")) _inThousandSeparator = Convert.ToString(tmp.Replace("inThousandSeparator=", "").Replace("'", ""));
                else if (tmp.Contains("decimals")) _decimals = Convert.ToDouble(tmp.Replace("decimals=", "").Replace("'", ""));
                else if (tmp.Contains("forceDecimals")) _forceDecimals = tmp.Replace("forceDecimals=", "").Replace("'", "") == "1" ? true : false;
                else if (tmp.Contains("yAxisValueDecimals")) _yAxisValueDecimals = Convert.ToDouble(tmp.Replace("yAxisValueDecimals=", "").Replace("'", ""));

                else if (tmp.Contains("baseFontSize")) _baseFontSize = Convert.ToDouble(tmp.Replace("baseFontSize=", "").Replace("'", ""));
                else if (tmp.Contains("baseFont")) _baseFont = new Font((tmp.Replace("baseFont=", "").Replace("'", "")), (float)_baseFontSize);
                else if (tmp.Contains("baseFontColor")) _baseFontColor = System.Drawing.ColorTranslator.FromHtml("#" + (tmp.Replace("baseFontColor=", "").Replace("'", "")));
                else if (tmp.Contains("outCnvBaseFont")) _outCnvBaseFont = Convert.ToString(tmp.Replace("outCnvBaseFont=", "").Replace("'", ""));
                else if (tmp.Contains("outCnvBaseFontSize")) _outCnvBaseFontSize = Convert.ToDouble(tmp.Replace("outCnvBaseFontSize=", "").Replace("'", ""));
                else if (tmp.Contains("outCnvBaseFontColor")) _outCnvBaseFontColor = System.Drawing.ColorTranslator.FromHtml("#" + (tmp.Replace("outCnvBaseFontColor=", "").Replace("'", "")));
            }
        }
        public string GetProperty()
        {
            string result = "";
            if (_chartname != null) result += " chartname='" + _chartname.ToString() + "'";
            if (_sheetChart != null) result += " sheetChart='" + _sheetChart.ToString() + "'";
            if (_dataRange != null) result += " dataRange='" + _dataRange.ToString() + "'";
            if (_captionRange != null) result += " captionRange='" + _captionRange.ToString() + "'";
            if (_subCaptionRange != null) result += " subCaptionRange='" + _subCaption.ToString() + "'";
            result = GetPropertyForChart(result);

            //if (result != "")
            //    result.Substring(1);
            return result;
        }

        public string GetPropertyForChart(string result)
        {
            result += " animation='" + Convert.ToInt32(_animation) + "'";
            if (_palette != 0) result += " palette='" + _palette.ToString() + "'";
            result += " showLabels='" + Convert.ToInt32(_showLabels) + "'";
            if (_labelDisplay != null) result += " labelDisplay='" + _labelDisplay.ToString() + "'";
            result += " rotateLabels='" + Convert.ToInt32(_rotateLabels) + "'";
            result += " slantLabels='" + Convert.ToInt32(_slantLabels) + "'";
            if (_labelStep != 0) result += " labelStep='" + _labelStep.ToString() + "'";
            if (_staggerLines != 0) result += " staggerLines='" + _staggerLines.ToString() + "'";
            result += " showValues='" + Convert.ToInt32(_showValues) + "'";
            result += " rotateValues='" + Convert.ToInt32(_rotateValues) + "'";
            result += " placeValuesInside='" + Convert.ToInt32(_placeValuesInside) + "'";
            result += " showYAxisValues='" + Convert.ToInt32(_showYAxisValues) + "'";
            result += " showLimits='" + Convert.ToInt32(_showLimits) + "'";
            result += " showDivLineValues='" + Convert.ToInt32(_showDivLineValues) + "'";
            if (_yAxisValuesStep != 0) result += " yAxisValuesStep='" + _yAxisValuesStep.ToString() + "'";
            result += " showShadow='" + Convert.ToInt32(_showShadow) + "'";
            result += " adjustDiv='" + Convert.ToInt32(_adjustDiv) + "'";
            result += " rotateYAxisName='" + Convert.ToInt32(_rotateYAxisName) + "'";
            if (_yAxisNameWidth != 0) result += " yAxisNameWidth='" + _yAxisNameWidth.ToString() + "'";
            if (_clickURL != null) result += " clickURL='" + _clickURL.ToString() + "'";
            result += " defaultAnimation='" + Convert.ToInt32(_defaultAnimation) + "'";
            if (_yAxisMinValue != 0) result += " yAxisMinValue='" + _yAxisMinValue.ToString() + "'";
            if (_yAxisMaxValue != 0) result += " yAxisMaxValue='" + _yAxisMaxValue.ToString() + "'";
            result += " setAdaptiveYMin='" + Convert.ToInt32(_setAdaptiveYMin) + "'";
          /*  result += " showAboutMenuItem='" + Convert.ToInt32(_showAboutMenuItem) + "'";
            if (_aboutMenuItemLink != null) result += " aboutMenuItemLink='" + _aboutMenuItemLink.ToString() + "'";
            if (_aboutMenuItemLabel != null) result += " aboutMenuItemLabel='" + _aboutMenuItemLabel.ToString() + "'";
            */
            if (_caption != null) result += " caption='" + _caption.ToString() + "'";
            if (_subCaption != null) result += " subCaption='" + _subCaption.ToString() + "'";
            if (_xAxisName != null) result += " xAxisName='" + _xAxisName.ToString() + "'";
            if (_yAxisName != null) result += " yAxisName='" + _yAxisName.ToString() + "'";

            if (_bgColor != Color.Empty) result += " bgColor='" + ColorToHexString(_bgColor) + "'";
            if (_bgAlpha != 0) result += " bgAlpha='" + _bgAlpha.ToString() + "'";
            if (_bgRatio != 0) result += " bgRatio='" + _bgRatio.ToString() + "'";
            if (_bgAngle != 0) result += " bgAngle='" + _bgAngle.ToString() + "'";
            if (_bgSWF != null) result += " bgSWF='" + _bgSWF.ToString() + "'";
            if (_bgSWFAlpha != 0) result += " bgSWFAlpha='" + _bgSWFAlpha.ToString() + "'";
            if (_canvasBgColor != Color.Empty) result += " canvasBgColor='" + ColorToHexString(_canvasBgColor) + "'";
            if (_canvasBgAlpha != 0) result += " canvasBgAlpha='" + _canvasBgAlpha.ToString() + "'";
            if (_canvasBgRatio != 0) result += " canvasBgRatio='" + _canvasBgRatio.ToString() + "'";
            if (_canvasBgAngle != 0) result += " canvasBgAngle='" + _canvasBgAngle.ToString() + "'";
            if (_canvasBorderColor != Color.Empty) result += " canvasBorderColor='" + ColorToHexString(_canvasBorderColor) + "'";
            if (_canvasBorderThickness != 0) result += " canvasBorderThickness='" + _canvasBorderThickness.ToString() + "'";
            if (_canvasBorderAlpha != 0) result += " canvasBorderAlpha='" + _canvasBorderAlpha.ToString() + "'";
            result += " showBorder='" + Convert.ToInt32(_showBorder) + "'";
            if (_borderColor != Color.Empty) result += " borderColor='" + ColorToHexString(_borderColor) + "'";
            if (_borderThickness != 0) result += " borderThickness='" + _borderThickness.ToString() + "'";
            if (_borderAlpha != 0) result += " borderAlpha='" + _borderAlpha.ToString() + "'";

            result += " useRoundEdges='" + Convert.ToInt32(_useRoundEdges) + "'";
            result += " showPlotBorder='" + Convert.ToInt32(_showPlotBorder) + "'";
            if (_plotBorderColor != Color.Empty) result += " plotBorderColor='" + ColorToHexString(_plotBorderColor) + "'";
            if (_plotBorderThickness != 0) result += " plotBorderThickness='" + _plotBorderThickness.ToString() + "'";
            if (_plotBorderAlpha != 0) result += " plotBorderAlpha='" + _plotBorderAlpha.ToString() + "'";
            result += " plotBorderDashed='" + Convert.ToInt32(_plotBorderDashed) + "'";
            if (_plotBorderDashLen != 0) result += " plotBorderDashLen='" + _plotBorderDashLen.ToString() + "'";
            if (_plotBorderDashGap != 0) result += " plotBorderDashGap='" + _plotBorderDashGap.ToString() + "'";
            if (_plotFillAngle != 0) result += " plotFillAngle='" + _plotFillAngle.ToString() + "'";
            if (_plotFillRatio != 0) result += " plotFillRatio='" + _plotFillRatio.ToString() + "'";
            if (_plotFillAlpha != 0) result += " plotFillAlpha='" + _plotFillAlpha.ToString() + "'";
            if (_plotGradientColor != Color.Empty) result += " plotGradientColor='" + ColorToHexString(_plotGradientColor) + "'";

            result += " showToolTip='" + Convert.ToInt32(_showToolTip) + "'";
            if (_toolTipBgColor != Color.Empty) result += " toolTipBgColor='" + ColorToHexString(_toolTipBgColor) + "'";
            if (_toolTipBorderColor != Color.Empty) result += " toolTipBorderColor='" + ColorToHexString(_toolTipBorderColor) + "'";
            if (_toolTipSepChar != null) result += " toolTipSepChar='" + _toolTipSepChar.ToString() + "'";

            if (_captionPadding != 0) result += " captionPadding='" + _captionPadding.ToString() + "'";
            if (_xAxisNamePadding != 0) result += " xAxisNamePadding='" + _xAxisNamePadding.ToString() + "'";
            if (_yAxisNamePadding != 0) result += " yAxisNamePadding='" + _yAxisNamePadding.ToString() + "'";
            if (_yAxisValuesPadding != 0) result += " yAxisValuesPadding='" + _yAxisValuesPadding.ToString() + "'";
            if (_labelPadding != 0) result += " labelPadding='" + _labelPadding.ToString() + "'";
            if (_valuePadding != 0) result += " valuePadding='" + _valuePadding.ToString() + "'";
            if (_plotSpacePercent != 0) result += " plotSpacePercent='" + _plotSpacePercent.ToString() + "'";
            if (_chartLeftMargin != 0) result += " chartLeftMargin='" + _chartLeftMargin.ToString() + "'";
            if (_chartRightMargin != 0) result += " chartRightMargin='" + _chartRightMargin.ToString() + "'";
            if (_chartTopMargin != 0) result += " chartTopMargin='" + _chartTopMargin.ToString() + "'";
            if (_chartBottomMargin != 0) result += " chartBottomMargin='" + _chartBottomMargin.ToString() + "'";

            result += " formatNumber='" + Convert.ToInt32(_formatNumber) + "'";
            result += " formatNumberScale='" + Convert.ToInt32(_formatNumberScale) + "'";
            if (_defaultNumberScale != null) result += " defaultNumberScale='" + _defaultNumberScale.ToString() + "'";
            if (_numberScaleUnit != null) result += " numberScaleUnit='" + _numberScaleUnit.ToString() + "'";
            if (_numberScaleValue != null) result += " numberScaleValue='" + _numberScaleValue.ToString() + "'";
            if (_numberPrefix != null) result += " numberPrefix='" + _numberPrefix.ToString() + "'";
            if (_numberSuffix != null) result += " numberSuffix='" + _numberSuffix.ToString() + "'";
            if (_decimalSeparator != null) result += " decimalSeparator='" + _decimalSeparator.ToString() + "'";
            if (_thousandSeparator != null) result += " thousandSeparator='" + _thousandSeparator.ToString() + "'";
            if (_inDecimalSeparator != null) result += " inDecimalSeparator='" + _inDecimalSeparator.ToString() + "'";
            if (_inThousandSeparator != null) result += " inThousandSeparator='" + _inThousandSeparator.ToString() + "'";
            if (_decimals != 0) result += " decimals='" + _decimals.ToString() + "'";
            result += " forceDecimals='" + Convert.ToInt32(_forceDecimals) + "'";
            if (_yAxisValueDecimals != 0) result += " yAxisValueDecimals='" + _yAxisValueDecimals.ToString() + "'";

            if (_baseFont != null) result += " baseFont='" + _baseFont.ToString() + "'";
            if (_baseFontSize != 0) result += " baseFontSize='" + _baseFontSize.ToString() + "'";
            if (_baseFontColor != Color.Empty) result += " baseFontColor='" + ColorToHexString(_baseFontColor) + "'";
            if (_outCnvBaseFont != null) result += " outCnvBaseFont='" + _outCnvBaseFont.ToString() + "'";
            if (_outCnvBaseFontSize != 0) result += " outCnvBaseFontSize='" + _outCnvBaseFontSize.ToString() + "'";
            if (_outCnvBaseFontColor != Color.Empty) result += " outCnvBaseFontColor='" + ColorToHexString(_outCnvBaseFontColor) + "'";
            return result;
        }
        static char[] hexDigits = {
         '0', '1', '2', '3', '4', '5', '6', '7',
         '8', '9', 'A', 'B', 'C', 'D', 'E', 'F'};
        public static string ColorToHexString(Color color)
        {
            byte[] bytes = new byte[3];
            bytes[0] = color.R;
            bytes[1] = color.G;
            bytes[2] = color.B;
            char[] chars = new char[bytes.Length * 2];
            for (int i = 0; i < bytes.Length; i++)
            {
                int b = bytes[i];
                chars[i * 2] = hexDigits[b >> 4];
                chars[i * 2 + 1] = hexDigits[b & 0xF];
            }
            return new string(chars);
        }
    }

    public class DisplayList : System.ComponentModel.StringConverter
    {
        string[] _display = new[] { "WRAP", "STAGGER", "ROTATE", "NONE" };
        public override bool GetStandardValuesSupported(
                           ITypeDescriptorContext context)
        {
            return true;
        }
        public override StandardValuesCollection
                     GetStandardValues(ITypeDescriptorContext context)
        {
            return new StandardValuesCollection(_display);
        }
        public override bool GetStandardValuesExclusive(
                           ITypeDescriptorContext context)
        {
            return false;
        }
    }
    public class ChartList : System.ComponentModel.StringConverter
    {
        string[] _display = new[] { "Area2D",
                                    "Bar2D",
                                    "Bubble",
                                    "Column2D",
                                    "Column3D",
                                    "Doughnut2D",
                                    "Doughnut3D",
                                    "Line2D",
                                    "MSArea2D",
                                    "MSBar2D",
                                    "MSBar3D",
                                    "MSColumn2D",
                                    "MSColumn3D",
                                    "MSColumn3DLineDY",
                                    "MSColumnLine3D",
                                    "MSCombi2D",
                                    "MSCombiDY2D",
                                    "MSLine",
                                    "MSStackedColumn2D",
                                    "MSStackedColumn2DLineDY",
                                    "Pie2D",
                                    "Pie3D",
                                    "Scatter",
                                    "ScrollArea2D",
                                    "ScrollColumn2D",
                                    "ScrollCombi2D",
                                    "ScrollCombiDY2D",
                                    "ScrollLine2D",
                                    "ScrollStackedColumn2D",
                                    "SSGrid",
                                    "StackedArea2D",
                                    "StackedBar2D",
                                    "StackedBar3D",
                                    "StackedColumn2D",
                                    "StackedColumn3D",
                                    "StackedColumn3DLineDY"
                                    };
        public override bool GetStandardValuesSupported(
                           ITypeDescriptorContext context)
        {
            return true;
        }
        public override StandardValuesCollection
                     GetStandardValues(ITypeDescriptorContext context)
        {
            return new StandardValuesCollection(_display);
        }
        public override bool GetStandardValuesExclusive(
                           ITypeDescriptorContext context)
        {
            return false;
        }
    }
}
