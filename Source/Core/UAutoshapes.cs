using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using System.Globalization;
using System.Drawing;
#if (MONOTOUCH)
  using Color = MonoTouch.UIKit.UIColor;
#endif

namespace FlexCel.Core
{
    #region ShapeProperties
    /// <summary>
    /// A class describing an Excel graphics object.
    /// </summary>
	public class TShapeProperties
	{
        private long FShapeId;
		private TRichString FText;
		private int FTextFlags;
		private int FTextRotation;
		private bool FFlipH;
		private bool FFlipV;
		private TClientAnchor FAnchor;
		private TShapePropertiesList FChildren;
		private TShapeType FShapeType;
		private TObjectType FObjectType;
		private string FShapeName;
		private string FObjectPathAbsolute;
		private TShapeOptionList FShapeOptions;
		private bool FPrint = true;
        private bool FVisible = true;
        internal bool FIsActiveX;
        internal bool FIsInternal;

        private TShapeGeom FShapeGeom;
        private TShapeFont FShapeThemeFont;

#if (!COMPACTFRAMEWORK)
		internal int zOrder; //internal use. Not valid if not explictly set.
#endif

		/// <summary>
		/// Creates a new instance.
		/// </summary>
		public TShapeProperties()
		{
			FChildren = new TShapePropertiesList();
			FObjectType = TObjectType.Undefined;
		}

		/// <summary>
		/// Type of shape. Note that this is not the same as <see cref="ObjectType"/>
		/// A comment might have a ShapeType=TShapeType.Rectangle, but its object type is TObjectType.Comment.
		/// A rectangle Autoshape will also have ShapeType=TShapeType.Rectangle, but its ObjectType will be
		/// TObjectType.Rectangle.
		/// </summary>
		public TShapeType ShapeType {get{return FShapeType;} set{FShapeType=value;}}

		/// <summary>
		/// Type of object. Note that this is not the same as <see cref="ShapeType"/>.
		/// A comment might have a ShapeType=TShapeType.Rectangle, but its object type is TObjectType.Comment.
		/// A rectangle Autoshape will also have ShapeType=TShapeType.Rectangle, but its ObjectType will be
		/// TObjectType.Rectangle.
		/// </summary>
		public TObjectType ObjectType {get{return FObjectType;} set{FObjectType=value;}}

		/// <summary>
		/// Name of the shape if it is named, null otherwise.
		/// </summary>
		public string ShapeName {get{return FShapeName;} set{FShapeName=value;}}

        /// <summary>
        /// This is an internal identified for the shape. It will remain the same once the file is loaded, but it might
        /// change when you load the same file at different times.
        /// </summary>
        public long ShapeId { get { return FShapeId; } set { FShapeId = value; } }

		/// <summary>
		/// Use this string to identify the shape when it is not the first on the hierarchy.
		/// For Example, imagine you have a Group Shape A with 2 children, B and C.
		/// If you want to change the text on shape C, you need to call SetObjectText(n,ObjectPath);
        /// <br></br>
        /// The object path can be of 2 types: Absolute or relative. Absolute object paths
        /// start with "\\" and include the parent object. Relative paths don't include the main group shape.
        /// <br></br>For example the absolute path "@1\\2\\3" is the same as accessing the object 1, with object path "2\\3"
        /// <b></b> This property returns the relative path, you can get the absolute path with <see cref="ObjectPathAbsolute"/>.
		/// </summary>
        public string ObjectPath
        {
            get
            { 
                if (FObjectPathAbsolute == null) return null;
                if (!FObjectPathAbsolute.StartsWith(FlxConsts.ObjectPathAbsolute)) return null;
                string s = FObjectPathAbsolute.Substring(1);
                int p = s.IndexOf(FlxConsts.ObjectPathSeparator);
                if (p <= 0) return null;
                return s.Substring(p + 1);
                
            }
        }

        /// <summary>
        /// Use this string to identify the shape when it is not the first on the hierarchy.
        /// For Example, imagine you have a Group Shape A with 2 children, B and C.
        /// If you want to change the text on shape C, you need to call SetObjectText(n,ObjectPath);
        /// <br></br>
        /// The object path can be of 2 types: Absolute or relative. Absolute object paths
        /// start with "\\" and include the parent object. Relative paths don't include the main group shape.
        /// <br></br>For example the absolute path "@1\\2\\3" is the same as accessing the object 1, with object path "2\\3"
        /// <b></b> This property returns the absolute path, you can get the relative path with <see cref="ObjectPath"/>.
        /// </summary>
        public string ObjectPathAbsolute { get { return FObjectPathAbsolute; } set { FObjectPathAbsolute = value; } }

        /// <summary>
        /// Returns the object path as a shape id. This is used mostly internally.
        /// </summary>
        public string ObjectPathShapeId { get { return FlxConsts.ObjectPathSpId + ShapeId.ToString(CultureInfo.InvariantCulture); } }
		
        /// <summary>
		/// Text of the shape if is has some, null otherwise.
		/// </summary>
		public TRichString Text {get {return FText;} set {FText = value;}}

		/// <summary>
		/// Option flags for the Text shape.
		/// Mask: 0x000E
		///        Horizontal text alignment:
		///             1 = left-aligned
		///             2 = centered
		///             3 = right-aligned
		///             4 = justified
		/// Mask: 0x0070
		///        Vertical text alignment:
		///             1 = top
		///             2 = center
		///             3 = bottom
		///             4 = justify
		/// Mask: 0x0200
		///         1 if the Lock Text option is on (Format Text Box dialog box, Protection tab)
		/// </summary>
		public int TextFlags {get {return FTextFlags;} set {FTextFlags = value;}}

		/// <summary>
		/// Text rotation, in degrees * 0xFFFF
		/// </summary>
		public int TextRotation {get {return FTextRotation;} set {FTextRotation = value;}}

		/// <summary>
		/// A lot of personalized settings, like shadow type fill color, line type, etc.
		/// </summary>
        public TShapeOptionList ShapeOptions { get { return FShapeOptions; } set { FShapeOptions = value; } }

		/// <summary>
		/// Coordinates of the shape. Note that when the shape is a group, this value is null and the real anchor is 
		/// retruned in the first child of the shape.
		/// To get the real Anchor of a first level object, use <see cref="NestedAnchor"/>
		/// </summary>
        public TClientAnchor Anchor { get { return FAnchor; } set { FAnchor = value; } }

        /// <summary>
        /// Geometry of the shape, if it is an xlsx shape. Shapes in xls will have this value null.
        /// </summary>
        internal TShapeGeom ShapeGeom { get { return FShapeGeom; } set { FShapeGeom = value; } }

        /// <summary>
        /// Theme Font used for the text in the shape. This property only affects xlsx files, in xls font is always arial black,
        /// and you can change the font ony inside the rich string, not in a global way. In xls files this will be null.
        /// </summary>
        public TShapeFont ShapeThemeFont { get { return FShapeThemeFont; } set { FShapeThemeFont = value; } }

		/// <summary>
		/// When the shape is a group, the real properties of the shape are in its first children. This method 
		/// returns the correct shape options.
		/// </summary>
		public TShapeOptionList NestedOptions {get {return GetNestedOptions(this);}}

		private static TShapeOptionList GetNestedOptions(TShapeProperties ShapeProps)
		{
			TClientAnchor Anchor = ShapeProps.Anchor;
			if (Anchor != null) return ShapeProps.ShapeOptions;

			if (ShapeProps.ChildrenCount > 0) return GetNestedOptions(ShapeProps.Children(1));  //Children(1) is the shape with info.
			return ShapeProps.ShapeOptions;
		}

		/// <summary>
		/// Real Coordinates of the shape. Note that when the shape is a group, the value in <see cref="Anchor"/> is null and the real anchor is 
		/// returned in the first child of the shape. 
		/// This method will get the real Anchor of a first level object.
		/// </summary>
        public TClientAnchor NestedAnchor { get { return GetNestedAnchor(this); } }

		private static TClientAnchor GetNestedAnchor(TShapeProperties ShapeProps)
		{
			TClientAnchor Result = ShapeProps.Anchor;
			if (Result == null)
			{
                if (ShapeProps.ChildrenCount > 1 && ShapeProps.Children(1).ShapeType == TShapeType.NotPrimitive)
                {
                    //This is shape that governs the others
                    Result = ShapeProps.Children(1).Anchor;
                }
			}
			return Result;
		}

        /// <summary>
        /// True if the shape is flipped horizontally.
        /// </summary>
        public bool FlipH {get{return FFlipH;} set {FFlipH = value;}}

        /// <summary>
        /// True if the shape is flipped vertically.
        /// </summary>
        public bool FlipV {get{return FFlipV;} set {FFlipV = value;}}

        /// <summary>
        /// True if the shape should be printed.
        /// </summary>
        public bool Print {get{return FPrint;} set {FPrint = value;}}

        /// <summary>
        /// True if the shape is visible.
        /// </summary>
        public bool Visible { get { return FVisible; } set { FVisible = value; } }

        /// <summary>
        /// Returns true if the object is an ActiveX object.
        /// </summary>
        public bool IsActiveX { get { return FIsActiveX; } }

        /// <summary>
        /// Returns true if the object is an internal object, like a comment or the arrow of an autofilter.
        /// Internal objects shouldn't be modified.
        /// </summary>
        public bool IsInternal { get { return FIsInternal; } }

        /// <summary>
        /// Number of shapes that are inside this shape.
        /// </summary>
        public int ChildrenCount{get {return FChildren.Count;}}

        /// <summary>
        /// Returns one of the shapes inside of this one.
        /// </summary>
        /// <param name="index">Index of the shape (1-based)</param>
        /// <returns>A child shape.</returns>
        public TShapeProperties Children(int index)
        {
            return FChildren[index-1];
        }

        /// <summary>
        /// Adds a new child for this autoshape.
        /// </summary>
        /// <param name="sp">Shape to be added.</param>
        public void AddChild(TShapeProperties sp)
        {
            FChildren.Add(sp);
        }
    }
    #endregion

    #region ShapePropertiesList
    /// <summary>
    /// A list of shapes.
    /// </summary>
    internal class TShapePropertiesList
    {
		private List<TShapeProperties> FList;

        public TShapePropertiesList()
        {
            FList = new List<TShapeProperties>();
        }

        public int Count
        {
            get
            {
                return FList.Count;
            }
        }

        public TShapeProperties this[int index]
        {
            get
            {
                return (TShapeProperties)FList[index];
            }
            set
            {
                FList[index] = value;
            }
        }

        public void Add (TShapeProperties sp)
        {
            FList.Add(sp);
        }

#if (FRAMEWORK20)
        public int BinarySearch(TShapeProperties Value, IComparer<TShapeProperties> aComparer)
#else
        public int BinarySearch(object Value, IComparer aComparer)
#endif
        {
            return FList.BinarySearch(0, FList.Count, Value, aComparer); //Only BinarySearch compatible with CF.
        }

        public void Insert(int index, TShapeProperties value)
        {
            FList.Insert(index, value);
        }
    }
    #endregion

    #region ShapeType
    /// <summary>
    /// Enumeration with the different shapes.
    /// </summary>
    public enum TShapeType
    {
        /// <summary></summary>
        NotPrimitive = 0,
        /// <summary></summary>
        Rectangle = 1,
        /// <summary></summary>
        RoundRectangle = 2,
        /// <summary></summary>
        Ellipse = 3,
        /// <summary></summary>
        Diamond = 4,
        /// <summary></summary>
        IsocelesTriangle = 5,
        /// <summary></summary>
        RightTriangle = 6,
        /// <summary></summary>
        Parallelogram = 7,
        /// <summary></summary>
        Trapezoid = 8,
        /// <summary></summary>
        Hexagon = 9,
        /// <summary></summary>
        Octagon = 10,
        /// <summary></summary>
        Plus = 11,
        /// <summary></summary>
        Star = 12,
        /// <summary></summary>
        Arrow = 13,
        /// <summary></summary>
        ThickArrow = 14,
        /// <summary></summary>
        HomePlate = 15,
        /// <summary></summary>
        Cube = 16,
        /// <summary></summary>
        Balloon = 17,
        /// <summary></summary>
        Seal = 18,
        /// <summary></summary>
        Arc = 19,
        /// <summary></summary>
        Line = 20,
        /// <summary></summary>
        Plaque = 21,
        /// <summary></summary>
        Can = 22,
        /// <summary></summary>
        Donut = 23,
        /// <summary></summary>
        TextSimple = 24,
        /// <summary></summary>
        TextOctagon = 25,
        /// <summary></summary>
        TextHexagon = 26,
        /// <summary></summary>
        TextCurve = 27,
        /// <summary></summary>
        TextWave = 28,
        /// <summary></summary>
        TextRing = 29,
        /// <summary></summary>
        TextOnCurve = 30,
        /// <summary></summary>
        TextOnRing = 31,
        /// <summary></summary>
        StraightConnector1 = 32,
        /// <summary></summary>
        BentConnector2 = 33,
        /// <summary></summary>
        BentConnector3 = 34,
        /// <summary></summary>
        BentConnector4 = 35,
        /// <summary></summary>
        BentConnector5 = 36,
        /// <summary></summary>
        CurvedConnector2 = 37,
        /// <summary></summary>
        CurvedConnector3 = 38,
        /// <summary></summary>
        CurvedConnector4 = 39,
        /// <summary></summary>
        CurvedConnector5 = 40,
        /// <summary></summary>
        Callout1 = 41,
        /// <summary></summary>
        Callout2 = 42,
        /// <summary></summary>
        Callout3 = 43,
        /// <summary></summary>
        AccentCallout1 = 44,
        /// <summary></summary>
        AccentCallout2 = 45,
        /// <summary></summary>
        AccentCallout3 = 46,
        /// <summary></summary>
        BorderCallout1 = 47,
        /// <summary></summary>
        BorderCallout2 = 48,
        /// <summary></summary>
        BorderCallout3 = 49,
        /// <summary></summary>
        AccentBorderCallout1 = 50,
        /// <summary></summary>
        AccentBorderCallout2 = 51,
        /// <summary></summary>
        AccentBorderCallout3 = 52,
        /// <summary></summary>
        Ribbon = 53,
        /// <summary></summary>
        Ribbon2 = 54,
        /// <summary></summary>
        Chevron = 55,
        /// <summary></summary>
        Pentagon = 56,
        /// <summary></summary>
        NoSmoking = 57,
        /// <summary></summary>
        Seal8 = 58,
        /// <summary></summary>
        Seal16 = 59,
        /// <summary></summary>
        Seal32 = 60,
        /// <summary></summary>
        WedgeRectCallout = 61,
        /// <summary></summary>
        WedgeRRectCallout = 62,
        /// <summary></summary>
        WedgeEllipseCallout = 63,
        /// <summary></summary>
        Wave = 64,
        /// <summary></summary>
        FoldedCorner = 65,
        /// <summary></summary>
        LeftArrow = 66,
        /// <summary></summary>
        DownArrow = 67,
        /// <summary></summary>
        UpArrow = 68,
        /// <summary></summary>
        LeftRightArrow = 69,
        /// <summary></summary>
        UpDownArrow = 70,
        /// <summary></summary>
        IrregularSeal1 = 71,
        /// <summary></summary>
        IrregularSeal2 = 72,
        /// <summary></summary>
        LightningBolt = 73,
        /// <summary></summary>
        Heart = 74,
        /// <summary></summary>
        PictureFrame = 75,
        /// <summary></summary>
        QuadArrow = 76,
        /// <summary></summary>
        LeftArrowCallout = 77,
        /// <summary></summary>
        RightArrowCallout = 78,
        /// <summary></summary>
        UpArrowCallout = 79,
        /// <summary></summary>
        DownArrowCallout = 80,
        /// <summary></summary>
        LeftRightArrowCallout = 81,
        /// <summary></summary>
        UpDownArrowCallout = 82,
        /// <summary></summary>
        QuadArrowCallout = 83,
        /// <summary></summary>
        Bevel = 84,
        /// <summary></summary>
        LeftBracket = 85,
        /// <summary></summary>
        RightBracket = 86,
        /// <summary></summary>
        LeftBrace = 87,
        /// <summary></summary>
        RightBrace = 88,
        /// <summary></summary>
        LeftUpArrow = 89,
        /// <summary></summary>
        BentUpArrow = 90,
        /// <summary></summary>
        BentArrow = 91,
        /// <summary></summary>
        Seal24 = 92,
        /// <summary></summary>
        StripedRightArrow = 93,
        /// <summary></summary>
        NotchedRightArrow = 94,
        /// <summary></summary>
        BlockArc = 95,
        /// <summary></summary>
        SmileyFace = 96,
        /// <summary></summary>
        VerticalScroll = 97,
        /// <summary></summary>
        HorizontalScroll = 98,
        /// <summary></summary>
        CircularArrow = 99,
        /// <summary></summary>
        NotchedCircularArrow = 100,
        /// <summary></summary>
        UturnArrow = 101,
        /// <summary></summary>
        CurvedRightArrow = 102,
        /// <summary></summary>
        CurvedLeftArrow = 103,
        /// <summary></summary>
        CurvedUpArrow = 104,
        /// <summary></summary>
        CurvedDownArrow = 105,
        /// <summary></summary>
        CloudCallout = 106,
        /// <summary></summary>
        EllipseRibbon = 107,
        /// <summary></summary>
        EllipseRibbon2 = 108,
        /// <summary></summary>
        FlowChartProcess = 109,
        /// <summary></summary>
        FlowChartDecision = 110,
        /// <summary></summary>
        FlowChartInputOutput = 111,
        /// <summary></summary>
        FlowChartPredefinedProcess = 112,
        /// <summary></summary>
        FlowChartInternalStorage = 113,
        /// <summary></summary>
        FlowChartDocument = 114,
        /// <summary></summary>
        FlowChartMultidocument = 115,
        /// <summary></summary>
        FlowChartTerminator = 116,
        /// <summary></summary>
        FlowChartPreparation = 117,
        /// <summary></summary>
        FlowChartManualInput = 118,
        /// <summary></summary>
        FlowChartManualOperation = 119,
        /// <summary></summary>
        FlowChartConnector = 120,
        /// <summary></summary>
        FlowChartPunchedCard = 121,
        /// <summary></summary>
        FlowChartPunchedTape = 122,
        /// <summary></summary>
        FlowChartSummingJunction = 123,
        /// <summary></summary>
        FlowChartOr = 124,
        /// <summary></summary>
        FlowChartCollate = 125,
        /// <summary></summary>
        FlowChartSort = 126,
        /// <summary></summary>
        FlowChartExtract = 127,
        /// <summary></summary>
        FlowChartMerge = 128,
        /// <summary></summary>
        FlowChartOfflineStorage = 129,
        /// <summary></summary>
        FlowChartOnlineStorage = 130,
        /// <summary></summary>
        FlowChartMagneticTape = 131,
        /// <summary></summary>
        FlowChartMagneticDisk = 132,
        /// <summary></summary>
        FlowChartMagneticDrum = 133,
        /// <summary></summary>
        FlowChartDisplay = 134,
        /// <summary></summary>
        FlowChartDelay = 135,
        /// <summary></summary>
        TextPlainText = 136,
        /// <summary></summary>
        TextStop = 137,
        /// <summary></summary>
        TextTriangle = 138,
        /// <summary></summary>
        TextTriangleInverted = 139,
        /// <summary></summary>
        TextChevron = 140,
        /// <summary></summary>
        TextChevronInverted = 141,
        /// <summary></summary>
        TextRingInside = 142,
        /// <summary></summary>
        TextRingOutside = 143,
        /// <summary></summary>
        TextArchUpCurve = 144,
        /// <summary></summary>
        TextArchDownCurve = 145,
        /// <summary></summary>
        TextCircleCurve = 146,
        /// <summary></summary>
        TextButtonCurve = 147,
        /// <summary></summary>
        TextArchUpPour = 148,
        /// <summary></summary>
        TextArchDownPour = 149,
        /// <summary></summary>
        TextCirclePour = 150,
        /// <summary></summary>
        TextButtonPour = 151,
        /// <summary></summary>
        TextCurveUp = 152,
        /// <summary></summary>
        TextCurveDown = 153,
        /// <summary></summary>
        TextCascadeUp = 154,
        /// <summary></summary>
        TextCascadeDown = 155,
        /// <summary></summary>
        TextWave1 = 156,
        /// <summary></summary>
        TextWave2 = 157,
        /// <summary></summary>
        TextWave3 = 158,
        /// <summary></summary>
        TextWave4 = 159,
        /// <summary></summary>
        TextInflate = 160,
        /// <summary></summary>
        TextDeflate = 161,
        /// <summary></summary>
        TextInflateBottom = 162,
        /// <summary></summary>
        TextDeflateBottom = 163,
        /// <summary></summary>
        TextInflateTop = 164,
        /// <summary></summary>
        TextDeflateTop = 165,
        /// <summary></summary>
        TextDeflateInflate = 166,
        /// <summary></summary>
        TextDeflateInflateDeflate = 167,
        /// <summary></summary>
        TextFadeRight = 168,
        /// <summary></summary>
        TextFadeLeft = 169,
        /// <summary></summary>
        TextFadeUp = 170,
        /// <summary></summary>
        TextFadeDown = 171,
        /// <summary></summary>
        TextSlantUp = 172,
        /// <summary></summary>
        TextSlantDown = 173,
        /// <summary></summary>
        TextCanUp = 174,
        /// <summary></summary>
        TextCanDown = 175,
        /// <summary></summary>
        FlowChartAlternateProcess = 176,
        /// <summary></summary>
        FlowChartOffpageConnector = 177,
        /// <summary></summary>
        Callout90 = 178,
        /// <summary></summary>
        AccentCallout90 = 179,
        /// <summary></summary>
        BorderCallout90 = 180,
        /// <summary></summary>
        AccentBorderCallout90 = 181,
        /// <summary></summary>
        LeftRightUpArrow = 182,
        /// <summary></summary>
        Sun = 183,
        /// <summary></summary>
        Moon = 184,
        /// <summary></summary>
        BracketPair = 185,
        /// <summary></summary>
        BracePair = 186,
        /// <summary></summary>
        Seal4 = 187,
        /// <summary></summary>
        DoubleWave = 188,
        /// <summary></summary>
        ActionButtonBlank = 189,
        /// <summary></summary>
        ActionButtonHome = 190,
        /// <summary></summary>
        ActionButtonHelp = 191,
        /// <summary></summary>
        ActionButtonInformation = 192,
        /// <summary></summary>
        ActionButtonForwardNext = 193,
        /// <summary></summary>
        ActionButtonBackPrevious = 194,
        /// <summary></summary>
        ActionButtonEnd = 195,
        /// <summary></summary>
        ActionButtonBeginning = 196,
        /// <summary></summary>
        ActionButtonReturn = 197,
        /// <summary></summary>
        ActionButtonDocument = 198,
        /// <summary></summary>
        ActionButtonSound = 199,
        /// <summary></summary>
        ActionButtonMovie = 200,
        /// <summary></summary>
        HostControl = 201,
        /// <summary></summary>
        TextBox = 202,

        //2007 shapes
        /// <summary></summary>
        LineInv = 1000,
        /// <summary></summary>
        NonIsoscelesTrapezoid = LineInv + 1,
        /// <summary></summary>
        Heptagon = LineInv + 2,
        /// <summary></summary>
        Decagon = LineInv + 3,
        /// <summary></summary>
        Dodecagon = LineInv + 4,


        /// <summary></summary>
        Star6 = LineInv + 6,
        /// <summary></summary>
        Star7 = LineInv + 7,
        /// <summary></summary>
        Star8 = LineInv + 8,
        /// <summary></summary>
        Star10 = LineInv + 9,
        /// <summary></summary>
        Star12 = LineInv + 10,

        /// <summary></summary>
        Round1Rect = LineInv + 11,
        /// <summary></summary>
        Round2SameRect = LineInv + 12,
        /// <summary></summary>
        Round2DiagRect = LineInv + 13,
        /// <summary></summary>
        SnipRoundRect = LineInv + 14,
        /// <summary></summary>
        Snip1Rect = LineInv + 15,
        /// <summary></summary>
        Snip2SameRect = LineInv + 16,
        /// <summary></summary>
        Snip2DiagRect = LineInv + 17,
        /// <summary></summary>
        Teardrop = LineInv + 18,
        /// <summary></summary>
        PieWedge = LineInv + 19,
        /// <summary></summary>
        Pie = LineInv + 20,

        /// <summary></summary>
        RightArrow = LineInv + 21,
        /// <summary></summary>
        LeftCircularArrow = LineInv + 22,
        /// <summary></summary>
        LeftRightCircularArrow = LineInv + 23,
        /// <summary></summary>
        Frame = LineInv + 24,
        /// <summary></summary>
        HalfFrame = LineInv + 25,
        /// <summary></summary>
        Corner = LineInv + 26,
        /// <summary></summary>
        DiagStripe = LineInv + 27,
        /// <summary></summary>
        Chord = LineInv + 28,
        /// <summary></summary>
        Cloud = LineInv + 29,

        /// <summary></summary>
        LeftRightRibbon = LineInv + 30,
        /// <summary></summary>
        Gear6 = LineInv + 31,
        /// <summary></summary>
        Gear9 = LineInv + 32,
        /// <summary></summary>
        Funnel = LineInv + 33,
        /// <summary></summary>
        MathPlus = LineInv + 34,
        /// <summary></summary>
        MathMinus = LineInv + 35,
        /// <summary></summary>
        MathMultiply = LineInv + 36,
        /// <summary></summary>
        MathDivide = LineInv + 37,
        /// <summary></summary>
        MathEqual = LineInv + 38,
        /// <summary></summary>
        MathNotEqual = LineInv + 39,

        /// <summary></summary>
        CornerTabs = LineInv + 40,
        /// <summary></summary>
        SquareTabs = LineInv + 41,
        /// <summary></summary>
        PlaqueTabs = LineInv + 42,
        /// <summary></summary>
        ChartX = LineInv + 43,
        /// <summary></summary>
        ChartStar = LineInv + 44,
        /// <summary></summary>
        ChartPlus = LineInv + 45,
        /// <summary></summary>
        SwooshArrow = LineInv + 46,
        /// <summary></summary>
        Trapezoid2007 = LineInv + 47,

        /// <summary></summary>
        Nil = 0x0FFF
    }

    #endregion

    #region ObjectType
    /// <summary>
    /// A type of object. Do not confuse this with a type of shape ( <see cref="TShapeType"/> ) This does not describe the shape
    /// form (like if it is a rectangle or a circle) but the shape kind (for example if it is a comment, an image or an autoshape)
    /// </summary>
    public enum TObjectType
    {
        /// <summary>Unknown object type.</summary>
        Undefined = -1,
        /// <summary></summary>
        Group = 0x00,
        /// <summary></summary>
        Line = 0x01,
        /// <summary></summary>
        Rectangle = 0x02,
        /// <summary></summary>
        Oval = 0x03,
        /// <summary></summary>
        Arc = 0x04,
        /// <summary></summary>
        Chart = 0x05,
        /// <summary></summary>
        Text = 0x06,
        /// <summary></summary>
        Button = 0x07,
        /// <summary>An image inserted on Excel</summary>
        Picture = 0x08,
        /// <summary></summary>
        Polygon = 0x09,
        /// <summary></summary>
        CheckBox = 0x0B,
        /// <summary></summary>
        OptionButton = 0x0C,
        /// <summary></summary>
        EditBox = 0x0D,
        /// <summary></summary>
        Label = 0x0E,
        /// <summary></summary>
        DialogBox = 0x0F,
        /// <summary></summary>
        Spinner = 0x10,
        /// <summary></summary>
        ScrollBar = 0x11,
        /// <summary></summary>
        ListBox = 0x12,
        /// <summary></summary>
        GroupBox = 0x13,
        /// <summary></summary>
        ComboBox = 0x14,
        /// <summary></summary>
        Comment = 0x19,
        /// <summary></summary>
        MicrosoftOfficeDrawing = 0x1E
    }
    #endregion

    #region ShapeOptions
    /// <summary>
    /// Many different configuration options for a shape.
    /// </summary>
    public enum TShapeOption
    {
        /// <summary>
        /// Not defined.
        /// </summary>
        None = 0,

        #region Transform
        /// <summary>
        /// Rotation in 1/65536 degrees.
        /// </summary>
        Rotation = 4,
        #endregion

        #region Protection
        /// <summary>
        /// No rotation
        /// </summary>
        LockRotation =	119,
 
        /// <summary>
        /// Don't allow changes in aspect ratio
        /// </summary>
        fLockAspectRatio = 120,
 
        /// <summary>
        /// Don't allow the shape to be moved
        /// </summary>
        fLockPosition = 121,
	
        /// <summary>
        /// Shape may not be selected
        /// </summary>
        fLockAgainstSelect = 122,
	
        /// <summary>
        /// No cropping this shape
        /// </summary>
        fLockCropping = 123,
	
        /// <summary>
        /// Edit Points not allowed
        /// </summary>
        fLockVertices = 124,
	
        /// <summary>
        /// Do not edit text
        /// </summary>
        fLockText = 125,
	
        /// <summary>
        /// Do not adjust
        /// </summary>
        fLockAdjustHandles = 126,
 
        /// <summary>
        /// Do not group this shape
        /// </summary>
        fLockAgainstGrouping = 127,
        #endregion

        #region Text
        /// <summary>
        /// id for the text, value determined by the host
        /// </summary>
        lTxid=128,
 
        /// <summary>
        /// margins relative to shape's inscribed text rectangle (in EMUs)
        /// 1/10 inch
        /// </summary>
        dxTextLeft = 129,
 
        /// <summary>
        /// margins relative to shape's inscribed text rectangle (in EMUs)
        /// 1/20 inch
        /// </summary>
        dyTextTop = 130,
 
        /// <summary>
        /// margins relative to shape's inscribed text rectangle (in EMUs)
        /// 1/10 inch
        /// </summary>
        dxTextRight = 131,
 
        /// <summary>
        /// margins relative to shape's inscribed text rectangle (in EMUs)
        /// 1/20 inch
        /// </summary>
        dyTextBottom = 132,

        /// <summary>
        /// 	Wrap text at shape margins
        /// </summary>
        WrapText = 133,
 
        /// <summary>
        /// Text zoom/scale (used if fFitTextToShape)
        /// </summary>
        scaleText = 134,
 
        /// <summary>
        /// How to anchor the text
        /// </summary>
        anchorText = 135,

        /// <summary>
        /// Text flow
        /// </summary>
        txflTextFlow = 136,
 
        /// <summary>
        /// Font rotation
        /// </summary>
        cdirFont = 137,
	
        /// <summary>
        /// ID of the next shape (used by Word for linked textboxes)
        /// </summary>
        hspNext = 138,
 
        /// <summary>
        /// Bi-Di Text direction
        /// </summary>
        txdir =	139,
	
        /// <summary>
        /// TRUE if single click selects text, FALSE if two clicks
        /// </summary>
        fSelectText = 187,

        /// <summary>
        /// use host's margin calculations
        /// </summary>
        fAutoTextMargin = 188,
 
        /// <summary>
        /// Rotate text with shape
        /// </summary>
        fRotateText = 189,
	
        /// <summary>
        /// Size shape to fit text size
        /// </summary>
        fFitShapeToText = 190,
 
        /// <summary>
        /// Size text to fit shape size
        /// </summary>
        fFitTextToShape = 191,

        #endregion

        #region GeoText

        /// <summary>
        /// UNICODE text string
        /// </summary>
        gtextUNICODE = 192,
 
        /// <summary>
        /// RTF text string
        /// </summary>
        gtextRTF = 193,
 
        /// <summary>
        /// alignment on curve
        /// </summary>
        gtextAlign = 194,
 
        /// <summary>
        /// default point size
        /// </summary>
        gtextSize = 195,
 
        /// <summary>
        /// fixed point 16.16
        /// </summary>
        gtextSpacing = 196,
 
        /// <summary>
        /// font family name
        /// </summary>
        gtextFont = 197,
 
        /// <summary>
        /// Reverse row order
        /// </summary>
        gtextFReverseRows = 240,
 
        /// <summary>
        /// Has text effect
        /// </summary>
        fGtext = 241,
 
        /// <summary>
        /// Rotate characters
        /// </summary>
        gtextFVertical = 242,
 
        /// <summary>
        /// Kern characters
        /// </summary>
        gtextFKern = 243,

        /// <summary>
        /// Tightening or tracking
        /// </summary>
        gtextFTight = 244,
 
        /// <summary>
        /// Stretch to fit shape
        /// </summary>
        gtextFStretch = 245,
 
        /// <summary>
        /// Char bounding box
        /// </summary>
        gtextFShrinkFit = 246,
 
        /// <summary>
        /// Scale text-on-path
        /// </summary>
        gtextFBestFit = 247,
 
        /// <summary>
        /// Stretch char height
        /// </summary>
        gtextFNormalize = 248,
 
        /// <summary>
        /// Do not measure along path
        /// </summary>
        gtextFDxMeasure = 249,
 
        /// <summary>
        /// Bold font
        /// </summary>
        gtextFBold = 250,
 
        /// <summary>
        /// Italic font
        /// </summary>
        gtextFItalic = 251,
 
        /// <summary>
        /// Underline font
        /// </summary>
        gtextFUnderline = 252,
 
        /// <summary>
        /// Shadow font
        /// </summary>
        gtextFShadow = 253,
 
        /// <summary>
        /// Small caps font
        /// </summary>
        gtextFSmallcaps = 254,
 
        /// <summary>
        /// Strike through font
        /// </summary>
        gtextFStrikethrough = 255,
        #endregion

        #region Blip
        /// <summary>
        /// 16.16 fraction times total image width or height, as appropriate.	
        /// </summary>
        cropFromTop	=	256,
	
        /// <summary>
        /// 16.16 fraction times total image width or height, as appropriate.	
        /// </summary>
		
        cropFromBottom	=	257,			
        /// <summary>
        /// 16.16 fraction times total image width or height, as appropriate.	
        /// </summary>
		
        cropFromLeft	=	258,			
        /// <summary>
        /// 16.16 fraction times total image width or height, as appropriate.	
        /// </summary>
		
        cropFromRight	=	259,	
		
        /// <summary>
        /// Blip to display	
        /// </summary>
        pib	=	260,

        /// <summary>
        /// Blip file name	
        /// </summary>
        pibName	=	261,

        /// <summary>
        /// Blip flags	
        /// </summary>
        pibFlags	=	262,

        /// <summary>
        /// transparent color (none if ~0UL)	
        /// </summary>
        pictureTransparent	=	263,

        /// <summary>
        /// contrast setting	
        /// </summary>
        pictureContrast	=	264,

        /// <summary>
        /// brightness setting	
        /// </summary>
        pictureBrightness	=	265,

        /// <summary>
        /// 16.16 gamma	
        /// </summary>
        pictureGamma	=	266,

        /// <summary>
        /// Host-defined ID for OLE objects (usually a pointer)	
        /// </summary>
        pictureId	=	267,

        /// <summary>
        /// Modification used if shape has double shadow	
        /// </summary>
        pictureDblCrMod	=	268,

        /// <summary>
        /// Blip file name	
        /// </summary>
        pibPrintName	=	272,

        /// <summary>
        /// Blip flags	
        /// </summary>
        pibPrintFlags	=	273,

        /// <summary>
        /// Do not hit test the picture	
        /// </summary>
        fNoHitTestPicture	=	316,

        /// <summary>
        /// grayscale display	
        /// </summary>
        pictureGray	=	317,

        /// <summary>
        /// bi-level display	
        /// </summary>
        pictureBiLevel	=	318,

        /// <summary>
        /// Server is active (OLE objects only)	
        /// </summary>
        pictureActive	=	319,
        #endregion

        #region Geometry
        /// <summary>
        /// Defines the G (geometry) coordinate space. 
        /// </summary>
        geoLeft = 320,

        /// <summary>
        /// 
        /// </summary>
        geoTop = 321,   

        /// <summary>
        /// 
        /// </summary>
        geoRight = 322,

        /// <summary>
        /// 
        /// </summary>
        geoBottom = 323,

        /// <summary>
        /// 
        /// </summary>
        shapePath = 324,

        /// <summary>
        /// 
        /// </summary>
        pVertices = 325,

        /// <summary>
        /// 
        /// </summary>
        pSegmentInfo = 326,

        /// <summary>
        /// Adjustment values corresponding to the positions of the adjust handles of the shape. The number of values used and their allowable ranges vary from shape type to shape type.
        /// </summary>
        adjustValue = 327,

        /// <summary>
        /// Adjustment values corresponding to the positions of the adjust handles of the shape. The number of values used and their allowable ranges vary from shape type to shape type.
        /// </summary>
        adjust2Value = 328,

        /// <summary>
        /// Adjustment values corresponding to the positions of the adjust handles of the shape. The number of values used and their allowable ranges vary from shape type to shape type.
        /// </summary>
        adjust3Value = 329,

        /// <summary>
        /// Adjustment values corresponding to the positions of the adjust handles of the shape. The number of values used and their allowable ranges vary from shape type to shape type.
        /// </summary>
        adjust4Value = 330,

        /// <summary>
        /// Adjustment values corresponding to the positions of the adjust handles of the shape. The number of values used and their allowable ranges vary from shape type to shape type.
        /// </summary>
        adjust5Value = 331,

        /// <summary>
        /// Adjustment values corresponding to the positions of the adjust handles of the shape. The number of values used and their allowable ranges vary from shape type to shape type.
        /// </summary>
        adjust6Value = 332,

        /// <summary>
        /// Adjustment values corresponding to the positions of the adjust handles of the shape. The number of values used and their allowable ranges vary from shape type to shape type.
        /// </summary>
        adjust7Value = 333,

        /// <summary>
        /// Adjustment values corresponding to the positions of the adjust handles of the shape. The number of values used and their allowable ranges vary from shape type to shape type.
        /// </summary>
        adjust8Value = 334,

        /// <summary>
        /// Adjustment values corresponding to the positions of the adjust handles of the shape. The number of values used and their allowable ranges vary from shape type to shape type.
        /// </summary>
        adjust9Value = 335,

        /// <summary>
        /// Adjustment values corresponding to the positions of the adjust handles of the shape. The number of values used and their allowable ranges vary from shape type to shape type.
        /// </summary>
        adjust10Value = 336,

        /// <summary>
        /// This property specifies an array of connection sites that a user can use to make a link between shapes. 
        /// </summary>
        pConnectionSites = 337,

        /// <summary>
        /// This property specifies an array of angles corresponding to the connection sites in the 
        /// pConnectionSites_complex property that are used to determine the direction that a connector links 
        /// to the corresponding connection site. 
        /// </summary>
        pConnectionSitesDir = 338,

        /// <summary>
        /// This property specifies an x coordinate above which limousine scaling is used in the horizontal 
        /// direction. This means that points whose x coordinate is greater than xLimo have their x coordinates 
        /// incremented rather than linearly scaled.
        /// </summary>
        xLimo = 339,

        /// <summary>
        /// This property specifies an y coordinate above which limousine scaling is used in the vertical 
        /// direction. This means that points whose y coordinate is greater than yLimo have their y coordinates
        /// incremented rather than linearly scaled.
        /// </summary>
        yLimo = 340,

        /// <summary>
        /// This property specifies an array of adjust handles which allow a user to manipulate the geometry of this shape. 
        /// </summary>
        pAdjustHandles = 341,

        /// <summary>
        ///  Array of guide formula for the shape which specify how the geometry of the shape changes as the adjust handles are dragged.
        /// </summary>
        pGuides = 342,

        /// <summary>
        /// This property specifies an array of rectangles specifying how text should be inscribed within this shape. 
        /// </summary>
        pInscribe = 343,

        /// <summary>
        /// Shadow may be set
        /// </summary>
        fShadowOK = 378,

        /// <summary>
        /// 3D may be set
        /// </summary>
        f3DOK = 379,

        /// <summary>
        /// Line style may be set
        /// </summary>
        fLineOK = 380,

        /// <summary>
        /// Text effect (WordArt) supported
        /// </summary>
        fGtextOK = 381,

        /// <summary>
        /// 
        /// </summary>
        fFillShadeShapeOK = 382,

        /// <summary>
        /// OK to fill the shape through the UI or VBAs
        /// </summary>
        fFillOK = 383,
        #endregion

        #region Fill Style
	
        /// <summary>
        /// Type of fill
        /// </summary>
        fillType = 384,
	
        /// <summary>
        /// Foreground color
        /// </summary>
        fillColor = 385,
	
        /// <summary>
        /// Fixed 16.16
        /// </summary>
        fillOpacity = 386,
	
        /// <summary>
        /// Background color
        /// </summary>
        fillBackColor = 387,
	
        /// <summary>
        /// Shades only
        /// </summary>
        fillBackOpacity = 388,
	
        /// <summary>
        /// Modification for BW views
        /// </summary>
        fillCrMod = 389,
	
        /// <summary>
        /// Pattern/texture
        /// </summary>
        fillBlip = 390,
	
        /// <summary>
        /// Blip file name
        /// </summary>
        fillBlipName = 391,
	
        /// <summary>
        /// Blip flags
        /// </summary>
        fillBlipFlags = 392,
	
        /// <summary>
        /// How big (A units) to make a metafile texture.
        /// </summary>
        fillWidth = 393,
	 
        /// <summary>
        /// 
        /// </summary>
        fillHeight = 394,

        /// <summary>
        /// Fade angle - degrees in 16.16
        /// </summary>
        fillAngle = 395,
	
        /// <summary>
        /// Linear shaded fill focus percent
        /// </summary>
        fillFocus = 396,
	
        /// <summary>
        /// Fraction 16.16
        /// </summary>
        fillToLeft = 397,
	
        /// <summary>
        /// Fraction 16.16
        /// </summary>
        fillToTop = 398,
	
        /// <summary>
        /// Fraction 16.16
        /// </summary>
        fillToRight = 399,
	
        /// <summary>
        /// Fraction 16.16
        /// </summary>
        fillToBottom = 400,
	
        /// <summary>
        /// For shaded fills, use the specified rectangle instead of the shape's bounding rect to define how large the fade is going to be.
        /// </summary>
        fillRectLeft = 401,

        /// <summary>
        /// </summary>
        fillRectTop = 402,

        /// <summary>
        /// </summary>
        fillRectRight = 403,
	
        /// <summary>
        /// </summary>
        fillRectBottom = 404,
		
        /// <summary>
        /// </summary>
        fillDztype = 405,

        /// <summary>
        /// Special shades
        /// </summary>
        fillShadePreset = 406,
	
        /// <summary>
        /// a preset array of colors
        /// </summary>
        fillShadeColors = 407,

        /// <summary>
        /// </summary>
        fillOriginX = 408,
 
        /// <summary>
        /// </summary>
        fillOriginY = 409,

        /// <summary>
        /// </summary>
        fillShapeOriginX = 410,
 
        /// <summary>
        /// </summary>
        fillShapeOriginY = 411,

        /// <summary>
        /// Type of shading, if a shaded (gradient) fill.
        /// </summary>
        fillShadeType = 412,
	
        /// <summary>
        /// Is shape filled?
        /// </summary>
        fFilled = 443,
	
        /// <summary>
        /// Should we hit test fill? 
        /// </summary>
        fHitTestFill = 444,
	
        /// <summary>
        /// Register pattern on shape
        /// </summary>
        fillShape = 445,
	
        /// <summary>
        /// Use the large rect?
        /// </summary>
        fillUseRect = 446,
	
        /// <summary>
        /// Hit test a shape as though filled
        /// </summary>
        fNoFillHitTest = 447,
	 
        #endregion

        #region Line Style
	
        /// <summary>
        /// Color of line
        /// </summary>
        lineColor = 448,
	 
        /// <summary>
        /// Not implemented
        /// </summary>
        lineOpacity = 449,
	
        /// <summary>
        /// Background color
        /// </summary>
        lineBackColor = 450,
	 
        /// <summary>
        /// Modification for BW views
        /// </summary>
        lineCrMod = 451,
	
        /// <summary>
        /// Type of line
        /// </summary>
        lineType = 452,
	
        /// <summary>
        /// Pattern/texture
        /// </summary>
        lineFillBlip = 453,
	
        /// <summary>
        /// Blip file name
        /// </summary>
        lineFillBlipName = 454,
	 
        /// <summary>
        /// Blip flags
        /// </summary>
        lineFillBlipFlags = 455,
	
        /// <summary>
        /// How big (A units) to make a metafile texture.
        /// </summary>
        lineFillWidth = 456,
	 
        /// <summary>
        /// </summary>
        lineFillHeight = 457,
	
        /// <summary>
        /// How to interpret fillWidth/Height numbers.
        /// </summary>
        lineFillDztype = 458,
	
        /// <summary>
        /// A units; 1pt == 12700 EMUs
        /// </summary>
        lineWidth = 459,
	
        /// <summary>
        /// ratio (16.16) of width
        /// </summary>
        lineMiterLimit = 460,
	
        /// <summary>
        /// Draw parallel lines?
        /// </summary>
        lineStyle = 461,
	
        /// <summary>
        /// Can be overridden by:
        /// </summary>
        lineDashing = 462,
	 
        /// <summary>
        /// As Win32 ExtCreatePen
        /// </summary>
        lineDashStyle = 463,
	
        /// <summary>
        /// Arrow at start
        /// </summary>
        lineStartArrowhead = 464,
	 
        /// <summary>
        /// Arrow at end
        /// </summary>
        lineEndArrowhead = 465,
	
        /// <summary>
        /// Arrow at start
        /// </summary>
        lineStartArrowWidth = 466,
	
        /// <summary>
        /// Arrow at end
        /// </summary>
        lineStartArrowLength = 467,
	
        /// <summary>
        /// Arrow at start
        /// </summary>
        lineEndArrowWidth = 468,
	
        /// <summary>
        /// Arrow at end
        /// </summary>
        lineEndArrowLength = 469,
	
        /// <summary>
        /// How to join lines
        /// </summary>
        lineJoinStyle = 470,
	
        /// <summary>
        /// How to end lines
        /// </summary>
        lineEndCapStyle = 471,
	 
        /// <summary>
        /// Allow arrowheads if prop. is set
        /// </summary>
        fArrowheadsOK = 507,
	
        /// <summary>
        /// Any line?
        /// </summary>
        fLine = 508,
	
        /// <summary>
        /// Should we hit test lines? 
        /// </summary>
        fHitTestLine = 509,
	 
        /// <summary>
        /// Register pattern on shape
        /// </summary>
        lineFillShape = 510,
	 
        /// <summary>
        /// Draw a dashed line if no line
        /// </summary>
        fNoLineDrawDash = 511,
	 

        #endregion

        #region Shadow Style
	
        /// <summary>
        /// Type of effect
        /// </summary>
        shadowType = 512,
	
        /// <summary>
        /// Foreground color
        /// </summary>
        shadowColor = 513,
	
        /// <summary>
        /// Embossed color
        /// </summary>
        shadowHighlight = 514,
	
        /// <summary>
        /// Modification for BW views
        /// </summary>
        shadowCrMod = 515,
	 
        /// <summary>
        /// Fixed 16.16
        /// </summary>
        shadowOpacity = 516,
	 
        /// <summary>
        /// Offset shadow
        /// </summary>
        shadowOffsetX = 517,	 
	
        /// <summary>
        /// Offset shadow
        /// </summary>
        shadowOffsetY = 518,
	
        /// <summary>
        /// Double offset shadow
        /// </summary>
        shadowSecondOffsetX = 519,
	 
        /// <summary>
        /// Double offset shadow
        /// </summary>
        shadowSecondOffsetY = 520,
	 
        /// <summary>
        /// 16.16
        /// </summary>
        shadowScaleXToX = 521,
	 
        /// <summary>
        /// 16.16
        /// </summary>
        shadowScaleYToX = 522,
	 
        /// <summary>
        /// 16.16
        /// </summary>
        shadowScaleXToY = 523,
	
        /// <summary>
        /// 16.16
        /// </summary>
        shadowScaleYToY = 524,
	
        /// <summary>
        /// 16.16 / weight
        /// </summary>
        shadowPerspectiveX = 525,
	 
        /// <summary>
        /// 16.16 / weight
        /// </summary>
        shadowPerspectiveY = 526,
	 
        /// <summary>
        /// scaling factor
        /// </summary>
        shadowWeight = 527,
	 
        /// <summary>
        /// </summary>
        shadowOriginX = 528,
 
        /// <summary>
        /// </summary>
        shadowOriginY = 529,
	
        /// <summary>
        /// Any shadow?
        /// </summary>
        fShadow = 574,
	 
        /// <summary>
        /// Excel5-style shadow
        /// </summary>
        fshadowObscured = 575,

        #endregion

        #region Perspective Style
 
        /// <summary>
        /// Where transform applies
        /// </summary>
        perspectiveType = 576,
 
        /// <summary>
        /// The LONG values define a transformation matrix, effectively, each value is scaled by the perspectiveWeight parameter.
        /// </summary>
        perspectiveOffsetX = 577,
 
        /// <summary>
        /// </summary>
        perspectiveOffsetY = 578,
 
        /// <summary>
        /// </summary>
        perspectiveScaleXToX = 579,
 
        /// <summary>
        /// </summary>
        perspectiveScaleYToX = 580,
 
        /// <summary>
        /// </summary>
        perspectiveScaleXToY = 581,
 
        /// <summary>
        /// </summary>
        perspectiveScaleYToY = 582,
 
        /// <summary>
        /// </summary>
        perspectivePerspectiveX = 583,
 
        /// <summary>
        /// </summary>
        perspectivePerspectiveY = 584,
 
        /// <summary>
        /// Scaling factor
        /// </summary>
        perspectiveWeight = 585,
 
        /// <summary>
        /// </summary>
        perspectiveOriginX = 586,
 
        /// <summary>
        /// </summary>
        perspectiveOriginY = 587,
 
        /// <summary>
        /// On/off
        /// </summary>
        fPerspective = 639,
 
        #endregion

        #region 3D Object
        /// <summary>
        /// Fixed-point 16.16
        /// </summary>
        c3DSpecularAmt = 640,
 
        /// <summary>
        /// Fixed-point 16.16
        /// </summary>
        c3DDiffuseAmt = 641,
 
        /// <summary>
        /// Default gives OK results
        /// </summary>
        c3DShininess = 642,
 
        /// <summary>
        /// Specular edge thickness
        /// </summary>
        c3DEdgeThickness = 643,
 
        /// <summary>
        /// Distance of extrusion in EMUs
        /// </summary>
        c3DExtrudeForward = 644,
 
        /// <summary>
        /// </summary>
        c3DExtrudeBackward = 645,
 
        /// <summary>
        /// Extrusion direction
        /// </summary>
        c3DExtrudePlane = 646,
 
        /// <summary>
        /// Basic color of extruded part of shape; the lighting model used will determine the exact shades used when rendering. 
        /// </summary>
        c3DExtrusionColor = 647,
 
        /// <summary>
        /// Modification for BW views
        /// </summary>
        c3DCrMod = 648,
 
        /// <summary>
        /// Does this shape have a 3D effect?
        /// </summary>
        f3D = 700,
 
        /// <summary>
        /// Use metallic specularity?
        /// </summary>
        fc3DMetallic = 701,
 
        /// <summary>
        /// </summary>
        fc3DUseExtrusionColor = 702,
 
        /// <summary>
        /// </summary>
        fc3DLightFace = 703,
        #endregion

        #region 3D Style
        /// <summary>
        /// degrees (16.16) about y axis
        /// </summary>
        c3DYRotationAngle = 704,
 
        /// <summary>
        /// degrees (16.16) about x axis
        /// </summary>
        c3DXRotationAngle = 705,
 
        /// <summary>
        /// These specify the rotation axis; only their relative magnitudes matter.
        /// </summary>
        c3DRotationAxisX = 706,
 
        /// <summary>
        /// </summary>
        c3DRotationAxisY = 707,
 
        /// <summary>
        /// </summary>
        c3DRotationAxisZ = 708,
 
        /// <summary>
        /// degrees (16.16) about axis
        /// </summary>
        c3DRotationAngle = 709,
 
        /// <summary>
        /// rotation center x (16.16 or g-units)
        /// </summary>
        c3DRotationCenterX = 710,
 
        /// <summary>
        /// rotation center y (16.16 or g-units)
        /// </summary>
        c3DRotationCenterY = 711,
 
        /// <summary>
        /// rotation center z (absolute (emus))
        /// </summary>
        c3DRotationCenterZ = 712,
 
        /// <summary>
        /// Full,wireframe, or bcube
        /// </summary>
        c3DRenderMode = 713,
 
        /// <summary>
        /// pixels (16.16)
        /// </summary>
        c3DTolerance = 714,
 
        /// <summary>
        /// X view point (emus)
        /// </summary>
        c3DXViewpoint = 715,
 
        /// <summary>
        /// Y view point (emus)
        /// </summary>
        c3DYViewpoint = 716,
 
        /// <summary>
        /// Z view distance (emus)
        /// </summary>
        c3DZViewpoint = 717,
 
        /// <summary>
        /// </summary>
        c3DOriginX = 718,
 
        /// <summary>
        /// </summary>
        c3DOriginY = 719,
 
        /// <summary>
        /// degree (16.16) skew angle
        /// </summary>
        c3DSkewAngle = 720,
 
        /// <summary>
        /// Percentage skew amount
        /// </summary>
        c3DSkewAmount = 721,
 
        /// <summary>
        /// Fixed point intensity
        /// </summary>
        c3DAmbientIntensity = 722,
 
        /// <summary>
        /// Key light source direction; only their relative
        /// </summary>
        c3DKeyX = 723,
 
        /// <summary>
        /// Key light source direction; only their relative
        /// </summary>
        c3DKeyY = 724,
 
        /// <summary>
        /// magnitudes matter
        /// </summary>
        c3DKeyZ = 725,
 
        /// <summary>
        /// Fixed point intensity
        /// </summary>
        c3DKeyIntensity = 726,
 
        /// <summary>
        /// Fill light source direction; only their relative
        /// </summary>
        c3DFillX = 727,
 
        /// <summary>
        /// Fill light source direction; only their relative
        /// </summary>
        c3DFillY = 728,
 
        /// <summary>
        /// magnitudes matter
        /// </summary>
        c3DFillZ = 729,
 
        /// <summary>
        /// Fixed point intensity
        /// </summary>
        c3DFillIntensity = 730,
 
        /// <summary>
        /// </summary>
        fc3DConstrainRotation = 763,
 
        /// <summary>
        /// </summary>
        fc3DRotationCenterAuto = 764,
 
        /// <summary>
        /// Parallel projection?
        /// </summary>
        fc3DParallel = 765,
 
        /// <summary>
        /// Is key lighting harsh?
        /// </summary>
        fc3DKeyHarsh = 766,

        /// <summary>
        /// </summary>
        fc3DFillHarsh = 767,

        /// <summary>
        ///  This property is present if the shape represents an equation generated by Office 2007 or later.  
        ///  The property is a string of XML representing a Word 2003 XML document. 
        ///  The original equation is stored within the oMathPara tag within the document. 
        /// </summary>
        wzEquationXML = 780,
        #endregion

        #region Shape
        /// <summary>
        /// master shape
        /// </summary>
        hspMaster = 769,
 
        /// <summary>
        /// Type of connector
        /// </summary>
        cxstyle = 771,
 
        /// <summary>
        /// Settings for modifications to be made when in different forms of black-and-white mode.
        /// </summary>
        bWMode = 772,
 
        /// <summary>
        /// </summary>
        bWModePureBW = 773,
 
        /// <summary>
        /// </summary>
        bWModeBW = 774,
 
        /// <summary>
        /// For OLE objects, whether the object is in icon form
        /// </summary>
        fOleIcon = 826,
 
        /// <summary>
        /// For UI only. Prefer relative resizing. 
        /// </summary>
        fPreferRelativeResize = 827,
 
        /// <summary>
        /// Lock the shape type (don't allow Change Shape)
        /// </summary>
        fLockShapeType = 828,
 
        /// <summary>
        /// </summary>
        fDeleteAttachedObject = 830,
 
        /// <summary>
        /// If TRUE, this is the background shape.
        /// </summary>
        fBackground = 831,
        #endregion

        #region CallOut
        /// <summary>
        /// Callout type  (TwoSegment)
        /// </summary>
        spcot = 832,
 
        /// <summary>
        /// Distance from box to first point.(EMUs) (1/12 inch)
        /// </summary>
        dxyCalloutGap = 833,
 
        /// <summary>
        /// Callout angle (Any)
        /// </summary>
        spcoa = 834,
    
        /// <summary>
        /// Callout drop type (Specified)
        /// </summary>
        spcod = 835,
    
        /// <summary>
        /// if msospcodSpecified, the actual drop distance (9 points)
        /// </summary>
        dxyCalloutDropSpecified = 836,
    
        /// <summary>
        /// if fCalloutLengthSpecified, the actual distance (0)
        /// </summary>
        dxyCalloutLengthSpecified = 837,
    
        /// <summary>
        /// Is the shape a callout? (FALSE)
        /// </summary>
        fCallout = 889,
    
        /// <summary>
        /// does callout have accent bar? (FALSE)
        /// </summary>
        fCalloutAccentBar = 890,
    
        /// <summary>
        /// does callout have a text border? (TRUE)
        /// </summary>
        fCalloutTextBorder = 891,
    
        /// <summary>
        /// (FALSE)
        /// </summary>
        fCalloutMinusX = 892,
 
        /// <summary>
        /// FALSE
        /// </summary>
        fCalloutMinusY = 893,
 
        /// <summary>
        /// If true, then we occasionally invert the drop distance (FALSE)
        /// </summary>
        fCalloutDropAuto = 894,
    
        /// <summary>
        /// if true, we look at dxyCalloutLengthSpecified (FALSE)
        /// </summary>
        fCalloutLengthSpecified = 895,
    
        #endregion

        #region Group Shape
 
        /// Shape Name (present only if explicitly set)
        wzName = 896,
 
        /// alternate text
        wzDescription = 897,
 
        /// The hyperlink in the shape.
        pihlShape = 898,
 
        /// The polygon that text will be wrapped around (Word)
        pWrapPolygonVertices = 899,
 
        /// Left wrapping distance from text (Word)
        dxWrapDistLeft = 900,
 
        /// Top wrapping distance from text (Word)
        dyWrapDistTop = 901,
 
        /// Right wrapping distance from text (Word)
        dxWrapDistRight = 902,
 
        /// Bottom wrapping distance from text (Word)
        dyWrapDistBottom = 903,
 
        /// Regroup ID 
        lidRegroup = 904,

        /// <summary>
        /// This property specifies the minimum sizes of the rows in a table.
        /// </summary>
        tableRowProperties = 928,
 
        /// Has the wrap polygon been edited?
        fEditedWrap = 953,
 
        /// Word-only (shape is behind text)
        fBehindDocument = 954,
  
        /// Notify client on a double click
        fOnDblClickNotify = 955,
  
        /// A button shape (i.e., clicking performs an action). Set for shapes with attached hyperlinks or macros.
        fIsButton = 956,
  
        /// 1D adjustment
        fOneD = 957,
  
        /// Do not display
        fHidden = 958,
  
        /// Print this shape
        fPrint = 959,
 
        #endregion

        #region Others
        /// <summary>
        /// This property specifies relationships in a diagram.
        /// </summary>
         pRelationTbl = 1284,

        /// <summary>
        /// Diagram constrain bounds
        /// </summary>
        dgmConstrainBounds = 1288,

        /// <summary>
        /// Custom dash style of the line.
        /// </summary>
        lineLeftDashStyle = 1359,
        
        /// <summary>
        /// Custom dash style of the line.
        /// </summary>
        lineRightDashStyle = 1487,
        
        /// <summary>
        /// Custom dash style of the line.
        /// </summary>
        lineTopDashStyle = 1423,

        /// <summary>
        /// Custom dash style of the line.
        /// </summary>
        lineBottomDashStyle = 1551
        
        #endregion
    }

	#endregion

	#region TShapeOptionList
	/// <summary>
	/// This class holds a list of key/values pairs specifying the options for a shape.
	/// To Get a value from it, use: ShapeOptionList[TShapeOption.xxx];
	/// </summary>
	public class TShapeOptionList: IEnumerable, ICloneable, IEnumerable<TShapeOption>
	{
        private Dictionary<TShapeOption, object> FList = new Dictionary<TShapeOption, object>();

		/// <summary>
		/// Gets the value for a key. Value can be a long or a string, depending on the type of property.
		/// </summary>
		public object this[TShapeOption key]
		{
			get
			{
				object Result = null;
                if (FList.TryGetValue(key, out Result))
                    return Result;
                return null;
			}
		}

		internal void Add(TShapeOption key, object value)
		{
			FList[key] = value;
		}

        internal static bool IsASCII(TShapeOption Id)
        {
            switch (Id)
            {
                case TShapeOption.gtextRTF: return true;
                default:
                    return false;
            }
        }

        internal static bool IsUTF8(TShapeOption Id)
        {
            switch (Id)
            {
                case TShapeOption.wzEquationXML: return true;
                default:
                    return false;
            }
        }

        /// <summary>
        /// Returns a long property if it exists, otherwise the default value. Note: This method will always assume a positive number.
        /// To get a signed int, use <see cref="AsSignedLong"/>
        /// </summary>
        /// <param name="key">Property Name.</param>
        /// <param name="Default">What to return if the property does not exist.</param>
        /// <returns></returns>
        public long AsLong(TShapeOption key, long Default)
        {
            object Result = this[key];
            if (Result == null || !(Result is long)) return Default;
            unchecked
            {
                return (long)Result;
            }
        }

        internal void SetLong(TShapeOption key, long value, long Default)
        {
            if (value == Default)
            {
                if (FList.ContainsKey(key)) FList.Remove(key);
                return;
            }

            FList[key] = value;
        }

        /// <summary>
        /// Returns a long property if it exists, otherwise the default value. Note: This method will return negative numbers if the number is bigger than 65536.
        /// To get an unsigned int, use <see cref="AsLong"/>
        /// </summary>
        /// <param name="key">Property Name.</param>
        /// <param name="Default">What to return if the property does not exist.</param>
        /// <returns></returns>
        public long AsSignedLong(TShapeOption key, long Default)
        {
            unchecked
            {
                return (int)AsLong(key, Default);
            }
        }

        /// <summary>
        /// Returns a float (Expressed as 16.16) property if it exists, otherwise the default value.
        /// </summary>
        /// <param name="key">Property Name.</param>
        /// <param name="Default">What to return if the property does not exist.</param>
        /// <returns></returns>
        public float As1616(TShapeOption key, float Default)
        {
            object Result = this[key];
            if (Result == null || !(Result is long)) return Default;
            return Get1616((long)Result);
        }

        internal static float Get1616(long value)
        {
            unchecked
            {
                return (short)((value >> 16)) + (value & 0xFFFF) / 65536f;
            }
        }

        internal void Set1616(TShapeOption key, double value, float Default)
        {
            if (value == Default)
            {
                if (FList.ContainsKey(key)) FList.Remove(key);
                return;
            }

#if (COMPACTFRAMEWORK)
            value = value > 0 ? Math.Floor(value) : Math.Ceiling(value);
#else
            value = Math.Truncate(value);
#endif
        
            long lg = ((long)value) << 16;
            lg += 0xFFFF & (long)((value - lg) * 65536f);
            FList[key] = lg;
        }

        /// <summary>
        /// Returns a bool property if it exists, otherwise the default value.
        /// </summary>
        /// <param name="key">Property Name.</param>
        /// <param name="Default">What to return if the property does not exist.</param>
        /// <param name="PositionInSet">Boolean properties are grouped so all properties on one set are in only
        /// one value. So, the last bool property on the set is the first bit, and so on. ONLY THE LAST PROPERTY
        /// ON THE SET IS PRESENT.</param>
        /// <returns></returns>
        public bool AsBool(TShapeOption key, bool Default, int PositionInSet)
        {
            object Result = this[key];
            if (Result == null || !(Result is long)) return Default;
            long r = (long)Result;
            if ((r & (1<<(16 + PositionInSet))) == 0) return Default; //property is not set.
            return (r & (1<<PositionInSet)) != 0;
        }

        /// <summary>
        /// Returns an unicode property if it exists, otherwise the default value.
        /// </summary>
        /// <param name="key">Property Name.</param>
        /// <param name="Default">What to return if the property does not exist.</param>
        /// <returns></returns>
        public string AsUnicodeString(TShapeOption key, string Default)
        {
            byte[] bResult = this[key] as byte[];
            if (bResult == null) return Default;
            string Result;
            if (IsASCII(key)) Result = Encoding.ASCII.GetString(bResult, 0, bResult.Length);
            else if (IsUTF8(key)) Result = Encoding.UTF8.GetString(bResult, 0, bResult.Length);
            else Result = Encoding.Unicode.GetString(bResult, 0, bResult.Length);
            int k = Result.Length - 1;
            while (k >= 0 && Result[k] == (char)0) k--;
            return Result.Substring(0, k + 1);

        }

		/// <summary>
		/// Returns an hyperlink property if it exists, otherwise the default value.
		/// You will normally want to use this property with <see cref="TShapeOption.pihlShape"/>
		/// since that is the property that holds the link for the objects.
		/// </summary>
		/// <param name="key">Property Name.</param>
		/// <param name="Default">What to return if the property does not exist.</param>
		/// <returns></returns>
		public THyperLink AsHyperLink(TShapeOption key, THyperLink Default)
		{
			byte[] bResult = this[key] as byte[];
			if (bResult == null) return Default;

			byte[] dResult = new byte[bResult.Length + 8];
			Array.Copy(bResult, 0, dResult, 8, bResult.Length);
			FlexCel.XlsAdapter.THLinkRecord lr = FlexCel.XlsAdapter.THLinkRecord.CreateFromBiff8(0, dResult);
			return lr.GetProperties();

		}

        #region Equals
        /// <summary>
        /// Returns true if 2 instances of this class have the same values.
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public override bool Equals(object obj)
        {
            TShapeOptionList o2 = obj as TShapeOptionList;
            if (o2 == null) return false;

            if (o2.FList.Count != o2.FList.Count) return false;
            foreach (TShapeOption shOpt in FList.Keys)
            {
                object v1 = FList[shOpt];
                object v2;
                if (!o2.FList.TryGetValue(shOpt, out v2)) return false;
                if (!object.Equals(v1, v2)) return false;
            }

            return true;
        }

        /// <summary>
        /// Hashcode for the object.
        /// </summary>
        /// <returns></returns>
        public override int GetHashCode()
        {
            return base.GetHashCode();
        }
        #endregion

        internal IEnumerable<TShapeOption> Keys { get { return FList.Keys; } }

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

		#region ICloneable Members

		/// <summary>
		/// Creates a deep copy of the object.
		/// </summary>
		/// <returns></returns>
		public object Clone()
		{
			TShapeOptionList Result = new TShapeOptionList();
			foreach (TShapeOption So in FList.Keys)
			{
				ICloneable Ic = FList[So] as ICloneable;
				if (Ic != null)
				{
					Result.Add(So, Ic.Clone());
				}
				else
				{
					Result.Add(So, FList[So]);
				}
			}

			return Result;
		}

		#endregion

        #region IEnumerable<TShapeOption> Members
#if (FRAMEWORK20 && !DELPHIWIN32)
        IEnumerator<TShapeOption> IEnumerable<TShapeOption>.GetEnumerator()
        {
            return FList.Keys.GetEnumerator();
        }
#endif
        #endregion
    }
	#endregion

    #region Fill Types
    /// <summary>
    /// Type of fill for an autoshape. (In xls files)
    /// </summary>
    public enum TFillType
    {
        /// <summary>
        /// Fill with a solid color
        /// </summary>
        Solid = 0,

        /// <summary>
        /// Fill with a pattern (bitmap)
        /// </summary>
        Pattern = 1,

        /// <summary>
        /// A texture (pattern with its own color map)
        /// </summary>
        Texture = 2,

        /// <summary>
        /// Center a picture in the shape
        /// </summary>
        Picture = 3,

        /// <summary>
        /// Shade from start to end points
        /// </summary>
        Shade =4,

        /// <summary>
        /// Shade from bounding rectangle to end point
        /// </summary>
        ShadeCenter = 5,

        /// <summary>
        /// Shade from shape outline to end point
        /// </summary>
        ShadeShape = 6,

        /// <summary>
        /// Similar to msofillShade, but the fillAngle
        /// is additionally scaled by the aspect ratio of
        /// the shape. If shape is square, it is the
        /// same as msofillShade.
        /// </summary>
        ShadeScale = 7,

        /// <summary>
        /// special type - shade to title ---  for PP 
        /// </summary>
        ShadeTitle = 8,

        /// <summary>
        /// the background fill color/pattern
        /// </summary>
        Background = 9
    }

#endregion

    #region Coordinate Type

    /// <summary>
    ///  specifies a formula used to calculate a value for use in a shape definition.
    /// </summary>
    internal enum TSgFormula
    {
        /// <summary>
        /// Addition and subtraction. param1 + param2 - param3 
        /// </summary>
        sgfSum = 0x0000,

        /// <summary>
        /// Multiplication and division. (param1*param2)/param3 
        /// </summary>
        sgfProduct = 0x0001,

        /// <summary>
        /// Simple average. (param1+param2)/2 
        /// </summary>
        sgfMid = 0x0002,

        /// <summary>
        /// Absolute value. abs(param1) 
        /// </summary>
        sgfAbsolute = 0x0003,

        /// <summary>
        /// The lesser of two values. min(param1, param2) 
        /// </summary>
        sgfMin = 0x0004,

        /// <summary>
        /// The greater of two values. max(param1, param2) 
        /// </summary>
        sgfMax = 0x0005,

        /// <summary>
        /// Conditional selection. param1 > 0 ? param2 : param3 
        /// </summary>
        sgfIf = 0x0006,

        /// <summary>
        /// Modulus. sqrt(param1^2 + param2^2 + param3^2) 
        /// </summary>
        sgfMod = 0x0007,

        /// <summary>
        /// Trigonometric arc tangent of a quotient. Angles in degrees, as 1616. atan2(param2,param1) 
        /// </summary>
        sgfATan2 = 0x0008,

        /// <summary>
        /// Sine. Angles in degrees, as 1616. param1*sin(param2) 
        /// </summary>
        sgfSin = 0x0009,

        /// <summary>
        /// Cosine. Angles in degrees, as 1616. param1*cos(param2) 
        /// </summary>
        sgfCos = 0x000A,

        /// <summary>
        /// Cosine and atan2 in one formula. param1*cos(atan2(param3,param2)) 
        /// </summary>
        sgfCosATan2 = 0x000B,

        /// <summary>
        /// Sine and atan2 in one formula. param1*sin(atan2(param3,param2)) 
        /// </summary>
        sgfSinATan2 = 0x000C,

        /// <summary>
        /// Square root. sqrt(param1) 
        /// </summary>
        sgfSqrt = 0x000D,

        /// <summary>
        /// Angles in degrees as 1616. param1 + param2*2^16 + param3*2^16 
        /// </summary>
        sgfSumAngle = 0x000E,

        /// <summary>
        /// The eccentricity formula for an ellipse, where param1 is the length of the semiminor axis and param2 is the length of the semimajor axis. param3*sqrt(1-(param1/param2)^2) 
        /// </summary>
        sgfEllipse = 0x000F,

        /// <summary>
        ///    Angles in degrees, as 1616. param1*tan(param2) 
        /// </summary>
        sgfTan = 0x0010
    }
    #endregion

    #region Checkbox state
    /// <summary>
    /// Possible values of a checkbox in a sheet.
    /// </summary>
    public enum TCheckboxState
    {
        /// <summary>
        /// Checkbox is not checked.
        /// </summary>
        Unchecked,

        /// <summary>
        /// Checkbox is checked.
        /// </summary>
        Checked,

        /// <summary>
        /// Checkbox is not set. An indeterminate control generally has a shaded appearance.
        /// </summary>
        Indeterminate
    }

    #endregion

    #region B&W Mode
    internal enum TBwMode
    {
        /// <summary>
        /// Object rendered in color.
        /// </summary>
        Color = 0x000000,

        /// <summary>
        /// Object rendered with automatic coloring. 
        /// </summary>
        Automatic = 0x00000001,

        /// <summary>
        /// Object rendered with gray coloring. 
        /// </summary>
        GrayScale = 0x00000002,

        /// <summary>
        /// Object rendered with light gray coloring. 
        /// </summary>
        LightGrayScale = 0x00000003,

        /// <summary>
        /// Object rendered with inverse gray coloring. 
        /// </summary>
        InverseGray = 0x00000004,

        /// <summary>
        /// Object rendered with gray and white coloring. 
        /// </summary>
        GrayOutline = 0x00000005,

        /// <summary>
        /// Object rendered with black and gray coloring. 
        /// </summary>
        BlackTextLine = 0x00000006,

        /// <summary>
        /// Object rendered with black and white coloring. 
        /// </summary>
        HighContrast = 0x00000007,

        /// <summary>
        /// Object rendered with black-only coloring. 
        /// </summary>
        Black = 0x00000008,

        /// <summary>
        /// Object rendered with white coloring. 
        /// </summary>
        White = 0x00000009,

        /// <summary>
        /// Object not rendered. 
        /// </summary>
        DontShow = 0x0000000A

    }

    #endregion

    #region Line Dash
    /// <summary>
    /// Line style (dashes, solid, etc).
    /// </summary>
    public enum TLineDashing
    {
        /// <summary>
        /// Solid (continuous) pen.
        /// </summary>
        Solid,              

        /// <summary>
        /// PS_DASH system   dash style.
        /// </summary>
        DashSys,            

        /// <summary>
        /// PS_DOT system   dash style.
        /// </summary>
        DotSys,             

        /// <summary>
        /// PS_DASHDOT system dash style.
        /// </summary>
        DashDotSys,         

        /// <summary>
        /// PS_DASHDOTDOT system dash style.
        /// </summary>
        DashDotDotSys,      

        /// <summary>
        /// square dot style.
        /// </summary>
        DotGEL,             

        /// <summary>
        /// dash style.
        /// </summary>
        DashGEL,       
     
        /// <summary>
        /// long dash style.
        /// </summary>
        LongDashGEL,        

        /// <summary>
        /// dash short dash.
        /// </summary>
        DashDotGEL,         

        /// <summary>
        /// long dash short dash.
        /// </summary>
        LongDashDotGEL,     

        /// <summary>
        /// longg dash short dash short dash.
        /// </summary>
        LongDashDotDotGEL
    };
    #endregion

    #region Fill & Line

    /// <summary>
    /// Contains the information for the fill of an autoshape.
    /// </summary>
    public class TShapeFill
    {
        bool FHasFill;
        TFillStyle FFillStyle;
        TDrawingColor? FThemeColor;
        TFormattingType FThemeStyle;
        bool FUseThemeBk;

        /// <summary>
        /// Creates a simple shape fill.
        /// </summary>
        /// <param name="aHasFill"></param>
        /// <param name="aFillStyle"></param>
        public TShapeFill(bool aHasFill, TFillStyle aFillStyle)
        {
            FFillStyle = aFillStyle;
            FHasFill = aHasFill;
        }

        internal TShapeFill(TFillStyle aFillStyle, bool aHasFill, TFormattingType aThemeStyle, TDrawingColor aThemeColor, bool aUseThemeBk):
            this(aHasFill, aFillStyle)
        {
            FThemeStyle = aThemeStyle;
            FThemeColor = aThemeColor;
            FUseThemeBk = aUseThemeBk;
        }

        /// <summary>
        /// Returns a deep copy of this object.
        /// </summary>
        /// <returns></returns>
        public TShapeFill Clone()
        {
            return (TShapeFill)MemberwiseClone();
        }

        /// <summary>
        /// True if the object has fill, false if it is transparent.
        /// </summary>
        public bool HasFill { get { return FHasFill; } set { FHasFill = value; } }

        /// <summary>
        /// Fill style for this object. This method will return <see cref="FillStyle"/> if it isn't null,
        /// or the default theme fill if it is.
        /// </summary>
        public TFillStyle GetFill(IFlexCelPalette aPalette)
        {
            if (FFillStyle != null) return FFillStyle;
#if (FRAMEWORK30)
            TThemeFormatScheme fs = aPalette.GetTheme().Elements.FormatScheme;
            TFillStyleList fsl = fs.FillStyleList;

            if (FUseThemeBk)
            {
                fsl = fs.BkFillStyleList;
            }

            return GetThemeFill(aPalette, fsl);
#else
            return null;
#endif
        }

        internal TFillStyle GetThemeFill(IFlexCelPalette aPalette, TFillStyleList fsl)
        {
#if (FRAMEWORK30 && !COMPACTFRAMEWORK)
            if (!FThemeColor.HasValue) return new TNoFill();
            if ((int)FThemeStyle >= fsl.Count) return TFillStyleList.GetDefaultFillStyle((int)FThemeStyle).ReplacePhClr(FThemeColor.Value);
            return fsl.GetRealFillStyle(FThemeStyle, FThemeColor.Value);
#else
            return null;
#endif
        }


        /// <summary>
        /// Fill for the shape. If this value is null, the fill specified in the theme will be used instead.
        /// To know the real fill style used even if this value is null, use <see cref="GetFill"/>
        /// </summary>
        public TFillStyle FillStyle { get { return FFillStyle; } set { FFillStyle = value; } }

        /// <summary>
        /// Fill taken from a theme. If <see cref="FillStyle"/> is null, this color here will be used along with the current theme.
        /// </summary>
        public TDrawingColor? ThemeColor { get { return FThemeColor; } set { FThemeColor = value; } }

        /// <summary>
        /// Style (subtle, normal, intense) from the theme used, when a theme is used.
        /// </summary>
        public TFormattingType ThemeStyle { get { return FThemeStyle; } set { FThemeStyle = value; } }

        /// <summary>
        /// If true and using a theme, the backgroubd fill from the theme will be used, if not, the normal fill from the theme will be used.
        /// </summary>
        public bool UseThemeBk { get { return FUseThemeBk; } set { FUseThemeBk = value; } }

        /// <summary>
        /// Returns true if both objects have the same data.
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public override bool Equals(object obj)
        {
            TShapeFill o2 = obj as TShapeFill;
            if (o2 == null) return false;
            return
                  FHasFill == o2.FHasFill
               && Object.Equals(FFillStyle, o2.FFillStyle)
               && Object.Equals(FThemeColor, o2.FThemeColor)
               && Object.Equals(FThemeStyle, o2.FThemeStyle)
               && FUseThemeBk == o2.FUseThemeBk;
        }

        /// <summary>
        /// Hash code for the object.
        /// </summary>
        /// <returns></returns>
        public override int GetHashCode()
        {
            return HashCoder.GetHashObj(HasFill, FFillStyle, FThemeColor, FThemeStyle, FUseThemeBk);
        }


        internal int GetIdx()
        {
            int Result = 0;
            if (HasFill) Result = (int)FThemeStyle + 1;
            if (FUseThemeBk) Result += 1000;
            return Result;
        }
    }

    /// <summary>
    /// Contains the information for the line style for an autoshape.
    /// </summary>
    public class TShapeLine
    {
        bool FHasLine;
        TLineStyle FLineStyle;
        TDrawingColor? FThemeColor;
        TFormattingType FThemeStyle;
    
        /// <summary>
        /// Creates a black simple line.
        /// </summary>
        public TShapeLine()
        {
            FLineStyle = new TLineStyle(new TSolidFill(Color.Black));
            FHasLine = true;
        }

        /// <summary>
        /// Creates a line with a line style and no theme.
        /// </summary>
        /// <param name="aHasLine">True if the shape has a line.</param>
        /// <param name="aLineStyle">Custom line style. If null, the theme line style will be used.</param>
        public TShapeLine(bool aHasLine, TLineStyle aLineStyle): this(aHasLine, aLineStyle, null, TFormattingType.Subtle)
        {
        }

        /// <summary>
        /// Creates a line with all the options.
        /// </summary>
        /// <param name="aHasLine">True if the shape has a line.</param>
        /// <param name="aLineStyle">Custom line style. If null, the theme line style will be used.</param>
        /// <param name="aThemeColor">Color to use in a theme line style. Will be used only if aLineStyle is null. 
        /// It has the color that will be used instead of the one in the theme. All other line properties come from the theme.</param>
        /// <param name="aThemeStyle">Theme that will be used for the line. This only affects the line if aLineStyle is null.</param>
        public TShapeLine(bool aHasLine, TLineStyle aLineStyle, TDrawingColor? aThemeColor, TFormattingType aThemeStyle)
        {
            FHasLine = aHasLine;
            FLineStyle = aLineStyle;
            FThemeColor = aThemeColor;
            FThemeStyle = aThemeStyle;
        }

        /// <summary>
        /// Returns a deep copy of this object.
        /// </summary>
        /// <returns></returns>
        public TShapeLine Clone()
        {
            TLineStyle NewLineStyle = FLineStyle == null ? null : FLineStyle.Clone();
            return new TShapeLine(FHasLine, NewLineStyle, FThemeColor, FThemeStyle);
        }

        /// <summary>
        /// True if the object has a border, false otherwise.
        /// </summary>
        public bool HasLine { get { return FHasLine; } set { FHasLine = value; } }

        /// <summary>
        /// Line style used to draw the line. Note that this can be null, and in this case, 
        /// a line style from the current theme is used. If this isn't null, the theme properties are ignored.
        /// </summary>
        public TLineStyle LineStyle { get { return FLineStyle; } set { FLineStyle = value; } }

        /// <summary>
        /// Theme used to draw the line. This property has effect only if <see cref="LineStyle"/> is null.
        /// </summary>
        public TFormattingType ThemeStyle { get { return FThemeStyle; } set { FThemeStyle = value; } }

        /// <summary>
        /// Color that will be used instead of the default in the theme, when using a theme to draw the line.
        /// This property has effect only if <see cref="LineStyle"/> is null.
        /// </summary>
        public TDrawingColor? ThemeColor { get { return FThemeColor; } set { FThemeColor = value; } }

        /// <summary>
        /// This is the color used to draw the line, even if <see cref="LineStyle"/> is null
        /// </summary>
        /// <param name="aPalette"></param>
        /// <returns></returns>
        public TFillStyle GetLineFill(IFlexCelPalette aPalette)
        {
            if (LineStyle != null && LineStyle.Fill != null) return LineStyle.Fill;
#if (FRAMEWORK30 && !COMPACTFRAMEWORK)
            TThemeFormatScheme fs = aPalette.GetTheme().Elements.FormatScheme;
            TLineStyleList lsl = fs.LineStyleList;


            return GetFill(aPalette, lsl);
#else
            return null;
#endif
        }

        internal TFillStyle GetFill(IFlexCelPalette aPalette, TLineStyleList lsl)
        {
#if (FRAMEWORK30 && !COMPACTFRAMEWORK)
            if (!FThemeColor.HasValue) return null;
            if ((int)FThemeStyle >= lsl.Count) return TLineStyleList.GetDefaultLineStyle((int)FThemeStyle).Fill.ReplacePhClr(FThemeColor.Value);
            return lsl.GetRealFillStyle(FThemeStyle, FThemeColor.Value);
#else
            return null;
#endif
        }

        private TLineStyle GetLineStyle(IFlexCelPalette aPalette, Func<TLineStyle, bool> TestNotNull)
        {
            if (LineStyle != null && TestNotNull(LineStyle)) return LineStyle;
#if (FRAMEWORK30 && !COMPACTFRAMEWORK)
            TThemeFormatScheme fs = aPalette.GetTheme().Elements.FormatScheme;
            TLineStyleList lsl = fs.LineStyleList;

            TLineStyle Result;
            if ((int)FThemeStyle >= lsl.Count)
            {
                Result = TLineStyleList.GetDefaultLineStyle((int)FThemeStyle);
            }
            else
            {
                Result = lsl[FThemeStyle];
            }
            if (TestNotNull(Result)) return Result; else return null;

#else
            return null;
#endif
        }

        /// <summary>
        /// Returns the line width, if there is a LineStyle this is the Line width, else if it is null it is the Theme line width.
        /// </summary>
        /// <returns></returns>
        public int GetWidth(IFlexCelPalette aPalette)
        {
#if (FRAMEWORK30)
            TLineStyle RealLineStyle = GetLineStyle(aPalette, (x)=> x.Width.HasValue);
#else
            TLineStyle RealLineStyle = GetLineStyle(aPalette, delegate(TLineStyle x) { return x.Width.HasValue; });
#endif
            if (RealLineStyle != null) return RealLineStyle.Width.Value;
            return 0;
        }


        /// <summary>
        /// Returns the line dashing, if there is a LineStyle this is the Line dashing, else if it is null it is the Theme line dashing.
        /// </summary>
        /// <returns></returns>
        public TLineDashing GetDashing(IFlexCelPalette aPalette)
        {
#if(FRAMEWORK30)
            TLineStyle RealLineStyle = GetLineStyle(aPalette, (x) => x.Dashing.HasValue);
#else
            TLineStyle RealLineStyle = GetLineStyle(aPalette, delegate(TLineStyle x) { return x.Dashing.HasValue; });
#endif
            if (RealLineStyle != null) return RealLineStyle.Dashing.Value;
            return TLineDashing.Solid;
        }

        /// <summary>
        /// Returns the line joining, if there is a LineStyle this is the Line join, else if it is null it is the Theme line join.
        /// </summary>
        /// <returns></returns>
        public TLineJoin GetJoin(IFlexCelPalette aPalette)
        {
#if(FRAMEWORK30)
            TLineStyle RealLineStyle = GetLineStyle(aPalette, (x)=> x.Join.HasValue);
#else
            TLineStyle RealLineStyle = GetLineStyle(aPalette, delegate(TLineStyle x) { return x.Join.HasValue; });
#endif
            if (RealLineStyle != null) return RealLineStyle.Join.Value;
            return TLineJoin.Miter;
        }

        /// <summary>
        /// Returns the line arrow for the head, if there is a LineStyle this is the Line arrow, else if it is null it is the Theme line arrow.
        /// </summary>
        /// <returns></returns>
        public TLineArrow GetHeadArrow(IFlexCelPalette aPalette)
        {
#if(FRAMEWORK30)
            TLineStyle RealLineStyle = GetLineStyle(aPalette, (x)=> x.HeadArrow.HasValue);
#else
            TLineStyle RealLineStyle = GetLineStyle(aPalette, delegate(TLineStyle x) { return x.HeadArrow.HasValue; });
#endif
            if (RealLineStyle != null) return RealLineStyle.HeadArrow.Value;
            return TLineArrow.None;
        }


        /// <summary>
        /// Returns the line arrow for the tail, if there is a LineStyle this is the Line arrow, else if it is null it is the Theme line arrow.
        /// </summary>
        /// <returns></returns>
        public TLineArrow GetTailArrow(IFlexCelPalette aPalette)
        {
#if(FRAMEWORK30)
            TLineStyle RealLineStyle = GetLineStyle(aPalette, (x)=> x.TailArrow.HasValue);
#else
            TLineStyle RealLineStyle = GetLineStyle(aPalette, delegate(TLineStyle x) { return x.TailArrow.HasValue; });
#endif
            if (RealLineStyle != null) return RealLineStyle.TailArrow.Value;
            return TLineArrow.None;
        }

        /// <summary>
        /// Returns true if both objects have the same data.
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public override bool Equals(object obj)
        {
            TShapeLine o2 = obj as TShapeLine;
            if (o2 == null) return false;
            return
                  FHasLine == o2.FHasLine
               && Object.Equals(FLineStyle, o2.FLineStyle)
               && FThemeColor == o2.FThemeColor
               && FThemeStyle == o2.FThemeStyle;
        }

        /// <summary>
        /// Hash code for the object.
        /// </summary>
        /// <returns></returns>
        public override int GetHashCode()
        {
            return HashCoder.GetHashObj(HasLine, FLineStyle, FThemeColor, FThemeStyle);
        }

        internal int GetIdx()
        {
            int Result = 0;
            if (HasLine) Result = (int)ThemeStyle + 1;
            return Result;
        }
    }

    /// <summary>
    /// Contains information for the font of an autoshape.
    /// </summary>
    public class TShapeFont
    {
        TFontScheme FThemeScheme;
        TDrawingColor FThemeColor;

        /// <summary>
        /// Creates a new instance.
        /// </summary>
        /// <param name="aThemeScheme"></param>
        /// <param name="aThemeColor"></param>
        public TShapeFont(TFontScheme aThemeScheme, TDrawingColor aThemeColor)
        {
            ThemeScheme = aThemeScheme;
            ThemeColor = aThemeColor;
        }

        /// <summary>
        /// Scheme used in the theme.
        /// </summary>
        public TFontScheme ThemeScheme { get { return FThemeScheme; } set { FThemeScheme = value; } }

        /// <summary>
        /// Color used for the font.
        /// </summary>
        public TDrawingColor ThemeColor { get { return FThemeColor; } set { FThemeColor = value; } }

        #region Equal
        /// <summary>
        /// Returns true if bth object are the same.
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public override bool Equals(object obj)
        {
            TShapeFont o2 = obj as TShapeFont;
            if (o2 == null) return false;
            if (ThemeScheme != o2.ThemeScheme) return false;
            if (ThemeColor != o2.ThemeColor) return false;
            return true;
        }

        /// <summary>
        /// Returns the hascode for the object.
        /// </summary>
        /// <returns></returns>
        public override int GetHashCode()
        {
            return HashCoder.GetHashObj(ThemeScheme, ThemeColor);
        }

        /// <summary>
        /// Returns a deep copy of this class.
        /// </summary>
        /// <returns></returns>
        public TShapeFont Clone()
        {
            return new TShapeFont(ThemeScheme, ThemeColor);
        }
        #endregion

    }

    /// <summary>
    /// Contains information for the effects of an autoshape.
    /// </summary>
    public class TShapeEffects
    {
        TFormattingType FThemeStyle;
        TDrawingColor FThemeColor;

        /// <summary>
        /// Creates a new instance.
        /// </summary>
        public TShapeEffects(TFormattingType aThemeStyle, TDrawingColor aThemeColor)
        {
            ThemeStyle = aThemeStyle;
            ThemeColor = aThemeColor;
        }

        /// <summary>
        /// Scheme used in the theme.
        /// </summary>
        public TFormattingType ThemeStyle { get { return FThemeStyle; } set { FThemeStyle = value; } }

        /// <summary>
        /// Color used for the effects.
        /// </summary>
        public TDrawingColor ThemeColor { get { return FThemeColor; } set { FThemeColor = value; } }

        #region Equal
        /// <summary>
        /// Returns true if bth object are the same.
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public override bool Equals(object obj)
        {
            TShapeEffects o2 = obj as TShapeEffects;
            if (o2 == null) return false;
            if (ThemeStyle != o2.ThemeStyle) return false;
            if (ThemeColor != o2.ThemeColor) return false;
            return true;
        }

        /// <summary>
        /// Returns the hascode for the object.
        /// </summary>
        /// <returns></returns>
        public override int GetHashCode()
        {
            return HashCoder.GetHashObj(ThemeStyle, ThemeColor);
        }

        /// <summary>
        /// Returns a deep copy of this class.
        /// </summary>
        /// <returns></returns>
        public TShapeEffects Clone()
        {
            return new TShapeEffects(ThemeStyle, ThemeColor);
        }
        #endregion

    }
    #endregion

    #region Line Join
    /// <summary>
    /// How a line joins with the next
    /// </summary>
    public enum TLineJoin
    {
        /// <summary>
        /// Bevel.
        /// </summary>
        Bevel,

        /// <summary>
        /// Miter.
        /// </summary>
        Miter,

        /// <summary>
        /// Round.
        /// </summary>
        Round,
    }
    #endregion
    #region Arrows
    /// <summary>
    /// Style of an arrow.
    /// </summary>
    public enum TArrowStyle
    {
        /// <summary>
        /// No arrow.
        /// </summary>
        None,

        /// <summary>
        /// Normal arrow.
        /// </summary>
        Normal,

        /// <summary>
        /// Stealth arrow.
        /// </summary>
        Stealth,

        /// <summary>
        /// Diamond-shaped arrow.
        /// </summary>
        Diamond,

        /// <summary>
        /// Oval shaped arrow.
        /// </summary>
        Oval,

        /// <summary>
        /// Line arrow. (no fill)
        /// </summary>
        Open,
    }

    /// <summary>
    /// Preset width for an arrow.
    /// </summary>
    public enum TArrowWidth
    {
        /// <summary>
        /// Small.
        /// </summary>
        Small,

        /// <summary>
        /// Medium.
        /// </summary>
        Medium,

        /// <summary>
        /// Large.
        /// </summary>
        Large
}

    /// <summary>
    /// The length of an arrow head.
    /// </summary>
    public enum TArrowLen
    {
        /// <summary>
        /// Small.
        /// </summary>
        Small,

        /// <summary>
        /// Medium.
        /// </summary>
        Medium,

        /// <summary>
        /// Large.
        /// </summary>
        Large,
    }

    /// <summary>
    /// Describes an arrow at the end of a line. This struct is immutable.
    /// </summary>
    public struct TLineArrow: IComparable, IComparable<TLineArrow>
    {
        readonly TArrowStyle FStyle;
        readonly TArrowLen FLen;
        readonly TArrowWidth FWidth;
        static readonly TLineArrow FNone = new TLineArrow(TArrowStyle.None, TArrowLen.Medium, TArrowWidth.Medium); 

        /// <summary>
        /// Creates a new arrow.
        /// </summary>
        /// <param name="aStyle">Style of the arrow.</param>
        /// <param name="aLen">Length of the arrow.</param>
        /// <param name="aWidth">Width of the arrow.</param>
        internal TLineArrow(TArrowStyle aStyle, TArrowLen aLen, TArrowWidth aWidth)
        {
            FStyle = aStyle;
            FLen = aLen;
            FWidth = aWidth;
        }

        /// <summary>
        /// Style of the arrow.
        /// </summary>
        public TArrowStyle Style { get { return FStyle; } }

        /// <summary>
        /// Length of the arrow.
        /// </summary>
        public TArrowLen Len { get { return FLen; } }

        /// <summary>
        /// Width of the arrow.
        /// </summary>
        public TArrowWidth Width { get { return FWidth; } }

        /// <summary>
        /// Returns a shared instance with no arrow.
        /// </summary>
        public static TLineArrow None { get { return FNone; } }

        #region IComparable Members

        /// <summary>
        /// Returns -1, 0 or 1 depending if the objects is smaller, equal or bigger than the other.
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public int CompareTo(TLineArrow obj)
        {
            int r;
            r = Style.CompareTo(obj.Style); if (r != 0) return r;
            r = Len.CompareTo(obj.Len); if (r != 0) return r;
            r = Width.CompareTo(obj.Width); if (r != 0) return r;

            return 0;
        }

        /// <summary></summary>
        public int CompareTo(object obj)
        {
            if (!(obj is TLineArrow)) return -1;
            return CompareTo((TLineArrow)obj);
        }


        /// <summary></summary>
        public override int GetHashCode()
        {
            return HashCoder.GetHashObj(Style, Len, Width);
        }

        /// <summary></summary>
        public override bool Equals(object obj)
        {
            if (!(obj is TLineArrow)) return false;
            return CompareTo((TLineArrow)obj) == 0;
        }

        /// <summary></summary>
        public static bool operator ==(TLineArrow f1, TLineArrow f2)
        {
            return f1.CompareTo(f2) == 0;
        }

        /// <summary></summary>
        public static bool operator !=(TLineArrow f1, TLineArrow f2)
        {
            return f1.CompareTo(f2) != 0;
        }

        /// <summary></summary>
        public static bool operator >(TLineArrow o1, TLineArrow o2)
        {
            return o1.CompareTo(o2) > 0;
        }

        /// <summary></summary>
        public static bool operator <(TLineArrow o1, TLineArrow o2)
        {
            return o1.CompareTo(o2) < 0;
        }

        #endregion
    }

    #endregion
}
