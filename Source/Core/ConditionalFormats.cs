using System;

namespace FlexCel.Core
{

    #region Condition Type
    /// <summary>
    /// A list of conditional operators that you can apply in a Conditional format.
    /// </summary>
    public enum TConditionType
    {
        /// <summary>
        /// Always evaluates to false.
        /// </summary>
        NoComparison = 0x00,

        /// <summary>
        /// Value between Formula1 and Formula2.
        /// </summary>
        Between = 0x01,

        /// <summary>
        /// Value not between Formula1 and Formula2.
        /// </summary>
        NotBetween = 0x02,

        /// <summary>
        /// Value equal to Formula1.
        /// </summary>
        Equal = 0x03,

        /// <summary>
        /// Value not equal to Formula1.
        /// </summary>
        NotEqual = 0x04,

        /// <summary>
        /// Value Greater than Formula1.
        /// </summary>
        GreaterThan = 0x05,

        /// <summary>
        /// Less than Formula1.
        /// </summary>
        LessThan = 0x06,

        /// <summary>
        /// Greater of Equal to Formula1.
        /// </summary>
        GreaterOrEqual = 0x07,

        /// <summary>
        /// Less of Equal to Formula1.
        /// </summary>
        LessOrEqual = 0x08
    }
    #endregion

    #region Conditional Format Definitions
    /// <summary>
    /// The format to apply when a <see cref="TConditionalFormatRule"/> is applied.
    /// </summary>
    public abstract class TConditionalFormatDef : ICloneable
    {
        #region ICloneable Members

        /// <summary>
        /// Returns a deep copy of this object.
        /// </summary>
        /// <returns>A deep copy o a TConditionalFormatDef instance.</returns>
        public virtual object Clone()
        {
            return MemberwiseClone();
        }

        #endregion

#if (COMPACTFRAMEWORK && !FRAMEWORK20)
		public static bool Equals(Object o1, Object o2)
		{
			return (o1 != null && o1.Equals(o2)) || (o1 == null && o2 == null);
		}
#endif
    }

    /// <summary>
    /// Defines a format to apply for cells when a rule evaluates to true.
    /// </summary>
    public class TConditionalFormatDefStandard : TConditionalFormatDef, ICloneable
    {
        #region Privates
        #region Font
        private bool FApplyFontSize20;
        private int FFontSize20;
        private bool FApplyFontColor;
        private TExcelColor FFontColor;

        private bool FApplyFontStyleBoldAndItalic;
        private bool FApplyFontStyleStrikeout;
        private bool FApplyFontStyleSubSuperscript;
        private TFlxFontStyles FFontStyle;

        private bool FApplyFontUnderline;
        private TFlxUnderline FFontUnderline;
        #endregion

        #region Pattern
        private bool FApplyPatternStyle;
        private TFlxPatternStyle FPatternStyle;

        private bool FApplyPatternFg;
        private TExcelColor FPatternFgColor;

        private bool FApplyPatternBg;
        private TExcelColor FPatternBgColor;


        #endregion

        #region Borders
        private bool FApplyBorderLeft;
        private TFlxOneBorder FBorderLeft;
        private bool FApplyBorderRight;
        private TFlxOneBorder FBorderRight;
        private bool FApplyBorderTop;
        private TFlxOneBorder FBorderTop;
        private bool FApplyBorderBottom;
        private TFlxOneBorder FBorderBottom;
        #endregion
        #endregion

        #region Constructors
        /// <summary>
        /// Creates an empty instance, where no format applies.
        /// </summary>
        public TConditionalFormatDefStandard()
        {
        }
        #endregion

        #region Publics
        #region Font

        /// <summary>
        /// When true the font size specified in <see cref="FontSize20"/> will be applied, else it will be ignored.
        /// </summary>
        public bool ApplyFontSize20 { get { return FApplyFontSize20; } set { FApplyFontSize20 = value; } }

        /// <summary>
        /// Font size in 1/20 of a point when <see cref="ApplyFontSize20"/> is true.
        /// </summary>
        public int FontSize20 { get { return FFontSize20; } set { FFontSize20 = value; } }

        /// <summary>
        /// When true the font color specified in <see cref="FontColor"/> will be applied, else it will be ignored.
        /// </summary>
        public bool ApplyFontColor { get { return FApplyFontColor; } set { FApplyFontColor = value; } }

        /// <summary>
        /// Font color index on the color palette when <see cref="ApplyFontColor"/> is true.
        /// </summary>
        public TExcelColor FontColor { get { return FFontColor; } set { FFontColor = value; } }

        /// <summary>
        /// When true, the font style on <see cref="FontStyle"/> for bold an italics will be applied. When false, it will be ignored.
        /// </summary>
        public bool ApplyFontStyleBoldAndItalic { get { return FApplyFontStyleBoldAndItalic; } set { FApplyFontStyleBoldAndItalic = value; } }

        /// <summary>
        /// When true, the font style on <see cref="FontStyle"/> for strikeout will be applied. When false, it will be ignored.
        /// </summary>
        public bool ApplyFontStyleStrikeout { get { return FApplyFontStyleStrikeout; } set { FApplyFontStyleStrikeout = value; } }

        /// <summary>
        /// When true, the font style on <see cref="FontStyle"/> for subscripts and superscripts will be applied. When false, it will be ignored.
        /// </summary>
        public bool ApplyFontStyleSubSuperscript { get { return FApplyFontStyleSubSuperscript; } set { FApplyFontStyleSubSuperscript = value; } }

        /// <summary>
        /// Style of the font, such as bold or italics when <see cref="ApplyFontStyleBoldAndItalic"/> is true and/or <see cref="ApplyFontStyleStrikeout"/> is true and/or <see cref="ApplyFontStyleSubSuperscript"/> is true. Underline is a different option.
        /// </summary>
        public TFlxFontStyles FontStyle { get { return FFontStyle; } set { FFontStyle = value; } }

        /// <summary>
        /// When true the font underline specified in <see cref="FontUnderline"/> will be applied, else it will be ignored.
        /// </summary>
        public bool ApplyFontUnderline { get { return FApplyFontUnderline; } set { FApplyFontUnderline = value; } }

        /// <summary>
        /// Underline type, when <see cref="ApplyFontUnderline"/> = true.
        /// </summary>
        public TFlxUnderline FontUnderline { get { return FFontUnderline; } set { FFontUnderline = value; } }

        #endregion

        #region Pattern
        /// <summary>
        /// When true the pattern style specified in <see cref="PatternStyle"/> will be applied, else it will be ignored.
        /// </summary>
        public bool ApplyPatternStyle { get { return FApplyPatternStyle; } set { FApplyPatternStyle = value; } }

        /// <summary>
        /// Pattern style to apply, when <see cref="ApplyPatternStyle"/> is true.
        /// </summary>
        public TFlxPatternStyle PatternStyle { get { return FPatternStyle; } set { FPatternStyle = value; } }

        /// <summary>
        /// When true the foreground color specified in <see cref="PatternFgColor"/> will be used, else it will be ignored.
        /// </summary>
        public bool ApplyPatternFg { get { return FApplyPatternFg; } set { FApplyPatternFg = value; } }

        /// <summary>
        /// Foreground color of the pattern. This value is the only color used when <see cref="PatternStyle"/> is solid. Only applies if <see cref="ApplyPatternFg"/> is true.
        /// </summary>
        public TExcelColor PatternFgColor { get { return FPatternFgColor; } set { FPatternFgColor = value; } }

        /// <summary>
        /// When true the background color specified in <see cref="PatternBgColor"/> will be used, else it will be ignored.
        /// </summary>
        public bool ApplyPatternBg { get { return FApplyPatternBg; } set { FApplyPatternBg = value; } }

        /// <summary>
        /// Background color of the pattern. This value is *NOT* used when <see cref="PatternStyle"/> is solid. Only applies if <see cref="ApplyPatternBg"/> is true.
        /// </summary>
        public TExcelColor PatternBgColor { get { return FPatternBgColor; } set { FPatternBgColor = value; } }

        #endregion

        #region Borders
        /// <summary>
        /// If true, the <see cref="BorderLeft"/> settings will be applied, esle they will be ignored.
        /// </summary>
        public bool ApplyBorderLeft { get { return FApplyBorderLeft; } set { FApplyBorderLeft = value; } }

        /// <summary>
        /// Color and style for the cell border.
        /// </summary>
        public TFlxOneBorder BorderLeft { get { return FBorderLeft; } set { FBorderLeft = value; } }

        /// <summary>
        /// If true, the <see cref="BorderRight"/> settings will be applied, esle they will be ignored.
        /// </summary>
        public bool ApplyBorderRight { get { return FApplyBorderRight; } set { FApplyBorderRight = value; } }

        /// <summary>
        /// Color and style for the cell border.
        /// </summary>
        public TFlxOneBorder BorderRight { get { return FBorderRight; } set { FBorderRight = value; } }

        /// <summary>
        /// If true, the <see cref="BorderTop"/> settings will be applied, esle they will be ignored.
        /// </summary>
        public bool ApplyBorderTop { get { return FApplyBorderTop; } set { FApplyBorderTop = value; } }

        /// <summary>
        /// Color and style for the cell border.
        /// </summary>
        public TFlxOneBorder BorderTop { get { return FBorderTop; } set { FBorderTop = value; } }

        /// <summary>
        /// If true, the <see cref="BorderBottom"/> settings will be applied, esle they will be ignored.
        /// </summary>
        public bool ApplyBorderBottom { get { return FApplyBorderBottom; } set { FApplyBorderBottom = value; } }

        /// <summary>
        /// Color and style for the cell border.
        /// </summary>
        public TFlxOneBorder BorderBottom { get { return FBorderBottom; } set { FBorderBottom = value; } }
        #endregion

        #region Utils
        /// <summary>
        /// Returns true if any font formatting is applied
        /// </summary>
        public bool HasFontBlock
        {
            get
            {
                return ApplyFontColor || ApplyFontSize20 || ApplyFontStyleBoldAndItalic
                    || ApplyFontStyleStrikeout || ApplyFontStyleSubSuperscript || ApplyFontUnderline;
            }
        }

        /// <summary>
        /// Returns true if any border formatting is applied.
        /// </summary>
        public bool HasBorderBlock
        {
            get
            {
                return ApplyBorderBottom || ApplyBorderLeft || ApplyBorderRight || ApplyBorderTop;
            }
        }

        /// <summary>
        /// Returns true if any pattern formatting is applied.
        /// </summary>
        public bool HasPatternBlock
        {
            get
            {
                return ApplyPatternBg || ApplyPatternFg || ApplyPatternStyle;
            }
        }

        /// <summary>
        /// Returns true if any format is applied.
        /// </summary>
        public bool HasFormat
        {
            get
            {
                return HasFontBlock || HasBorderBlock || HasPatternBlock;
            }
        }

        #endregion

        #endregion

        #region ICloneable Members

        /// <summary>
        /// Returns a deep copy of the object.
        /// </summary>
        /// <returns></returns>
        public override object Clone()
        {
            return MemberwiseClone();  //no classes here.
        }

        #endregion

        /// <summary>
        /// Returns true if this object is equal to obj.
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public override bool Equals(object obj)
        {
            TConditionalFormatDefStandard r2 = obj as TConditionalFormatDefStandard;
            if (r2 == null) return false;

            #region Font
            if (FApplyFontSize20 != r2.FApplyFontSize20) return false;
            if (FApplyFontSize20 && FFontSize20 != r2.FFontSize20) return false;
            if (FApplyFontColor != r2.FApplyFontColor) return false;
            if (FApplyFontColor && FontColor != r2.FontColor) return false;

            if (FApplyFontStyleBoldAndItalic != r2.FApplyFontStyleBoldAndItalic) return false;
            if (FApplyFontStyleStrikeout != r2.FApplyFontStyleStrikeout) return false;
            if (FApplyFontStyleSubSuperscript != r2.FApplyFontStyleSubSuperscript) return false;

            if ((FApplyFontStyleBoldAndItalic || FApplyFontStyleStrikeout || FApplyFontStyleSubSuperscript) && FFontStyle != r2.FFontStyle) return false;

            if (FApplyFontUnderline != r2.FApplyFontUnderline) return false;
            if (FApplyFontUnderline && FFontUnderline != r2.FFontUnderline) return false;
            #endregion

            #region Pattern
            if (FApplyPatternStyle != r2.FApplyPatternStyle) return false;
            if (FApplyPatternStyle && FPatternStyle != r2.FPatternStyle) return false;

            if (FApplyPatternFg != r2.FApplyPatternFg) return false;
            if (FApplyPatternFg && PatternFgColor != r2.PatternFgColor) return false;

            if (FApplyPatternBg != r2.FApplyPatternBg) return false;
            if (FApplyPatternBg && PatternBgColor != r2.PatternBgColor) return false;


            #endregion

            #region Borders
            if (FApplyBorderLeft != r2.FApplyBorderLeft) return false;
            if (FApplyBorderLeft && FBorderLeft != r2.FBorderLeft) return false;
            if (FApplyBorderRight != r2.FApplyBorderRight) return false;
            if (FApplyBorderRight && FBorderRight != r2.FBorderRight) return false;
            if (FApplyBorderTop != r2.FApplyBorderTop) return false;
            if (FApplyBorderTop && FBorderTop != r2.FBorderTop) return false;
            if (FApplyBorderBottom != r2.FApplyBorderBottom) return false;
            if (FApplyBorderBottom && FBorderBottom != r2.FBorderBottom) return false;
            #endregion

            return true;
        }

        /// <summary>
        /// Returns the hashcode for this object.
        /// </summary>
        /// <returns></returns>
        public override int GetHashCode()
        {
            return HashCoder.GetHash(
                FApplyFontSize20.GetHashCode(),
                FApplyFontSize20 ? FFontSize20.GetHashCode() : 0,
                FApplyFontColor.GetHashCode(),
                FApplyFontColor ? FontColor.GetHashCode() : 0,

                FApplyFontStyleBoldAndItalic.GetHashCode(),
                FApplyFontStyleSubSuperscript.GetHashCode(),
                FApplyFontStyleStrikeout.GetHashCode(),
                FApplyFontStyleBoldAndItalic || FApplyFontStyleSubSuperscript || FApplyFontStyleStrikeout ? FFontStyle.GetHashCode() : 0,

                FApplyFontUnderline.GetHashCode(),
                FApplyFontUnderline ? FFontUnderline.GetHashCode() : 0,

                FApplyPatternStyle.GetHashCode(),
                FApplyPatternStyle ? FPatternStyle.GetHashCode() : 0,

                FApplyPatternFg.GetHashCode(),
                FApplyPatternFg ? FPatternFgColor.GetHashCode() : 0,

                FApplyPatternBg.GetHashCode(),
                FApplyPatternBg ? PatternBgColor.GetHashCode() : 0,


                FApplyBorderLeft.GetHashCode(),
                FApplyBorderLeft ? FBorderLeft.GetHashCode() : 0,
                FApplyBorderRight.GetHashCode(),
                FApplyBorderRight ? FBorderRight.GetHashCode() : 0,
                FApplyBorderTop.GetHashCode(),
                FApplyBorderTop ? FBorderTop.GetHashCode() : 0,
                FApplyBorderBottom.GetHashCode(),
                FApplyBorderBottom ? FBorderBottom.GetHashCode() : 0
                );

        }
    }

    //Pending for Excel 2007: Conditional formats for gradient bars, etc.

    #endregion

    #region Conditional format rules
    /// <summary>
    /// A rule specifying a conditional format. You cannot create instances of this class, only of their children.
    /// </summary>
    public abstract class TConditionalFormatRule : ICloneable
    {
        private bool FStopIfTrue;
        private TConditionalFormatDef FFormatDef;

        /// <summary>
        /// Creates a new instance of TConditionalFormatRule, with the corrsponding Format definition.
        /// </summary>
        /// <param name="aFormatDef">Format definition for the rule.</param>
        /// <param name="aStopIfTrue">Only valid on Excel 2007. If true, no more conditional format rules after this one will be applied if this rule applies.</param>
        protected TConditionalFormatRule(TConditionalFormatDef aFormatDef, bool aStopIfTrue)
        {
            FStopIfTrue = aStopIfTrue;
            FFormatDef = aFormatDef;
        }

        /// <summary>
        /// When true, rules after this one will not evaluate if this one applies. Only Applies to Excel 2007
        /// </summary>
        public bool StopIfTrue { get { return FStopIfTrue; } set { FStopIfTrue = value; } }

        /// <summary>
        /// Format to apply when the rule evaluates to true.
        /// </summary>
        public TConditionalFormatDef FormatDef { get { return FFormatDef; } set { FFormatDef = value; } }

        /// <summary>
        /// Returns true if this object is equal to obj.
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public override bool Equals(object obj)
        {
            TConditionalFormatRule r2 = obj as TConditionalFormatRule;
            if (r2 == null) return false;

            if (r2.StopIfTrue != StopIfTrue) return false;
            if (!TConditionalFormatDef.Equals(FormatDef, r2.FormatDef)) return false;

            return true;

        }

        /// <summary>
        /// Returns the hashcode for this object.
        /// </summary>
        /// <returns></returns>
        public override int GetHashCode()
        {
            return base.GetHashCode();
        }

        #region ICloneable Members

        /// <summary>
        /// Returns a deep copy of the object.
        /// </summary>
        /// <returns></returns>
        public virtual object Clone()
        {
            TConditionalFormatRule Result = (TConditionalFormatRule)MemberwiseClone();
            Result.FFormatDef = (TConditionalFormatDef)FFormatDef.Clone();
            return null;
        }

        #endregion
    }

    /// <summary>
    /// A conditional format rule specified by a formula.
    /// </summary>
    public class TConditionalFormulaRule : TConditionalFormatRule, ICloneable
    {
        private string FFormula;

        /// <summary>
        /// Creates a new instace of TConditionalFormulaRule.
        /// </summary>
        /// <param name="aStopIfTrue">See <see cref="TConditionalFormatRule.StopIfTrue"/></param>
        /// <param name="aFormula">See <see cref="Formula"/></param>
        /// <param name="aFormatDef">See <see cref="TConditionalFormatRule.FormatDef"/></param>
        public TConditionalFormulaRule(TConditionalFormatDefStandard aFormatDef, bool aStopIfTrue, string aFormula)
            : base(aFormatDef, aStopIfTrue)
        {
            FFormula = aFormula;
        }

        /// <summary>
        /// The formula to be evaluated. The conditional format will be applied when it evaluates to true.
        /// <br/>Note that with <b>relative</b> references, we always consider "A1" to be the cell where the format is. This means that the formula:
        /// "=$A$1 + A1" when evaluated in Cell B8, will read "=$A$1 + B8". To provide a negative offset, you need to wrap the formula.
        /// For example "=A1048575" will evaluate to B7 when evaluated in B8.
        /// </summary>
        public string Formula { get { return FFormula; } set { FFormula = value; } }

        /// <summary>
        /// Returns true if this object is equal to obj.
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public override bool Equals(object obj)
        {
            if (!base.Equals(obj)) return false;
            TConditionalFormulaRule r2 = obj as TConditionalFormulaRule;
            if (r2 == null) return false;
            return String.Equals(r2.Formula, Formula, StringComparison.CurrentCultureIgnoreCase);
        }

        /// <summary>
        /// Returns the hashcode for this object.
        /// </summary>
        /// <returns></returns>
        public override int GetHashCode()
        {
            return base.GetHashCode();
        }


        #region ICloneable Members

        /// <summary>
        /// Returns a deep copy of the object.
        /// </summary>
        /// <returns></returns>
        public override object Clone()
        {
            return base.Clone();
        }

        #endregion
    }

    /// <summary>
    /// A conditional format rule specified by the value on a cell (less than, equal, etc).
    /// </summary>
    public class TConditionalCellValueRule : TConditionalFormatRule, ICloneable
    {
        private TConditionType FConditionType;
        private string FFormula1;
        private string FFormula2;

        /// <summary>
        /// Creates a new instace of TConditionalFormulaRule.
        /// </summary>
        /// <param name="aStopIfTrue">See <see cref="TConditionalFormatRule.StopIfTrue"/></param>
        /// <param name="aConditionType">See <see cref="ConditionType"/></param>
        /// <param name="aFormula1">See <see cref="Formula1"/></param>
        /// <param name="aFormula2">See <see cref="Formula2"/></param>
        /// <param name="aFormatDef">See <see cref="TConditionalFormatRule.FormatDef"/></param>
        public TConditionalCellValueRule(TConditionalFormatDefStandard aFormatDef, bool aStopIfTrue, TConditionType aConditionType, string aFormula1, string aFormula2)
            : base(aFormatDef, aStopIfTrue)
        {
            FConditionType = aConditionType;
            FFormula1 = aFormula1;
            FFormula2 = aFormula2;
        }

        /// <summary>
        /// Condition to apply for the rule.
        /// </summary>
        public TConditionType ConditionType { get { return FConditionType; } set { FConditionType = value; } }

        /// <summary>
        /// The first formula to be evaluated. When the condition needs only one parameter (for example when condition is "Equal") this is the only formula that is used.
        /// <br/>Note that with <b>relative</b> references, we always consider "A1" to be the cell where the format is. This means that the formula:
        /// "=$A$1 + A1" when evaluated in Cell B8, will read "=$A$1 + B8". To provide a negative offset, you need to wrap the formula.
        /// For example "=A1048575" will evaluate to B7 when evaluated in B8.
        /// </summary>
        public string Formula1 { get { return FFormula1; } set { FFormula1 = value; } }

        /// <summary>
        /// The second formula to be evaluated. Note that this formula is only used if the condition needs more than one parameter (for example when condition is "Between").
        /// If using a condition with only one parameter, you can leave this formula to null.
        /// <br/>Note that with <b>relative</b> references, we always consider "A1" to be the cell where the format is. This means that the formula:
        /// "=$A$1 + A1" when evaluated in Cell B8, will read "=$A$1 + B8". To provide a negative offset, you need to wrap the formula.
        /// For example "=A1048575" will evaluate to B7 when evaluated in B8.
        /// </summary>
        public string Formula2 { get { return FFormula2; } set { FFormula2 = value; } }

        /// <summary>
        /// Returns true if this object is equal to obj.
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public override bool Equals(object obj)
        {
            if (!base.Equals(obj)) return false;
            TConditionalCellValueRule r2 = obj as TConditionalCellValueRule;
            if (r2 == null) return false;
            if (r2.ConditionType != ConditionType) return false;
            if (!String.Equals(r2.Formula1, Formula1, StringComparison.CurrentCultureIgnoreCase)) return false;
            if (String.Equals(r2.Formula2, Formula2, StringComparison.CurrentCultureIgnoreCase)) return false;

            return true;
        }

        /// <summary>
        /// Returns the hashcode for this object.
        /// </summary>
        /// <returns></returns>
        public override int GetHashCode()
        {
            return base.GetHashCode();
        }

        #region ICloneable Members

        /// <summary>
        /// Returns a deep copy of the object.
        /// </summary>
        /// <returns></returns>
        public override object Clone()
        {
            return base.Clone();
        }

        #endregion
    }

    #endregion

}
