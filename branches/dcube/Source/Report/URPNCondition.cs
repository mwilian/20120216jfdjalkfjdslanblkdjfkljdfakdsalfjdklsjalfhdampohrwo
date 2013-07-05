using System;
using FlexCel.Core;
using System.Globalization;

namespace FlexCel.Report
{
	/// <summary>
	/// Holds a definition for a #IF or #evaluate tag.
	/// </summary>
	internal class TRPNExpression
	{
        private TParsedTokenList Fmla;
		private TWorkbookInfo wi;

		internal TRPNExpression(string Condition, ExcelFile Xls, TBand CurrentBand, FlexCelReport fr, TStackData Stack)
		{
			if (Condition != null && !Condition.StartsWith(TFormulaMessages.TokenString(TFormulaToken.fmStartFormula))) Condition = TFormulaMessages.TokenString(TFormulaToken.fmStartFormula) + Condition;

			TFormulaConvertTextWithTagsToInternal Parser = new TFormulaConvertTextWithTagsToInternal(Xls, Condition, CurrentBand, fr, Stack);
			wi = new TWorkbookInfo(Xls, Xls.ActiveSheet, 0, 0, 0, 0, 0, 0, false);
            Parser.Parse();
            Fmla = Parser.GetTokens();
		}

        internal bool IsTrue(int RowAbs, int ColAbs, int RowOfs, int ColOfs, TDebugStack aDebugStack, int FullDataSetIndex)
        {
			wi.Row = RowAbs; wi.Col = ColAbs;
			wi.RowOfs = RowOfs; wi.ColOfs = ColOfs;
			wi.DebugStack = aDebugStack;
			wi.FullDataSetIndex = FullDataSetIndex;
            return Convert.ToBoolean(Fmla.EvaluateAll(wi, new TCalcState(), new TCalcStack()), CultureInfo.CurrentCulture);
        }

		internal object Evaluate(int RowAbs, int ColAbs, int RowOfs, int ColOfs, TDebugStack aDebugStack, int FullDataSetIndex)
		{
			wi.Row = RowAbs; wi.Col = ColAbs;
			wi.RowOfs = RowOfs; wi.ColOfs = ColOfs;
			wi.FullDataSetIndex = FullDataSetIndex;
			wi.DebugStack = aDebugStack;
            return Fmla.EvaluateAll(wi, new TCalcState(), new TCalcStack());
		}

	}

    /// <summary>
    /// A #tag.
    /// </summary>
    internal class TTagToken: TBaseParsedToken
    {
        private TOneCellValue Data;
        internal TTagToken(TOneCellValue aData): base(0){Data=aData;}

        internal override ptg GetId  //Tags don't have real ids.
        {
            get { return (ptg) (0xFF); }
        }
        internal override object Evaluate(TParsedTokenList FTokenList, TWorkbookInfo wi, TBaseAggregate f, TCalcState STotal, TCalcStack CalcStack)
        {
            TValueAndXF val= new TValueAndXF();
			val.FullDataSetColumnIndex = wi.FullDataSetIndex;
			val.Workbook = wi.Xls;
			val.DebugStack = wi.DebugStack;
            Data.Evaluate(wi.Row, wi.Col, wi.RowOfs, wi.ColOfs, val); 
            return ConvertToAllowedObject(val.Value);
        }

        internal override bool Same(TBaseParsedToken aBaseParsedToken)
        {
            return false; //we don't care about this in tags.
        }
    }


    /// <summary>
    /// Specialization of TEvaluateFormula that recognizes "&lt;#data&gt;" for using with &lt;#if&gt; and &lt;#evaluate&gt;
    /// </summary>
    internal class TFormulaConvertTextWithTagsToInternal: TFormulaConvertTextToInternal
    {
        private TBand CurrentBand;
        private FlexCelReport fr;
		private TStackData Stack;

        internal TFormulaConvertTextWithTagsToInternal(ExcelFile aXls, string aFormulaText, TBand aCurrentBand, FlexCelReport afr, TStackData aStack) :
            base(aXls, aXls.ActiveSheet, false, aFormulaText, true)
		{
			MaxFormulaLen=0xFFFF;
			CurrentBand=aCurrentBand;
			fr=afr;
			Stack = aStack;
		}
       
        protected override bool DoExtraToken(char c)
        {
            if (base.DoExtraToken (c)) return true;

			int SaveParsePos=ParsePos;
			SkipWhiteSpace();
            string s= RemainingFormula;

            int z=0;
			if (s.Length>ReportTag.StrOpen.Length+ReportTag.StrClose.Length &&
				s.StartsWith(ReportTag.StrOpen)
				)
			{
				//z+=ReportTag.StrOpen.Length;

				TCellParser.GetSection(s, ref z, true);
				TOneCellValue v= TCellParser.GetCellValue(s.Substring(0, z+1), Xls, Stack, -1, CurrentBand, fr );
				for (int i=0;i<z+1;i++) NextChar();

				Push(new TTagToken(v));
				return true;
			}
			else
			{
				UndoSkipWhiteSpace(SaveParsePos);
			}
            return false;

        }

    }


}
