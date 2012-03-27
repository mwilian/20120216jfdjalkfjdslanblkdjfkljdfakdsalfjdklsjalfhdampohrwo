using System;

namespace FlexCel.Core
{
    internal class TRowAndCols
    {
        internal int Row;
        internal int RowCount;
        internal int Col;
        internal int ColCount;

        internal TRowAndCols(int aRow, int aRowCount, int aCol, int aColCount)
        {
            Row = aRow;
            RowCount = aRowCount;
            Col = aCol;
            ColCount = aColCount;
        }
    }

    /// <summary>
    /// A class for parsing a formula into RPN format for evaluation.
    /// </summary>
    internal class TFormulaConvertTextToInternal: TBaseFormulaParser
    {
		protected TParsedTokenListBuilder FParsedDataBuilder;
		protected TParsedTokenList FParsedData;
        private string FormulaText;
		private bool FThrowExceptions;
        private int LastRefOp;

        private bool Force3D; //For named ranges
        private bool RelativeAddress; //for conditional formats. We will always consider the start of a relative ref at cell RelStartRow, RelStartCol. Names always are 3d, so no "N" tokens for them.

        private string Default3DExternSheet;

        private bool FHasErrors;
        public TFormulaConvertTextToInternal(ExcelFile aXls, int aWorkingSheet, bool aCanModifyXls, string aFormulaText, bool aThrowExceptions)
            : base(aXls, aCanModifyXls, aFormulaText, TFmReturnType.Value, false, aWorkingSheet)
        {
            FormulaText=aFormulaText;
			FThrowExceptions = aThrowExceptions;
        }

        //This method is used by CondFmts and reading xlsx, neither needs R1C1 offsets.
        public TFormulaConvertTextToInternal(ExcelFile aXls, int aWorkingSheet, bool aCanModifyXls, string aFormulaText, bool aThrowExceptions, bool aRelativeAddress)
            : this(aXls, aWorkingSheet, aCanModifyXls, aFormulaText, aThrowExceptions)
        {
            RelativeAddress = aRelativeAddress;
        }      

        //This method doesn't know about the cell it is being entered. 
        //It either has RelativeAddress true (condfmt/dataval - so it will use Ntokens and the cell doesn't matter) 
        // or ForceAbsolute = true (charts - so relative doesn't matter)
        //or it is reading xlsx (where we only use A1 notation)
        public TFormulaConvertTextToInternal (ExcelFile aXls, int aWorkingSheet, bool aCanModifyXls, string aFormulaText, bool aThrowExceptions, bool aRelativeAddress, bool aForce3D, string aDefault3DExternSheet, TFmReturnType ReturnType, bool aForceAbsolute): 
            base(aXls, aCanModifyXls, aFormulaText, ReturnType, false, aWorkingSheet)
        {
            FormulaText=aFormulaText;
            FThrowExceptions = aThrowExceptions;
            Force3D = aForce3D;
            Default3DExternSheet = aDefault3DExternSheet;
            ForceAbsolute = aForceAbsolute;
            RelativeAddress = aRelativeAddress;
        }

        
		#region Error Handling
        /// <summary>
        /// This can be used both on evaluating formulas or report expressions. On the first case we just want to return an error code. On the second, throw an exception.
        /// </summary>
        /// <param name="Err"></param>
        /// <param name="ArgCount"></param>
        private void DoError(FlxErr Err, int ArgCount)
        {
            FHasErrors = true;
            if (FThrowExceptions)
                FlxMessages.ThrowException(Err, FormulaText);
            else 
				Push (new TUnsupportedToken(ArgCount, (ptg)0));

        }

        private void DoError(FlxErr Err, int ArgCount, bool DoPush, params object[] args)
        {
            FHasErrors = true;
            if (FThrowExceptions)
                FlxMessages.ThrowException(Err, args);
            else 
                if (DoPush) Push (new TUnsupportedToken(ArgCount, (ptg)0));

        }
		#endregion

        #region StartRowCol
        private int RelStartRow(bool RowAbs)
        {
            return RowAbs ? 0 : CurrentRow;
        }

        private int RelStartCol(bool ColAbs)
        {
            return ColAbs ? 0 : CurrentCol;
        }

        public void SetStartForRelativeRefs(int aRelStartRow, int aRelStartCol)
        {
            CurrentRow = aRelStartRow;
            CurrentCol = aRelStartCol;
        }

        internal void SetReadingXlsx()
        {
            ReadingXlsx = true;
            R1C1 = false;
        }

        #endregion

        #region ConvertValueRef
        private ptg GetLastRefOp()
		{
            return FParsedDataBuilder[LastRefOp].GetId;
		}

		private void SetLastRefOp(ptg aptg, TFmReturnType RefMode)
		{

            byte newptg = (byte) aptg;
			if ((((byte)aptg) & 0x60) != 0)
				{
					switch (RefMode)
					{
						case TFmReturnType.Ref:
							newptg = (byte) (newptg & 0x9F | 0x20);
							break;
						case TFmReturnType.Value:
							newptg = (byte) (newptg & 0x9F | 0x40);
							break;
						case TFmReturnType.Array:
							newptg = (byte) (newptg | 0x60);
							break;
					}
				}

            FParsedDataBuilder[LastRefOp] = FParsedDataBuilder[LastRefOp].SetId((ptg)newptg);
        }

		private static TFmReturnType GetPtgMode(ptg aptg)
		{
			TFmReturnType PtgMode = TFmReturnType.Value;
			if (aptg == ptg.Range || aptg == ptg.Isect || aptg == ptg.Union) //binary operators with ref results.
			{
				PtgMode = TFmReturnType.Ref;
			}

			switch (((byte)aptg) & 0x60)
			{
				case 0x20: 
					PtgMode = TFmReturnType.Ref;
					break;
				case 0x60: 
					PtgMode = TFmReturnType.Array;
					break;
			}
			return PtgMode;
		}

		protected override void ConvertLastRefValueType(TFmReturnType RefMode, TParseState ParseState, bool IgnoreArray)
		{
			if (LastRefOp < 0)  //we should always call this when an operand has been added.
			{
				char c = ' ';
				if (ParsePos - 1 >= 0 && ParsePos - 1 < FormulaText.Length) c = FormulaText[ParsePos - 1];
				FlxMessages.ThrowException(FlxErr.ErrUnexpectedChar, c, ParsePos, FormulaText);
			}

			ptg aptg = GetLastRefOp();
			TFmReturnType PtgMode = GetPtgMode(aptg);

			switch (RefMode)
			{
				case TFmReturnType.Ref:	
					if (ParseState.ForcedArrayClass && ParseState.Level > 0 && PtgMode == TFmReturnType.Value)
						SetLastRefOp(aptg, TFmReturnType.Array);
					break;  

				case TFmReturnType.Value:
					if (ParseState.ForcedArrayClass && ParseState.Level > 0)
						SetLastRefOp(aptg, TFmReturnType.Array); 
					else
						if (!IgnoreArray || PtgMode != TFmReturnType.Array)
							SetLastRefOp(aptg, TFmReturnType.Value); 
					break;

				case TFmReturnType.Array:
					SetLastRefOp(aptg, TFmReturnType.Array); 
					break;

			}
			
		}

		protected override bool LastIsReference()
		{
			if (LastRefOp < 0) return false;
			ptg aptg = GetLastRefOp();
			TFmReturnType PtgMode = GetPtgMode(aptg);
			return PtgMode == TFmReturnType.Ref;
		}

		#endregion

        #region AddParsed
        protected override void AddParsedUInt16(int w)
        {
            Push(new TIntDataToken(w));
        }

        protected override void AddParsed(double d)
        {
            if (d >= 0 && d <= 0xFFFF && (UInt16)d == d)
            {
                AddParsedUInt16((UInt16)d);
            }
            else
                Push(new TNumDataToken(d));
        }

        protected override void AddParsed(string s)
        {
            Push(new TStrDataToken(s, FormulaText, (Xls.ErrorActions & TExcelFileErrorActions.ErrorOnFormulaConstantTooLong) != 0));
        }

        protected override void AddParsed(bool b)
        {
            Push(new TBoolDataToken(b));
        }

        protected override void AddParsed(TFlxFormulaErrorValue err)
        {
            Push(new TErrDataToken(err));
        }

        protected override void AddParsedName(int NamePos)
        {
            Push(new TNameToken(GetRealPtg(ptg.Name, TFmReturnType.Ref), NamePos + 1));
        }

		protected override void AddParsedExternName(int ExternSheet, int ExternName)
		{
            Push(new TNameXToken(GetRealPtg(ptg.NameX, TFmReturnType.Ref), ExternSheet, ExternName + 1));
        }

		protected override void AddParsedExternName(string ExternSheet, string ExternName)
		{
            bool IsLocal;
            int Sheet;
            int extsheet = Xls.GetExternSheet(ExternSheet, false, ReadingXlsx, out IsLocal, out Sheet);
            int extname = -1;
            if (IsLocal) 
            {
                extname = GetNamedRangeIndex(ExternName, Sheet + 1, true);
                if (ReadingXlsx && extname < 0)
                {
                    extname = Xls.AddEmptyName(ExternName, Sheet + 1);
                }
            }
            else
            {
                extname = Xls.EnsureExternName(extsheet, ExternName);
            }

            if (extname < 0)
                FlxMessages.ThrowException(FlxErr.ErrInvalidRef, ExternSheet + fts(TFormulaToken.fmExternalRef) + ExternName);

            AddParsedExternName(extsheet, extname);
        }



        protected override void AddParsedRef(int Row, int Col, bool RowAbs, bool ColAbs)
        {
            if (Force3D)
            {
                AddParsed3dRef(Default3DExternSheet, Row, Col, RowAbs, ColAbs);
                return;
            }

            if (RelativeAddress)
            {
                Push(new TRefNToken(GetRealPtg(ptg.RefN, TFmReturnType.Ref), Row - RelStartRow(RowAbs), Col - RelStartCol(ColAbs), RowAbs, ColAbs, false));
            }
            else
            {
                Push(new TRefToken(GetRealPtg(ptg.Ref, TFmReturnType.Ref), Row, Col, RowAbs, ColAbs));
            }
        }

        protected override void AddParsedArea(int Row1, int Row2, int Col1, int Col2, bool RowAbs1, bool RowAbs2, bool ColAbs1, bool ColAbs2)
        {
            if (Force3D)
            {
                AddParsed3dArea(Default3DExternSheet, Row1, Row2, Col1, Col2, RowAbs1, RowAbs2, ColAbs1, ColAbs2);
                return;
            }

            if (RelativeAddress)
            {
                Push(new TAreaNToken(GetRealPtg(ptg.AreaN, TFmReturnType.Ref), Row1 - RelStartRow(RowAbs1), Col1 - RelStartCol(ColAbs1), RowAbs1, ColAbs1, Row2 - RelStartRow(RowAbs2), Col2 - RelStartCol(ColAbs2), RowAbs2, ColAbs2, false));
            }
            else
            {
                Push(new TAreaToken(GetRealPtg(ptg.Area, TFmReturnType.Ref), Row1, Col1, RowAbs1, ColAbs1, Row2, Col2, RowAbs2, ColAbs2));
            }
        }

        protected override void AddParsed3dRef(string ExternSheet, int Row, int Col, bool RowAbs, bool ColAbs)
        {
            if (RelativeAddress)
            {
                Push(new TRef3dNToken(GetRealPtg(ptg.Ref3d, TFmReturnType.Ref), Xls.GetExternSheet(ExternSheet, ReadingXlsx), Row - RelStartRow(RowAbs), Col - RelStartCol(ColAbs), RowAbs, ColAbs, false));
            }
            else
            {
                Push(new TRef3dToken(GetRealPtg(ptg.Ref3d, TFmReturnType.Ref), Xls.GetExternSheet(ExternSheet, ReadingXlsx), Row, Col, RowAbs, ColAbs));
            }
        }

        protected override void AddParsed3dRefErr(string ExternSheet)
        {
            Push(new TRef3dToken(GetRealPtg(ptg.Ref3dErr, TFmReturnType.Ref), Xls.GetExternSheet(ExternSheet, ReadingXlsx), 0, 0, false, false));
        }

        protected override void AddParsedRefErr()
        {
            if (Force3D)
            {
                Push(new TRef3dToken(GetRealPtg(ptg.Ref3dErr, TFmReturnType.Ref), 0xFFFF, 0, 0, false, false));
            }
            else
            {
                Push(new TRefToken(GetRealPtg(ptg.RefErr, TFmReturnType.Ref), 0, 0, false, false));
            }
        }

        protected override void AddParsed3dArea(string ExternSheet, int Row1, int Row2, int Col1, int Col2, bool RowAbs1, bool RowAbs2, bool ColAbs1, bool ColAbs2)
        {
            if (RelativeAddress)
            {
                //Really we should get an offset here, but as there is no active cell, we will assume the cell where the cursor is on is A1. So Row and Col are the offsets.
                Push(new TArea3dNToken(GetRealPtg(ptg.Area3d, TFmReturnType.Ref), Xls.GetExternSheet(ExternSheet, ReadingXlsx), Row1 - RelStartRow(RowAbs1), Col1 - RelStartCol(ColAbs1), RowAbs1, ColAbs1, Row2 - RelStartRow(RowAbs1), Col2 - RelStartCol(ColAbs2), RowAbs2, ColAbs2, false));
            }
            else
            {
                Push(new TArea3dToken(GetRealPtg(ptg.Area3d, TFmReturnType.Ref), Xls.GetExternSheet(ExternSheet, ReadingXlsx), Row1, Col1, RowAbs1, ColAbs1, Row2, Col2, RowAbs2, ColAbs2));
            }
        }

        protected override void AddParsedSpace(byte Count, FormulaAttr Kind)
        {
            AddParsedNoPop(new TAttrSpaceToken(Kind, Count, false));
        }   

        protected override void AddParsedParen()
        {
            Push(TParenToken.Instance);
        }

		protected override void AddParsedSep(byte b)
		{
			switch ((ptg)b)
			{
				case ptg.Isect: Push(TISectToken.Instance); break;
				case ptg.Union: Push(TUnionToken.Instance); break;
				case ptg.Range: Push(TRangeToken.Instance); break;
				default: DoError(FlxErr.ErrFormulaInvalid, 2);break;
			}   
		}

		protected override void AddParsedOp(TOperator op)
		{
			TBaseParsedToken OpToken = TParsedTokenListBuilder.GetParsedOp(op);
			if (OpToken is TUnsupportedToken)
			{
				DoError(FlxErr.ErrFormulaInvalid, 2);			
			}
			else
			{
				Push(OpToken);
			}
		}

		protected override void AddParsedMissingArg()
		{
            Push(TMissingArgDataToken.Instance);
		}

		protected override void AddParsedArray(object[,] ArrayData)
		{
            Push(new TArrayDataToken(GetRealPtg(ptg.Array, TFmReturnType.Array), ArrayData));
		}


        protected override void AddParsedFunction(TCellFunctionData Func, byte ArgCount)
        {
            ptg FmlaPtg;
            if (Func.MinArgCount != Func.MaxArgCount || Func.FutureInXls)
            {
                FmlaPtg = GetRealPtg(ptg.FuncVar, Func.ReturnType);
            }
            else
            {
                FmlaPtg = GetRealPtg(ptg.Func, Func.ReturnType);
            }

            TBaseParsedToken FmlaToken = TParsedTokenListBuilder.GetParsedFormula(FmlaPtg, Func, ArgCount);         
            Push(FmlaToken);
        }
        
        #endregion

        protected void Push (TBaseParsedToken obj)
        {
            PopWhiteSpace();
            AddParsedNoPop(obj);
        }

        private void AddParsedNoPop(TBaseParsedToken obj)
        {
            if (obj.GetId != ptg.Paren && obj.GetId != ptg.Attr) //Those are "transparent" for reference ops.
            {
                LastRefOp = FParsedDataBuilder.Count;
            }
            FParsedDataBuilder.Add(obj);
        }


        protected override TCellFunctionData FuncNameArray(string FuncName)
        {
            return TXlsFunction.GetData(FuncName);
        }

        private static ptg GetRealPtg(ptg PtgBase, TFmReturnType ReturnType)
        {
            switch (ReturnType)
            {
                case TFmReturnType.Array: return (ptg)(PtgBase + 0x40);
                case TFmReturnType.Ref: return (ptg)PtgBase;

                default: return (ptg)(PtgBase + 0x20);
            } //case
        }

        
        #region Public
        public virtual void Parse()
        {
            LastRefOp = -1;
            FHasErrors = false;
            FParsedDataBuilder = new TParsedTokenListBuilder();
			try
			{
				Go();
				FParsedData = FParsedDataBuilder.ToParsedTokenList();
                FParsedData.TextLenght = FormulaText.Length;
			}
			finally
			{
				FParsedDataBuilder = null;
			}

            //Try to decode what we encoded
            //something like "= >" will be encoded nicely, but will crash when decoded
			/*try
			{
				FParsedData.ResetPositionToLast();
				FParsedData.Flush();
				if (!FParsedData.Bof())FlxMessages.ThrowException(FlxErr.ErrFormulaInvalid, FormulaText);
			}
			catch (Exception)
			{
				FlxMessages.ThrowException(FlxErr.ErrFormulaInvalid,FormulaText);
			}*/

        }

        public TParsedTokenList GetTokens()
        {
            return FParsedData;
        }

        public bool HasErrors { get { return FHasErrors; } }
        #endregion
    }
}
