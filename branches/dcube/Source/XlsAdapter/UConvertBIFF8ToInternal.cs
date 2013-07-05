using System;
using FlexCel.Core;
using System.Diagnostics;
using System.Collections.Generic;


namespace FlexCel.XlsAdapter
{
    internal class TTokenOffset
    {
        private Dictionary<int, int> FList;
        internal TTokenOffset()
        {
            FList = new Dictionary<int, int>();           
        }

        internal void Add(int CurrentTokenOffs, int TokenPos)
        {
            FList.Add(CurrentTokenOffs, TokenPos);
        }

        /// <summary>
        /// Throws an exception if not found.
        /// </summary>
        internal int this[int index]
        {
            get { int r; if (!FList.TryGetValue(index, out r)) XlsMessages.ThrowException(XlsErr.ErrExcelInvalid); return r; }
            set { FList[index] = value; }
        }


        internal Dictionary<int,int>.KeyCollection Keys
        {
            get
            {
                return FList.Keys;
            }
        }

        internal void RemoveToken(int tPos)
        {
            int RemoveK = -1;
            List<int> myKeys = new List<int>(FList.Keys);
            foreach (int key in myKeys)
            {
                if (FList[key] == tPos) RemoveK = key;
                else if (FList[key] > tPos) FList[key]--;
            }

            if (RemoveK >= 0) FList.Remove(RemoveK);
        }
    }

    /// <summary>
    /// A class for parsing a RPN formula for evaluation.
    /// Note: RPN is *always* BIFF8, so no need for more than 256 columns or 65536 rows.
    /// </summary>
	internal class TFormulaConvertBiff8ToInternal
	{
		private bool Relative3dRanges; //We need this for names, since they are relative but the tokens they use are the same, ptgRef3d and ptgAread3d. For shared formulas, cond formats and data validations there is no need, since 3d refs are not allowed, and simple refs use ptfAreaN and ptgRefN.
        private bool IsFmlaObject;  //Object formulas use ptgtbl differently.

		private const int ColMask = 255;
		public TFormulaConvertBiff8ToInternal ()
		{
        }        
                
		#region Public
		public TParsedTokenList ParseRPN(TNameRecordList Names, int aRow, int aCol, byte[] Data, int atPos, int fmlaLen, bool aRelative3dRanges, out bool HasSubtotal, out bool HasAggregate, bool aIsFmlaObject)
		{
			HasSubtotal = false;
            HasAggregate = false;
			Relative3dRanges = aRelative3dRanges;
            IsFmlaObject = aIsFmlaObject;
			TParsedTokenListBuilder TokenBuilder = new TParsedTokenListBuilder();

            if (!DoRPN(TokenBuilder,Names, Data, atPos, atPos + fmlaLen, ref HasSubtotal, ref HasAggregate)) XlsMessages.ThrowException(XlsErr.ErrBadFormula, aRow + 1, aCol + 1, 0);
            return TokenBuilder.ToParsedTokenList();
		}

		internal TParsedTokenList ParseRPN(TNameRecordList Names, int aCol, int aRow, byte[] Data, int p, int FLen,  bool aRelative3dRanges)
		{
            bool HasSubtotal; bool HasAggregate;
			return ParseRPN(Names, aCol, aRow, Data, p, FLen, aRelative3dRanges, out HasSubtotal, out HasAggregate, false);
		}


		#endregion

		#region AddData
		protected static void Push(TParsedTokenListBuilder TokenBuilder, TBaseParsedToken obj)
		{
			TokenBuilder.Add(obj);
		}

		#endregion

		#region AddParsed
        protected static void AddParsed16(TParsedTokenListBuilder TokenBuilder, int t)
		{
			Push(TokenBuilder, new TIntDataToken(t));
		}

        protected static void AddParsed(TParsedTokenListBuilder TokenBuilder, double d)
		{
			Push(TokenBuilder, new TNumDataToken(d));
		}

        protected static void AddParsed(TParsedTokenListBuilder TokenBuilder, string s)
		{
			Push(TokenBuilder, new TStrDataToken(s, null, false));
		}

        protected static void AddParsed(TParsedTokenListBuilder TokenBuilder, bool b)
		{
			Push(TokenBuilder, new TBoolDataToken(b));
		}

        protected static void AddParsed(TParsedTokenListBuilder TokenBuilder, TFlxFormulaErrorValue err)
		{
			Push(TokenBuilder, new TErrDataToken(err));
		}

        protected static void AddMissingArg(TParsedTokenListBuilder TokenBuilder)
		{
			Push(TokenBuilder, TMissingArgDataToken.Instance);
		}

        protected static void AddParsedRef(TParsedTokenListBuilder TokenBuilder, ptg aId, int Rw1, int grBit1)
		{
			Push(TokenBuilder, new TRefToken(aId, Biff8Utils.ExpandBiff8Row(Rw1), Biff8Utils.ExpandBiff8Col(grBit1 & ColMask), (grBit1 & 0x8000) == 0, (grBit1 & 0x4000) == 0));
		}


        private static void GetRelativeRowAndCol(int Rw1, int grBit1, out bool RowAbs, out bool ColAbs, out int Row, out int Col)
		{
			RowAbs = (grBit1 & 0x8000) == 0;
			ColAbs = (grBit1 & 0x4000) == 0;

			unchecked
			{
				Row = RowAbs ? Biff8Utils.ExpandBiff8Row(Rw1) : (Int16)Rw1;
				Col = ColAbs ? Biff8Utils.ExpandBiff8Col(grBit1 & ColMask) : (sbyte)(grBit1 & ColMask);
			}

		}

        protected static void AddParsedRefN(TParsedTokenListBuilder TokenBuilder, ptg aId, int Rw1, int grBit1)
		{
			bool RowAbs; bool ColAbs; int Row; int Col;
			GetRelativeRowAndCol(Rw1, grBit1, out RowAbs, out ColAbs, out Row, out Col);
			Push(TokenBuilder, new TRefNToken(aId, Row, Col, RowAbs, ColAbs, true));
		}

        protected static void AddParsedArea(TParsedTokenListBuilder TokenBuilder, ptg aId, int Rw1, int Rw2, int grBit1, int grBit2)
		{
			Push(TokenBuilder, new TAreaToken(aId, Biff8Utils.ExpandBiff8Row(Rw1), Biff8Utils.ExpandBiff8Col(grBit1 & ColMask), (grBit1 & 0x8000) == 0, (grBit1 & 0x4000) == 0,
				Biff8Utils.ExpandBiff8Row(Rw2), Biff8Utils.ExpandBiff8Col(grBit2 & ColMask), (grBit2 & 0x8000) == 0, (grBit2 & 0x4000) == 0));
		}

        protected static void AddParsedAreaN(TParsedTokenListBuilder TokenBuilder, ptg aId, int Rw1, int Rw2, int grBit1, int grBit2)
		{
			bool RowAbs1; bool ColAbs1; int Row1; int Col1;
			GetRelativeRowAndCol(Rw1, grBit1, out RowAbs1, out ColAbs1, out Row1, out Col1);
			bool RowAbs2; bool ColAbs2; int Row2; int Col2;
			GetRelativeRowAndCol(Rw2, grBit2, out RowAbs2, out ColAbs2, out Row2, out Col2);

			Push(TokenBuilder, new TAreaNToken(aId, Row1, Col1, RowAbs1, ColAbs1, Row2, Col2, RowAbs2, ColAbs2, true));
		}

		private void AddParsed3dRef(TParsedTokenListBuilder TokenBuilder, ptg RealToken, int ExternSheet, int Rw1, int grBit1)
		{
			if (Relative3dRanges)
			{
				bool RowAbs1; bool ColAbs1; int Row1; int Col1;
				GetRelativeRowAndCol(Rw1, grBit1, out RowAbs1, out ColAbs1, out Row1, out Col1);
				Push(TokenBuilder, new TRef3dNToken(RealToken, ExternSheet, Row1, Col1, RowAbs1, ColAbs1, true));
			}
			else
			{
				Push(TokenBuilder, new TRef3dToken(RealToken, ExternSheet, Biff8Utils.ExpandBiff8Row(Rw1),
					Biff8Utils.ExpandBiff8Col(grBit1 & ColMask), (grBit1 & 0x8000) == 0, (grBit1 & 0x4000) == 0));
			}
		}

		private void AddParsed3dArea(TParsedTokenListBuilder TokenBuilder, ptg RealToken, int ExternSheet, int Rw1, int Rw2, int grBit1, int grBit2)
		{
			if (Relative3dRanges)
			{
				bool RowAbs1; bool ColAbs1; int Row1; int Col1;
				GetRelativeRowAndCol(Rw1, grBit1, out RowAbs1, out ColAbs1, out Row1, out Col1);
				bool RowAbs2; bool ColAbs2; int Row2; int Col2;
				GetRelativeRowAndCol(Rw2, grBit2, out RowAbs2, out ColAbs2, out Row2, out Col2);

				Push(TokenBuilder, new TArea3dNToken(RealToken, ExternSheet, Row1, Col1, RowAbs1, ColAbs1, Row2, Col2, RowAbs2, ColAbs2, true));
			}
			else
			{
				Push(TokenBuilder, new TArea3dToken(RealToken, ExternSheet, Biff8Utils.ExpandBiff8Row(Rw1),
					Biff8Utils.ExpandBiff8Col(grBit1 & ColMask), (grBit1 & 0x8000) == 0, (grBit1 & 0x4000) == 0, Biff8Utils.ExpandBiff8Row(Rw2), Biff8Utils.ExpandBiff8Col(grBit2 & ColMask),
					(grBit2 & 0x8000) == 0, (grBit2 & 0x4000) == 0));
			}
		}


        protected static void AddParsedSep(TParsedTokenListBuilder TokenBuilder, ptg b)
		{
			switch (b)
			{
				case ptg.Isect: Push(TokenBuilder, TISectToken.Instance); break;
				case ptg.Union: Push(TokenBuilder, TUnionToken.Instance); break;
				case ptg.Range: Push(TokenBuilder, TRangeToken.Instance); break;
				default: Push(TokenBuilder, new TUnsupportedToken(2, b));break;
			}   
		}

        protected static void AddParsedOp(TParsedTokenListBuilder TokenBuilder, TOperator op)
		{
			Push(TokenBuilder, TParsedTokenListBuilder.GetParsedOp(op));
		}


        protected static void AddParsedFormula(TParsedTokenListBuilder TokenBuilder, ptg FmlaPtg, TCellFunctionData Func, byte ArgCount)
		{
			TBaseParsedToken FmlaToken = TParsedTokenListBuilder.GetParsedFormula(FmlaPtg, Func, ArgCount);
			Push(TokenBuilder, FmlaToken); //Always push unsupported.
		}

        private void AddParsedTable(TParsedTokenListBuilder TokenBuilder, int Row, int Col)
		{
            if (IsFmlaObject) Push(TokenBuilder, new TTableObjToken(Row, Col));
            else Push(TokenBuilder, new TTableToken(Row, Col));
		}

        private static void AddParsedExp(TParsedTokenListBuilder TokenBuilder, int Row, int Col)
		{
			Push(TokenBuilder, new TExp_Token(Row, Col));
		}


		#endregion

		#region Evaluate RPN
        private bool DoRPN(TParsedTokenListBuilder TokenBuilder, TNameRecordList Names, byte[] RPN, int atPos, int afPos, ref bool HasSubtotal, ref bool HasAggregate)
		{
			int tPos=atPos;
			int fPos=afPos;
			int ArrayPos=fPos;

            TTokenOffset TokenOffset = new TTokenOffset();
			while (tPos<fPos) 
			{
                TokenOffset.Add(tPos, TokenBuilder.Count);
                byte RealToken= RPN[tPos];
				ptg BaseToken=TBaseParsedToken.CalcBaseToken((ptg)RealToken);
				TUnsupportedFormulaErrorType ErrType = TUnsupportedFormulaErrorType.FormulaTooComplex;
				string ErrName = null;
				if (!Evaluate(TokenBuilder, TokenOffset, Names, BaseToken, (ptg)RealToken, RPN, ref tPos, ref ArrayPos, ref ErrType, ref ErrName, ref HasSubtotal, ref HasAggregate))
				{
					TokenBuilder.Clear();                   
					return false;
				}
				tPos++;
			} //while

			TokenOffset.Add(tPos, TokenBuilder.Count); //eof

            FixGotosAndMemTokens(TokenBuilder, TokenOffset);
			return true;
		}

        private bool Evaluate(TParsedTokenListBuilder TokenBuilder, TTokenOffset TokenOffset, TNameRecordList Names, ptg BaseToken, ptg RealToken, 
            byte[]RPN, ref int tPos, ref int ArrayPos, ref TUnsupportedFormulaErrorType ErrType, ref string ErrName, ref bool HasSubtotal, ref bool HasAggregate)
		{
			switch (BaseToken)
			{
				case ptg.Exp:
					AddParsedExp(TokenBuilder, BitOps.GetWord(RPN, tPos + 1), BitOps.GetWord(RPN, tPos + 3));
					tPos += 4;
					break;
				case ptg.Tbl:
					AddParsedTable(TokenBuilder, BitOps.GetWord(RPN, tPos + 1), BitOps.GetWord(RPN, tPos + 3));
					tPos += 4;
					break;
				case ptg.Add: 
				case ptg.Sub: 
				case ptg.Mul: 
				case ptg.Div: 
				case ptg.Power: 
				case ptg.Concat: 
				case ptg.LT: 
				case ptg.LE: 
				case ptg.EQ: 
				case ptg.GE: 
				case ptg.GT: 
				case ptg.NE: 
				case ptg.Uminus: 
				case ptg.Percent:
				case ptg.Uplus:
					AddParsedOp(TokenBuilder, (TOperator)BaseToken); break;

				case ptg.MissArg: AddMissingArg(TokenBuilder);break;


				case ptg.Isect: AddParsedSep(TokenBuilder, BaseToken); break;
				case ptg.Union: AddParsedSep(TokenBuilder, BaseToken); break;
				case ptg.Range: AddParsedSep(TokenBuilder, BaseToken); break;

				case ptg.Paren: Push(TokenBuilder, TParenToken.Instance); break;
				case ptg.Str: 
					long sl=0;
					string Result=null;
					StrOps.GetSimpleString(false, RPN, tPos+1, false, 0, ref Result, ref sl);
					AddParsed(TokenBuilder, Result); 
					tPos+=(int)sl; 
					break;
				case ptg.Attr: int AttrLen=0; if (!ProcessAttr(TokenBuilder, RPN, tPos+1, ref AttrLen)) return false; tPos+= AttrLen; break;
				case ptg.Sheet:  return false;
				case ptg.EndSheet:  return false;
				case ptg.Err: AddParsed(TokenBuilder, (TFlxFormulaErrorValue)RPN[tPos+1]); tPos++; break;
				case ptg.Bool: AddParsed(TokenBuilder, RPN[tPos+1]==1); tPos++; break;
				case ptg.Int: AddParsed16(TokenBuilder, BitOps.GetWord(RPN, tPos+1)); tPos+=2; break;
				case ptg.Num: AddParsed(TokenBuilder, BitConverter.ToDouble(RPN, tPos+1)); tPos+=8; break;
				case ptg.Array: Push(TokenBuilder, GetArrayDataToken(RealToken, RPN, ref ArrayPos)); tPos+=7; break;
				case ptg.Func: 
					bool Result1;
					int index=BitOps.GetWord(RPN, tPos+1);
					TCellFunctionData fd=TXlsFunction.GetData(index, out Result1);
					if (!Result1) return false;

					Debug.Assert(fd.MinArgCount == fd.MaxArgCount, "On a fixed formula the min count of arguments should be the same as the max");
					AddParsedFormula(TokenBuilder, RealToken, fd, (byte)fd.MinArgCount);
					tPos+=2; 
					break;

				case ptg.FuncVar:
					bool Result2;
					int index2=BitOps.GetWord(RPN, tPos+2);
					TCellFunctionData fd2=TXlsFunction.GetData(index2, out Result2);
					if (!Result2) return false;

					if (fd2.Index==344) // SubTotal
						HasSubtotal = true;

					int np=RPN[tPos+1] & 0x7F;

                    if (fd2.Index == 255) CheckFutureFunction(TokenBuilder, ref np, Names, ref fd2, TokenOffset, ref HasAggregate);

                    
					AddParsedFormula(TokenBuilder, RealToken, fd2, (byte) np);

					tPos+=3; 
					break;
				case ptg.Name:
					Push(TokenBuilder, new TNameToken(RealToken, BitOps.GetWord(RPN, tPos+1)));
					tPos += 4;
					break;

				case ptg.NameX:
					Push(TokenBuilder, new TNameXToken(RealToken, BitOps.GetWord(RPN, tPos + 1), BitOps.GetWord(RPN, tPos + 3)));
					tPos += 6;
					break;

				case ptg.RefErr:
				case ptg.Ref: 
					AddParsedRef(TokenBuilder, RealToken, BitOps.GetWord(RPN, tPos+1),BitOps.GetWord(RPN, tPos+2+1)); tPos+=4; 
					break;

				case ptg.RefN:
					AddParsedRefN(TokenBuilder, RealToken, BitOps.GetWord(RPN, tPos + 1), BitOps.GetWord(RPN, tPos + 2 + 1)); tPos += 4;
					break;
                    
				case ptg.AreaErr:
				case ptg.Area: 
					AddParsedArea(TokenBuilder, RealToken, BitOps.GetWord(RPN, tPos+1),BitOps.GetWord(RPN, tPos+2+1),BitOps.GetWord(RPN, tPos+4+1),BitOps.GetWord(RPN, tPos+6+1)); tPos+=8; 
					break;

				case ptg.AreaN:
					AddParsedAreaN(TokenBuilder, RealToken, BitOps.GetWord(RPN, tPos + 1), BitOps.GetWord(RPN, tPos + 2 + 1), BitOps.GetWord(RPN, tPos + 4 + 1), BitOps.GetWord(RPN, tPos + 6 + 1)); tPos += 8;
					break;

				case ptg.MemArea:
				{
					int ArrayLen;
					Push(TokenBuilder, new TMemAreaToken(RealToken, GetMemControl(RPN, ArrayPos, out ArrayLen), tPos + 1 + 6 + BitOps.GetWord(RPN, tPos + 1 + 4)));
					ArrayPos += ArrayLen; tPos += 6;
					break; //this is an optimization, but we don't need it. 
				}

				case ptg.MemErr:
				{
					Push(TokenBuilder, new TMemErrToken(RealToken, (TFlxFormulaErrorValue) RPN[tPos + 1], tPos + 1 + 6 + BitOps.GetWord(RPN, tPos + 1 + 4)));
					tPos += 6;
					break; //this is an optimization, but we don't need it. 
				}

				case ptg.MemNoMem:
				{
					Push(TokenBuilder, new TMemNoMemToken(RealToken, tPos + 1 + 6 + BitOps.GetWord(RPN, tPos + 1 + 4)));
					tPos += 6;
					break; //this is an optimization, but we don't need it. 
				}

				case ptg.MemFunc:
				{
					Push(TokenBuilder, new TMemFuncToken(RealToken, tPos + 1 + 2 + BitOps.GetWord(RPN, tPos + 1)));
					tPos += 2;
					break; //this is an optimization, but we don't need it. 
				}

				case ptg.MemAreaN:
				{
					Push(TokenBuilder, new TMemAreaNToken(RealToken, tPos + 1 + 2 + BitOps.GetWord(RPN, tPos + 1)));
					tPos += 2;
					break; //this is an optimization, but we don't need it. 
				}

				case ptg.MemNoMemN:
				{
					Push(TokenBuilder, new TMemNoMemNToken(RealToken, tPos + 1 + 2 + BitOps.GetWord(RPN, tPos + 1)));
					tPos += 2;
					break; //this is an optimization, but we don't need it. 
				}
                  

				case ptg.Ref3dErr:
				case ptg.Ref3d:
				{
					int grBit1 = BitOps.GetWord(RPN, tPos + 4 + 1);
					int ExternSheet = BitOps.GetWord(RPN, tPos + 1);
					int Row1 = BitOps.GetWord(RPN, tPos + 2 + 1);
					AddParsed3dRef(TokenBuilder, RealToken, ExternSheet, Row1, grBit1);

					tPos += 6;
					break;
				}

				case ptg.Area3dErr:
				case ptg.Area3d:
				{
					int ExternSheet = BitOps.GetWord(RPN, tPos + 1);
					int Row1 = BitOps.GetWord(RPN, tPos + 2 + 1);
					int Row2 = BitOps.GetWord(RPN, tPos + 4 + 1);
					int grBit1 = BitOps.GetWord(RPN, tPos + 6 + 1);
					int grBit2 = BitOps.GetWord(RPN, tPos + 8 + 1);
					AddParsed3dArea(TokenBuilder, RealToken, ExternSheet, Row1, Row2, grBit1, grBit2);

					tPos += 10; 
					break;
				}

				default: return false;
			}
			return true;
		}

        private void CheckFutureFunction(TParsedTokenListBuilder TokenBuilder, ref int np, TNameRecordList Names, ref TCellFunctionData fd2, TTokenOffset TokenOffset, ref bool HasAggregate)
        {
            //We need to recursively read parameters in back order to find out the name, which is stored as far as possible from the function :(
            TParsedTokenList ParsedList = TokenBuilder.ToParsedTokenList();
            ParsedList.ResetPositionToLast();
            for (int i = 0; i < np; i++) //np is +1. But we will move below the name, then inc 1, so we know there isn't a "neutral" token like parent or memarea, intead of the real thing.
            {
                ParsedList.Flush();                
            }

            ParsedList.MoveBack();
            TNameToken bp = ParsedList.ForwardPop() as TNameToken;
            if (bp is TNameXToken) return; //This name isn't an internal 2007 name.

            if (bp == null) return;
            if (bp.NameIndex <= 0 || bp.NameIndex > Names.Count) return;
            string FunctionName = Names[bp.NameIndex - 1].Name;
            if (FunctionName.StartsWith("_xlfn.", StringComparison.InvariantCultureIgnoreCase))
            {
                TCellFunctionData fn = TXlsFunction.GetData(FunctionName.Substring("_xlfn.".Length));
                if (fn != null)
                {
                    if (fn.Index == (int)TFutureFunctions.Aggregate) HasAggregate = true;

                    fd2 = fn;
                    int tPos = ParsedList.SavePosition();
                    TokenBuilder.RemoveAt(tPos);
                    TokenOffset.RemoveToken(tPos);
                    np--;
                }
            }
        }

        private TRefRange[] GetMemControl(byte[] RPN, int ArrayPos, out int ArrayLen)
        {
			int Count = BitOps.GetWord(RPN, ArrayPos);
            ArrayPos += 2;
            ArrayLen = 2 + Count * 8;

            TRefRange[] Result = new TRefRange[Count];
            for (int i = 0; i < Count; i++)
            {
                Result[i].FirstRow = BitOps.GetWord(RPN, ArrayPos); ArrayPos += 2;
                Result[i].LastRow = BitOps.GetWord(RPN, ArrayPos); ArrayPos += 2;
                Result[i].FirstCol = BitOps.GetWord(RPN, ArrayPos); ArrayPos += 2;
                Result[i].LastCol = BitOps.GetWord(RPN, ArrayPos); ArrayPos += 2;
            }
            return Result;
        }

		private static TArrayDataToken GetArrayDataToken(ptg aId, byte[] RPN, ref int ArrayPos)
		{
			int Columns= RPN[ArrayPos]+1;
			int Rows=BitOps.GetWord(RPN, ArrayPos+1)+1;
			ArrayPos+=3;
			object[,] Result = new object[Rows, Columns];

			for (int r=0; r<Rows; r++)
			{
				for (int c=0; c<Columns;c++)
				{
					Result[r,c] = GetArrayObject(RPN, ref ArrayPos);
				}
			}
			return new TArrayDataToken(aId, Result);
		}

		private static TFlxFormulaErrorValue GetError(byte err)
		{
			return (TFlxFormulaErrorValue) err;
		}

		private static bool GetBool(byte b)
		{
			if (b==0) return false;
			else return true;
		}

		private static object GetArrayObject(byte[] RPN, ref int ArrayPos)
		{
			byte ValueType=RPN[ArrayPos];
			ArrayPos++;
			switch (ValueType)
			{
                case 0x00: ArrayPos += 8; return null;
				case 0x01: ArrayPos+=8;return BitConverter.ToDouble(RPN, ArrayPos-8);
				case 0x02: long sl=0; string s = null;
									  StrOps.GetSimpleString(true, RPN, ArrayPos, false, 0, ref s, ref sl);
					                  ArrayPos+=(int)sl; return s;
				case 0x04: ArrayPos+=8; return GetBool(RPN[ArrayPos-8]);
				case 0x10: ArrayPos+=8; return GetError(RPN[ArrayPos-8]);
				default: XlsMessages.ThrowException(XlsErr.ErrBadToken, ValueType);break;
			} //case
			return String.Empty;
		}

        private static int[] GetOptChoose(byte[] RPN, int start, int endpos)
        {
            int[] Result = new int[BitOps.GetWord(RPN, start) + 1];
            for (int i = 0; i < Result.Length; i++)
            {
                Result[i] = endpos + BitOps.GetWord(RPN, start + 2 * (i + 1));
            }

            return Result;
        }

        private bool ProcessAttr(TParsedTokenListBuilder TokenBuilder, byte[] RPN, int tPos, ref int AttrLen)
        {
            AttrLen = 3;

            if ((RPN[tPos] & 0x10) == 0x10)
            { //optimized sum
                Push(TokenBuilder, TAttrSumToken.Instance);
                return true;
            }

            if ((RPN[tPos] & 0x40) == 0x40)  //Spaces. As we can have volatile spaces, this should go before the check for volatile.
            {
                Push(TokenBuilder, new TAttrSpaceToken((FormulaAttr)RPN[tPos+1], RPN[tPos+2], (RPN[tPos] & 0x1) == 0x1));
                return true;
            }

            if ((RPN[tPos] & 0x1) == 0x1) //volatile
            {
                Push(TokenBuilder, TAttrVolatileToken.Instance);
                return true;
            }


            if ((RPN[tPos] & 0x2) == 0x2) //Optimized if
            {
                //It works the same by ignoring this and the optif tokens, but it is slower.
                //On a normal if you have:  [cond], [trueexpr], [falseexpr], [if]
                //so, it doesn't matter condition is false or true, both truexpr and falsexpr will be evaluated.
                //On an optimized if you have: [Cond], [AttrIf], [true], [Goto],[false], [Goto], [if]
                //                                                          |---------------|--------->
                //                                        |-------------------->
                //
                //So, you evaluate Cond. Next token is AttrIf, it Cond is true if will continue, if false jump before [false]
                //Now evaluate true or false, and the goto goes after if.

                Push(TokenBuilder, new TAttrOptIfToken(tPos + AttrLen + BitOps.GetWord(RPN, tPos + 1)));//temporarily store the stream offset in the token, it will be fixed by FixTokens
                return true;
            }

            if ((RPN[tPos] & 0x04) == 0x04)
            {
                //This works like the optimized if.
                //For example:
                // = Choose(Index,3/0,4)
                //would be parsed as:
                // Index, AttrChoose, 3, 0, /, Goto, 4, Goto, Choose.
                //
                // Then we can get from the index the direct position of the token, and after it there is a goto to the end.

                AttrLen += (BitOps.GetWord(RPN, tPos + 1) + 1) * 2;
                Push(TokenBuilder, new TAttrOptChooseToken(GetOptChoose(RPN, tPos + 1, tPos + 3))); //distance goes from start of jump table.
                return true;
            }

            if ((RPN[tPos] & 0x8) == 0x8 || RPN[tPos] == 0x0) //Goto.  0x0 is not documented, but might happen and means goto.
            {
                //Goto can't really be implemented with our engine.
                //The Excel engine has a stack and travels from left to right. Say you have the RPN 5,7,+:
                //Excel would push 5 and 7 to the stack, then the "+" would pop those values and push 12.
                //We use a recursive approach where the stack is implicit, so we start by "+", and it recusively gets
                //its 2 arguments. Bt if we add a jmp before the 5, we are never going to see it this way. Excel would.
                //It doesn't really matter since gotos are optimizations and can be ignored.

                Push(TokenBuilder, new TAttrGotoToken(1 + tPos + AttrLen + BitOps.GetWord(RPN, tPos + 1))); //temporarily store the stream offset in the token, it will be fixed by FixTokens
                return true;
            }

            //return false;
            return true;
        }

        private void FixGotosAndMemTokens(TParsedTokenListBuilder TokenBuilder, TTokenOffset TokenOffset)
        {
            for (int i = TokenBuilder.Count - 1; i >= 0; i--)
            {
                TAttrGotoToken tk = TokenBuilder[i] as TAttrGotoToken;
                if (tk != null)
                {
                    tk.PositionOfNextPtg = TokenOffset[tk.PositionOfNextPtg];
                    continue;
                }

                TAttrOptChooseToken ochoose = TokenBuilder[i] as TAttrOptChooseToken;
                if (ochoose != null)
                {
                    for (int z = 0; z < ochoose.PositionOfNextPtg.Length; z++)
                    {
                        ochoose.PositionOfNextPtg[z] = TokenOffset[ochoose.PositionOfNextPtg[z]];
                    }
                    continue;
                }

				TSimpleMemToken mem = TokenBuilder[i] as TSimpleMemToken;
				if (mem != null)
				{
					mem.PositionOfNextPtg = TokenOffset[mem.PositionOfNextPtg];
				}
            }
        }

        #endregion

    }
}
