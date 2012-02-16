using System;
using System.Text;
using System.Diagnostics;
using FlexCel.Core;
using System.Globalization;

namespace FlexCel.XlsAdapter
{
    /// <summary>
    /// This static class takes an RPN formula and converts it to a text representation.
    /// </summary>
    internal sealed class TFormulaConvertInternalToText
    {
        private TFormulaConvertInternalToText()
        {
        }

        internal static string AsString(TParsedTokenList Tokens, int CellRow, int CellCol, ICellList CellList)
        {
            TWorkbookGlobals Globals = null;
            if (CellList != null) Globals = CellList.Globals;
            return AsString(Tokens, CellRow, CellCol, CellList, Globals, -1, false);
        }

        internal static string AsString(TParsedTokenList Tokens, int CellRow, int CellCol, ICellList CellList, TWorkbookGlobals Globals, int MaxStringConstantLen, bool WritingXlsx)
        {
            return AsString(Tokens, CellRow, CellCol, CellList, Globals, MaxStringConstantLen, WritingXlsx, false);

        }

        internal static string AsString(TParsedTokenList Tokens, int CellRow, int CellCol, 
            ICellList CellList, TWorkbookGlobals Globals, int MaxStringConstantLen, bool WritingXlsx, bool SkipEqual)
        {
            string StartFormula = fts(TFormulaToken.fmStartFormula); //Formulas do not always begin with "=". Array formulas begin with "{", and they override this var to empty.
			TFormulaStack ParsedStack=new TFormulaStack();

            bool R1C1 = false;
            if (Globals != null && !WritingXlsx)
            {
                R1C1 = Globals.Workbook.FormulaReferenceStyle == TReferenceStyle.R1C1;
            }

            Tokens.ResetPositionToStart();
			while (!Tokens.Eof()) 
			{
				TBaseParsedToken Token = Tokens.ForwardPop();
				ptg BaseToken = Token.GetBaseId;
				Evaluate(Token, R1C1, CellRow, CellCol, BaseToken, Tokens, CellList, Globals, ParsedStack, ref StartFormula, MaxStringConstantLen, WritingXlsx);
			} //while


            if (WritingXlsx || SkipEqual)
            {
                StartFormula = String.Empty;
            }

            if (ParsedStack.Count == 0) return String.Empty;  //StartFormula + TFormulaMessages.ErrString(TFlxFormulaErrorValue.ErrRef);  This is needed for deleted named ranges.
            return StartFormula + ParsedStack.Pop();			
		}
    
        private static string fts(TFormulaToken t)
        {
            return TFormulaMessages.TokenString(t);   
        }

        private static void Evaluate(TBaseParsedToken Token, bool R1C1, int CellRow, int CellCol, ptg BaseToken, TParsedTokenList RPN, ICellList CellList, TWorkbookGlobals Globals, TFormulaStack ParsedStack, ref string StartFormula, int MaxStringConstantLen, bool WritingXlsx)
        {
            string s1; string s2; string s3;
            switch (BaseToken)
            {
                case ptg.Exp: AddArray(Token,  CellRow, CellCol, CellList, Globals, ParsedStack, MaxStringConstantLen, WritingXlsx);
                    StartFormula=String.Empty; 
                    break;
                case ptg.Tbl: AddTable(R1C1, Token, CellList, ParsedStack);
                    StartFormula=String.Empty;
                    break;
                case ptg.Add: s2=ParsedStack.Pop(); s1=ParsedStack.Pop(); ParsedStack.Push(s1+ParsedStack.FmSpaces+fts(TFormulaToken.fmPlus)+s2); break;
                case ptg.Sub: s2=ParsedStack.Pop(); s1=ParsedStack.Pop(); ParsedStack.Push(s1+ParsedStack.FmSpaces+fts(TFormulaToken.fmMinus)+s2); break;
                case ptg.Mul: s2=ParsedStack.Pop(); s1=ParsedStack.Pop(); ParsedStack.Push(s1+ParsedStack.FmSpaces+fts(TFormulaToken.fmMul)+s2); break;
                case ptg.Div: s2=ParsedStack.Pop(); s1=ParsedStack.Pop(); ParsedStack.Push(s1+ParsedStack.FmSpaces+fts(TFormulaToken.fmDiv)+s2); break;
                case ptg.Power: s2=ParsedStack.Pop(); s1=ParsedStack.Pop(); ParsedStack.Push(s1+ParsedStack.FmSpaces+fts(TFormulaToken.fmPower)+s2); break;
                case ptg.Concat: s2=ParsedStack.Pop(); s1=ParsedStack.Pop(); ParsedStack.Push(s1+ParsedStack.FmSpaces+fts(TFormulaToken.fmAnd)+s2); break;
                case ptg.LT: s2=ParsedStack.Pop(); s1=ParsedStack.Pop(); ParsedStack.Push(s1+ParsedStack.FmSpaces+fts(TFormulaToken.fmLT)+s2); break;
                case ptg.LE: s2=ParsedStack.Pop(); s1=ParsedStack.Pop(); ParsedStack.Push(s1+ParsedStack.FmSpaces+fts(TFormulaToken.fmLE)+s2); break;
                case ptg.EQ: s2=ParsedStack.Pop(); s1=ParsedStack.Pop(); ParsedStack.Push(s1+ParsedStack.FmSpaces+fts(TFormulaToken.fmEQ)+s2); break;
                case ptg.GE: s2=ParsedStack.Pop(); s1=ParsedStack.Pop(); ParsedStack.Push(s1+ParsedStack.FmSpaces+fts(TFormulaToken.fmGE)+s2); break;
                case ptg.GT: s2=ParsedStack.Pop(); s1=ParsedStack.Pop(); ParsedStack.Push(s1+ParsedStack.FmSpaces+fts(TFormulaToken.fmGT)+s2); break;
                case ptg.NE: s2=ParsedStack.Pop(); s1=ParsedStack.Pop(); ParsedStack.Push(s1+ParsedStack.FmSpaces+fts(TFormulaToken.fmNE)+s2); break;
                case ptg.Isect: s2=ParsedStack.Pop(); s1=ParsedStack.Pop(); ParsedStack.Push(s1+ParsedStack.FmSpaces+fts(TFormulaToken.fmIntersect)+s2); break;
                case ptg.Union: s2=ParsedStack.Pop(); s1=ParsedStack.Pop(); ParsedStack.Push(s1+ParsedStack.FmSpaces+fts(TFormulaToken.fmUnion)+s2); break;
                case ptg.Range: s2=ParsedStack.Pop(); s1=ParsedStack.Pop(); ParsedStack.Push(s1+ParsedStack.FmSpaces+fts(TFormulaToken.fmRangeSep)+s2); break;
                case ptg.Uplus: s1=ParsedStack.Pop(); ParsedStack.Push(ParsedStack.FmSpaces+fts(TFormulaToken.fmPlus)+s1); break;
                case ptg.Uminus: s1=ParsedStack.Pop(); ParsedStack.Push(ParsedStack.FmSpaces+fts(TFormulaToken.fmMinus)+s1); break;
                case ptg.Percent: s1=ParsedStack.Pop(); ParsedStack.Push(s1+ParsedStack.FmSpaces+fts(TFormulaToken.fmPercent)); break;
                case ptg.Paren: s1=ParsedStack.Pop(); ParsedStack.Push(ParsedStack.FmPreSpaces+fts(TFormulaToken.fmOpenParen)+s1+ParsedStack.FmPostSpaces+fts(TFormulaToken.fmCloseParen)); break;
                case ptg.MissArg: ParsedStack.Push(ParsedStack.FmSpaces); break;
                case ptg.Str: 
                    ParsedStack.Push(ParsedStack.FmSpaces+fts(TFormulaToken.fmStr)
                        +GetString(((TStrDataToken)Token).GetData(), MaxStringConstantLen)+fts(TFormulaToken.fmStr)); 
                    break;
                case ptg.Attr: ProcessAttr(Token, ParsedStack); break;
                case ptg.Sheet:  break;
                case ptg.EndSheet:  break;
                case ptg.Err: ParsedStack.Push(ParsedStack.FmSpaces+GetErrorText(((TErrDataToken)Token).GetData())); break;
                case ptg.Bool: ParsedStack.Push(ParsedStack.FmSpaces + GetBoolText(((TBoolDataToken)Token).GetData())); break;
                case ptg.Int: ParsedStack.Push(ParsedStack.FmSpaces+ TFormulaMessages.FloatToString(((TIntDataToken)Token).GetData())); break;
                case ptg.Num: ParsedStack.Push(ParsedStack.FmSpaces + TFormulaMessages.FloatToString(((TNumDataToken)Token).GetData())); break;
                case ptg.Array: ParsedStack.Push(ParsedStack.FmSpaces + GetArrayText(((TArrayDataToken)Token).GetData, MaxStringConstantLen)); break;

                case ptg.Func:
                case ptg.FuncVar:
                    int ArgCount;
                    bool IsAddin;
                    StringBuilder sb = new StringBuilder();
                    TBaseFunctionToken FunctionTk = Token as TBaseFunctionToken;

                    s3 = ParsedStack.FmSpaces + GetFuncName(FunctionTk, out ArgCount, WritingXlsx, out IsAddin); 

                    if (ArgCount > 0) sb.Append(ParsedStack.Pop()); 
                    for (int i=2;i <= ArgCount;i++)  
                    {
                        s1=ParsedStack.Pop(); 
                        sb.Insert(0,s1+fts(TFormulaToken.fmFunctionSep)); 
                    }

                    if (IsAddin) s3 += ConvertInternalFunctionName(Globals, ParsedStack.Pop());
                    ParsedStack.Push( s3+fts(TFormulaToken.fmOpenParen)+ParsedStack.FmPreSpaces+sb.ToString()+ParsedStack.FmPostSpaces+fts(TFormulaToken.fmCloseParen));  
                    break;
                case ptg.Name: ParsedStack.Push(ParsedStack.FmSpaces+GetName(((TNameToken)Token).NameIndex, -1, Globals, WritingXlsx)); break;

                case ptg.RefN:
                case ptg.Ref: ParsedStack.Push(ParsedStack.FmSpaces + GetRef(R1C1, (TRefToken)Token, CellRow, CellCol)); break;

                case ptg.AreaN:
                case ptg.Area: ParsedStack.Push(ParsedStack.FmSpaces + GetArea(R1C1, (TAreaToken)Token, CellRow, CellCol)); break;

                case ptg.MemArea: break;
                case ptg.MemErr: break;
                case ptg.MemNoMem: break;
                case ptg.MemFunc: break;
                case ptg.RefErr: ParsedStack.Push(ParsedStack.FmSpaces+TFormulaMessages.ErrString(TFlxFormulaErrorValue.ErrRef)); break;
                case ptg.AreaErr: ParsedStack.Push(ParsedStack.FmSpaces+TFormulaMessages.ErrString(TFlxFormulaErrorValue.ErrRef)); break;
                case ptg.MemAreaN: break;
                case ptg.MemNoMemN: break;
                case ptg.NameX: ParsedStack.Push(ParsedStack.FmSpaces+GetNameX((TNameXToken)Token, Globals, WritingXlsx)); break;
                case ptg.Ref3d: ParsedStack.Push(ParsedStack.FmSpaces+GetRef3D(R1C1, (TRef3dToken)Token, CellRow, CellCol, Globals, false, WritingXlsx)); break;
                case ptg.Area3d: ParsedStack.Push(ParsedStack.FmSpaces+GetArea3D(R1C1, (TArea3dToken)Token, CellRow, CellCol, Globals, false, WritingXlsx)); break;
                case ptg.Ref3dErr: ParsedStack.Push(ParsedStack.FmSpaces + GetRef3D(R1C1, (TRef3dToken)Token, -1, -1, Globals, true, WritingXlsx)); break;
                case ptg.Area3dErr: ParsedStack.Push(ParsedStack.FmSpaces+ GetArea3D(R1C1, (TArea3dToken)Token, CellRow, CellCol, Globals, true, WritingXlsx));break;
                default: XlsMessages.ThrowException(XlsErr.ErrBadToken, Token);break;
            }
        }

        private static string ConvertInternalFunctionName(TWorkbookGlobals Globals, string InternalName)
        {
            TUserDefinedFunctionContainer fn = Globals.Workbook.GetUserDefinedFunction(InternalName);
            if (fn != null) return fn.Function.Name;
            return InternalName;
        }

        private static void AddTable(bool R1C1, TBaseParsedToken Token, ICellList CellList, TFormulaStack ParsedStack)
        {
            if (CellList == null || Token is TTableObjToken)
            {
                ParsedStack.Push(" <Table> ");
                return;
            }
            
            ParsedStack.Push(fts(TFormulaToken.fmOpenArray) + fts(TFormulaToken.fmStartFormula)
                + fts(TFormulaToken.fmTableText) + fts(TFormulaToken.fmOpenParen)
                + GetTableText(R1C1, CellList.TableFormula(((TTableToken)Token).Row, ((TTableToken)Token).Col))
                + fts(TFormulaToken.fmCloseParen) + fts(TFormulaToken.fmCloseArray));
        }

        private static void AddArray(TBaseParsedToken Token, int CellRow, int CellCol, ICellList CellList, TWorkbookGlobals Globals, TFormulaStack ParsedStack, int MaxStringConstantLen, bool WritingXlsx)
        {
            if (CellList == null)
            {
                ParsedStack.Push(" <Ref> ");
                return;
            }

            string Start = "";
            string Stop = "";
            if (!WritingXlsx)
            {
                Start = fts(TFormulaToken.fmOpenArray);
                Stop = fts(TFormulaToken.fmCloseArray);
            }
            ParsedStack.Push(Start
                + TFormulaConvertInternalToText.AsString(CellList.ArrayFormula(((TExp_Token)Token).Row, ((TExp_Token)Token).Col), CellRow, CellCol, CellList, Globals, MaxStringConstantLen, WritingXlsx)
                + Stop);
        }
        
		/// <summary>
        /// Returns an string token
        /// </summary>
        private static string GetString(string s, int MaxStringConstantLen)
        {
            string Result = s.Replace("\"", "\"\"");
            if (MaxStringConstantLen > 0 && s.Length > MaxStringConstantLen)
            {
                FlxMessages.ThrowException(FlxErr.ErrStringConstantInFormulaTooLong, Result, String.Empty);
            }
            return Result;
        }

        private static string GetErrorText(TFlxFormulaErrorValue err)
        {
            return TFormulaMessages.ErrString(err);
        }

        private static string GetBoolText(bool b)
        {
            if (b) return fts(TFormulaToken.fmTrue);
            else return fts(TFormulaToken.fmFalse);
        }

        private static string ReadCachedValueTxt(object obj, int MaxStringConstantLen)
        {
            if (obj is double) return TFormulaMessages.FloatToString((double)obj);
            if (obj is string) { string s = fts(TFormulaToken.fmStr) + GetString(FlxConvert.ToString(obj), MaxStringConstantLen) + fts(TFormulaToken.fmStr); return s; }
            if (obj is bool) return GetBoolText((bool)obj);
            if (obj is TFlxFormulaErrorValue) return GetErrorText((TFlxFormulaErrorValue)obj);

            XlsMessages.ThrowException(XlsErr.ErrBadToken, FlxConvert.ToString(obj));
            return String.Empty;
        }

        private static string GetArrayText(object[,] dt, int MaxStringConstantLen)
        {
            StringBuilder Result= new StringBuilder( fts(TFormulaToken.fmOpenArray));
            string fmArrayColSep=fts(TFormulaToken.fmArrayColSep);
            string fmArrayRowSep=fts(TFormulaToken.fmArrayRowSep);
            string Sep=String.Empty;
            for (int r=0; r< dt.GetLength(0); r++)
            {
                for (int c=0; c< dt.GetLength(1);c++)
                {
                    Result.Append(Sep+ReadCachedValueTxt(dt[r,c], MaxStringConstantLen));
                    if (c==0) Sep=fmArrayColSep;
                }
                Sep=fmArrayRowSep;
            }
            Result.Append(fts(TFormulaToken.fmCloseArray));
            return Result.ToString();
        }

        private static string Get1Ref(bool R1C1, int Row, int Col, int CellRow, int CellCol, bool RowAbs, bool ColAbs)
        {
            if (Row < 1 || Col < 1 || Row > FlxConsts.Max_Rows + 1 || Col > FlxConsts.Max_Columns + 1) return TFormulaMessages.ErrString(TFlxFormulaErrorValue.ErrRef);

            if (R1C1) return Get1R1C1Ref(Row, Col, CellRow, CellCol, RowAbs, ColAbs);

            string Result="";  //We won't use a string builder here as it would probably have more overhead.
            if (ColAbs) Result=fts(TFormulaToken.fmAbsoluteRef);
            Result= Result+TCellAddress.EncodeColumn(Col);

            if (RowAbs) Result= Result+fts(TFormulaToken.fmAbsoluteRef);
            return Result+Row.ToString();
        }

        internal static string Get1R1C1Ref(int Row, int Col, int CellRow, int CellCol, bool RowAbs, bool ColAbs)
        {
            return TCellAddress.GetR1C1Ref(Row, Col, CellRow, CellCol, RowAbs, ColAbs);
        }

        private static string GetRef(bool R1C1, TRefToken Ref, int CellRow, int CellCol)
        {
            return Get1Ref(R1C1, Ref.GetRow1(CellRow) + 1, Ref.GetCol1(CellCol) + 1, CellRow, CellCol, Ref.RowAbs, Ref.ColAbs);
        }

        private static string GetRowRange(bool R1C1, int Row1, int Row2, int CellRow, bool Abs1, bool Abs2)
        {
            if (R1C1) return GetFullR1C1Range(TFormulaToken.fmR1C1_R, Row1, Row2, CellRow, Abs1, Abs2);
            string Result = "";

            if (Abs1) Result=fts(TFormulaToken.fmAbsoluteRef);
            Result = Result + Row1.ToString() + fts(TFormulaToken.fmRangeSep);
            if (Abs2) Result = Result + fts(TFormulaToken.fmAbsoluteRef);
            return Result+Row2.ToString();
        }

        private static string GetFullR1C1Range(TFormulaToken RC, int RowCol1, int RowCol2, int CellRowCol, bool Abs1, bool Abs2)
        {
            string Result = fts(RC);
            Result += GetR1SimpleRef(RowCol1, CellRowCol, Abs1);
            if (Abs2 == Abs1 && RowCol2 == RowCol1) return Result;
            Result = Result + RowCol1.ToString() + fts(TFormulaToken.fmRangeSep);
            Result = fts(RC);
            Result += GetR1SimpleRef(RowCol2, CellRowCol, Abs2);
            return Result;
        }

        private static string GetR1SimpleRef(int RowCol, int CellRowCol, bool Abs)
        {
            return TCellAddress.GetR1SimpleRef(RowCol, CellRowCol, Abs);
        }

        private static string GetColRange(bool R1C1, int Col1, int Col2, int CellCol, bool Abs1, bool Abs2)
        {
            if (R1C1) return GetFullR1C1Range(TFormulaToken.fmR1C1_C, Col1, Col2, CellCol, Abs1, Abs2);

            string Result="";
            if (Abs1) Result=fts(TFormulaToken.fmAbsoluteRef);
            Result=Result+TCellAddress.EncodeColumn(Col1)+fts(TFormulaToken.fmRangeSep);
            if (Abs2) Result=Result+fts(TFormulaToken.fmAbsoluteRef);
            return Result+TCellAddress.EncodeColumn(Col2);
        }

        private static string GetArea(bool R1C1, TAreaToken Area, int CellRow, int CellCol)
        {
            int Row1 = Area.GetRow1(CellRow) + 1;
            int Row2 = Area.GetRow2(CellRow) + 1;
            int Col1 = Area.GetCol1(CellCol) + 1;
            int Col2 = Area.GetCol2(CellCol) + 1;
          
            if (Col1 ==1 && Col2 == FlxConsts.Max_Columns + 1)
                return GetRowRange(R1C1, Row1, Row2, CellRow, Area.RowAbs1, Area.RowAbs2); 
            if (Row1==1 && Row2==FlxConsts.Max_Rows+1)
                return GetColRange(R1C1, Col1, Col2, CellCol, Area.ColAbs1, Area.ColAbs2);
 
            return Get1Ref(R1C1, Row1, Col1, CellRow, CellCol, Area.RowAbs1, Area.ColAbs1)+fts(TFormulaToken.fmRangeSep)
                + Get1Ref(R1C1, Row2, Col2, CellRow, CellCol, Area.RowAbs2, Area.ColAbs2);
        }

        private static string GetName(int NameIndex, int ExternSheet, TWorkbookGlobals Globals, bool WritingXlsx)
        {
            if (Globals == null) return " <Name> ";
            return Globals.References.GetName(ExternSheet, NameIndex - 1, Globals, WritingXlsx);
        }

        private static string GetNameX(TNameXToken Token, TWorkbookGlobals Globals, bool WritingXlsx)
        {
            if (Globals == null) return " <Name> ";
            int ExternSheetIndex = Token.FExternSheet; //This index is *not* used to get the sheet. We use the externname instead.

            string Sheet = Globals.References.GetSheetFromName(ExternSheetIndex, Token.NameIndex - 1, Globals);
            int SupBook = Globals.References.GetSupBook(ExternSheetIndex);
            string ExternalString = Globals.References.GetSheetName(SupBook, Sheet, Globals, WritingXlsx);

            return ExternalString + GetName(Token.NameIndex, ExternSheetIndex, Globals, WritingXlsx);
        }

        private static string GetRef3D(bool R1C1, TRef3dToken Token, int CellRow, int CellCol, TWorkbookGlobals Globals, bool IsErr, bool WritingXlsx)
        {
            if (Globals == null) return " <Ref> ";
            int ExternSheet = Token.FExternSheet;
            if (ExternSheet == 0xFFFF) return TFormulaMessages.ErrString(TFlxFormulaErrorValue.ErrRef);
            int Row = IsErr ? -1 : Token.GetRow1(CellRow) + 1;
            int Col = IsErr? -1: Token.GetCol1(CellCol) + 1;
            return GetSheetName(ExternSheet, Globals, WritingXlsx) + Get1Ref(R1C1, Row, Col, CellRow, CellCol, Token.RowAbs, Token.ColAbs);
        }

        private static string GetArea3D(bool R1C1, TArea3dToken Token, int CellRow, int CellCol, TWorkbookGlobals Globals, bool IsErr, bool WritingXlsx)
        {
            if (Globals == null) return " <Ref> ";

            int ExternSheet = Token.FExternSheet;
            if (ExternSheet == 0xFFFF) return TFormulaMessages.ErrString(TFlxFormulaErrorValue.ErrRef);
            int Row1 = IsErr ? -1 : Token.GetRow1(CellRow) + 1;
            int Row2 = IsErr ? -1 : Token.GetRow2(CellRow) + 1;
            int Col1 = IsErr ? -1 : Token.GetCol1(CellCol) + 1;
            int Col2 = IsErr ? -1 : Token.GetCol2(CellCol) + 1;
            
			if (Col1 ==1 && Col2 == FlxConsts.Max_Columns + 1)
                return GetSheetName(ExternSheet, Globals, WritingXlsx) + GetRowRange(R1C1, Row1, Row2, CellRow, Token.RowAbs1, Token.RowAbs2); 
			if (Row1==1 && Row2==FlxConsts.Max_Rows + 1)
                return GetSheetName(ExternSheet, Globals, WritingXlsx) + GetColRange(R1C1, Col1, Col2, CellCol, Token.ColAbs1, Token.ColAbs2);

            string RestOfArea = IsErr? String.Empty: fts(TFormulaToken.fmRangeSep) + Get1Ref(R1C1, Row2, Col2, CellRow, CellCol, Token.RowAbs2, Token.ColAbs2);
            return GetSheetName(ExternSheet, Globals, WritingXlsx) + Get1Ref(R1C1, Row1, Col1, CellRow, CellCol, Token.RowAbs1, Token.ColAbs1) + RestOfArea;
        }

        private static string GetSheetName(int ExternSheet, TWorkbookGlobals Globals, bool WritingXlsx)
        {
            if (ExternSheet == 0xFFFF) return TFormulaMessages.ErrString(TFlxFormulaErrorValue.ErrRef);
            return Globals.References.GetSheetName(ExternSheet, Globals, WritingXlsx);
        }

        private static string GetFuncName(TBaseFunctionToken Token, out int ArgCount, bool WritingXlsx, out bool IsAddin)
        {
            TCellFunctionData fd = Token.GetFunctionData();
            ArgCount = Token.ArgumentCount;

            IsAddin = Token is TUserDefinedToken;
            if (IsAddin)
            {
                ArgCount--;
                return String.Empty;
            }

            if (WritingXlsx && fd.FutureInXlsx) return fd.FutureName;
            return fd.Name;
        }

 
        private static string GetTableText(bool R1C1, TTableRecord Table)
        {
            bool IsRow = ((Table.OptionFlags & 4) != 0) || ((Table.OptionFlags & 8) != 0);
            string before = IsRow ? String.Empty : fts(TFormulaToken.fmUnion);
            string after = !IsRow ? String.Empty : fts(TFormulaToken.fmUnion);
            //In R1C1, refs are absolute. In A1 they are relative.
            string Result = before + Get1Ref(R1C1, Table.RwInpRw + 1, Table.ColInpRw + 1, 0, 0, R1C1, R1C1)+ after; //Yes.. Excel leaves an empty comma at the end for one entry tables
            if ((Table.OptionFlags & 0x08) == 0x08) //two entry table
                return Result+Get1Ref(R1C1, Table.RwInpCol + 1, Table.ColInpCol + 1, 0, 0, R1C1, R1C1);
            else return Result;
        }

        private static void ProcessAttr(TBaseParsedToken Token, TFormulaStack ParsedStack)
        {
            TAttrSpaceToken SpaceToken = Token as TAttrSpaceToken;
            if (SpaceToken != null)
            { 
                switch (SpaceToken.SpaceType)
                {
                    case FormulaAttr.bitFSpace: ParsedStack.FmSpaces += new String(' ', SpaceToken.SpaceCount); break;
                    case FormulaAttr.bitFEnter: ParsedStack.FmSpaces += new String('\u000D', SpaceToken.SpaceCount); break;
                    case FormulaAttr.bitFPreSpace: ParsedStack.FmPreSpaces += new String(' ', SpaceToken.SpaceCount); break;
                    case FormulaAttr.bitFPreEnter: ParsedStack.FmPreSpaces += new String('\u000D', SpaceToken.SpaceCount); break;
                    case FormulaAttr.bitFPostSpace: ParsedStack.FmPostSpaces += new String(' ', SpaceToken.SpaceCount); break;
                    case FormulaAttr.bitFPostEnter: ParsedStack.FmPostSpaces += new String('\u000D', SpaceToken.SpaceCount); break;
                    case FormulaAttr.bitFPreFmlaSpace: break;//not handled;
                } //case
            }

            if (Token is TAttrSumToken)
            {
                string s=ParsedStack.Pop();
                ParsedStack.Push(ParsedStack.FmSpaces + TXlsFunction.GetData(4).Name + fts(TFormulaToken.fmOpenParen) + ParsedStack.FmPreSpaces + s + ParsedStack.FmPostSpaces + fts(TFormulaToken.fmCloseParen));
            }
        }
    }
}
