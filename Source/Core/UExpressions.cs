using System;
namespace FlexCel.Core
{
    /// <summary>
    /// Operation between ranges.
    /// </summary>
    public enum TRangeOp
    {
        /// <summary>
        /// An union of the 2 ranges.
        /// </summary>
        Union,

        /// <summary>
        /// The intersection of the 2 ranges.
        /// </summary>
        Intersection,

        /// <summary>
        /// Creates a range from the Left-Top and Right-Bottom coordinates provided by the opearnds.
        /// </summary>
        Range
    }

        

    /// <summary>
    /// Enumerates the different kind of tokens you can find in a formula
    /// </summary>
    public enum TTokenType
    {
        /// <summary>
        /// A single cell address in the same sheet or to other sheet. See also CellRange for other tokens that can contain addresses.
        /// </summary>
        CellAddress,


        /// <summary>
        /// A range of cells. See CellAddress for single cell references.
        /// </summary>
        CellRange,

        /// <summary>
        /// A function like "Sum" or "If"
        /// </summary>
        Function,

        /// <summary>
        /// An operator like "+" or "-"
        /// </summary>
        Operator,

        /// <summary>
        /// Operator in ranges, like "Union" or "Intersection"
        /// </summary>
        RangeOp,

        /// <summary>
        /// Whitespace like " "
        /// </summary>
        Whitespace,

        /// <summary>
        /// A parenthesis surrounding the last token. Note that this token is not used in the formula (since RPN doesn't need parenthesis)
        /// but it is there so Excel can display them.
        /// </summary>
        Parethesis,

        /// <summary>
        /// A named range
        /// </summary>
        Name,

        /// <summary>
        /// Constant data, like "Hello" and 1 in the formula: "=IF(A1 = 1,,"Hello")"
        /// </summary>
        Data,

        /// <summary>
        /// A missing argument for a function. For example the second parameter in the formula: "=IF(A1 = 1,,"Hello")"
        /// </summary>
        MissingArgument,

        /// <summary>
        /// This token is not supported by FlexCel.
        /// </summary>
        Unsupported
    }

    /// <summary>
    /// This class and all its decendants represent a token in a formula. You can use these classes to modify formulas
    /// without having to parse the text in them, which can be a difficult task.
    /// </summary>
    public abstract class TToken
    {
        private readonly TTokenType FTokenType;

        /// <summary>
        /// Creates a new token.
        /// </summary>
        /// <param name="aTokenType"></param>
        protected TToken(TTokenType aTokenType)
        {
			FTokenType = aTokenType;
        }

        /// <summary>
        /// Returns the type of token.
        /// </summary>
        public TTokenType TokenType { get { return FTokenType; } }

    }

        /// <summary>
        /// A single cell address in the same sheet or to other sheet. See also TTokenCellRange for other tokens that can contain addresses.
        /// </summary>
        public class TTokenCellAddress : TToken
        {
            TCellAddress FAddress;
            /// <summary>
            /// Creates a new Cell address token.
            /// </summary>
            /// <param name="aAddress">Cell where this token will point to.</param>
            public TTokenCellAddress(TCellAddress aAddress): base (TTokenType.CellAddress)
            {
                Address = aAddress;
            }

            /// <summary>
            /// Cell address where this reference points to.
            /// </summary>
            public TCellAddress Address 
            {
                get {return FAddress;} 
                set
                {
                    if (value == null) FlxMessages.ThrowException(FlxErr.ErrNullParameter, "Address");
                    FAddress = value;
                }
            }
        }


        /// <summary>
        /// A range of cells. See TTokenCellAddress for single cell references.
        /// </summary>
        public class TTokenCellRange : TToken
        {
            TXlsCellRange FRange;
         
            /// <summary>
            /// Creates a new Cell range token.
            /// </summary>
            /// <param name="aRange">Range where this token will point to.</param>
            public TTokenCellRange(TXlsCellRange aRange): base (TTokenType.CellRange)
            {
                Range = aRange;
            }

            /// <summary>
            /// Cell address where this reference points to.
            /// </summary>
            public TXlsCellRange Range 
            {
                get {return FRange;} 
                set
                {
                    if (value == null) FlxMessages.ThrowException(FlxErr.ErrNullParameter, "Range");
                    FRange = value;
                }
            }
        }

        /// <summary>
        /// A function like "Sum" or "If"
        /// </summary>
        public class TTokenFunction : TToken
        {
            string FFunctionName;
            int FArgumentCount;
         
            /// <summary>
            /// Creates a new function token.
            /// </summary>
            /// <param name="aFunctionName">Name of the function.</param>
            /// <param name="aArgumentCount">Number of arguments for this function. Note that if the function has a fixed
            /// number of arguments, this parameter is ignored.</param>
            public TTokenFunction(string aFunctionName, int aArgumentCount): base (TTokenType.Function)
            {
                FunctionName = aFunctionName;
                ArgumentCount = aArgumentCount;
            }

            /// <summary>
            /// Name of the function represented by this token.
            /// </summary>
            public string FunctionName 
            {
                get {return FFunctionName;} 
                set
                {
                    if (value == null) FlxMessages.ThrowException(FlxErr.ErrNullParameter, "FunctionName");
                    if (TXlsFunction.GetData(value) == null) FlxMessages.ThrowException(FlxErr.ErrFunctionNotFound, value, string.Empty);
                    FFunctionName = value;
                }
            }

            /// <summary>
            /// Number of arguments for this function. Note that if the function has a fixed
            /// number of arguments, this parameter is ignored.
            /// </summary>
            public int ArgumentCount
            {
                get { return FArgumentCount; }
                set
                {
                    if (value < 0 || value > FlxConsts.Max_FormulaArguments2007) FlxMessages.ThrowException(FlxErr.ErrInvalidValue, "ArgumentCount", 0, FlxConsts.Max_FormulaArguments2007);
                    FArgumentCount = value;
                }
            }
        }

        /// <summary>
        /// An operator like "+" or "-"
        /// </summary>
        public class TTokenOperator : TToken
        {
            TOperator FOperator;
         
            /// <summary>
            /// Creates a new operator token.
            /// </summary>
            /// <param name="aOperator">Type of operator.</param>
            public TTokenOperator(TOperator aOperator): base (TTokenType.Operator)
            {
                Operator = aOperator;
            }

            /// <summary>
            /// Operator represented by this token.
            /// </summary>
            public TOperator Operator 
            {
                get {return FOperator;} 
                set
                {
                    if (!Enum.IsDefined(typeof(TOperator), value)) FlxMessages.ThrowException(FlxErr.ErrInvalidEnum);
                    FOperator = value;
                }
            }
        }

        /// <summary>
        /// An operator that operates in ranges, like union or intersection.
        /// </summary>
        public class TTokenRangeOp : TToken
        {
            TRangeOp FOperator;

            /// <summary>
            /// Creates a new range operator token.
            /// </summary>
            /// <param name="aOperator">Type of operator.</param>
            public TTokenRangeOp(TRangeOp aOperator)
                : base(TTokenType.RangeOp)
            {
                Operator = aOperator;
            }

            /// <summary>
            /// Operator represented by this token.
            /// </summary>
            public TRangeOp Operator
            {
                get { return FOperator; }
                set
                {
                    if (!Enum.IsDefined(typeof(TRangeOp), value)) FlxMessages.ThrowException(FlxErr.ErrInvalidEnum);
                    FOperator = value;
                }
            }
        }

        /// <summary>
        /// Whitespace like " ". This is not used in calculation, but it is use by Excel to show the formula as it was entered.
        /// </summary>
        public class TTokenWhitespace : TToken
        {
            int FWhitespaceCount;
         
            /// <summary>
            /// Creates a new whitespace token.
            /// </summary>
            /// <param name="aWhitespaceCount">Whitespace in the formula.</param>
            public TTokenWhitespace(int aWhitespaceCount): base (TTokenType.Whitespace)
            {
                WhitespaceCount = aWhitespaceCount;
            }

            /// <summary>
            /// Number of whitespace characters in this token.
            /// </summary>
            public int WhitespaceCount
            {
                get {return FWhitespaceCount;} 
                set
                {
                    if (value <= 0) FlxMessages.ThrowException(FlxErr.ErrInvalidValue, "WhitespaceCount", 1, Int32.MaxValue);
                    FWhitespaceCount = value;
                }
            }
        }

        /// <summary>
        /// A parenthesis surrounding the last token. Note that this token is not used in the formula (since RPN doesn't need parenthesis)
        /// but it is there so Excel can display them.
        /// </summary>
        public class TTokenParethesis : TToken
        {
            /// <summary>
            /// Creates a new parenthesis token.
            /// </summary>
            public TTokenParethesis(): base(TTokenType.Parethesis){}
        }

        /// <summary>
        /// A named range
        /// </summary>
        public class TTokenName : TToken
        {
            string FName;
         
            /// <summary>
            /// Creates a new name token.
            /// </summary>
            /// <param name="aName">Named range.</param>
            public TTokenName(string aName): base (TTokenType.Name)
            {
                Name = aName;
            }

            /// <summary>
            /// Named range represented by this token.
            /// </summary>
            public string Name 
            {
                get {return FName;} 
                set
                {
                    if (value == null) FlxMessages.ThrowException(FlxErr.ErrNullParameter, "Name");
                    bool IsInternal;
                    if (!TXlsNamedRange.IsValidRangeName(value, out IsInternal)) FlxMessages.ThrowException(FlxErr.ErrInvalidName, value);
                    FName = value;
                }
            }
        }        

        /// <summary>
        /// Constant data, like "Hello" and 1 in the formula: "=IF(A1 = 1,,"Hello")"
        /// </summary>
        public class TTokenData : TToken
        {
            object FData;
         
            /// <summary>
            /// Creates a new data token.
            /// </summary>
            /// <param name="aData">Data.</param>
            public TTokenData(object aData): base (TTokenType.Data)
            {
                Data = aData;
            }

            /// <summary>
            /// Data represented by this token.
            /// </summary>
            public object Data 
            {
                get {return FData;} 
                set
                {
                    if (value == null) FlxMessages.ThrowException(FlxErr.ErrNullParameter, "Data");
                    FData = value;
                }
            }
        }


        /// <summary>
        /// A missing argument for a function. For example the second parameter in the formula: "=IF(A1 = 1,,"Hello")"
        /// </summary>
        public class TTokenMissingArgument : TToken
        {
            /// <summary>
            /// Creates a new Missing argument token.
            /// </summary>
            public TTokenMissingArgument(): base(TTokenType.MissingArgument){}
        }

        /// <summary>
        /// This token is not supported by FlexCel.
        /// </summary>
        public class TTokenUnsupported : TToken
        {
            /// <summary>
            /// Creates a new unsupported token.
            /// </summary>
            public TTokenUnsupported(): base(TTokenType.Unsupported){}
        }


        internal static class TFormulaConverterInternalToToken
        {
            static TToken[] Convert(TParsedTokenList Tokens)
            {
                TToken[] Result = new TToken[Tokens.Count];
                Tokens.ResetPositionToStart();
                int i = 0; 
                while (!Tokens.Eof())
                {
                    TBaseParsedToken tk = Tokens.ForwardPop();
                    Result[i] = GetPublicToken(tk);
                    i++;
                }

                return Result;
            }

            private static TToken GetPublicToken(TBaseParsedToken tk)
            {
               /* ptg BaseToken = tk.GetBaseId;

                
                switch (BaseToken)
                {
                    case ptg.Exp: //must be array, can't be shared formula
                        return;
                    case ptg.Tbl: AddTable(R1C1, Token, CellList, ParsedStack);
                        StartFormula = String.Empty;
                        break;
                    case ptg.Add: return new TTokenOperator(TOperator.Add); 
                    case ptg.Sub: return new TTokenOperator(TOperator.Sub);
                    case ptg.Mul: return new TTokenOperator(TOperator.Mul);
                    case ptg.Div: return new TTokenOperator(TOperator.Div);
                    case ptg.Power: return new TTokenOperator(TOperator.Power);
                    case ptg.Concat: return new TTokenOperator(TOperator.Concat);
                    case ptg.LT: return new TTokenOperator(TOperator.LT);
                    case ptg.LE: return new TTokenOperator(TOperator.LE);
                    case ptg.EQ: return new TTokenOperator(TOperator.EQ);
                    case ptg.GE: return new TTokenOperator(TOperator.GE);
                    case ptg.GT: return new TTokenOperator(TOperator.GT);
                    case ptg.NE: return new TTokenOperator(TOperator.NE);
                    
                    case ptg.Isect: return new TTokenRangeOp(TRangeOp.Intersection);
                    case ptg.Union: return new TTokenRangeOp(TRangeOp.Union);
                    case ptg.Range: return new TTokenRangeOp(TRangeOp.Range);
                    
                    case ptg.Uplus: return new TTokenOperator(TOperator.UPlus);
                    case ptg.Uminus: return new TTokenOperator(TOperator.Neg);
                    case ptg.Percent: return new TTokenOperator(TOperator.Percent);
                    
                    case ptg.Paren: return new TTokenParethesis();
                    case ptg.MissArg: return new TTokenMissingArgument(); 
                    case ptg.Str: return new TTokenData(((TStrDataToken)tk).GetData());

                    case ptg.Attr: ProcessAttr(Token, ParsedStack); break;
                    case ptg.Sheet: return new TTokenUnsupported();
                    case ptg.EndSheet: new TTokenUnsupported();
                    case ptg.Err: return new TTokenData(((TErrDataToken)tk).GetData()); 
                    case ptg.Bool: return new TTokenData(((TBoolDataToken)tk).GetData());
                    case ptg.Int: return new TTokenData(((TIntDataToken)tk).GetData());
                    case ptg.Num: return new TTokenData(((TNumDataToken)tk).GetData());
                    case ptg.Array: ParsedStack.Push(ParsedStack.FmSpaces + GetArrayText(((TArrayDataToken)Token).GetData, MaxStringConstantLen)); break;

                    case ptg.Func:
                    case ptg.FuncVar:
                        int ArgCount;
                        bool IsAddin;
                        TBaseFunctionToken Function = tk as TBaseFunctionToken;

                        string FuncName = GetFuncName(Function, out ArgCount, false, out IsAddin);
                        if (IsAddin) FuncName = ConvertInternalFunctionName(Globals, ParsedStack.Pop());
                        return new TTokenFunction(FuncName, ArgCount);

                    case ptg.Name: ParsedStack.Push(ParsedStack.FmSpaces + GetName(((TNameToken)Token).NameIndex, -1, Globals, WritingXlsx)); break;

                    case ptg.RefN:
                    case ptg.Ref: ParsedStack.Push(ParsedStack.FmSpaces + GetRef(R1C1, (TRefToken)Token, CellRow, CellCol)); break;

                    case ptg.AreaN:
                    case ptg.Area: ParsedStack.Push(ParsedStack.FmSpaces + GetArea(R1C1, (TAreaToken)Token, CellRow, CellCol)); break;

                    case ptg.MemArea: break;
                    case ptg.MemErr: break;
                    case ptg.MemNoMem: break;
                    case ptg.MemFunc: break;
                    case ptg.RefErr: ParsedStack.Push(ParsedStack.FmSpaces + TFormulaMessages.ErrString(TFlxFormulaErrorValue.ErrRef)); break;
                    case ptg.AreaErr: ParsedStack.Push(ParsedStack.FmSpaces + TFormulaMessages.ErrString(TFlxFormulaErrorValue.ErrRef)); break;
                    case ptg.MemAreaN: break;
                    case ptg.MemNoMemN: break;
                    case ptg.NameX: ParsedStack.Push(ParsedStack.FmSpaces + GetNameX((TNameXToken)Token, Globals, WritingXlsx)); break;
                    case ptg.Ref3d: ParsedStack.Push(ParsedStack.FmSpaces + GetRef3D(R1C1, (TRef3dToken)Token, CellRow, CellCol, Globals, false, WritingXlsx)); break;
                    case ptg.Area3d: ParsedStack.Push(ParsedStack.FmSpaces + GetArea3D(R1C1, (TArea3dToken)Token, CellRow, CellCol, Globals, false, WritingXlsx)); break;
                    case ptg.Ref3dErr: ParsedStack.Push(ParsedStack.FmSpaces + GetRef3D(R1C1, (TRef3dToken)Token, -1, -1, Globals, true, WritingXlsx)); break;
                    case ptg.Area3dErr: ParsedStack.Push(ParsedStack.FmSpaces + GetArea3D(R1C1, (TArea3dToken)Token, CellRow, CellCol, Globals, true, WritingXlsx)); break;
                    default: XlsMessages.ThrowException(XlsErr.ErrBadToken, Token); break;
                }*/
                return null;
            }
        }


}
