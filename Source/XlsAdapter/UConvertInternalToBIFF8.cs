using System;
using FlexCel.Core;
using System.Diagnostics;
using System.IO;
using System.Collections.Generic;

namespace FlexCel.XlsAdapter
{
    internal enum TFormulaType
    {
        Normal,
        Name,
        DataValidation,
        CondFmt,
        Chart
    }



    /// <summary>
    /// A class for parsing a RPN formula for evaluation.
    /// Note: RPN is *always* BIFF8, so no need for more than 256 columns or 65536 rows.
    /// </summary>
    internal class TFormulaConvertInternalToBiff8
    {
        private const int ColMask = 255;
        private TFormulaConvertInternalToBiff8()
        {
        }


		/// <summary>
		/// FmlaWithoutArrayLen might be less than the total array len
		/// </summary>
        internal static byte[] GetTokenData(TNameRecordList Names, TParsedTokenList Tokens, TFormulaType FmlaType, out int FmlaLenWithoutArray)
        {
            // Remember to ungrow maxrow and maxcol from 2007 too biff8
            // Tokens that grow are ptgref, area, 3dref and 3darea. All of them (and nothing else) should ungrow.

            using (MemoryStream Data = new MemoryStream())
            {
                using (MemoryStream ArrayData = new MemoryStream())
                {
                    int[] TokenOffset = new int[Tokens.Count + 1];
                    TTokenOffset StreamPos = new TTokenOffset();
                    Int32List FuturePos = new Int32List();

                    Tokens.ResetPositionToStart();
                    while (!Tokens.Eof())
                    {
                        TBaseParsedToken Token = Tokens.ForwardPop();
						TokenOffset[Tokens.SavePosition()] = (int)Data.Position;
						ptg BaseToken = Token.GetBaseId;
						Add(Data, ArrayData, Token, BaseToken, Tokens, StreamPos, FmlaType, FuturePos);
                    } //while

					TokenOffset[Tokens.Count] = (int)Data.Position;
                    FixFutureFunctions(Names, FuturePos, Tokens, Data, TokenOffset, ref StreamPos);
                    FixGotoAndMemTokens(Tokens, Data, TokenOffset, StreamPos);
                    ArrayData.Position = 0;
					Data.Position = Data.Length; //FixGoto will change this.
					FmlaLenWithoutArray = (int)Data.Length;
                    ArrayData.WriteTo(Data);
                    return Data.ToArray();
                }
            }
        }

        private static void FixFutureFunctions(TNameRecordList Names, Int32List FuturePos, TParsedTokenList Tokens, MemoryStream Data, int[] TokenOffset, ref TTokenOffset StreamPos)
        {
            if (FuturePos.Count == 0) return;


            List<byte> NewData = new List<byte>(Data.ToArray()); //we need to insert in any order

            for (int i = 0; i < FuturePos.Count; i++)
            {
                Tokens.MoveTo(FuturePos[i]);
                TBaseFunctionToken FuncToken = (TBaseFunctionToken)Tokens.GetToken(FuturePos[i]);
                for (int k = 0; k < FuncToken.ArgumentCount; k++) //this doesn't include the name.
                {
                    Tokens.Flush();
                }

                int TokPos = Tokens.SavePosition();
                while (TokPos > 0)
                {
                    if (!(Tokens.GetToken(TokPos - 1) is TIgnoreInCalcToken) || Tokens.GetToken(TokPos - 1) is TAttrToken) break;
                    TokPos--;
                }

                int ofs = TokenOffset[TokPos];
                WriteFutureName(NewData, ofs, FindName(Names, FuncToken.GetFunctionData().FutureName));
                
                for (int k = TokPos; k < TokenOffset.Length; k++)
                {
                    TokenOffset[k] += 5;
                }

                TTokenOffset NewStreamPos = new TTokenOffset();
                foreach (int streamofs in StreamPos.Keys)
                {
                    int sofs = streamofs;
                    if (sofs >= ofs) sofs += 5;
                    NewStreamPos.Add(sofs, StreamPos[streamofs]);
                }

                StreamPos = NewStreamPos;
            }

            Data.SetLength(0);
            Data.Write(NewData.ToArray(), 0, NewData.Count);
        }

        private static int FindName(TNameRecordList Names, string Name)
        {
            if (Names == null) return 0; //Calculating length.

            for (int i = 0; i < Names.Count; i++)
            {
                if (String.Equals(Name, Names[i].Name, StringComparison.CurrentCultureIgnoreCase) && (Names[i].IsAddin)) return i + 1;
            }

            FlxMessages.ThrowException(FlxErr.ErrCantFindNamedRange, Name);
            return -1;
        }

        private static void WriteFutureName(List<byte> Data, int ofs, int NameIndex)
        {
            byte[] NameBits = new byte[5];
            NameBits[0] = (byte)ptg.Name;
            unchecked
            {
                NameBits[1] = (byte)(NameIndex & 0xFF);
                NameBits[2] = (byte)((NameIndex >> 8) & 0xFF);
            }

            Data.InsertRange(ofs, NameBits);
        }

        private static void FixGotoAndMemTokens(TParsedTokenList Tokens, MemoryStream Data, int[] TokenOffset, TTokenOffset StreamPos)
        {
            foreach (int streamofs in StreamPos.Keys)
            {
                TBaseParsedToken Token = Tokens.GetToken(StreamPos[streamofs]);
                Data.Position = streamofs;
                
                TAttrOptIfToken oiftk = Token as TAttrOptIfToken;
                if (oiftk != null)
                {
                    WriteWord(Data, TokenOffset[oiftk.PositionOfNextPtg] - (streamofs + 2));
                    continue;
                }

                TAttrOptChooseToken ctk = Token as TAttrOptChooseToken;
                if (ctk != null)
                {
                    for (int i = 0; i < ctk.PositionOfNextPtg.Length; i++)
                    {
                        WriteWord(Data, TokenOffset[ctk.PositionOfNextPtg[i]] - streamofs);
                    }

                    continue;
                }

                TAttrGotoToken gtk = Token as TAttrGotoToken;
                if (gtk != null)
                {
                    WriteWord(Data, TokenOffset[gtk.PositionOfNextPtg] - (streamofs + 2) - 1);
                    continue;
                }

				TSimpleMemToken memtk = Token as TSimpleMemToken;
				if (memtk != null)
				{
					WriteWord(Data, TokenOffset[memtk.PositionOfNextPtg] - (streamofs + 2));
					continue;
				}

            }
        }

        private static void Add(Stream Data, Stream ArrayData, TBaseParsedToken Token, ptg BaseToken, TParsedTokenList Tokens, TTokenOffset StreamPos, TFormulaType FmlaType, Int32List FuturePos)
        {
            if (IsError(BaseToken, FmlaType))
            {
                Data.WriteByte(0x1C); //PtgErr
                Data.WriteByte(0x2A);
                return;
            }


            Data.WriteByte((byte)Token.GetId);
            switch (BaseToken)
            {
                case ptg.Exp:
                    TExp_Token exp = (TExp_Token)Token;
                    Biff8Utils.CheckRow(exp.Row);
                    Biff8Utils.CheckCol(exp.Col);
                    WriteWord(Data, exp.Row);
                    WriteWord(Data, exp.Col);
                    break;
                case ptg.Tbl:
                    TTableObjToken tblo = Token as TTableObjToken;
                    if (tblo != null)
                    {
                        //no biff8 checks here. This numbers might be anything.
                        WriteWord(Data, tblo.Row);
                        WriteWord(Data, tblo.Col);
                    }
                    else
                    {
                        TTableToken tbl = (TTableToken)Token;
                        Biff8Utils.CheckRow(tbl.Row);
                        Biff8Utils.CheckCol(tbl.Col);
                        WriteWord(Data, tbl.Row);
                        WriteWord(Data, tbl.Col);
                    }
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
                case ptg.Isect:
                case ptg.Union:
                case ptg.Range:
                case ptg.Uplus:
                case ptg.Uminus:
                case ptg.Percent:
                case ptg.Paren:
                case ptg.MissArg:
                    break;


                case ptg.Str:
                    {
                        TStrDataToken tkd = Token as TStrDataToken;
                        string s = tkd.GetData();
                        if (s.Length > FlxConsts.Max_FormulaStringConstant)
                        {
                            FlxMessages.ThrowException(FlxErr.ErrStringConstantInFormulaTooLong, s, String.Empty);
                        }
                        TExcelString Xs = new TExcelString(TStrLenLength.is8bits, s, null, false);
                        byte[] b = new byte[Xs.TotalSize()];
                        Xs.CopyToPtr(b, 0);
                        Data.Write(b, 0, b.Length);
                        break;
                    }

                case ptg.Attr:
                    WriteAttr(Data, Token, StreamPos, Tokens.SavePosition());

                    break;

                /*
                case ptg.Sheet: //Can't happen, ConvertBiff8ToInternal skips them.
                case ptg.EndSheet: //Can't happen, ConvertBiff8ToInternal skips them.
                    break;*/

                case ptg.Err:
                    {
                        Data.WriteByte((byte)((TErrDataToken)Token).GetData());
                        break;
                    }
                case ptg.Bool:
                    {
                        if (((TBoolDataToken)Token).GetData()) Data.WriteByte(1); else Data.WriteByte(0);
                        break;
                    }
                case ptg.Int:
                    {
                        UInt16 a = (UInt16)((TIntDataToken)Token).GetData();
                        WriteWord(Data, a);
                        break;
                    }
                case ptg.Num:
                    {
                        double d = ((TNumDataToken)Token).GetData();
                        Data.Write(BitConverter.GetBytes(d), 0, 8);
                        break;
                    }
                
                
				case ptg.Array:
				{
					Data.Write(new byte[7], 0, 7);
					TArrayDataToken tk = (TArrayDataToken)Token;
					object[,] Arr = tk.GetData;

					int ColCount = Arr.GetLength(1) - 1;
					if (ColCount < 0 || ColCount > FlxConsts.Max_Columns97_2003) 
						FlxMessages.ThrowException(FlxErr.ErrInvalidCols, ColCount, FlxConsts.Max_Columns97_2003 + 1);
					ArrayData.WriteByte((byte)(ColCount));
				
					int RowCount = Arr.GetLength(0)-1;
					if (RowCount < 0 || RowCount > FlxConsts.Max_Rows) 
						FlxMessages.ThrowException(FlxErr.ErrInvalidRows, RowCount, FlxConsts.Max_Rows + 1);
					WriteWord(ArrayData, RowCount);

					for (int r = 0; r <= RowCount; r++)
					{
						for (int c = 0; c <= ColCount; c++)
						{
							WriteArrayObject(ArrayData, Arr[r,c]);
						}
					}

					break;
				}
                
                case ptg.Func:
                    {
                        TBaseFunctionToken ft = (TBaseFunctionToken)Token;
                        TCellFunctionData fd = ft.GetFunctionData();
                        WriteWord(Data, fd.Index);
                        break;
                    }
                case ptg.FuncVar:
                    {
                        TBaseFunctionToken ft = (TBaseFunctionToken)Token;
                        TCellFunctionData fd = ft.GetFunctionData();
                        if (!BackFromFutureToUserDef(Data, Tokens, FuturePos, ft, fd))
                        {
                            Data.WriteByte((byte)ft.ArgumentCount);
                            WriteWord(Data, fd.Index);
                        }
                        break;
                    }

                case ptg.Name:
                    WriteWord(Data, ((TNameToken)Token).NameIndex);
                    WriteWord(Data, 0);
                    break;

                case ptg.Ref:
                case ptg.RefN:
                case ptg.RefErr:
                    {
                        TRefToken reft = (TRefToken)Token;
                        WriteRef(Data, reft.CanHaveRelativeOffsets, reft.Row, reft.RowAbs, reft.Col, reft.ColAbs);
                        break;
                    }

                case ptg.Area:
                case ptg.AreaN:
                case ptg.AreaErr:
                    {
                        TAreaToken areat = (TAreaToken)Token;
                        WriteArea(Data, areat.CanHaveRelativeOffsets, areat.Row1, areat.RowAbs1, areat.Col1, areat.ColAbs1, areat.Row2, areat.RowAbs2, areat.Col2, areat.ColAbs2);
                        break;
                    }

				case ptg.MemArea:
				{
					WriteWord(Data, 0);
					WriteWord(Data, 0);
					StreamPos.Add((int)Data.Position, Tokens.SavePosition());
					WriteWord(Data, 0);

					TRefRange[] Range = ((TMemAreaToken)Token).Data;
					WriteWord(ArrayData, Range.Length);
					foreach (TRefRange OneRef in Range)
					{
						int r1 = Biff8Utils.CheckAndContractBiff8Row(OneRef.FirstRow);
						int r2 = Biff8Utils.CheckAndContractBiff8Row(OneRef.LastRow);
						int c1 = Biff8Utils.CheckAndContractBiff8Col(OneRef.FirstCol);
						int c2 = Biff8Utils.CheckAndContractBiff8Col(OneRef.LastCol);

						WriteWord(ArrayData, r1);
						WriteWord(ArrayData, r2);
						WriteWord(ArrayData, c1);
						WriteWord(ArrayData, c2);
					}

					break;
				}

                case ptg.MemErr:
                case ptg.MemNoMem:
					WriteWord(Data, 0);
					WriteWord(Data, 0);
					StreamPos.Add((int)Data.Position, Tokens.SavePosition());
					WriteWord(Data, 0);
					break;

				case ptg.MemFunc:
				case ptg.MemAreaN:
				case ptg.MemNoMemN:
					StreamPos.Add((int)Data.Position, Tokens.SavePosition());
					WriteWord(Data, 0);
                    break;
                
                case ptg.NameX:
                    TNameXToken NameX = (TNameXToken)Token;
                    WriteWord(Data, NameX.FExternSheet);
                    WriteWord(Data, NameX.NameIndex);
                    WriteWord(Data, 0);
                    break;

                case ptg.Ref3dErr:
                case ptg.Ref3d:
                    TRef3dToken reft3d = (TRef3dToken)Token;
                    WriteWord(Data, reft3d.FExternSheet);
                    WriteRef(Data, reft3d.CanHaveRelativeOffsets, reft3d.Row, reft3d.RowAbs, reft3d.Col, reft3d.ColAbs);
                    break;
                
                case ptg.Area3d:
                case ptg.Area3dErr:
                    TArea3dToken areat3d = (TArea3dToken)Token;
                    WriteWord(Data, areat3d.FExternSheet);
                    WriteArea(Data, areat3d.CanHaveRelativeOffsets, areat3d.Row1, areat3d.RowAbs1, areat3d.Col1, areat3d.ColAbs1, areat3d.Row2, areat3d.RowAbs2, areat3d.Col2, areat3d.ColAbs2);
                    break;

				default:
					XlsMessages.ThrowException(XlsErr.ErrInternal); //All tokens here should exist
					break;

            }  
        }

        private static bool BackFromFutureToUserDef(Stream Data, TParsedTokenList Tokens, Int32List FuturePos, TBaseFunctionToken ft, TCellFunctionData fd)
        {
            string FutureName = fd.FutureName;
            if (FutureName != null)
            {
                Data.WriteByte((byte)(ft.ArgumentCount + 1));
                WriteWord(Data, 0xFF); //User def...
                FuturePos.Add(Tokens.SavePosition());
                return true;
            }

            return false;
        }

        private static bool IsError(ptg BaseToken, TFormulaType FmlaType)
        {
            switch (FmlaType)
            {
                case TFormulaType.Normal:
                    return false;

                case TFormulaType.Name:
                    return (BaseToken == ptg.Tbl || BaseToken == ptg.Exp
                        || BaseToken == ptg.Ref || BaseToken == ptg.RefN || BaseToken == ptg.RefErr
                        || BaseToken == ptg.Area || BaseToken == ptg.AreaN || BaseToken == ptg.AreaErr);

                case TFormulaType.DataValidation:
                    return (BaseToken == ptg.Tbl || BaseToken == ptg.Exp
                        || BaseToken == ptg.Isect || BaseToken == ptg.Union || BaseToken == ptg.Array
                        || BaseToken == ptg.Ref3d || BaseToken == ptg.Area3d || BaseToken == ptg.Ref3dErr || BaseToken == ptg.Area3dErr
                        || BaseToken == ptg.NameX || BaseToken == ptg.MemArea || BaseToken == ptg.MemNoMem
                        );

                case TFormulaType.CondFmt:
                    return (BaseToken == ptg.Tbl || BaseToken == ptg.Exp
                        || BaseToken == ptg.Isect || BaseToken == ptg.Union || BaseToken == ptg.Array
                        || BaseToken == ptg.Ref3d || BaseToken == ptg.Ref3dErr
                        || BaseToken == ptg.NameX || BaseToken == ptg.MemArea || BaseToken == ptg.MemNoMem
                        );

                case TFormulaType.Chart:
                    return BaseToken != ptg.Paren && BaseToken != ptg.Union && BaseToken != ptg.Ref3d && BaseToken != ptg.Ref3dErr &&
                        BaseToken != ptg.Area3d && BaseToken != ptg.Area3dErr && BaseToken != ptg.NameX && BaseToken != ptg.MemFunc
                        ;
            }

            return false;
        }


        #region Write objects
        private static void WriteWord(Stream Data, int p)
        {
            unchecked
            {
                Data.WriteByte((byte)p);
                Data.WriteByte((byte)(p >> 8));
            }
        }

        private static void WriteArrayObject(Stream ArrayData, object o)
        {
            if (o == null)
            {
                ArrayData.Write(new byte[9], 0, 9);
                return;
            }

            if (o is double)
            {
                ArrayData.WriteByte(0x01);
                byte[] b = BitConverter.GetBytes((double)o);
                ArrayData.Write(b, 0, b.Length);
                return;
            }

            string s = o as string;
            if (s != null)
            {
                if (s.Length > 255) s = s.Substring(0, 255);
                TExcelString Xs = new TExcelString(TStrLenLength.is16bits, s, null, false);
                ArrayData.WriteByte(0x02);
                byte[] b = new byte[Xs.TotalSize()];
                Xs.CopyToPtr(b, 0);
                ArrayData.Write(b, 0, b.Length);
                return;
            }

            if (o is bool)
            {
                ArrayData.WriteByte(0x04);
                byte e = ((bool)o) ? (byte)1 : (byte)0;
                ArrayData.WriteByte(e);
                byte[] b = new byte[7];
                ArrayData.Write(b, 0, b.Length);
                return;
            }

            if (o is TFlxFormulaErrorValue)
            {
                ArrayData.WriteByte(0x10);
                ArrayData.WriteByte((byte)((TFlxFormulaErrorValue)o));
                byte[] b = new byte[7];
                ArrayData.Write(b, 0, b.Length);
                return;
            }

            XlsMessages.ThrowException(XlsErr.ErrBadToken, Convert.ToString(o.GetType()));
        }

        private static void WriteRow(Stream Data, bool CanHaveRelativeOffsets, int r, bool rabs)
        {
            if (CanHaveRelativeOffsets && !rabs)
                r = Biff8Utils.CheckAndContractRelativeBiff8Row(r);
            else
                r = Biff8Utils.CheckAndContractBiff8Row(r);
            WriteWord(Data, r);
        }

        private static void WriteCol(Stream Data, bool CanHaveRelativeOffsets, bool rabs, int c, bool cabs)
        {
            if (CanHaveRelativeOffsets && !cabs)
                c = Biff8Utils.CheckAndContractRelativeBiff8Col(c);
            else
                c = Biff8Utils.CheckAndContractBiff8Col(c);

            if (!rabs) c |= 0x8000;
            if (!cabs) c |= 0x4000;
            WriteWord(Data, c);
        }

        private static void WriteRef(Stream Data, bool CanHaveRelativeOffsets, int r, bool rabs, int c, bool cabs)
        {
            WriteRow(Data, CanHaveRelativeOffsets, r, rabs);
            WriteCol(Data, CanHaveRelativeOffsets, rabs, c, cabs);
        }

        private static void WriteArea(Stream Data, bool CanHaveRelativeOffsets, int r1, bool rabs1, int c1, bool cabs1, int r2, bool rabs2, int c2, bool cabs2)
        {
            WriteRow(Data, CanHaveRelativeOffsets, r1, rabs1);
			WriteRow(Data, CanHaveRelativeOffsets, r2, rabs2);
			WriteCol(Data, CanHaveRelativeOffsets, rabs1, c1, cabs1);
            WriteCol(Data, CanHaveRelativeOffsets, rabs2, c2, cabs2);
        }

        private static void WriteAttr(Stream Data, TBaseParsedToken Token, TTokenOffset StreamPos, int TokenPos)
        {
            if (Token is TAttrVolatileToken)
            {
                Data.WriteByte(0x01);
                WriteWord(Data, 0);
                return;
            }

            if (Token is TAttrSumToken)
            {
                Data.WriteByte(0x10);
                WriteWord(Data, 0);
                return;
            }

            if (Token is TAttrSpaceToken)
            {
                int Id = 0x40;
                TAttrSpaceToken sp = (TAttrSpaceToken)Token;
                if (sp.Volatile) Id |= 0x01;

                Data.WriteByte((byte)Id);
                Data.WriteByte((byte)sp.SpaceType);
                Data.WriteByte((byte)sp.SpaceCount);
                return;
            }

            TAttrOptIfToken oiftk = Token as TAttrOptIfToken;
            if (oiftk != null)
            {
                Data.WriteByte(0x02);
                StreamPos.Add((int)Data.Position, TokenPos);
                WriteWord(Data, 0);
                return;
            }

            TAttrOptChooseToken ctk = Token as TAttrOptChooseToken;
            if (ctk != null)
            {
                Data.WriteByte(0x04);
                WriteWord(Data, ctk.PositionOfNextPtg.Length - 1);
                StreamPos.Add((int)Data.Position, TokenPos);
                for (int i = 0; i < ctk.PositionOfNextPtg.Length; i++)
                {
                    WriteWord(Data, 0);
                }

                return;
            }

            TAttrGotoToken gtk = Token as TAttrGotoToken;
            if (gtk != null)
            {
                Data.WriteByte(0x08);
                StreamPos.Add((int)Data.Position, TokenPos);
                WriteWord(Data, 0);
                return;
            }

            XlsMessages.ThrowException(XlsErr.ErrInternal); //All tokens here should exist
                
        }


        #endregion

    }
}
