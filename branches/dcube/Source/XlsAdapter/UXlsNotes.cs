using System;
using FlexCel.Core;
using System.Collections.Generic;

namespace FlexCel.XlsAdapter
{
	/// <summary>
	/// A cell Comment.
	/// </summary>
	internal class TNoteRecord: TBaseRowColRecord, IComparable
	{
        internal TEscherClientDataRecord Dwg;
		TAbsoluteAnchorRect NoteTextBox; //we will save it apart from the coords on the actual box, so we can move it with the comment.
        UInt16 OptionFlags;
        UInt16 ObjId;
        string Author;

		internal TNoteRecord(int aId, byte[] aData): base(aId, BitOps.GetWord(aData,2))
        {
            OptionFlags=BitConverter.ToUInt16(aData,4);
            ObjId=BitConverter.ToUInt16(aData,6);

            long aSize = 0;
            StrOps.GetSimpleString(true, aData, 8, false, 0, ref Author, ref aSize); 
        }

        /// <summary>
        /// CreateFromData
        /// </summary>
        internal TNoteRecord(int aRow, int aCol, TRichString aTxt, string aAuthor, TDrawing Drawing, TImageProperties Properties,
            ExcelFile xls, TSheet sSheet, bool ReadFromXlsx)
            : base((int)xlr.NOTE, aCol)
        {
            if ((aCol < 0) || (aCol > FlxConsts.Max_Columns)) XlsMessages.ThrowException(XlsErr.ErrXlsIndexOutBounds, aCol, "Column", 0, FlxConsts.Max_Columns);

            Dwg = Drawing.AddNewComment(xls, Properties, sSheet, ReadFromXlsx);
            TEscherImageAnchorRecord Anchor = GetImageRecord();
            if (Anchor != null) NoteTextBox = Anchor.SaveCommentCoords(sSheet, aRow, aCol);

            Col = aCol;
            OptionFlags = 0;   //option flags
            unchecked
            {
                ObjId = (UInt16)Dwg.ObjId;   //object id
            }

            Author = aAuthor;

            SetText(aTxt);
        }

        internal override void Destroy()
        {
            base.Destroy();
            if (Dwg!=null)
            {
                if(Dwg.Patriarch()==null) XlsMessages.ThrowException(XlsErr.ErrLoadingEscher);
                Dwg.Patriarch().ContainedRecords.Remove(Dwg.FindRoot());
            }

        }

		internal override bool AllowCopyOnOnlyFormula
		{
			get
			{
				return true;
			}
		}


        internal TRichString GetText()
        {
            if (Dwg == null) return new TRichString();
            else
            {
                TEscherRecord R = Dwg.FindRoot();
                if (R == null) return new TRichString();
                else
                {
                    R = R.FindRec<TEscherClientTextBoxRecord>();
                    if (R == null) return new TRichString(); else return ((TEscherClientTextBoxRecord)R).GetValue();
                }
            }
        }

        internal void SetText(TRichString value)
        {
            if (Dwg == null) return;
            else
            {
                TEscherRecord R = Dwg.FindRoot();
                if (R == null) return;
                else
                {
                    R = R.FindRec<TEscherClientTextBoxRecord>();
                    if (R == null) return; else ((TEscherClientTextBoxRecord)R).SetValue(value);
                }
            }
        }

        internal TEscherClientDataRecord GetDwg()
        {
            return Dwg;
        }

        internal TEscherClientTextBoxRecord GetClientTextBox()
        {
            if (Dwg == null) return null;
            else
            {
                TEscherRecord R = Dwg.FindRoot();
                if (R == null) return null;
                else
                {
                    R = R.FindRec<TEscherClientTextBoxRecord>();
                    return (TEscherClientTextBoxRecord)R;
                }
            }
        }

        internal TEscherOPTRecord GetOpt()
        {
            if (Dwg == null) return null;
            else
            {
                TEscherRecord R = Dwg.FindRoot();
                if (R == null) return null;
                else
                {
                    R = R.FindRec<TEscherOPTRecord>();
                    return (TEscherOPTRecord)R;
                }
            }
        }

        protected override TBaseRecord DoCopyTo(TSheetInfo SheetInfo)
        {
            TNoteRecord Result = (TNoteRecord)base.DoCopyTo(SheetInfo);
            Result.Dwg = Dwg;
            Result.OptionFlags = OptionFlags;
            Result.ObjId =  0; // we will fix this later.
            Result.Author = Author;
            return Result;
        }

        private void AdaptSize(TEscherImageAnchorRecord Anchor, int Row, TSheet dSheet)
        {
			Anchor.RestoreCommentCoords(NoteTextBox, dSheet, Row, Col);
        }

        internal override void ArrangeCopyRange(TXlsCellRange SourceRange, int Row, int RowOffset, int ColOffset, TSheetInfo SheetInfo)
        {
            FixDwg(RowOffset, ColOffset, SheetInfo);
            base.ArrangeCopyRange (SourceRange, Row, RowOffset, ColOffset, SheetInfo); //This must be last, so we dont modify row
        }

        private void FixDwg(int RowOffset, int ColOffset, TSheetInfo SheetInfo)
        {
            if (Dwg != null)
            {
                if (Dwg.Patriarch() == null) XlsMessages.ThrowException(XlsErr.ErrLoadingEscher);
                SheetInfo.IncCopiedGen();
                //We only copy DWG if we are copying rows/columns, when we copy sheets we dont have to.
                //now we have to copy dwg even in different sheets.
                Dwg = (TEscherClientDataRecord)Dwg.CopyDwg(RowOffset, ColOffset, SheetInfo);
                unchecked
                {
                    ObjId = (UInt16)Dwg.ObjId;   //object id
                }
            }
        }

        private TEscherImageAnchorRecord GetImageRecord()
        {
            if (Dwg == null) return null;
            else
            {
                TEscherRecord R = Dwg.FindRoot();
                if (R == null) return null;
                else
                {
                    R = R.FindRec<TEscherImageAnchorRecord>();
                    if (R == null) return null;
                    else
                    {
                        TEscherImageAnchorRecord Anchor = ((TEscherImageAnchorRecord)R);
                        return Anchor;
                    }
                }
            }
        }

        internal TClientAnchor GetAnchor(int Row, TSheet dSheet)
        {
            TEscherImageAnchorRecord Anchor = GetImageRecord();
            if (Anchor == null) return new TClientAnchor();
			AdaptSize(Anchor, Row, dSheet);
			return Anchor.GetAnchor();
        }

/*Now we keep range to insert on dwgofs. Inserting here could lead to inserting rows on the cells (and moving the red triangle down), but not on the note. Both must move together.
 *         internal override void ArrangeInsertRange(int Row, TXlsCellRange CellRange, int aRowCount, int aColCount, TSheetInfo SheetInfo)
        {
            if ((SheetInfo.InsSheet<0) || (SheetInfo.SourceFormulaSheet != SheetInfo.InsSheet)) return;  //this is also on the base.arrangeinsertrange, but we want to keep both inserts together. If we insert on the red triangle, we insert on the shape. If we don't we don't no matter that the shape is on range to insert.

            base.ArrangeInsertRange (Row, CellRange, aRowCount, aColCount, SheetInfo);
            if ((Dwg!=null) && (Dwg.FindRoot()!=null)) 
            {
                Dwg.FindRoot().ArrangeInsertRange(CellRange, aRowCount, aColCount, SheetInfo, true);
            }
        }
*/
        internal void FixDwgIds(TDrawing Drawing, int Row, TSheet sSheet, bool UseObjId, TCopiedGen CopiedGen)
        {
            if (UseObjId)
            {
                Dwg = Drawing.FindObjId(ObjId);
            }
            else
            {
                if (Dwg != null) Dwg = Dwg.CopiedTo(CopiedGen) as TEscherClientDataRecord;
                if (Dwg != null) ObjId = (UInt16)Dwg.ObjId;
            }
			TEscherImageAnchorRecord Anchor = GetImageRecord();
			if (Anchor != null) NoteTextBox = Anchor.SaveCommentCoords(sSheet, Row, Col);
       }

        internal void FixDwgOfs(int Row, TSheet dSheet)
        {
            TEscherImageAnchorRecord Anchor = GetImageRecord();
            if (Anchor != null)
            {
                AdaptSize(Anchor, Row, dSheet);
            }
        }

        internal override void SaveToStream(IDataStream Workbook, TSaveData SaveData, int Row)
        {
            base.SaveToStream(Workbook, SaveData, Row);
            Workbook.Write16(OptionFlags);
            Workbook.Write16(ObjId);

            byte[] bAuthor = Biff8Author();
            Workbook.Write(bAuthor, bAuthor.Length);
        }

        private byte[] Biff8Author()
        {
            TExcelString Xs = new TExcelString(TStrLenLength.is16bits, GetAuthor(), null, false);
            byte[] bAuthor = new byte[Xs.TotalSize() + 1];
            Xs.CopyToPtr(bAuthor, 0);
            return bAuthor;
        }

        public string GetAuthor()
        {
            if (Author == null || Author.Length == 0)
                return string.Empty;
            else
                return Author.Substring(0, Math.Min(Author.Length, FlxConsts.Max_CommentAuthor));
        }

        internal override int TotalSizeNoHeaders()
        {
            return base.TotalSizeNoHeaders() + 4 + Biff8Author().Length;
        }

		internal override void LoadIntoSheet(TSheet ws, int rRow, TBaseRecordLoader RecordLoader, ref TLoaderInfo Loader)
		{
			ws.Notes.AddRecord(this, rRow);
		}



        #region IComparable Members

        public int CompareTo(object obj)
        {
            TNoteRecord obj2= (TNoteRecord)obj;
            if (Col<obj2.Col) return -1; else if (Col<obj2.Col) return 1; else
                return 0;
        }

        #endregion

        internal void FillAuthors(TNoteAuthorList Authors)
        {
            Authors.AddAuthor(GetAuthor());
        }
    }

    /// <summary>
    /// A list of note records
    /// </summary>
    internal class TNoteRecordList: TBaseRowColRecordList<TNoteRecord>
    {
        #region Generics
        internal new TNoteRecord this[int index] 
        {
            get {return FList[index];} 
            set {SetThis(value, index);}
        }

        #endregion

        internal void FixDwgIds(TDrawing Drawing, int Row, TSheet sSheet, bool UseObjId, TCopiedGen CopiedGen)
        {
            for (int i=0; i< Count;i++)
                this[i].FixDwgIds(Drawing, Row, sSheet, UseObjId, CopiedGen);
        }

        internal void FixDwgOfs(int Row, TSheet dSheet)
        {
            for (int i = 0; i < Count; i++)
                this[i].FixDwgOfs(Row, dSheet);
        }

        internal void FillAuthors(TNoteAuthorList Authors)
        {
            for (int i = 0; i < Count; i++)
                this[i].FillAuthors(Authors);
        }
    }

    
    /// <summary>
    /// A list of TNoteRecordList lists.
    /// </summary>
    internal class TNoteList: TBaseRowColList<TNoteRecord, TNoteRecordList>
    {
        internal TFutureStorage FutureStorage;

        #region Generics
        internal void AddRecord(TNoteRecord aRecord, int aRow)
        {
            for (int i = Count; i <= aRow; i++)
                FList.Add(CreateRecord());
            this[aRow].Add(aRecord);
        }

        #endregion


        public TNoteAuthorList GetAuthors()
        {
            TNoteAuthorList Result = new TNoteAuthorList();
            for (int i = 0; i < Count; i++)
                this[i].FillAuthors(Result);

            return Result;
        }

        protected override TNoteRecordList CreateRecord()
        {
            return new TNoteRecordList();
        }

        internal void FixDwgIds(TDrawing Drawing, TSheet sSheet, bool UseObjId, TCopiedGen CopiedGen)
        {
            for (int i = 0; i < Count; i++)
                this[i].FixDwgIds(Drawing, i, sSheet, UseObjId, CopiedGen);
        }

        internal void FixDwgOfs(TSheet dSheet)
        {
            for (int i=0; i< Count;i++)
                this[i].FixDwgOfs(i, dSheet);
        }
   
        /* Now this is handled on DwgOfs
        internal override void ArrangeInsertRange(TXlsCellRange CellRange, int aRowCount, int aColCount, TSheetInfo SheetInfo)
        {
            base.ArrangeInsertRange (CellRange, aRowCount, aColCount, SheetInfo);
            //As it is now, base.ArrangeInsertRange will fix shapes when inserting columns, but not when inserting rows.

            if ((SheetInfo.InsSheet<0) || (SheetInfo.SourceFormulaSheet != SheetInfo.InsSheet)) return;

            if (aRowCount > 0)
            {
                int aCount= Count;
                for (int r= CellRange.Bottom; r< aCount;r++)
                {
                    TNoteRecordList RowNotes = this[r];
                    int bCount = RowNotes.Count;
                    for (int c = 0; c < bCount; c++)
                    {
                        RowNotes[c].ArrangeInsertRange(r, CellRange, aRowCount, 0, SheetInfo);
                    }
                }

            }
        }*/

		internal bool HasNotes(int Row)
		{
			return Row >= 0 && Row < Count && FList[Row] != null && this[Row].Count > 0;
		}

        internal void AddNewComment(int Row, int Col, TRichString Txt, string Author, TDrawing Drawing, TImageProperties Properties, 
            ExcelFile xls, TSheet sSheet, bool ReadFromXlsx)
        {
            TNoteRecord R= new TNoteRecord(Row, Col, Txt, Author, Drawing, Properties, xls, sSheet, ReadFromXlsx); //.CreateFromData
            AddRecord(R, Row);
        }

        internal void AddFutureStorage(TFutureStorageRecord R)
        {
            TFutureStorage.Add(ref FutureStorage, R);
        }
    }

    internal class TNoteAuthorList
    {
        Dictionary<string, int> ListByAuthors;
        List<string> ListById;

        internal TNoteAuthorList()
        {
            ListByAuthors = new Dictionary<string, int>();
            ListById = new List<string>();
        }

        internal void AddAuthor(string Author)
        {
            if (ListByAuthors.ContainsKey(Author)) return;
            ListByAuthors.Add(Author, ListByAuthors.Count);
            ListById.Add(Author);
        }

        internal List<string> AuthorsById()
        {
            return ListById;
        }

        public int Count
        {
            get
            {
                return ListById.Count;
            }
        }

        internal int GetId(string Author)
        {
            return ListByAuthors[Author];
        }
    }

}
