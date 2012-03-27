using System;
using System.Collections.Generic;

using FlexCel.Core;

namespace FlexCel.Report
{

	internal enum TBandType
	{
		Static = 1000,
		RowFull = TFlxInsertMode.ShiftRowDown,
		RowRange = TFlxInsertMode.ShiftRangeDown,
		ColFull = TFlxInsertMode.ShiftColRight,
		ColRange = TFlxInsertMode.ShiftRangeRight,
		FixedRow = TFlxInsertMode.NoneDown,
		FixedCol = TFlxInsertMode.NoneRight,
		Ignore = 1001
	}

	internal enum TBandMoveType
	{
		/// <summary>
		/// Only this band, not its children. Faster, but use it with care.
		/// </summary>
		Alone,
		/// <summary>
		/// Only direct children of the master band (1 level)
		/// </summary>
		DirectChildren,
		/// <summary>
		/// Children and children of the children...
		/// </summary>
		AllChildren
	}

	//1 based.
	internal interface ITopLeft
	{
		int Top{get;}
		int Left{get;}
	}

	internal class TBand: IComparable, ITopLeft, IDisposable
	{
		private TFlexCelDataSource FDataSource;

		internal int Refs;

		private TBand FMasterBand;
		private TBand FSearchBand;  //Used to search for different datasets.

		private TBandList FDetailBands;

		private TXlsCellRange FCellRange;

		private TBandType FBandType;

		private string FName;
		private string FDataSourceName;

		private bool FDeleteLastRow;

		internal TOneRowValue[] Rows;
		internal TBandImages Images;
		internal bool HasObjects;

		//these values will be reset each time we copy the range.
        internal int TmpExpandedRows;
        internal int TmpExpandedCols;
        internal int TmpPartialRows;
        internal int TmpPartialCols;

		internal TAddedRowColList ChildTmpExpandedRows;
		internal TAddedRowColList ChildTmpExpandedCols;

        internal bool Preprocessed;

        internal int FixedOfs;

		internal TBand(TFlexCelDataSource aDataSource, TBand aMasterBand, TBand aSearchBand, 
			TXlsCellRange aCellRange, string aName, TBandType aBandType, bool aDeleteLastRow, string aDataSourceName)
		{
			FDataSource=aDataSource;
			FMasterBand=aMasterBand;
			FSearchBand=aSearchBand;
			FDetailBands= new TBandList();
			if (aCellRange != null) FCellRange= (TXlsCellRange) aCellRange.Clone(); else FCellRange = null;
			FBandType=aBandType;
			FName=aName;
			FDataSourceName=aDataSourceName;
			FDeleteLastRow=aDeleteLastRow;
		}

		internal TBand(TFlexCelDataSource aDataSource, TBand aMasterBand, TXlsCellRange aCellRange, 
			string aName, TBandType aBandType, bool aDeleteLastRow, string aDataSourceName):
			this(aDataSource, aMasterBand, aMasterBand, aCellRange, aName, aBandType, aDeleteLastRow, aDataSourceName)
		{
			//On an normal band, Searchband=Masterband. The only case this doesn't happen is on Sheet bands, 
			//where we search on all other sheets but they are not in master/detail relationship.
		}

		internal TFlexCelDataSource DataSource
		{
			get
			{
				return FDataSource;
			}
			set
			{
				FDataSource=value;
			}
		}

		internal int RecordCount
		{
			get
			{
				if (DataSource==null) return 1; else return DataSource.RecordCount; //Bands with no datasource are repeated *1* time, not 0.
			}
		}

        internal int RealRecordCount
        {
            get
            {
                int Result = RecordCount - FixedOfs;
                if (Result < 1) Result = 1;
                return Result;
            }
        }


		internal void AddTmpExpandedRows(int n, int Left, int Right)
		{
			TmpExpandedRows += n;
            if (MasterBand != null)
            {
                if (MasterBand.ChildTmpExpandedRows == null) MasterBand.ChildTmpExpandedRows = new TAddedRowColList(MasterBand.CellRange.Left);
                MasterBand.ChildTmpExpandedRows.Add(Left, Right, n);
            }
        }

		internal void AddTmpExpandedCols(int n, int Top, int Bottom)
		{
			TmpExpandedCols += n;
            if (MasterBand != null)
            {
                if (MasterBand.ChildTmpExpandedCols == null) MasterBand.ChildTmpExpandedCols = new TAddedRowColList(MasterBand.CellRange.Top);
                MasterBand.ChildTmpExpandedCols.Add(Top, Bottom, n);
            }
        }
      
		internal void MoveFirst(TBandMoveType MoveType)
		{
			if (DataSource!=null) 
			{         
				DataSource.First();
			}
			if (MoveType==TBandMoveType.Alone) return;
			TBandMoveType NewMoveType=MoveType;
			if (MoveType==TBandMoveType.DirectChildren) NewMoveType=TBandMoveType.Alone;

            for (int i = 0; i < DetailBands.Count; i++)
            {
                TFlexCelDataSource fs = DetailBands[i].DataSource;
                if (fs != null) fs.MoveMasterRecord();
                DetailBands[i].MoveFirst(NewMoveType);
            }
		}

        internal void MoveNext(TBandMoveType MoveType)
        {
            if (DataSource == null) return;
            DataSource.Next();

            if (MoveType == TBandMoveType.Alone) return;
            TBandMoveType NewMoveType = MoveType;
            if (MoveType == TBandMoveType.DirectChildren) NewMoveType = TBandMoveType.Alone;

            for (int i = 0; i < DetailBands.Count; i++)
            {
                TFlexCelDataSource fs = DetailBands[i].DataSource;
                if (fs != null) fs.MoveMasterRecord();
                DetailBands[i].MoveFirst(NewMoveType);  //Detail bands return to first record when moving the parent.
            }
        }

		internal bool Eof()
		{
			if (DataSource==null) return true;
			else return DataSource.Eof();
		}

		internal void SetAllHasObjects(bool value)
		{
			HasObjects = value;
			if (DetailBands != null)
			{
				for (int i = DetailBands.Count-1; i >=0; i--) DetailBands[i].SetAllHasObjects(value);
			}
		}

		internal TXlsCellRange CellRange
		{
			get
			{
				return FCellRange;
			}
		}

		internal TBand MasterBand
		{
			get
			{
				return FMasterBand;
			}
			set
			{
				FMasterBand = value;
			}
		}

		internal TBand SearchBand
		{
			get
			{
				return FSearchBand;
			}
			set
			{
				FSearchBand = value;
			}
		}
        
		internal TBandList DetailBands
		{
			get
			{
				return FDetailBands;
			}

		}

		internal TBandType BandType
		{
			get
			{
				return FBandType;
			}
			set
			{
				FBandType=value;
			}
		}

		internal bool DeleteLastRow
		{
			get
			{
				return FDeleteLastRow;
			}
		}

		public string DataSourceName {get {return FDataSourceName;} set {FDataSourceName=value;}}
		public string Name {get {return FName;} set {FName=value;}}

		#region IComparable Members

		public int CompareTo(object obj)
		{
			TXlsCellRange c2= ((TBand)obj).CellRange;

			if (FCellRange.Left<c2.Left) return -1;
			else if (FCellRange.Left>c2.Left) return 1;

			if (FCellRange.Top<c2.Top) return -1;
			else if (FCellRange.Top>c2.Top) return 1;
            
			if (FCellRange.Right>c2.Right) return -1;
			else if (FCellRange.Right<c2.Right) return 1;
            
			if (FCellRange.Bottom>c2.Bottom) return -1;
			else if (FCellRange.Bottom<c2.Bottom) return 1;
            
			return 0;
		}

		#endregion

		#region ITopLeft Members

		public int Top
		{
			get
			{
				return CellRange.Top;
			}
		}

		public int Left
		{
			get
			{
				return CellRange.Left;
			}
		}

		#endregion

		#region IDisposable Members

		public void Dispose()
		{
			if (Refs > 0)
			{
				Refs --;
				return;
			}
			if (FDataSource != null) FDataSource.Dispose();
			if (Rows != null) 
				foreach (TOneRowValue r in Rows)
				{
					if (r != null) r.Dispose();
				}

			if (Images != null) Images.Dispose();

			for (int i = 0; i < DetailBands.Count; i++)
			{
				TBand db = DetailBands[i];
				db.Dispose();
			}
            GC.SuppressFinalize(this);
        }

		#endregion
	}


	//For ordering by column and row.
	internal class TRowColComparer: IComparer<ITopLeft>
	{
        #region IComparer<ITopLeft> Members

        public int Compare(ITopLeft x, ITopLeft y)
        {
            if (x.Top != y.Top) return x.Top.CompareTo(y.Top);
            else return x.Left.CompareTo(y.Left);
        }

        #endregion
    }

	internal class TBandList
	{
		protected List<TBand> FList;
		internal readonly static TRowColComparer RowColComparerMethod=new TRowColComparer(); //STATIC*

		internal TBandList()
		{
			FList = new List<TBand>();
		}

		#region Generics
		internal void Add (TBand a)
		{
			FList.Add(a);
		}

		internal TBand this[int index] 
		{
			get {return FList[index];} 
			set {FList[index]=value;}
		}

		#endregion

		internal int Count
		{
			get {return FList.Count;}
		}

		internal void Delete(int index)
		{
			FList.RemoveAt(index);
		}

		internal void Sort()
		{
			FList.Sort();
		}

	}

#if (FRAMEWORK20)
    internal class TBandImages: Dictionary<TRichString, TOneCellValue>, IDisposable
    {
#else
	internal class TBandImages: Hashtable, IDisposable
	{
		public bool TryGetValue(TRichString key, out TOneCellValue Result)
		{
			Result = (TOneCellValue)this[key];	
			return Result != null;
		}
#endif

		#region IDisposable Members

		public void Dispose()
		{
			foreach (object cv in this)
			{
				IDisposable disp = cv as IDisposable;
				if (disp != null) disp.Dispose();
			}
            GC.SuppressFinalize(this);
        }
		#endregion
	}

#if (FRAMEWORK20)
    internal class TBandSheetList: Dictionary<int, TBand>, IDisposable
    {
		public new TBand this[int index]
		{
			get
			{
				TBand Result;
				if (!TryGetValue(index, out Result)) return null;
				return Result;
			}
			set
			{
				base[index] = value;
			}
		}
#else
	internal class TBandSheetList: Hashtable, IDisposable
	{
		public TBand this[int index]
		{
			get
			{
				return (TBand)base[index];
			}
			set
			{
				base[index] = value;
			}
		}
#endif

		#region IDisposable Members

		public void Dispose()
		{
			foreach (object cv in this)
			{
				IDisposable disp = cv as IDisposable;
				if (disp != null) disp.Dispose();
			}
            GC.SuppressFinalize(this);
		}
		#endregion
	}


#if (FRAMEWORK20)
    internal class TBoolArray: Dictionary<int, bool>
    {
		public new bool this[int index]
		{
			get
			{
				bool Result;
				if (!TryGetValue(index, out Result)) return false;
				return Result;
			}
			set
			{
				base[index] = value;
			}
		}
#else
	internal class TBoolArray: Hashtable
	{
		public bool this[int index]
		{
			get
			{
				object Result = base[index];
				if (Result == null) return false;
				return (bool)Result;
			}
			set
			{
				base[index] = value;
			}
		}
#endif

	}
}
