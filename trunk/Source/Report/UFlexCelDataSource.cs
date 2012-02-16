using System;
using System.Data;
using System.Text;
using System.Globalization;
using System.Diagnostics;

using FlexCel.Core;
using System.Collections.Generic;

namespace FlexCel.Report
{
    internal interface IDataTableFinder
    {
        TDataSourceInfo TryGetDataTable(string parent);
    }

    internal class TMasterDetailLinkList: List<TMasterDetailLink>
    {
        public void AddRelation(VirtualDataTableState aParentDataSource, TRelation Relation)
        {
            int[] Pc=Relation.ParentColumns;
            int[] Cc=Relation.ChildColumns;

            if (Pc == null && Cc == null) //Intrinsic relationship
            {
                Add(new TMasterDetailLink(aParentDataSource, 0, null));
                return;
            }

            if (Pc.Length!=Cc.Length) return; //Invalid relation.

            for (int i = 0; i < Pc.Length; i++)
            {
                string ChildName = Relation.ChildTable.GetColumnName(Cc[i]);
                if (!Find(ChildName))
                    Add(new TMasterDetailLink(aParentDataSource, Pc[i], ChildName));
            }
        }

        private bool Find(string CName)
        {
            for (int i = 0; i < Count; i++)
            {
                TMasterDetailLink key = base[i] as TMasterDetailLink;
                if (key == null) continue;
                if (String.Equals(key.ChildFieldName, CName, StringComparison.CurrentCultureIgnoreCase)) return true;
            }
            return false;
        }
    }

    /// <summary>
    /// Encapsulates a datasource to be used with FlexCelReport.
    /// It can have a position, master detail, etc.
    /// </summary>
    internal class TFlexCelDataSource: IDisposable
    {
        private VirtualDataTable Data;
        private VirtualDataTableState DataState;
        private TMasterDetailLink[] MasterDetailLinks;	
        internal TSplitLink SplitLink;

        private string FName;

        public TFlexCelDataSource(string dtName, VirtualDataTable aData, TRelationshipList ExtraRelations, TRelationshipList StaticRelations,
            TBand MasterBand, string Sort, IDataTableFinder TableFinder)
        {
            SplitLink = null;

            TBand band = MasterBand;
            TMasterDetailLinkList MasterDetailLinkList = new TMasterDetailLinkList();

            while (band != null)
            {
                if (band.DataSource != null)
                {
                    TRelation RelToMaster = band.DataSource.Data.GetRelationWith(aData);
                    if (RelToMaster != null)
                    {
                        MasterDetailLinkList.AddRelation(band.DataSource.DataState, RelToMaster);
                    }

                    //Create the splitlink.
                    TMasterSplitDataTableState SplitMaster = band.DataSource.DataState as TMasterSplitDataTableState;
                    if (SplitMaster != null && String.Equals(SplitMaster.DetailName, dtName, StringComparison.CurrentCultureIgnoreCase))
                    {
                        if (SplitLink != null)
                            FlxMessages.ThrowException(FlxErr.ErrSplitNeedsOnlyOneMaster, dtName, SplitMaster.TableName, SplitLink.ParentDataSource.TableName);
                        SplitLink = new TSplitLink(SplitMaster, SplitMaster.SplitCount);
                    }

                    AddExtraRelationships(TableFinder, aData, ExtraRelations, band, MasterDetailLinkList);
                    AddExtraRelationships(TableFinder, aData, StaticRelations, band, MasterDetailLinkList);
                }
                band = band.MasterBand;
            }

            MasterDetailLinks = MasterDetailLinkList.ToArray();

            Data = aData;
            DataState = aData.CreateState(Sort, MasterDetailLinks, SplitLink);
            DataState.FTableName = dtName;
            FName = dtName;
        }

        private static void AddExtraRelationships(IDataTableFinder TableFinder, VirtualDataTable aData, TRelationshipList Relations, TBand band, TMasterDetailLinkList MasterDetailLinkList)
        {
            foreach (TRelationship tr in Relations)
            {
                TDataSourceInfo dParent = TableFinder.TryGetDataTable(tr.ParentTable);
                if (dParent == null || dParent.Table == null)
                    FlxMessages.ThrowException(FlxErr.ErrInvalidManualRelationshipDatasetNull, tr.ParentTable);

                TDataSourceInfo dChild = TableFinder.TryGetDataTable(tr.ChildTable);
                if (dChild == null || dChild.Table == null)
                    FlxMessages.ThrowException(FlxErr.ErrInvalidManualRelationshipDatasetNull, tr.ChildTable);

                if (SameTable(dChild.Table, aData) && SameTable(dParent.Table, band.DataSource.Data))
                {
                    GetColIndexes(dParent.Table, tr.ParentColumns, true); //we do it here so it checks all columns exist in the parent. They might not exist in our table, and that is ok if our table is a distinct for example)
                    GetColIndexes(dChild.Table, tr.ChildColumns, true);
                    int[] ChildCols = GetColIndexes(aData, tr.ChildColumns, false);
                    int[] ParentCols = GetColIndexes(band.DataSource.Data, tr.ParentColumns, false);

                    if (ChildCols != null && ParentCols != null)
                    {
                        MasterDetailLinkList.AddRelation(band.DataSource.DataState,
                            new TRelation(band.DataSource.Data, aData, ParentCols,
                               ChildCols));
                    }
                }
            }
        }

        private static int[] GetColIndexes(VirtualDataTable Table, string[] ColNames, bool ThrowExceptions)
        {
            int[] ColIdx = new int[ColNames.Length];

            for (int i = 0; i < ColIdx.Length; i++)
            {
                int cs = Table.GetColumn(ColNames[i]);
                if (cs < 0)
                {
                    if (ThrowExceptions) FlxMessages.ThrowException(FlxErr.ErrColumNotFound, ColNames[i], Table.TableName);
                    else return null;
                }

                ColIdx[i] = cs;
            }

            return ColIdx;
        }

        private static bool SameTable(VirtualDataTable relatedData, VirtualDataTable data)
        {
            if (data == null) return relatedData == null;
            if (relatedData == null) return false;
            if (relatedData.TableName == data.TableName) return true;
            return SameTable(relatedData, data.CreatedBy);
        }
    

        public int GetColumn(string ColumnName)
        {
            int dc=Data.GetColumn(ColumnName);
            if (dc < 0) FlxMessages.ThrowException(FlxErr.ErrColumNotFound, ColumnName, Name);
            return dc;
        }

        public int GetColumnWithoutException(string ColumnName)
        {
            return Data.GetColumn(ColumnName);
        }

        public VirtualDataTable GetTable()
        {
            return Data;
        }


        public CultureInfo Locale
        {
            get
            {
                return Data.Locale;
            }
        }


        public TMasterSplitDataTable SplitMaster
        {
            get
            {
                return Data as TMasterSplitDataTable;
            }
        }
        
        internal int FilteredRowCount()
        {
            return DataState.FilteredRowCount(MasterDetailLinks);
        }

        public void MoveMasterRecord()
        {
            DataState.MoveMasterRecord(MasterDetailLinks, SplitLink);
        }


        public int RecordCount
        {
            get
            {
                return DataState.RowCount;
            }
        }

        public int Position
        {
            get
            {
                return DataState.Position;
            }
        }

        public void First()
        {
            DataState.DoMoveFirst();
        }

        public void Next()
        {
            DataState.DoMoveNext();
        }

        public bool Eof()
        {
            return DataState.Eof();
        }

       
        public object GetValue(int ColumnIndex)
        {
            if ((TPseudoColumn)ColumnIndex == TPseudoColumn.RowCount) return DataState.RowCount; //First so it does not check for eof.
            if (Eof()) FlxMessages.ThrowException(FlxErr.ErrReadAfterEOF);
            if (ColumnIndex < 0)
            {
                switch ((TPseudoColumn)ColumnIndex)
                {
                    case TPseudoColumn.RowPos: return DataState.Position;
                    default: FlxMessages.ThrowException(FlxErr.ErrInternal); break; //Should never come here.
                }
            }
            return DataState.GetValue(ColumnIndex);
        }

        public object GetValueForRow(int Row, int ColumnIndex)
        {
            if ((TPseudoColumn)ColumnIndex == TPseudoColumn.RowCount) return DataState.RowCount; //First so it does not check for eof.
            if (Row >= RecordCount) FlxMessages.ThrowException(FlxErr.ErrReadAfterEOF);
            if (Row < 0) FlxMessages.ThrowException(FlxErr.ErrReadBeforeBOF);

            if (ColumnIndex >= ColumnCount) FlxMessages.ThrowException(FlxErr.ErrInvalidColumn, ColumnIndex);
    
            if (ColumnIndex < 0)
            {
                switch ((TPseudoColumn)ColumnIndex)
                {
                    case TPseudoColumn.RowPos: return DataState.Position;
                    default: FlxMessages.ThrowException(FlxErr.ErrInvalidColumn, ColumnIndex); break;
                }
            }

            if (Row == DataState.Position) return DataState.GetValue(ColumnIndex); //avoid an exception if movetorecord is not supported and we really don't need to move to record.
            int SavePosition = DataState.Position;
            try
            {
                DataState.DoMoveToRecord(Row);
                return DataState.GetValue(ColumnIndex);
            }
            finally
            {
                DataState.DoMoveToRecord(SavePosition);
            }
        }

        public int ColumnCount
        {
            get
            {
                return Data.ColumnCount;
            }
        }

        public string ColumnCaption(int ColumnIndex)
        {
            return Data.GetColumnCaption(ColumnIndex);
        }

        public string Name
        {
            get
            {
                return FName;
            }
        }

        #region IDisposable Members

        public void Dispose()
        {
            if (DataState != null) DataState.Dispose();
            DataState = null;
            GC.SuppressFinalize(this);
        }

        #endregion

        internal VirtualDataTable GetDetail(string DataTableName, VirtualDataTable DataTable)
        {
            return Data.GetDetail(DataTableName, DataTable);
        }

        internal bool TryAggregate(TAggregateType AggregateType, int colIndex, out double? ResultValue)
        {
            return DataState.TryAggregate(AggregateType, colIndex, out ResultValue);
        }


        internal bool AllRecordsUsed()
        {
            return DataState.Eof();
        }
    }

    internal class TDataSourceInfo
    {
        #region Privates
        private VirtualDataTable FTable;
        private string FSort;
        private bool FDisposeVirtualTable;
        private bool FTempTable;
        IDataTableFinder TableFinder;
        #endregion

        internal TDataSourceInfo(string aName, VirtualDataTable aTable, string aRowFilter, string aSort, bool aDisposeVirtualTable, bool aTempTable, IDataTableFinder aTableFinder)
        {
            Filter(ref aTable, aRowFilter, ref aDisposeVirtualTable, aName);
            FTable = aTable;
            FSort = aSort;
            FDisposeVirtualTable = aDisposeVirtualTable;
            FTempTable = aTempTable;
            TableFinder = aTableFinder;
        }

        internal TFlexCelDataSource CreateDataSource(TBand MasterBand, TRelationshipList ExtraRelations, TRelationshipList StaticRelations)
        {
            VirtualDataTable LinkedTable = FindLinkedTable(MasterBand, Name, FTable);
            return new TFlexCelDataSource(LinkedTable.TableName, LinkedTable, ExtraRelations, StaticRelations, MasterBand, FSort, TableFinder);
        }

        internal static VirtualDataTable FindLinkedTable(TBand MasterBand, string Name, VirtualDataTable Table)
        {
            TBand band = MasterBand;
            while (band != null)
            {
                if (band.DataSource != null)
                {
                    VirtualDataTable vt = band.DataSource.GetDetail(Name, Table);
                    {
                        if (vt != null) return vt;
                    }
                }
                band = band.MasterBand;
            }
            return Table;
        }

        public string Name {get {return FTable.TableName;}}
        public VirtualDataTable Table { get { return FTable; } set { FTable = value; } }


        private static void Filter(ref VirtualDataTable aTable, string RowFilter, ref bool TableNeedsDispose, string aName)
        {
            RowFilter = RowFilter.Trim();
            if (string.IsNullOrEmpty(RowFilter))
            {
                if (aTable.TableName != aName)
                {
                    VirtualDataTable nTable = aTable.FilterData(aName, null);
                    if (aTable != null && TableNeedsDispose)
                    {
                        aTable.Dispose();
                    }
                    aTable = nTable;
                    TableNeedsDispose = true;
                }
                return;
            }

            int fPos = RowFilter.IndexOf(ReportTag.StrOpenParen);
            if (fPos>2)
            { 
                if (String.Equals(RowFilter.Substring(0, fPos).Trim(), ReportTag.ConfigTag(ConfigTagEnum.Distinct), StringComparison.InvariantCultureIgnoreCase))
                {
                    if (RowFilter[RowFilter.Length - 1] != ReportTag.StrCloseParen)
                        FlxMessages.ThrowException(FlxErr.ErrInvalidFilterParen, RowFilter);

                    VirtualDataTable bTable = aTable.GetDistinct(aName, GetColumnsForDistinct(aTable, RowFilter.Substring(fPos + 1, RowFilter.Length - fPos - 2)));
                    if (aTable != null && TableNeedsDispose)
                    {
                        aTable.Dispose();
                    }
                    aTable = bTable;

                    TableNeedsDispose = true;
                    return;
                }
            }
            
            VirtualDataTable cTable = aTable.FilterData(aName, RowFilter);
            if (aTable != null && TableNeedsDispose)
            {
                aTable.Dispose();
            }
            aTable = cTable;
            TableNeedsDispose = true;
        }

        private static int[] GetColumnsForDistinct(VirtualDataTable aTable, string SortString)
        {
            if (SortString.Trim().Length==0)
                FlxMessages.ThrowException(FlxErr.ErrInvalidDistinctParams);

            int start = 0;
            int count = 0;
            do 
            {
                start = SortString.IndexOf(ReportTag.ParamDelim, start);
                if (start > 0)
                {
                    count++;
                    start++;
                }
            }
            while (start>0);

            start = 0;
            int oldStart = 0;
            int[] Result = new int[count + 1];
            for (int i = 0; i< count; i++)
            {
                start = SortString.IndexOf(ReportTag.ParamDelim, oldStart);
                Result[i] = aTable.GetColumn(SortString.Substring(oldStart, start - oldStart).Trim());
                if (Result[i] < 0)
                    FlxMessages.ThrowException(FlxErr.ErrColumNotFound, SortString.Substring(oldStart, start - oldStart).Trim(), aTable.TableName);
                if (i < count) oldStart = start + 1;
            }

            Result[count] = aTable.GetColumn(SortString.Substring(oldStart).Trim());
            if (Result[count] < 0)
                FlxMessages.ThrowException(FlxErr.ErrColumNotFound, SortString.Substring(oldStart).Trim(), aTable.TableName);

            return Result;
        }



        /// <summary>
        /// Call this method ONLY if the collection owns the objects. 
        /// As this is not always the case, we cannot follow a more common Dispose pattern.
        /// </summary>
        public void DestroyTables()
        {
            if (FTable != null && FDisposeVirtualTable) FTable.Dispose();
            FTable = null;
        }

        public bool TempTable { get { return FTempTable; } }
        public bool VirtualTableNeedsDispose{get {return FDisposeVirtualTable;}}

    }

    internal class TDataSourceInfoList: IDisposable
    {
        private bool OwnsObjects;
        IDataTableFinder TableFinder;

        #region Privates
#if (FRAMEWORK20)
        Dictionary<string, TDataSourceInfo> FList = null;
#else
        Hashtable FList=null;
#endif
        #endregion

        internal TDataSourceInfoList(bool aOwnsObjects, IDataTableFinder aTableFinder)       
        {
            OwnsObjects = aOwnsObjects;
            FList = new Dictionary<string, TDataSourceInfo>(StringComparer.InvariantCultureIgnoreCase);
            TableFinder = aTableFinder;
        }

        internal void Clear()   
        {
            if (OwnsObjects)
            {
                foreach (TDataSourceInfo info in FList.Values)
                {
                    info.DestroyTables();
                }
            }
            FList.Clear();
        }

        internal void DeleteTempTables()
        {
            if (!OwnsObjects) return;
            List<string> keys = new List<string>(FList.Keys);
            foreach (string s in keys)
            {
                DestroyItem(s, true);
            }

        }

        private void DestroyItem(string key)
        {
            DestroyItem(key, false);
        }

        private void DestroyItem(string key, bool OnlyTemp)
        {
            if (!OwnsObjects) return;

            TDataSourceInfo di = this[key];
            if (di != null && (di.TempTable || (!OnlyTemp && di.VirtualTableNeedsDispose)))
            {
                di.DestroyTables();
                FList.Remove(key);
            }
        }

        internal void Add(string dtName, DataTable dt, bool DataTableNeedsDispose)
        {
            DestroyItem(dtName);
            FList[dtName]= new TDataSourceInfo(dtName, new TAdoDotNetDataTable(dtName, null, dt, DataTableNeedsDispose), "", "", true, DataTableNeedsDispose, TableFinder);
        }

        internal void Add(string dvName, DataView dv, bool DataViewNeedsDispose)
        {
            DestroyItem(dvName);
            FList[dvName]= new TDataSourceInfo(dvName, new TAdoDotNetDataTable(dvName, null, dv.Table, DataViewNeedsDispose), dv.RowFilter, dv.Sort, true, DataViewNeedsDispose, TableFinder);
        }

        internal void Add(string dvName, TDataSourceInfo di)
        {
            DestroyItem(dvName);
            FList[di.Name]=di;
        }

        internal void Add(string dtName, VirtualDataTable dt, bool VirtualTableNeedsDispose)
        {
            DestroyItem(dtName);
            FList[dtName]= new TDataSourceInfo(dtName, dt, "", "", VirtualTableNeedsDispose, VirtualTableNeedsDispose, TableFinder);
        }

        internal TDataSourceInfo this[string key]
        {
            get
            {
                TDataSourceInfo Result = null;
                if (FList.TryGetValue(key, out Result))
                    return Result;
                return null;
            }
        }

        internal int Count
        {
            get
            {
                return FList.Count;
            }
        }

        #region IDisposable Members

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (!OwnsObjects) return;
                foreach (TDataSourceInfo info in FList.Values)
                {
                    info.DestroyTables();
                }
            }
        }

        #endregion

        public Dictionary<string, TDataSourceInfo>.ValueCollection Values
        {
            get
            {
                return FList.Values;
            }
        }


    }

}
