using System;
using System.Data;
using System.Text;
using System.Globalization;
using System.Diagnostics;

using FlexCel.Core;
using System.Collections.Generic;
#if (FRAMEWORK30 && !COMPACTFRAMEWORK)
using System.Linq;
#endif

namespace FlexCel.Report
{
    #region ADO.NET
    /// <summary>
    /// An implementation of a VirtualDataTable using ADO.NET.
    /// </summary>
    internal class TAdoDotNetDataTable: VirtualDataTable
    {
        #region Privates
        internal DataTable Data;
        private bool DataNeedsDispose;
        private TLookupCacheList  LookupDataViews;
        internal DataTable OriginalTable; //Data relations are bound to this one.

        #endregion

        #region Constructors
        public TAdoDotNetDataTable(string aTableName, VirtualDataTable aCreatedBy, DataTable aData, bool aDataNeedsDispose): this(aTableName, aCreatedBy, aData, aDataNeedsDispose, aData)
        {
        }
        
        public TAdoDotNetDataTable(string aTableName, VirtualDataTable aCreatedBy, DataTable aData, bool aDataNeedsDispose, DataTable aOriginalTable): base(aTableName, aCreatedBy)
        {
            Data = aData;
            DataNeedsDispose = aDataNeedsDispose;
            LookupDataViews = new TLookupCacheList();
            OriginalTable = aOriginalTable;
        }


        #endregion

        #region Settings
        public override CultureInfo Locale
        {
            get
            {
                return Data.Locale;
            }
        }
        #endregion

        #region Create State Dataset
        public override VirtualDataTableState CreateState(string sort, TMasterDetailLink[] masterDetailLinks, TSplitLink splitLink)
        {
            return new TAdoDotNetDataTableState(this, Data, sort, masterDetailLinks, splitLink);
        }

        #endregion

        #region Columns
        public override int ColumnCount
        {
            get
            {
                return Data.Columns.Count;
            }
        }


        public override int GetColumn(string columnName)
        {
            DataColumn dc=Data.Columns[columnName];
            if (dc==null) return -1;
            return dc.Ordinal;
        }

        public override string GetColumnName(int columnIndex)
        {
            return Data.Columns[columnIndex].ColumnName;
        }

        public override string GetColumnCaption(int columnIndex)
        {
            return Data.Columns[columnIndex].Caption;
        }
        #endregion

        #region Filter
        public override VirtualDataTable FilterData(string newDataName, string rowFilter)
        {
            if (string.IsNullOrEmpty(rowFilter)) return new TAdoDotNetDataTable(newDataName, this, Data, false);

            DataTable FilteredData = null;

            try
            {
                DataView dv = new DataView(Data, rowFilter, String.Empty, DataViewRowState.CurrentRows);
                try
                {
                    FilteredData= Data.Clone();  //create a copy of the filtered data.
                    foreach (DataRowView rv in dv)
                        FilteredData.ImportRow(rv.Row);
                }
                finally
                {
                    TCompactFramework.DisposeDataView(dv);
                }
                return new TAdoDotNetDataTable(newDataName, this, FilteredData, true, OriginalTable);
            }
            catch
            {
                if (FilteredData != null) TCompactFramework.DisposeDataTable(FilteredData);
                throw;
            }
        }

        #endregion

        #region Distinct
        public override VirtualDataTable GetDistinct(string newDataName, int[] filterFields)
        {
            DataColumn[] SortFields = GetSortFields(Data, filterFields);

#if (FRAMEWORK20)
            if (Data.CaseSensitive)  //this does not seem to work on non case sensitives
            {
                DataTable NewData20 = null;
                try
                {
                    DataView dv = new DataView(Data);
                    try
                    {
                        string[] Columns = new string[SortFields.Length];
                        for (int i = 0; i < Columns.Length; i++)
                        {
                            Columns[i] = SortFields[i].ColumnName;
                        }
                        NewData20 = dv.ToTable(Data.TableName, true, Columns);
                    }
                    finally
                    {
                        TCompactFramework.DisposeDataView(dv);
                    }

                    return new TAdoDotNetDataTable(newDataName, this, NewData20, true, OriginalTable);
                }
                catch
                {
                    TCompactFramework.DisposeDataTable(NewData20);
                    throw;
                }
            }
#endif
            string Sort = GetSortString(SortFields);
            DataTable NewData = new DataTable();
            try
            {
                NewData.Locale = Data.Locale;
                NewData.CaseSensitive = Data.CaseSensitive;

                CopyColumns(SortFields, NewData);
                DataView dv = new DataView(Data, String.Empty, Sort, DataViewRowState.CurrentRows);
                try
                {
                    object[] LastRow = null;

                    NewData.BeginLoadData();
                    foreach (DataRowView dr in dv)
                    {
                        CopyRow(SortFields, dr, NewData, ref LastRow);
                    }
                    NewData.EndLoadData();
                }
                finally
                {
                    TCompactFramework.DisposeDataView(dv);
                }
                return new TAdoDotNetDataTable(newDataName, this, NewData, true, OriginalTable);
            }
            catch
            {
                TCompactFramework.DisposeDataTable(NewData);
                throw;
            }
        }

        private static DataColumn[] GetSortFields(DataTable Source, int[] SortColumns)
        {
            DataColumn[] Result = new DataColumn[SortColumns.Length];
            for (int i = 0; i< Result.Length; i++)
            {
                Result[i] = Source.Columns[SortColumns[i]];
            }
            return Result;
        }

#if (!FRAMEWORK20__)
        private static string GetSortString(DataColumn[] Fields)
        {
            if (Fields.Length <= 0) return String.Empty;
            StringBuilder Result = new StringBuilder(Fields[0].ColumnName);

            for (int i = 1; i < Fields.Length; i++)
            {
                Result.Append(",");
                Result.Append(Fields[i].ColumnName);
            }
            return Result.ToString();
        }

        private static void CopyColumns(DataColumn[] Columns, DataTable Dest)
        {
            DataColumn[] pk = new DataColumn[Columns.Length];
            for (int i = 0; i < Columns.Length; i++)
            {
                DataColumn dc = Columns[i];
                pk[i] = new DataColumn(dc.ColumnName, dc.DataType, dc.Expression, dc.ColumnMapping);
                Dest.Columns.Add(pk[i]);
            }
    
            Dest.PrimaryKey = pk;
        }

        private static bool ColumnsAreEqual(object a, object b, DataTable dt)
        {
            if ( a == DBNull.Value && b == DBNull.Value ) 
                return true;
            if ( a == DBNull.Value || b == DBNull.Value ) 
                return false;
            if ( a == null) 
            {
                if (b == null) return true;
                return false;
            }

            
            //we cannot use this directly compare since "a" != "a " and the datatable comparer thinks "a"=="a  "
            string sa = a as string;
            if (sa != null)
            {
                string sb = b as String;
                if (sb != null)
                {
                    sa = sa.Trim();
                    sb = sb.Trim();
                    CompareOptions Options = CompareOptions.IgnoreNonSpace | CompareOptions.IgnoreWidth | CompareOptions.IgnoreKanaType;
                    if (!dt.CaseSensitive) Options |=  CompareOptions.IgnoreCase;
                    return  dt.Locale.CompareInfo.Compare(sa, sb, Options) == 0;
                }
            }

            return a.Equals(b);  
        }
    
        private static bool RowsAreEqual (object[] Row1, object[] Row2, DataTable dt)
        {
            Debug.Assert(Row1.Length == Row2.Length, "Row lengths should be the same");
            for (int i = 0; i < Row1.Length; i++)
            {
                if (!ColumnsAreEqual(Row1[i], Row2[i], dt))
                    return false;
            }

            return true;
        }

        private static void CopyRow(DataColumn[] Col, DataRowView dr, DataTable NewData, ref object[] LastRow)
        {
            object[] NewRow = new object[Col.Length];
            for (int i = 0; i < NewRow.Length; i++)
            {
                NewRow[i] = dr[Col[i].ColumnName];  //Col[i].Ordinal will not work on CF!
            }

            if (LastRow != null && RowsAreEqual(LastRow, NewRow, NewData)) return;

            NewData.LoadDataRow(NewRow, true);
            LastRow = NewRow;
        }
#endif
        #endregion

        #region Master Detail
        public override TRelation GetRelationWith(VirtualDataTable aDetail)
        {
            TAdoDotNetDataTable AdoDetail = aDetail as TAdoDotNetDataTable; //DataRelationships only happen between DataTables.
            if (AdoDetail != null && OriginalTable.ChildRelations != null)
            {
                foreach (DataRelation dr in OriginalTable.ChildRelations)
                {
                    if (dr.ChildTable == AdoDetail.OriginalTable)
                    {
                        return new TRelation(this, AdoDetail, GetColumnOrdinals(dr.ParentColumns, this), GetColumnOrdinals(dr.ChildColumns, AdoDetail));
                    }
                }
            }

            return null;
        }

        private static int[] GetColumnOrdinals(DataColumn[] dc, TAdoDotNetDataTable TargetTable)
        {
            int[] Result = new int[dc.Length];
            for (int i = 0; i < Result.Length; i++)
            {
                Result[i] = TargetTable.GetColumn(dc[i].ColumnName);
            }
            return Result;
        }


        #endregion

        #region Lookup
        public override object Lookup(int column, string keyNames, object[] keyValues)
        {
            lock (LookupDataViews) //Make sure no 2 threads try to use LookupDataViews at the same time.
            {
                TLookupCache LookupDV;
                if (!LookupDataViews.TryGetValue(keyNames, out LookupDV))
                {
                    if (LookupDataViews.Count > 10) LookupDataViews.Clear(); //Avoid too many indexes.
                    LookupDV = new TLookupCache(new DataView(Data, String.Empty, keyNames, DataViewRowState.CurrentRows));
                    LookupDataViews.Add(keyNames, LookupDV);
                }
            
                if (LookupDV.EqualKeys(keyValues))
                {
                    return LookupDV.LastRow.Row[column];
                }

                int row = LookupDV.dv.Find(keyValues);
                if (row < 0) return null;
            
                LookupDV.LastRow = LookupDV.dv[row];
                LookupDV.LastKeyValues = keyValues;
                return LookupDV.LastRow.Row[column];
            }
        }
        #endregion

        #region IDisposable Members
        protected override void Dispose(bool disposing)
        {
            try
            {
                //only managed resources
                if (disposing)
                {
                    if (DataNeedsDispose && Data != null) TCompactFramework.DisposeDataTable(Data);
                    Data = null;

                    foreach (TLookupCache dv in LookupDataViews.Values)
                    {
                        dv.Dispose();
                    }
                    LookupDataViews = null;
                }
            }
            finally
            {
                //last call.
                base.Dispose (disposing);
            }
        }

        #endregion
    }

    #region Lookup Cache
    /// <summary>
    /// This class is used to avoid repeated searches when many lookups are placed on the same row.
    /// </summary>
    internal class TLookupCache: IDisposable
    {
        internal DataView dv;
        internal DataRowView LastRow;
        internal object[] LastKeyValues;

        internal TLookupCache(DataView aDataView)
        {
            dv = aDataView;
        }

        internal bool EqualKeys(object[] keyValues)
        {
            if (LastKeyValues == null) return false; //cache might have not been initialized.
            if (keyValues == null || LastKeyValues.Length != keyValues.Length) return false;
            for (int i = 0; i < LastKeyValues.Length; i++)
            {
                if (LastKeyValues[i] == null)
                {
                    if (keyValues[i] == null) continue;
                    return false;
                }
                if (!(LastKeyValues[i].Equals(keyValues[i]))) return false;
            }

            return true;
        }

        #region IDisposable Members

        public void Dispose()
        {
            if (dv != null) dv.Dispose();
            dv = null;
            GC.SuppressFinalize(this);
        }

        #endregion
    }

#if (FRAMEWORK20)
    internal sealed class TLookupCacheList : Dictionary<string, TLookupCache>
    {
        public TLookupCacheList(): base(StringComparer.InvariantCultureIgnoreCase)
        {
        }
    }
#else
    internal class TLookupCacheList: Hashtable
    {
        public TLookupCacheList(): base(FormatComparer.HashProvider, FormatComparer.HashComparer)
        {
        }

        public bool TryGetValue(string key, out TLookupCache LookupDV)
        {
            LookupDV = (TLookupCache)this[key];	
            return LookupDV != null;
        }


    }
#endif

    #endregion

    internal class TAdoDotNetDataTableState: VirtualDataTableState
    {
        #region Privates
        private DataView SourceForFilteredData; //The dataview with the original data.
        private DataTable FilteredData; //We need it to isolate us from the original table, and use this datatable as master for others.
        private DataView SortedFilteredData; //Again a dataview, used to sort FilteredData.

        private bool FilteredDataNeedsDispose; 
        #endregion
        
        #region Constructors
        public TAdoDotNetDataTableState(VirtualDataTable aTableData, DataTable Data, string sort, TMasterDetailLink[] masterDetailLinks, TSplitLink splitLink): base(aTableData)
        {
            if (masterDetailLinks.Length > 0 || splitLink != null)
            {
                SourceForFilteredData = new DataView(Data, String.Empty, GetSortColumns(masterDetailLinks), DataViewRowState.CurrentRows);
                FilteredData = Data.Clone();  //we will have to fill this table each time.
                FilteredDataNeedsDispose = true;
            }
            else
            {
                FilteredData = Data;  //The table is the same as the original.
            }
            
            SortedFilteredData = new DataView(FilteredData);
            SortedFilteredData.Sort = sort;

        }
        #endregion

        #region Rows and data
        public override int RowCount
        {
            get
            {
                return FilteredData.Rows.Count;
            }
        }

        public override object GetValue(int column)
        {
            return SortedFilteredData[Position].Row[column]; // SortedFilteredData[Position][ColumnIndex] is not supported on CF
        }
        #endregion

        #region Move
        public override void MoveFirst()
        {
        }

        public override void MoveNext()
        {
        }

        public override void MoveToRecord(int aPosition)
        {
        }

        #endregion

        #region Relationships
        public override void MoveMasterRecord(TMasterDetailLink[] masterDetailLinks, TSplitLink splitLink)
        {
            if (masterDetailLinks.Length>0 || splitLink != null) //This dataset has a master, so rows need to be filtered.
                FillChildDataSet(masterDetailLinks, splitLink);
        }

        public override int FilteredRowCount(TMasterDetailLink[] masterDetailLinks)
        {
            if (masterDetailLinks.Length <= 0) return SourceForFilteredData.Count;
            object[] KeyValues = GetKeyValues(masterDetailLinks);
            return SourceForFilteredData.FindRows(KeyValues).Length;
        }


        #endregion

        #region Private implementation
        
        private static string GetSortColumns(TMasterDetailLink[] MasterDetailLinks)
        {
            StringBuilder Result= new StringBuilder();
            foreach (TMasterDetailLink key in MasterDetailLinks)
            {
                if (Result.Length>0) Result.Append(",");
                Result.Append(key.ChildFieldName);
            }
            return Result.ToString();
        }

        private static object[] GetKeyValues(TMasterDetailLink[] MasterDetailLinks)
        {
            object[] Result = new object[MasterDetailLinks.Length];
            for (int i = 0; i < MasterDetailLinks.Length; i++)
            {
                if (MasterDetailLinks[i].ParentDataSource.RowCount > 0)
                    Result[i] = MasterDetailLinks[i].ParentDataSource.GetValue(MasterDetailLinks[i].ParentField);
            }
            return Result;
        }

        private void FillChildDataSet(TMasterDetailLink[] MasterDetailLinks, TSplitLink SplitLink)
        {
            Debug.Assert(MasterDetailLinks.Length > 0 || SplitLink != null, "This method can only be called on master detail.");
            Debug.Assert(FilteredDataNeedsDispose == true, "We need to own FilteredData to modify it.");
            FilteredData.Clear();

            DataRowView[] drv = null;
            if (MasterDetailLinks.Length > 0) 
            {
                object[] KeyValues = GetKeyValues(MasterDetailLinks);
                drv = SourceForFilteredData.FindRows(KeyValues);
            }

            if (SplitLink != null)
            {
                int ParentPos = SplitLink.ParentDataSource.Position * SplitLink.SplitCount;
                for (int row =  ParentPos; row < ParentPos + SplitLink.SplitCount; row++)
                {
                    if (drv != null)
                    {
                        if (row >= drv.Length) break;
                    }
                    else
                    {
                        if (row >= SourceForFilteredData.Count) break;
                    }
                    ImportRows(FilteredData, drv, SourceForFilteredData, row);
                }
            }
            
            else // If splitlink = null, MasterdetailLinks.Length > 0, so drv != null.
            {
                foreach (DataRowView r in drv)
                    FilteredData.ImportRow(r.Row);
            }
        }

        private static void ImportRows(DataTable aFilteredData, DataRowView[] drv, DataView Data, int row)
        {
            if (drv != null)
            {
                aFilteredData.ImportRow(drv[row].Row);
            }
            else
            {
                aFilteredData.ImportRow(Data[row].Row);
            }
        }

        #endregion

        #region IDisposable Members

        protected override void Dispose(bool disposing)
        {
            try
            {
                if (disposing)
                {
                    if (SourceForFilteredData != null) TCompactFramework.DisposeDataView(SourceForFilteredData);
                    SourceForFilteredData = null;
                    if (SortedFilteredData != null) TCompactFramework.DisposeDataView(SortedFilteredData);
                    SortedFilteredData = null;
                    if (FilteredData != null && FilteredDataNeedsDispose) TCompactFramework.DisposeDataTable(FilteredData);
                    FilteredData = null;
                }
            }
            finally
            {
                //last call.
                base.Dispose (disposing);
            }
        }

        #endregion
    }
    #endregion

}
