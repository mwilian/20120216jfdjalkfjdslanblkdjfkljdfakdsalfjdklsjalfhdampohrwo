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
    #region MasterSplitDataTable
    /// <summary>
    /// An implementation of a VirtualDataTable that is used as a master table on a splitted dataset.
    /// </summary>
    internal class TMasterSplitDataTable: VirtualDataTable
    {
        #region Privates
        internal string DetailName;
        internal TFlexCelDataSource DetailData; //Will be filled outside.
        internal int SplitCount;
        #endregion

        #region Constructors
        public TMasterSplitDataTable(string aTableName, VirtualDataTable aCreatedBy, string aDetailName, int aSplitCount) : base(aTableName, aCreatedBy)
        {
            DetailName = aDetailName;
            SplitCount = aSplitCount;
        }
        #endregion

        #region Settings
        public override CultureInfo Locale
        {
            get
            {
                return DetailData.Locale;
            }
        }

        #endregion

        #region Create State Dataset
        public override VirtualDataTableState CreateState(string sort, TMasterDetailLink[] masterDetailLinks, TSplitLink splitLink)
        {
            return new TMasterSplitDataTableState(this);
        }

        #endregion

        #region Columns
        public override int ColumnCount
        {
            get
            {
                return 1;
            }
        }

        public override int GetColumn(string columnName)
        {
            if (String.Equals(columnName, ReportTag.StrRowPosColumn, StringComparison.CurrentCultureIgnoreCase)) return 0;
            return -1;
        }

        public override string GetColumnName(int columnIndex)
        {
            return ReportTag.StrRowPosColumn;
        }

        public override string GetColumnCaption(int columnIndex)
        {
            return ReportTag.StrRowPosColumn;
        }

        #endregion

        //We do not implement filter or distinct.
        //We do not implement lookup.
        
    }
    
    internal class TMasterSplitDataTableState: VirtualDataTableState
    {
        #region Privates
        private int LastPosition;
        private int LastRecordCount;
        #endregion

        #region Constructors
        public TMasterSplitDataTableState(VirtualDataTable aTableData) : base(aTableData)
        {
            LastPosition = -1;
        }

        #endregion

        #region Rows and Data
        public override int RowCount
        {
            get
            {
                if (LastPosition != Position)
                {
                    int SplitDetailRecordCount = DetailData.FilteredRowCount();
                    int SplitCount = DetailData.SplitLink.SplitCount;  //should never be null.

                    if (SplitDetailRecordCount <= 0 || SplitCount <= 0) 
                    {
                        LastRecordCount = 0;
                    }
                    else
                    {
                        LastRecordCount = (SplitDetailRecordCount - 1) / SplitCount + 1;
                    }
                    LastPosition = Position;
                }
                return LastRecordCount;
            }
        }

        public override object GetValue(int column)
        {
            return Position + 1;
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
            LastPosition = -1;
        }

        public override int FilteredRowCount(TMasterDetailLink[] masterDetailLinks)
        {
            return RowCount; //this dataset should be never filtered.
        }
        #endregion

        #region Private Implementation
        internal TFlexCelDataSource DetailData
        {
            get
            {
                return ((TMasterSplitDataTable)TableData).DetailData;
            }
        }

        internal string DetailName
        {
            get
            {
                return ((TMasterSplitDataTable)TableData).DetailName;
            }
        }

        internal int SplitCount
        {
            get
            {
                return ((TMasterSplitDataTable)TableData).SplitCount;
            }
        }

        #endregion

        #region IDisposable Members
        #endregion

    }
    #endregion

    #region TopDataTable
    /// <summary>
    /// An implementation of a VirtualDataTable that is used to limit the records to the "top n".
    /// </summary>
    internal class TTopDataTable: VirtualDataTable
    {
        #region Privates
        internal VirtualDataTable ActualData;
        internal int TopCount;
        #endregion

        #region Constructors
        public TTopDataTable(string aTableName, VirtualDataTable aCreatedBy, VirtualDataTable aActualData, int aTopCount)
            : base(aTableName, aCreatedBy)
        {
            ActualData = aActualData;
            TopCount = aTopCount;
        }
        #endregion

        #region Settings
        public override CultureInfo Locale
        {
            get
            {
                return ActualData.Locale;
            }
        }

        #endregion

        #region Create State Dataset
        public override VirtualDataTableState CreateState(string sort, TMasterDetailLink[] masterDetailLinks, TSplitLink splitLink)
        {
            return new TTopDataTableState(this, TopCount, ActualData.CreateState(sort, masterDetailLinks, splitLink));
        }

        #endregion

        #region Columns
        public override int ColumnCount
        {
            get
            {
                return ActualData.ColumnCount;
            }
        }

        public override int GetColumn(string columnName)
        {
            return ActualData.GetColumn(columnName);
        }

        public override string GetColumnName(int columnIndex)
        {
            return ActualData.GetColumnName(columnIndex);
        }

        public override string GetColumnCaption(int columnIndex)
        {
            return ActualData.GetColumnCaption(columnIndex);
        }

        #endregion

        public override VirtualDataTable FilterData(string newDataName, string rowFilter)
        {
            return new TTopDataTable(newDataName, this, ActualData.FilterData(newDataName, rowFilter), TopCount);
        }

        public override VirtualDataTable GetDistinct(string newDataName, int[] filterFields)
        {
            return new TTopDataTable(newDataName, this, ActualData.GetDistinct (newDataName, filterFields), TopCount);
        }

        public override object Lookup(int column, string keyNames, object[] keyValues)
        {
            return ActualData.Lookup (column, keyNames, keyValues);
        }	
    }
    
    internal class TTopDataTableState: VirtualDataTableState
    {
        #region Privates
        private VirtualDataTableState ActualDataState;
        private int TopCount;
        #endregion

        #region Constructors
        public TTopDataTableState(VirtualDataTable aTableData, int aTopCount, VirtualDataTableState aActualDataState) : base(aTableData)
        {
            ActualDataState = aActualDataState;
            TopCount = aTopCount;
        }

        #endregion

        #region Rows and Data
        public override int RowCount
        {
            get
            {
                if (ActualDataState.RowCount < TopCount) return ActualDataState.RowCount; else return TopCount;
            }
        }

        public override object GetValue(int column)
        {
            return ActualDataState.GetValue(column);
        }
        #endregion

        #region Move
        public override void MoveFirst()
        {
            ActualDataState.DoMoveFirst();
        }

        public override void MoveNext()
        {
            ActualDataState.DoMoveNext();
        }

        public override void MoveToRecord(int aPosition)
        {
            ActualDataState.DoMoveToRecord(aPosition);
        }
        #endregion


        #region Relationships
        public override void MoveMasterRecord(TMasterDetailLink[] masterDetailLinks, TSplitLink splitLink)
        {
            ActualDataState.MoveMasterRecord(masterDetailLinks, splitLink);
        }

        public override int FilteredRowCount(TMasterDetailLink[] masterDetailLinks)
        {
            int Result = ActualDataState.FilteredRowCount(masterDetailLinks);
            if (Result < TopCount) return Result; else return TopCount;
        }
        #endregion

        #region IDisposable Members
        #endregion

    }
    #endregion

    #region NRowsDataTable
    /// <summary>
    /// An implementation of a VirtualDataTable that is used return a table with n records. It has no columns.
    /// </summary>
    internal class TNRowsDataTable : VirtualDataTable
    {
        #region Privates
        internal int RecordCount;
        #endregion

        #region Constructors
        public TNRowsDataTable(string aTableName, VirtualDataTable aCreatedBy, int aRecordCount)
            : base(aTableName, aCreatedBy)
        {
            RecordCount = aRecordCount;
        }
        #endregion

        #region Settings
        public override CultureInfo Locale
        {
            get
            {
                return CultureInfo.CurrentCulture;
            }
        }

        #endregion

        #region Create State Dataset
        public override VirtualDataTableState CreateState(string sort, TMasterDetailLink[] masterDetailLinks, TSplitLink splitLink)
        {
            return new TNRowsDataTableState(this, RecordCount);
        }

        #endregion

        #region Columns
        public override int ColumnCount
        {
            get
            {
                return 0;
            }
        }

        public override int GetColumn(string columnName)
        {
            FlxMessages.ThrowException(FlxErr.ErrInvalidColumn, columnName);
            return 0;
        }

        public override string GetColumnName(int columnIndex)
        {
            FlxMessages.ThrowException(FlxErr.ErrInvalidColumn, columnIndex);
            return String.Empty;
        }

        public override string GetColumnCaption(int columnIndex)
        {
            FlxMessages.ThrowException(FlxErr.ErrInvalidColumn, columnIndex);
            return String.Empty;
        }

        #endregion
    }

    internal class TNRowsDataTableState : VirtualDataTableState
    {
        #region Privates
        private int RecordCount;
        #endregion

        #region Constructors
        public TNRowsDataTableState(VirtualDataTable aTableData, int aRecordCount)
            : base(aTableData)
        {
            RecordCount = aRecordCount;
        }

        #endregion

        #region Rows and Data

        public override int RowCount
        {
            get
            {
                return RecordCount;
            }
        }

        public override object GetValue(int column)
        {
            return null;
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

        #endregion

        #region IDisposable Members
        #endregion

    }
    #endregion

    #region ColumnsDataTable
    struct TPositionAndName
    {
        int FPosition;
        string FName;
        public int Position { get { return FPosition; } set { FPosition = value; } }
        public string Name { get { return FName; } set { FName = value; } }

        internal TPositionAndName(int aPosition, string aName)
        {
            FPosition = aPosition;
            FName = aName;
        }
    }
    /// <summary>
    /// An implementation of a VirtualDataTable that is used return a table with the column names of other table.
    /// </summary>
    internal class TColumnsDataTable : VirtualDataTable
    {
        #region Privates
        internal TPositionAndName[] Columns;
        #endregion

        #region Constructors
        public TColumnsDataTable(string aTableName, VirtualDataTable aCreatedBy, VirtualDataTable MasterTable)
            : base(aTableName, aCreatedBy)
        {
            Columns = new TPositionAndName[MasterTable.ColumnCount];
            for (int i = 0; i < Columns.Length; i++)
            {
                Columns[i] = new TPositionAndName(i, MasterTable.GetColumnName(i));
            }
        }

        public TColumnsDataTable(string aTableName, VirtualDataTable aCreatedBy, TPositionAndName[] aColumns)
            : base(aTableName, aCreatedBy)
        {
            Columns = aColumns;
        }

        #endregion

        #region Settings
        public override CultureInfo Locale
        {
            get
            {
                return CultureInfo.CurrentCulture;
            }
        }

        #endregion

        #region Create State Dataset
        public override VirtualDataTableState CreateState(string sort, TMasterDetailLink[] masterDetailLinks, TSplitLink splitLink)
        {
            return new TColumnsDataTableState(this, Columns);
        }

        #endregion

        #region Columns
        public override int ColumnCount
        {
            get
            {
                return 2;
            }
        }

        public override int GetColumn(string columnName)
        {
#if (COMPACTFRAMEWORK)
            switch (columnName.ToLower(CultureInfo.InvariantCulture))
#else
            switch (columnName.ToLowerInvariant())
#endif
            {
                case "position": return 0;
                case "name": return 1;
                default:
                    break;
            }
            FlxMessages.ThrowException(FlxErr.ErrInvalidColumn, columnName);
            return 0;

        }

        public override string GetColumnName(int columnIndex)
        {
            switch (columnIndex)
            {
                case 0: return "Position";
                case 1: return "Name";
            }

            FlxMessages.ThrowException(FlxErr.ErrInvalidColumn, columnIndex);
            return String.Empty;
        }

        public override string GetColumnCaption(int columnIndex)
        {
            switch (columnIndex)
            {
                case 0: return "Position";
                case 1: return "Name";
            }
            FlxMessages.ThrowException(FlxErr.ErrInvalidColumn, columnIndex);
            return String.Empty;
        }

        #endregion

        #region Filters
        public override VirtualDataTable FilterData(string newDataName, string rowFilter)
        {
            if (rowFilter == null) return new TColumnsDataTable(newDataName, this, Columns);
#if(FRAMEWORK30 && !COMPACTFRAMEWORK)
            var sfp = new TSimpleFilterParser<TPositionAndName>(rowFilter, 
                new Dictionary<string,int>(StringComparer.InvariantCultureIgnoreCase)
                 {{"Position", 0},{"Name", 1}}, TableName);
            var f = sfp.ToExpression();
            return new TColumnsDataTable(newDataName, this, (f(Columns.AsQueryable())).ToArray());
#else
 	        return base.FilterData(newDataName, rowFilter);
#endif
        }
        #endregion
    }

    internal class TColumnsDataTableState : VirtualDataTableState
    {
        #region Privates
        private TPositionAndName[] Columns;
        #endregion

        #region Constructors
        public TColumnsDataTableState(VirtualDataTable aTableData, TPositionAndName[] aColumns)
            : base(aTableData)
        {
            Columns = aColumns;
        }

        #endregion

        #region Rows and Data

        public override int RowCount
        {
            get
            {
                return Columns.Length;
            }
        }

        public override object GetValue(int column)
        {
            switch (column)
            {
                case 0: return Position;
                case 1: return Columns[Position].Name;
            }
            return null;
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

        #endregion

        #region IDisposable Members
        #endregion

    }
    #endregion

}
