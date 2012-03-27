using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;

using FlexCel.Core;

namespace FlexCel.Report
{
	/// <summary>
	/// Base class for a table used on FlexCelReport.
	/// Inherit from this class and <see cref="VirtualDataTableState"/> to create your custom sources of data.
	/// Make sure you read the documentation on <b>UsingFlexCelReports.pdf</b> for more information.
	/// </summary>
	public abstract class VirtualDataTable: IDisposable
	{
		#region Privates
		internal string FTableName;
        internal VirtualDataTable FCreatedBy;
		#endregion

		#region Constructors
		/// <summary>
		/// Creates a new virtual datatable instance and assigns a name to it.
		/// </summary>
		/// <param name="aTableName">Name for the virtual data table. Note that this name is *not* used anywhere in FlexCel code, except to report errors.
		/// The Table names that are used on reports are the ones in <see cref="VirtualDataTableState"/></param>
        /// <param name="aCreatedBy">Table that created this table (via a filter, distinct, etc), or null if this table wasn't created from another VirtualDataTable.</param>
		protected VirtualDataTable(string aTableName, VirtualDataTable aCreatedBy)
		{
			FTableName = aTableName;
            FCreatedBy = aCreatedBy;
		}
		#endregion

		#region Settings
		/// <summary>
		/// Name for the virtual data table. Note that this name is *not* used anywhere in FlexCel code, except to report errors.
		/// The Table names that are used on reports are the ones in <see cref="VirtualDataTableState"/>
		/// </summary>
		public string TableName {get {return FTableName;}}

		/// <summary>
		/// Locale for this dataset. This might be needed to create datatables with data and the same locale.
		/// </summary>
		public abstract CultureInfo Locale{get;}
		#endregion

		#region Create State Dataset
		/// <summary>
		/// Creates a VirtualDataSetState to be used in a report. Make sure you override this method on your derived classes
		/// and point it to the correct VirtualDataSet descendant.
		/// </summary>
		/// <param name="sort">A string showing how to sort this dataset. This string might be null, empty, or whatever the user wrote on the config sheet.</param>
		/// <param name="masterDetailLinks">A list of the the master datatables and relation fields on the bands outside this one. 
		/// You can pass this parameter to the VirtualDataSetState so it can create indexes on the required fields. 
		/// This parameter will be an empty array if no master detail relationships apply to the VirtualDataSetState, but it will not be null.
		/// </param>
		/// <param name="splitLink">A link to a parent Split datasource with the number of records to split, or null if there is no parent split datasource.</param>
		/// <returns></returns>
		public abstract VirtualDataTableState CreateState(string sort, TMasterDetailLink[] masterDetailLinks, TSplitLink splitLink);
		#endregion

		#region Columns

		/// <summary>
		/// Returns the number of columns of the table.
		/// </summary>
		public abstract int ColumnCount{get;}

		/// <summary>
		/// Returns a column indentifier that you can later use on <see cref="VirtualDataTableState.GetValue"/>. 
		/// Return -1 if the column does not exist, and make sure this search is case insensitive.
		/// </summary>
		/// <param name="columnName">Name of the column to search.</param>
		/// <returns>Column identifier if found, -1 if not found.</returns>
		public abstract int GetColumn(string columnName);

		/// <summary>
		/// Returns the column name for a column identifier. This method is the reverse of <see cref="GetColumn"/>
		/// </summary>
		/// <param name="columnIndex">Column index returned by <see cref="GetColumn"/></param>
		/// <returns>The name of the column at columnIndex. If the columnIndex is not valid it should throw an Exception, since
		/// this method will only be called for columnIndexes returned by <see cref="GetColumn"/></returns>
		public abstract string GetColumnName(int columnIndex);

		/// <summary>
		/// Returns the column caption for a column identifier. This method is used on generic dataset to write the header column.
		/// For most uses, <see cref="GetColumnName"/> will be used.
		/// </summary>
		/// <param name="columnIndex">Column index returned by <see cref="GetColumn"/></param>
		/// <returns>The name of the column at columnIndex. If the columnIndex is not valid it should throw an Exception, since
		/// this method will only be called for columnIndexes returned by <see cref="GetColumn"/></returns>
		public abstract string GetColumnCaption(int columnIndex);
		#endregion

		#region Filter
		/// <summary>
		/// This method should return a new VirtualDataTable instance with the data filtered. If RowFilter is null,
        /// this method should return a copy of the dataset with a different name.
		/// Note that you might have the same data with different states, so this method might be called more than once.
		/// </summary>
		/// <remarks>You do not need to implement this method if you do not want to let users create new filtered datasets on the config sheet.</remarks>
		/// <param name="newDataName">How this new VirtualDataSet will be named. This is what the user wrote on the config sheet,
		/// when creating a filtered dataset. Note that as with all the VirtualDataSets, this name is meaningless except for error messages.</param>
		/// <param name="rowFilter">Filter for the new data. This can be null, an empty string, or whatever the user wrote on the "filter" column
		/// on the config sheet. Note that if the Filter is &quot;distinct()&quot; this method will not be called, but <see cref="GetDistinct"/> instead.
		/// This method will not be called if the filter is &quot;Split()&quot; either.</param>
		/// <returns>A new VirtualDataTable with the filtered data and the new name.</returns>
		public virtual VirtualDataTable FilterData(string newDataName, string rowFilter)
		{
			FlxMessages.ThrowException(FlxErr.ErrTableDoesNotSupportFilter, FTableName);
			return null; //just to compile.
		}
		#endregion

		#region Distinct
		/// <summary>
		/// Override this method to return a new VirtualDataSet with unique values.
		/// Note that the returned dataset will not have all the columns this one has, only the ones defined on &quot;filterFields&quot;
		/// </summary>
		/// <remarks>
		/// You do not need to implement this method if you do not want to let users create &quot;Distinct()&quot; filters on the config sheet.
		/// </remarks>
		/// <param name="newDataName">How this new VirtualDataSet will be named. This is what the user wrote on the config sheet,
		/// when creating the distinct dataset. Note that as with all the VirtualDataSets, this name is meaningless except for error messages.</param>
		/// <param name="filterFields">Fields where to apply the &quot;distinct&quot; condition.</param>
		/// <returns>A new VirtualDataTable with the filtered data and the new name.</returns>
		public virtual VirtualDataTable GetDistinct(string newDataName, int[] filterFields)
		{
			FlxMessages.ThrowException(FlxErr.ErrTableDoesNotSupportTag, FTableName, ReportTag.ConfigTag(ConfigTagEnum.Distinct));
			return null; //just to compile.
		}
		#endregion

        #region Detail tables
        /// <summary>
        /// Override this method if the table has linked tables that you can use for master detail relationships
        /// instead of normal relationships. This is the case for example in "Entity framework". 
        /// </summary>
        /// <param name="dataTableName">Name of the detail dataset we are looking for.</param>
        /// <param name="dataTable">Detail dataset that we are looking for.</param>
        /// <returns>The dataset if it is a detail dataset, null otherwise.</returns>
        public virtual VirtualDataTable GetDetail(string dataTableName, VirtualDataTable dataTable)
        {
            return null;
        }

        /// <summary>
        /// Override this method if the datatable has intrinsic relationships that you want to use.
        /// For example DataSets have DataRelationships, or Entity Framework tables are related as properties from the 
        /// master to the detail. All those relationships that are not explicitly defined in the report should be returned here.
        /// </summary>
        /// <param name="aDetail">Detail table from where we want to get the relationship.</param>
        /// <returns></returns>
        public virtual TRelation GetRelationWith(VirtualDataTable aDetail)
        {
            return null;
        }

        /// <summary>
        /// Returns the table that created this one (by a filter, distinct, etc), or null if this table 
        /// was not created from another VirtualDataTable.
        /// </summary>
        public VirtualDataTable CreatedBy { get { return FCreatedBy; } }

        #endregion

		#region Lookup
		/// <summary>
		/// Looks for a key on this dataset and returns the corresponding value.
		/// Note: Remember that VirtualDataSet is stateless, so if you use any caching here, make sure you appropiately lock()
		/// this method so there is no possibility of one thread reading the cache when the other is updating it.
		/// </summary>
		/// <remarks>
		/// You do not need to implement this method if you do not want to let your users use the &lt;#Lookup&gt; tag.
		/// </remarks>
		/// <param name="column">Column with the value to be returned.</param>
		/// <param name="keyNames">A list of column names, as the user wrote them on the &lt;#Lookup&gt; tag</param>
		/// <param name="keyValues">A list of the values for the keys, that you should use to locate the right record.</param>
		/// <returns>The value at "column" , for the record where the columns on "keyNames" have the "keyValues" values.
		/// If there is more than one record where "keyNames" is equal to "keyValues" you might opt to throw an Exception or just return any of the 
		/// valid values, depending on the behavior you want for lookup.</returns>
		public virtual object Lookup(int column, string keyNames, object[] keyValues)
		{
			FlxMessages.ThrowException(FlxErr.ErrTableDoesNotSupportLookup, FTableName);
			return null; //just to compile.
		}

		#endregion

		#region IDisposable Members

		/// <summary>
		/// Dispose whatever is needed on the children here.
		/// </summary>
		public void Dispose()
		{
			Dispose(true);
			GC.SuppressFinalize(this);
		}

		/// <summary>
		/// Override this method on derived classes.
		/// </summary>
		/// <param name="disposing"></param>
		protected virtual void Dispose(bool disposing) 
		{
			//Nothing here.
		}

		#endregion
    }


	/// <summary>
	/// A table that corresponds to a band on the report.
	/// Make sure you read the documentation on <b>UsingFlexCelReports.pdf</b> for more information.
	/// </summary>
	public abstract class VirtualDataTableState: IDisposable
	{
		#region Privates
		/// <summary>
		/// This value should not be modified externally.
		/// </summary>
		private int FPosition;
		internal string FTableName;
		private VirtualDataTable FTableData;
		#endregion

		#region Constructors
		/// <summary>
		/// Constructs a new VirtualDataTableState with the specified name. Note that you will not call
		/// this constructor directly, new VirtualDataTableState instances will be created only by <see cref="VirtualDataTable.CreateState"/>
		/// </summary>
		/// <param name="aTableData">VirtualDataTable that created this instance.</param>
		protected VirtualDataTableState(VirtualDataTable aTableData)
		{
			FTableData = aTableData;
		}
		#endregion

        #region Settings
        /// <summary>
		/// The VirtualDataTable that created this instance.
		/// </summary>
		public VirtualDataTable TableData {get {return FTableData;}}

		/// <summary>
		/// Returns the active row on the table. (0 based) 
		/// You should use this value to return the values on <see cref="GetValue"/>
		/// </summary>
		public int Position
		{
			get
			{
				return FPosition;
			}
		}

		/// <summary>
		/// Returns the table name assigned on the template to this dataset. Note that this name is the one on the bands in the template.
		/// </summary>
		public string TableName
		{
			get
			{
				return FTableName;
			}
		}

		#endregion

        #region Aggregate
        /// <summary>
        /// This method is used by the "AGGREGATE" tag in a FlexCel report to calculate the mximum/minimum/average/etc of 
        /// the values in the table. If you don't implement this method, FlexCel will still calculate those values by looping
        /// through the dataset, but if you have a faster way to do it (like with a "select max(field) from table") then implement this
        /// method and return true.
        /// </summary>
        /// <param name="aggregateType">Which operation to do on the dataset. (Max/Min/etc)</param>
        /// <param name="colIndex">Index of the filed in which we want to aggregate.</param>
        /// <param name="resultValue">Returns the result of the operation in the dataset.</param>
        /// <returns>True if this method is implemented, false if not. Note that even if we return false here,
        /// FlexCel will still calculate the aggregate by looping through all the records.</returns>
        public virtual bool TryAggregate(TAggregateType aggregateType, int colIndex, out double? resultValue)
        {
            resultValue = 0;
            return false;
        }

        #endregion
 
		#region Rows and Data
		/// <summary>
		/// Returns the number of rows available on the dataset, for the current state. Note that this method can be called many times, so it should be fast.
		/// Usa a cache if necessary. Do *not* use something like "return select count(*) from table" here, it would be too slow.
		/// </summary>
		public abstract int RowCount{get;}

        /// <summary>
        /// This method returns if we have reached the last record in the table. The default implementation
        /// just sees if <see cref="Position"/> = <see cref="RowCount"/>. If RowCount is slow and you have a 
        /// faster way to know if you are at the end, override this method.
        /// </summary>
        /// <returns>True if we are at the end of the datatable.</returns>
        public virtual bool Eof()
        {
            return Position >= RowCount;
        }

		/// <summary>
		/// Returns the value for row <see cref="Position"/>, at the column "column"
		/// </summary>
		/// <param name="column">Column identifier returned by <see cref="VirtualDataTable.GetColumn"/></param>
		/// <returns></returns>
		public abstract object GetValue(int column);
		#endregion

        #region Move
        internal void DoMoveFirst()
        {
            FPosition = 0;
            MoveFirst();
        }

        /// <summary>
        /// This method is called when we want to move to the first record. 
        /// You can always know the current record with <see cref="Position"/>
        /// </summary>
        public virtual void MoveFirst()
        {
        }


        internal void DoMoveNext()
        {
            FPosition++;
            MoveNext();
        }

        /// <summary>
        /// This method is called when we want to move to the next record.
        /// You can always know the current record with <see cref="Position"/>
        /// </summary>
        public abstract void MoveNext();

        internal void DoMoveToRecord(int aPosition)
        {
            FPosition = aPosition;
            MoveToRecord(aPosition);
        }

        /// <summary>
        /// This method is called when you move to a random record and it is used by the DbValue tag. 
        /// By default it will raise an exception. If your VirtualDataSource supportes random lookup, override this method
        /// and make it not raise the exception.
        /// </summary>
        /// <param name="aPosition">Position in the dataset where we want to move.</param>
        public virtual void MoveToRecord(int aPosition)
        {
            FlxMessages.ThrowException(FlxErr.ErrTableDoesNotSupportTag, FTableName, "DBVALUE");
        }

        #endregion

        #region Relationships
        /// <summary>
		/// This method will be called each time that the master datasource moves its position. Use it to filter the data returned
		/// if this is used on a master-detail relationship.
		/// </summary>
		/// <remarks>
		/// You do not need to implement this method if you are not using Master-Detail or Split relationships.</remarks>
		/// <param name="masterDetailLinks">List of all the master tables that are related to this one.
		/// If there are no parents on this VirtualDataTableState, this will be an empty array, not null.
		/// Use it on <see cref="GetValue"/> to filter the data and then return only the records that satisfy the master-detail relationships on <see cref="GetValue"/>
		/// </param>
		/// <param name="splitLink">Parent Split table if this dataset is on a Split relationship, or null if there is none.
		/// Use it to know how many records you should retun on <see cref="RowCount"/>. Note that a table might be on Master-Detail relationship 
		/// *and* split relationship. In this case you need to first filter the records that are on the master detail relationship, and then apply the split to them.
		/// </param>
		public virtual void MoveMasterRecord(TMasterDetailLink[] masterDetailLinks, TSplitLink splitLink)
		{
			if (masterDetailLinks == null) return;
			for (int i = 0; i < masterDetailLinks.Length; i++)
			{
				if (masterDetailLinks[i].ParentDataSource.RowCount != 0)  //Avoid throwing an exception if there is actually only one master record.
					FlxMessages.ThrowException(FlxErr.ErrTableDoesNotSupportMasterDetail, FTableData.TableName);
			}
		}

		/// <summary>
		/// This method will be called when a Split master wants to know how many records its detail has. For example, if the detail has 30 records 
		/// and the split is at 10, the Split master will call this method to find out that it has to return 3 on its own record count.
		/// You need to filter the data here depending on the master detail relationships, but not on the splitLink.
		/// </summary>
		/// <remarks>
		/// You do not need to implement this method if you are not using Split relationships.
		/// </remarks>
		/// <param name="masterDetailLinks">List of all the master tables that are related to this one.
		/// If there are no parents on this VirtualDataTableState, this will be an empty array, not null.
		/// </param>
		/// <returns>The count of records for a specific position of the master datsets.</returns>
		public virtual int FilteredRowCount(TMasterDetailLink[] masterDetailLinks)
		{
			FlxMessages.ThrowException(FlxErr.ErrTableDoesNotSupportTag, FTableData.TableName, ReportTag.ConfigTag(ConfigTagEnum.Split));
			return -1; //just to compile.
		}
		#endregion

		#region IDisposable Members

		/// <summary>
		/// Dispose whatever is needed on the children here.
		/// </summary>
		public void Dispose()
		{
			Dispose(true);
			GC.SuppressFinalize(this);
		}

		/// <summary>
		/// Override this method on derived classes.
		/// </summary>
		/// <param name="disposing"></param>
		protected virtual void Dispose(bool disposing) 
		{
			//Nothing here.
		}

        /// <summary>
        /// Finalizer for the class.
        /// </summary>
        ~VirtualDataTableState()      
        {
            Dispose(false);
        }

		#endregion
    }


	/// <summary>
	/// A parent table and a parent field used on a master-detail relationship.
	/// </summary>
	public class TMasterDetailLink
	{
		private VirtualDataTableState FParentDataSource;
		private int FParentField;
		private string FChildFieldName;

		/// <summary>
		/// Creates a new instance of TMasterDetailLink with the given values.
		/// </summary>
		/// <param name="aParentDataSource">See <see cref="ParentDataSource"/></param>
		/// <param name="aParentField">See <see cref="ParentField"/></param>
		/// <param name="aChildFieldName">See <see cref="ChildFieldName"/></param>
		public TMasterDetailLink(VirtualDataTableState aParentDataSource, int aParentField, string aChildFieldName)
		{
			FParentDataSource=aParentDataSource;
			FParentField=aParentField;
			FChildFieldName=aChildFieldName;
		}


        /// <summary>
        /// A table that is acting as a master on a master detail relationship.
        /// </summary>
		public VirtualDataTableState ParentDataSource {get {return FParentDataSource;}}

		/// <summary>
		/// Column index of the key that acts as primary key on the relationship.
		/// </summary>
		public int ParentField {get {return FParentField;}}

		/// <summary>
		/// Column name on the detail dataset that is related with <see cref="ParentField"/>
		/// </summary>
		public string ChildFieldName {get {return FChildFieldName;}}

	}


	/// <summary>
	/// Specifies a "Split" relation between two tables.
	/// </summary>
	public class TSplitLink
	{
		internal VirtualDataTableState FParentDataSource;
		private int FSplitCount;

		/// <summary>
		/// Create a new TSplitLink with the given values.
		/// </summary>
		/// <param name="aParentDataSource">See <see cref="ParentDataSource"/></param>
		/// <param name="aSplitCount">See <see cref="SplitCount"/></param>
		public TSplitLink(VirtualDataTableState aParentDataSource, int aSplitCount)
		{
			FParentDataSource=aParentDataSource;
			FSplitCount = aSplitCount;
		}

		/// <summary>
		/// Split the detail every &quot;SplitCount&quot; number of records. If for example Splitcount is 5 and 
		/// the detail table has 14 registers, it should be split as 5 records, 5 records, 4 records.
		/// </summary>
		public int SplitCount {get {return FSplitCount;}}

		/// <summary>
		/// Table that acts as a parent on the split relationship. Read its <see cref="VirtualDataTableState.Position"/> to know
		/// which group of splitted records to return.
		/// </summary>
		public VirtualDataTableState ParentDataSource {get {return FParentDataSource;}}

	}

    	/// <summary>
	/// A data relation between two tables. Different from standard .NET datarelations, this class is not tied
	/// to ADO.NET, and allows you to specify relationships between any arbitrary VirtualDataTable objects.
	/// </summary>
	public class TRelation
	{
		private VirtualDataTable FParentTable;
		private VirtualDataTable FChildTable;
		private int[] FParentColumns;
		private int[] FChildColumns;

		/// <summary>
		/// Creates a new relation with the given values.
		/// </summary>
		/// <param name="aParentTable">See <see cref="ParentTable"/></param>
		/// <param name="aChildTable">See <see cref="ChildTable"/></param>
		/// <param name="aParentColumns">See <see cref="ParentColumns"/></param>
		/// <param name="aChildColumns">See <see cref="ChildColumns"/></param>
		public TRelation(VirtualDataTable aParentTable, VirtualDataTable aChildTable, int[] aParentColumns, int[] aChildColumns)
		{
			FParentTable = aParentTable;
			FChildTable = aChildTable;
			FParentColumns = aParentColumns;
			FChildColumns = aChildColumns;
		}

		/// <summary>
		/// And array of colum indexes on the Master table that are related to the detail
		/// ChildColumns[0] is related to ParentColumns[0], ChildColumns[1] to ParentColunns[1], and so on.
		/// </summary>
		public int[] ParentColumns {get {return FParentColumns;}}

		/// <summary>
		/// And array of colum indexes on the Detail table that are related to the master.
		/// ChildColumns[0] is related to ParentColumns[0], ChildColumns[1] to ParentColunns[1], and so on.
		/// </summary>
		public int[] ChildColumns {get {return FChildColumns;}}

		/// <summary>
		/// Table that acts as a master on a Master-Detail relationship.
		/// </summary>
		public VirtualDataTable ParentTable {get {return FParentTable;}}

		/// <summary>
		/// Table that acts as a detail on a Master-Detail relationship.
		/// </summary>
		public VirtualDataTable ChildTable {get {return FChildTable;}}

	}

	/// <summary>
	/// A data relation between two tables defined in a report. Different from standard .NET datarelations, this class is not tied
	/// to ADO.NET, and allows you to specify relationships between any arbitrary tables. By holding strings and not virtualdatatables,
    /// this allows us to set the relationship before tables are added.
	/// </summary>
	internal class TRelationship
	{
		private string FParentTable;
		private string FChildTable;
		private string[] FParentColumns;
		private string[] FChildColumns;

		/// <summary>
		/// Creates a new relation with the given values.
		/// </summary>
		/// <param name="aParentTable">See <see cref="ParentTable"/></param>
		/// <param name="aChildTable">See <see cref="ChildTable"/></param>
		/// <param name="aParentColumns">See <see cref="ParentColumns"/></param>
		/// <param name="aChildColumns">See <see cref="ChildColumns"/></param>
		public TRelationship(string aParentTable, string aChildTable, string[] aParentColumns, string[] aChildColumns)
		{
			FParentTable = aParentTable;
			FChildTable = aChildTable;
			FParentColumns = aParentColumns;
			FChildColumns = aChildColumns;
		}

		/// <summary>
		/// And array of columns on the Master table that are related to the detail
		/// ChildColumns[0] is related to ParentColumns[0], ChildColumns[1] to ParentColunns[1], and so on.
		/// </summary>
		public string[] ParentColumns {get {return FParentColumns;}}

		/// <summary>
		/// And array of columns on the Detail table that are related to the master.
		/// ChildColumns[0] is related to ParentColumns[0], ChildColumns[1] to ParentColunns[1], and so on.
		/// </summary>
		public string[] ChildColumns {get {return FChildColumns;}}

		/// <summary>
		/// Table that acts as a master on a Master-Detail relationship.
		/// </summary>
		public string ParentTable {get {return FParentTable;}}

		/// <summary>
		/// Table that acts as a detail on a Master-Detail relationship.
		/// </summary>
		public string ChildTable {get {return FChildTable;}}

	}

	/// <summary>
	/// A list of <see cref="TRelation"/> classes, that you can use to group all the relations on a VirtualDataTable.
	/// </summary>
	internal sealed class TRelationshipList: IEnumerable<TRelationship>
    {
        List<TRelationship> FList;

		/// <summary>
		/// Creates a new TRelationshipList instance.
		/// </summary>
		public TRelationshipList()
		{
            FList = new List<TRelationship>();
		}

		/// <summary>
		/// Clears all relations on this table.
		/// </summary>
		public void Clear()   
		{
			FList.Clear();
		}

		/// <summary>
		/// Adds a new relation to the list.
		/// </summary>
		/// <param name="relation">Relation to add.</param>
		public void Add(TRelationship relation)
		{
			FList.Add(relation);
		}

		/// <summary>
		/// Returns the specified Relation at index (0 based)
		/// </summary>
		public TRelationship this[int index]
		{
			get
			{
                return FList[index];
			}
		}

		/// <summary>
		/// Returns the number of relations on this list.
		/// </summary>
		public int Count
		{
			get
			{
				return FList.Count;
			}
		}

        #region IEnumerable<TRelationship> Members
        IEnumerator<TRelationship> IEnumerable<TRelationship>.GetEnumerator()
        {
            return FList.GetEnumerator();
        }
        #endregion

        #region IEnumerable Members

        IEnumerator IEnumerable.GetEnumerator()
        {
            return FList.GetEnumerator();
        }

        #endregion
    }



}
