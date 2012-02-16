#if (FRAMEWORK30 && !DELPHIWIN32)
using System;
using System.Data;
using System.Linq;
using System.Text;
using System.Globalization;
using System.Diagnostics;
using System.Collections;

using FlexCel.Core;
using System.Collections.Generic;
using System.Reflection;
using System.Linq.Expressions;

namespace FlexCel.Report
{
    #region LINQ
    /// <summary>
    /// Interface for generic methods in ILinqDataTable.
    /// </summary>
    public interface ILinqDataTable
    {
        /// <summary>
        /// Returns true if the database has a linked datatable with name dataTableName.
        /// </summary>
        /// <param name="dataTableName">Name for the linked datatable that we are searching for.</param>
        /// <returns></returns>
        bool HasDetail(string dataTableName);

        /// <summary>
        /// Returns the base table this one is created from. This is used in relationships, 
        /// to find out the relationships of the master table.
        /// </summary>
        /// <returns></returns>
        string FindOrigTableName();
    }

    /// <summary>
    /// A class to store compiled expressions to access fields in a <see cref="TLinqDataTable{T}"/>. 
    /// This is faster than reading the field values with reflection. Field expressions will be created on demand,
    /// so if you never use a field, an expression for it won't be created.
    /// </summary>
    /// <typeparam name="T"></typeparam>
    public class TFieldExpressions<T>
    {
        Func<T, object>[] Funcs;
        private PropertyInfo[] Fields;

        /// <summary>
        /// Creates a new instance and assigns it to an array of PropertyInfo. 
        /// </summary>
        /// <param name="aFields"></param>
        public TFieldExpressions(PropertyInfo[] aFields)
        {
            Fields = aFields;
            Funcs = new Func<T, object>[aFields.Length];
        }

        /// <summary>
        /// Gets the compiled expression for the needed field.
        /// </summary>
        /// <param name="index"></param>
        /// <returns></returns>
        public Func<T, object> this[int index]
        {
            get
            {
                if (Funcs[index] == null) Funcs[index] = GenerateFunc(index);
                return Funcs[index];
            }
        }

        private Func<T, object> GenerateFunc(int index)
        {
            var x = Expression.Parameter(typeof(T), "x");
            var pc = Expression.Property(x, Fields[index]);
            var po = Expression.Convert(pc, typeof(object));
            return Expression.Lambda<Func<T, object>>(po, x).Compile();
        }

    }

    /// <summary>
    /// An implementation of a VirtualDataTable using IEnumerator and Linq. 
    /// This class is abstract so to use it you should inherit from it and override the needed methods.
    /// In normal cases you won't need to override this class, as FlexCel already has an implementation of
    /// it that deals with IEnumerator tables. Inheriting fro this class is only for very special cases.
    /// </summary>
    public abstract class TLinqDataTable<T> : VirtualDataTable, ILinqDataTable
    {
        #region Privates
        /// <summary>
        /// Data on this table. It might be null, if this table is linked to another.
        /// </summary>
        protected IQueryable<T> Data { get; set; }

        /// <summary>
        /// Type definitions for the properties of the class this table holds.
        /// </summary>
        protected PropertyInfo[] Fields { get; set; }

        /// <summary>
        /// A list of compiled field expressions that can be used to find the values of the fileds without using reflection.
        /// </summary>
        protected TFieldExpressions<T> FieldExpressions { get; set; }

        Dictionary<string, int> FFieldsByName;
        /// <summary>
        /// A dictionary to find the field position in the <see cref="Fields"/> array given its name.
        /// </summary>
        protected Dictionary<string, int> FieldsByName { get { return FFieldsByName; } }

        /// <summary>
        /// An array of filter strings that must be applied to this dataset.
        /// </summary>
        protected string[] Filters { get; set; }

        TLinqDataTable<T> OrigTable;
        #endregion

        #region Constructors
        /// <summary>
        /// Creates a new table with "tablename" name and a source data.
        /// Data can't be null when using this constructor.
        /// </summary>
        /// <param name="aTableName">Name of the table that will be created.</param>
        /// <param name="aData">Data. Can't be null in this particular constructor.</param>
        /// <param name="aCreatedBy">Table that created this table (via a filter, distinct, etc), or null if this table wasn't created from another VirtualDataTable.</param>
        protected TLinqDataTable(string aTableName, VirtualDataTable aCreatedBy, IQueryable<T> aData)
            : base(aTableName, aCreatedBy)
        {
            if (aData == null) FlxMessages.ThrowException(FlxErr.ErrDataSetNull);
            Filters = new string[0];
            Data = aData;
            ReadFields();
            OrigTable = this;
        }

        //Note that this method is called by Activator.
        /// <summary>
        /// This constructor is used to create a nested detail table. Data might be null when calling this constructor
        /// when it will be read from the master table every time the master record changes.
        /// <b>IMPORTANT: When creating a descendand of this class, you always need to provide this constructor as it will
        /// be called automatically when creating filtered of detail tables.</b>
        /// </summary>
        /// <param name="aTableName">Name for the new table.</param>
        /// <param name="aData">Data for the table. Might be null if in a master-detail relationship.</param>
        /// <param name="aOrigTable">Real table from where the data comes. We will use this to find the relationships.</param>
        /// <param name="aFilters">A string with all the filters that need to be applied to this table.</param>
        /// <param name="aCreatedBy">Table that created this table (via a filter, distinct, etc), or null if this table wasn't created from another VirtualDataTable.</param>
        protected TLinqDataTable(string aTableName, VirtualDataTable aCreatedBy, IQueryable<T> aData, TLinqDataTable<T> aOrigTable,
            string[] aFilters)
            : base(aTableName, aCreatedBy)
        {
            Data = aData;
            Filters = aFilters;
            OrigTable = aOrigTable;
            ReadFields();
        }

        /// <summary>
        /// Override this method when inheriting from this class to return the correct object.
        /// To do so, you will need to create a constructor in your class with parameters (aTableName, aCreator, aOrigTable, aFilters),
        /// call the inherited constructor in this class with the same parameters, and then in this method create a new
        /// instance of your class with the above constructor. 
        /// </summary>
        /// <param name="aTableName">Name for the new table.</param>
        /// <param name="aCreatedBy">Table that created this table (via a filter, distinct, etc), or null if this table wasn't created from another VirtualDataTable.</param>
        /// <param name="aData">Data for the table. Might be null if in a master detail relationship.</param>
        /// <param name="aOrigTable">Real table from where the data comes. We will use this to find the relationships.</param>
        /// <param name="aFilters">A string with all the filters that need to be applied to this table.</param>
        /// <returns>A new instance of the class.</returns>
        protected abstract VirtualDataTable CreateNewDataTable(string aTableName, VirtualDataTable aCreatedBy, IQueryable<T> aData, TLinqDataTable<T> aOrigTable, string[] aFilters);

        /// <summary>
        /// Original name of the table. This will be used in relationships.
        /// </summary>
        /// <returns></returns>
        public string FindOrigTableName()
        {
            if (OrigTable == null) return null;
            return OrigTable.TableName;
        }

        private void ReadFields()
        {
            Fields = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance);
            FieldExpressions = new TFieldExpressions<T>(Fields);
            FFieldsByName = new Dictionary<string, int>(StringComparer.CurrentCultureIgnoreCase);

            for (int i = 0; i < Fields.Length; i++)
            {
                FieldsByName[Fields[i].Name] = i;
            }
        }
        #endregion

        #region Settings
        ///<inheritdoc />
        public override CultureInfo Locale
        {
            get
            {
                return CultureInfo.CurrentCulture;
            }
        }
        #endregion

        #region Columns
        ///<inheritdoc />
        public override int ColumnCount
        {
            get
            {
                return Fields.Length;
            }
        }

        ///<inheritdoc />
        public override int GetColumn(string columnName)
        {
            int Result;
            if (FieldsByName.TryGetValue(columnName, out Result)) return Result;
            return -1;
        }

        ///<inheritdoc />
        public override string GetColumnName(int columnIndex)
        {
            return Fields[columnIndex].Name;
        }

        ///<inheritdoc />
        public override string GetColumnCaption(int columnIndex)
        {
            return Fields[columnIndex].Name;
        }
        #endregion

        #region Filter
        ///<inheritdoc />
        public override VirtualDataTable FilterData(string newDataName, string rowFilter)
        {
            if (String.IsNullOrEmpty(rowFilter)) return CreateNewDataTable(newDataName, this, Data, OrigTable, Filters);
            //We can't filter here as we might not have the data yet. So we will just add the filter to a list.
            var NewFilters = Filters.ToList();
            NewFilters.Add(rowFilter);
            return CreateNewDataTable(newDataName, this, Data, OrigTable, NewFilters.ToArray());
        }
        #endregion

        #region Master-Detail
        ///<inheritdoc />
        public override VirtualDataTable GetDetail(string dataTableName, VirtualDataTable dataTable)
        {
            Type[] typeParameters = ParseDetail(dataTableName);
            if (typeParameters == null) return null;
            Type nt = this.GetType().GetGenericTypeDefinition().MakeGenericType(typeParameters[0]);

            return (VirtualDataTable)Activator.CreateInstance(nt, dataTableName, this, null, dataTable,
                Filters); //Here data is null because we will calculate it when the master moves. This new dataset can't be called outside a master-detail relationship.
        }

        private Type[] ParseDetail(string dataTableName)
        {
            Type childTableType = null;
            int position;
            if (dataTableName == null) return null;
            if (!FieldsByName.TryGetValue(dataTableName, out position)) return null;

            PropertyInfo Field = Fields[position];
            childTableType = Field.PropertyType;
            if (!childTableType.IsGenericType) return null;
            Type[] typeParameters = childTableType.GetGenericArguments();
            if (typeParameters == null || typeParameters.Length != 1) return null;

            return typeParameters;
        }

        /// <summary>
        /// Returns true if the database has a linked datatable with name dataTableName.
        /// </summary>
        /// <param name="dataTableName">Name for the linked datatable that we are searching for.</param>
        /// <returns></returns>
        public bool HasDetail(string dataTableName)
        {
            return ParseDetail(dataTableName) != null;
        }

        ///<inheritdoc />
        public override TRelation GetRelationWith(VirtualDataTable aDetail)
        {
            ILinqDataTable LinqDetail = aDetail as ILinqDataTable;
            if (LinqDetail == null) return null;
            if (OrigTable.HasDetail(LinqDetail.FindOrigTableName()))
            {
                return new TRelation(this, aDetail, null, null);
            }
            return null;
        }

        #endregion

        #region Distinct
        /*        
        ///<inheritdoc />
        public override VirtualDataTable GetDistinct(string newDataName, int[] filterFields)
        {
            var DistinctData = from p in Data select p;
        }
*/
        #endregion

        #region Lookup
        ///<inheritdoc />
        public override object Lookup(int column, string keyNames, object[] keyValues)
        {
            var sourceParam = Expression.Parameter(typeof(T), "x");
            int[] KeyColumns = ParseColumns(keyNames);
            if (keyValues == null || KeyColumns.Length != keyValues.Length) FlxMessages.ThrowException(FlxErr.ErrValuesAndKeysMismatch);
            Expression colReference = null;
            int k = -1;
            foreach (int KeyCol in KeyColumns)
            {
                k++;
                if (keyValues[k] == null || Convert.IsDBNull(keyValues[k])) return null;
                var propertyReference = Expression.Property(sourceParam, Fields[KeyCol]);
                var sourceColParam = Expression.Constant(keyValues[k]);
                var Relationship = Expression.Equal(propertyReference, sourceColParam);
                if (colReference == null) colReference = Relationship; else colReference = Expression.AndAlso(colReference, Relationship);
            }
            var whereSelector = Expression.Lambda<Func<T, bool>>(colReference, sourceParam);

            var LookedUp = OrigTable.Data.FirstOrDefault(whereSelector);
            if (LookedUp == null) return null;

            return FieldExpressions[column](LookedUp);

        }

        private int[] ParseColumns(string keyNames)
        {
            TColumnAndOrder[] cols = TColumnAndOrder.GetSortFields(keyNames);
            if (cols == null || cols.Length == 0) FlxMessages.ThrowException(FlxErr.ErrEmptyKeyNames);

            return (from p in cols select FieldsByName[p.Column]).ToArray();


        }
        #endregion

        #region IDisposable Members
        ///<inheritdoc />
        protected override void Dispose(bool disposing)
        {
            try
            {
                //only managed resources
                if (disposing)
                {
                }
            }
            finally
            {
                //last call.
                base.Dispose(disposing);
            }
        }
        #endregion
    }

    /// <summary>
    /// This interface is used by <see cref="TLinqDataTableState{T}"/> for calling generic methods.
    /// </summary>
    public interface ILinqDataTableState
    {
        /// <summary>
        /// Returns a detail table. This is used in master-detail relationships to get the nested table for the detail band.
        /// </summary>
        /// <typeparam name="Q"></typeparam>
        /// <param name="Detail">Detail data table calling this method.</param>
        /// <param name="IsEmptyList">Return true if the enumerator has 0 records.</param>
        /// <returns></returns>
        IEnumerable<Q> GetDetail<Q>(TLinqDataTable<Q> Detail, out bool IsEmptyList);
    }

    /// <summary>
    /// This class implements support for Linq queries in FlexCelReport. 
    /// Inherit this class if you want to implement some non-standard Linq query provider.
    /// For normal cases, you don't need to do anything since FlexCel already implements a descendant of this
    /// class with the default functionality.
    /// </summary>
    /// <typeparam name="T">Type of the object that will be enumerated.</typeparam>
    public abstract class TLinqDataTableState<T> : VirtualDataTableState, IDisposable, ILinqDataTableState
    {
        #region Privates
        private IQueryable<T> FData;
        private IQueryable<T> FUnFilteredData;
        private bool NeedsSorting;

        /// <summary>
        /// Data for the table. Note that this might change in a master-detail report, every time the master
        /// changes its active record, Data in the detail will change to the records for the master.
        /// </summary>
        protected IQueryable<T> Data
        {
            get
            {
                if (NeedsSorting)
                {
                    SortAndFilter(FData == null);
                    NeedsSorting = false;
                }
                return FData;
            }
        }


        /// <summary>
        /// A list of then public fields in the datatype.
        /// </summary>
        protected PropertyInfo[] Fields { get; set; }

        /// <summary>
        /// Cached expressions to access the fields faster than with reflection.
        /// </summary>
        private TFieldExpressions<T> FieldExpressions { get; set; }

        Dictionary<string, int> FFieldsByName;
        /// <summary>
        /// A dictionary to find the field position in the <see cref="Fields"/> array given its name.
        /// </summary>
        protected Dictionary<string, int> FieldsByName { get { return FFieldsByName; } }

        /// <summary>
        /// A list of filter strings that must be applied to the data.
        /// </summary>
        protected string[] Filters { get; set; }

        /// <summary>
        /// Split link.
        /// </summary>
        protected TSplitLink SplitLink { get; private set; }

        Func<IQueryable<T>, IQueryable<T>>[] FilterActions;
        Func<IEnumerable<T>, IQueryable<T>> DataConverter;
        bool SearchedForDataConverter;
        IEnumerator<T> FilteredDataEnumerator;
        T FilteredDataCurrent;
        bool EofReached, NextIsEof, AtBof;

        int CachedCount = -1;
        string FSortString;
        #endregion

        #region Constructors
        /// <summary>
        /// Creates a new LinqDataTableState.
        /// </summary>
        /// <param name="aTableData"></param>
        /// <param name="aData"></param>
        /// <param name="aFields"></param>
        /// <param name="aFieldExpressions"></param>
        /// <param name="aFieldsByName"></param>
        /// <param name="aFilters"></param>
        /// <param name="sort"></param>
        /// <param name="masterDetailLinks"></param>
        /// <param name="splitLink"></param>
        protected TLinqDataTableState(VirtualDataTable aTableData, IQueryable<T> aData, PropertyInfo[] aFields,
            TFieldExpressions<T> aFieldExpressions, Dictionary<string, int> aFieldsByName,
            string[] aFilters,
            string sort, TMasterDetailLink[] masterDetailLinks, TSplitLink splitLink)
            : base(aTableData)
        {
            Fields = aFields;
            FieldExpressions = aFieldExpressions;
            FFieldsByName = aFieldsByName;
            Filters = aFilters;
            FData = aData;
            FUnFilteredData = aData;
            FSortString = sort;

            NeedsSorting = true;

        }

        #endregion

        #region Sort
        /// <summary>
        /// Sorts the data according to SortStr.
        /// </summary>
        /// <param name="SortStr">Sort string the user wrote in the report.</param>
        public virtual void SortData(string SortStr)
        {
            if (string.IsNullOrEmpty(SortStr)) return;
            if (SortStr.Trim().Length == 0) return;

            TColumnAndOrder[] Columns = TColumnAndOrder.GetSortFields(SortStr);
            var First = true;
            foreach (var col in Columns)
            {
                int ColIndex = GetSortField(col.Column);
                if (ColIndex < 0) continue;

                Action<int, bool, bool> SortMethod = SortDataByField<int>;
                MethodInfo method = this.GetType().GetMethod(SortMethod.Method.Name); //avoid using the string "SortDataByField" because it could confuse obfuscators.
                MethodInfo generic = method.MakeGenericMethod(Fields[ColIndex].PropertyType);
                generic.Invoke(this, new object[] { ColIndex, First, col.Descending });
                First = false;
            }
        }

        /// <summary>
        /// This method is used by the current <see cref="SortData"/> implementation, and it creates an expression to sort the data.
        /// </summary>
        /// <typeparam name="Q"></typeparam>
        /// <param name="SortField"></param>
        /// <param name="First"></param>
        /// <param name="Descending"></param>
        public void SortDataByField<Q>(int SortField, bool First, bool Descending)
        {
            var sourceParam = Expression.Parameter(typeof(T), "x");
            var propertyReference = Expression.Property(sourceParam, Fields[SortField]);
            var orderBySelector = Expression.Lambda<Func<T, Q>>(propertyReference, sourceParam);

            IOrderedQueryable<T> fg;

            if (Descending)
            {
                if (First)
                {
                    fg = FData.OrderByDescending<T, Q>(orderBySelector);
                }
                else
                {
                    fg = ((IOrderedQueryable<T>)FData).ThenByDescending<T, Q>(orderBySelector);
                }
            }
            else
            {
                if (First)
                {
                    fg = FData.OrderBy<T, Q>(orderBySelector);
                }
                else
                {
                    fg = ((IOrderedQueryable<T>)FData).ThenBy<T, Q>(orderBySelector);
                }
            }

            FData = fg.AsQueryable<T>();

        }

        private int GetSortField(string SortField)
        {
            if (string.IsNullOrEmpty(SortField)) return -1;
            if (SortField.Trim().Length == 0) return -1;

            int index;
            if (!FieldsByName.TryGetValue(SortField, out index))
            {
                FlxMessages.ThrowException(FlxErr.ErrInvalidSortString, TableName, SortField);
            }
            return index;
        }
        #endregion

        #region Aggregate
        ///<inheritdoc />
        public override bool TryAggregate(TAggregateType aggregateType, int colIndex, out double? resultValue)
        {
            resultValue = null;
            string MethodName = GetAggMethodName(aggregateType);
            if (MethodName == null) return false;

            Type tp = Fields[colIndex].PropertyType;
            if (aggregateType == TAggregateType.Min || aggregateType == TAggregateType.Max) tp = null; //Min and Max functions are Func<T,Q> instead of Func<T,int> as sum or avg.
            MethodInfo method = GetAggMethod(MethodName, tp);
            if (method == null) return false;

            MethodInfo generic;
            if (tp == null)
            {
                generic = method.MakeGenericMethod(typeof(T), Fields[colIndex].PropertyType);
            }
            else
            {
                generic = method.MakeGenericMethod(typeof(T));
            }

            var sourceParam = Expression.Parameter(typeof(T), "x");
            var propertyReference = Expression.Property(sourceParam, Fields[colIndex]);

            var selector = Expression.Lambda<Func<T, int>>(propertyReference, sourceParam);

            object r = generic.Invoke(null, new object[] { Data, selector });
            if (r == null) resultValue = null; else resultValue = Convert.ToDouble(r, CultureInfo.InvariantCulture);
            return true;



        }

        private static MethodInfo GetAggMethod(string MethodName, Type tp)
        {
            var x = from methodInfo in typeof(Queryable).GetMethods()
                    where methodInfo.Name == MethodName
                    let parameterInfo = methodInfo.GetParameters()
                    where parameterInfo.Length == 2
                    && parameterInfo[0].ParameterType.GetGenericTypeDefinition() == typeof(IQueryable<>)
                    && parameterInfo[1].ParameterType.GetGenericTypeDefinition() == typeof(Expression<>)
                    && parameterInfo[1].ParameterType.GetGenericArguments().Length == 1
                    && parameterInfo[1].ParameterType.GetGenericArguments()[0].GetGenericArguments().Length == 2
                    && (tp == null || parameterInfo[1].ParameterType.GetGenericArguments()[0].GetGenericArguments()[1] == tp)
                    select
                           methodInfo;

            if (x.Count() == 0) return null;
            return x.Single();
        }

        private static string GetAggMethodName(TAggregateType aggregateType)
        {
            switch (aggregateType)
            {
                case TAggregateType.Sum:
                    return "Sum";

                case TAggregateType.Average:
                    return "Average";

                case TAggregateType.Max:
                    return "Max";

                case TAggregateType.Min:
                    return "Min";
            }

            return null;
        }

        #endregion

        #region Rows and data
        ///<inheritdoc />        
        public override int RowCount
        {
            get
            {
                EnsureCachedCount();

                if (SplitLink != null)
                {
                    int Remaining = CachedCount - SplitLink.ParentDataSource.Position * SplitLink.SplitCount;
                    if (Remaining < SplitLink.SplitCount) return Remaining; else return SplitLink.SplitCount;
                }
                return CachedCount;
            }
        }

        private void EnsureCachedCount()
        {
            if (CachedCount == -1)
            {
                int cidx;
                if (FieldsByName.TryGetValue(ReportTag.StrColWithRowCount, out cidx))
                {
                    object ox = GetValue(cidx);
                    if (ox == null) CachedCount = 0; else CachedCount = Convert.ToInt32(ox);
                }
                else
                {
                    //IQueryable<T>.Count will read all the data and count it, while IEnumerable<T>.Count will use SQL. 
                    StartEnumerator(); //not needed really, but makes sure connection doesn't close.
                    CachedCount = Data.Count();
                }
            }
        }

        /// <summary>
        /// Returns true when at the end of the table.
        /// </summary>
        /// <returns></returns>
        public override bool Eof()
        {
            StartEnumerator();
            return EofReached;
        }

        ///<inheritdoc />
        public override object GetValue(int column)
        {
            StartEnumerator();
            if (FilteredDataCurrent == null) return null; //empty
            return FieldExpressions[column](FilteredDataCurrent);
        }
        #endregion

        #region Move
        ///<inheritdoc />
        public override void MoveFirst()
        {
            if (SplitLink != null && SplitLink.ParentDataSource.Position > 0) return; //don't reset the enum if inside a split.

            if (AtBof) return; //Avoid recreating a valid enumerator.

            //Not all enumerators allow reset, and in most cases movefirst will be called just once at the start.
            //It might be called more than once if you have a master-detail report where the detail is not linked
            //to the master, so detail has to be fully dumped for every master record. But that report doesn't make much sense anyway.
            DisposeEnumerator();
            //we will use lazy loading, so if movefirst is called many times and nothing is done, the sql is not sent to the server.
        }

        ///<inheritdoc />
        public override void MoveNext()
        {
            StartEnumerator();
            FilteredDataCurrent = FilteredDataEnumerator.Current;
            AtBof = false;
            EofReached = NextIsEof;
            if (!EofReached)
            {
                NextIsEof = !FilteredDataEnumerator.MoveNext();
            }
        }
        #endregion

        #region Relationships
        /// <summary>
        /// Returns a detail table. This is used in master-detail relationships to get the nested table for the detail band.
        /// </summary>
        /// <typeparam name="Q"></typeparam>
        /// <param name="Detail">Detail data table calling this method.</param>
        /// <param name="IsEmptyList">Return true if the enumerator has 0 records.</param>
        /// <returns></returns>
        public IEnumerable<Q> GetDetail<Q>(TLinqDataTable<Q> Detail, out bool IsEmptyList)
        {
            IsEmptyList = false;
            StartEnumerator();
            if (EofReached) //empty
            {
                IsEmptyList = true;
                return new List<Q>();
            }
            int i;
            if (!FieldsByName.TryGetValue(Detail.FindOrigTableName(), out i)) return null;
            IEnumerable<Q> en = GetValue(i) as IEnumerable<Q>; //detail classes are not IQueryable
            if (en == null) return null;
            return en;
        }

        ///<inheritdoc />
        public override void MoveMasterRecord(TMasterDetailLink[] masterDetailLinks, TSplitLink splitLink)
        {
            SplitLink = splitLink;
            if (SplitLink != null && SplitLink.ParentDataSource.Position > 0)
            {
                MoveNext();
                return; //don't reset the enum if inside a split.
            }

            if (masterDetailLinks.Length > 0)
            {
                DisposeEnumerator();
                FData = FUnFilteredData; //reset the data.
                CachedCount = -1;

                NeedsSorting = false;
                for (int i = 0; i < masterDetailLinks.Length; i++)
                {
                    if (IsNestedRelationship(masterDetailLinks[i]))
                    {
                        bool IsEmptyList = MdNestedTable(masterDetailLinks[i]);
                        if (IsEmptyList)
                        {
                            break;
                        }
                        else
                        {
                            NeedsSorting = true;
                        }
                    }
                    else
                    {
                        FData = FilterMD(masterDetailLinks[i]);
                        NeedsSorting = true;
                    }

                }

            }

            if (FData == null) FlxMessages.ThrowException(FlxErr.ErrInvalidLinQDetail, TableName);
        }

        private static bool IsNestedRelationship(TMasterDetailLink masterDetailLink)
        {
            return masterDetailLink.ChildFieldName == null;
        }

        private bool MdNestedTable(TMasterDetailLink masterDetailLink)
        {
            ILinqDataTableState MasterState = masterDetailLink.ParentDataSource as ILinqDataTableState;
            if (MasterState == null) FlxMessages.ThrowException(FlxErr.ErrInvalidLinQDetail, TableName);
            bool IsEmptyList = false;
            var FilteredDataEn = MasterState.GetDetail<T>(TableData as TLinqDataTable<T>, out IsEmptyList);
            if (IsEmptyList)
            {
                FData = FilteredDataEn.AsQueryable<T>();
            }
            else
            {
                ConvertDataToIQueryable(FilteredDataEn);
            }
            return IsEmptyList;
        }

        private IQueryable<T> FilterMD(TMasterDetailLink masterDetailLink)
        {
            //FData.Where((x)=> x.CustomerId = masterValue)
            var sourceParam = Expression.Parameter(typeof(T), "x");
            var propertyReference = Expression.Property(sourceParam, masterDetailLink.ChildFieldName);
            int f0 = FieldsByName[masterDetailLink.ChildFieldName];
            Type ObjType = Fields[f0].PropertyType;
            var po = Expression.Convert(propertyReference, ObjType);

            var masterValue = Expression.Constant(masterDetailLink.ParentDataSource.GetValue(masterDetailLink.ParentField), ObjType);
            var comp = Expression.Equal(po, masterValue);
            var whereSelector = Expression.Lambda<Func<T, bool>>(comp, sourceParam);

            return FData.Where(whereSelector);

        }

        private void SortAndFilter(bool IsEmptyList)
        {
            if (!IsEmptyList)
            {
                FilterData();
                SortData(FSortString); //only here FilteredData is guaranteed not to be null.
            }
        }


        private void ConvertDataToIQueryable(IEnumerable<T> FilteredDataEn)
        {
            EnsureDataConverter(FilteredDataEn);
            if (DataConverter != null)
            {
                FData = DataConverter(FilteredDataEn); //We can't apply "where(string)" in a child table.
            }
            else
            {
                FData = FilteredDataEn.AsQueryable<T>();
            }
        }

        private void EnsureDataConverter(object oData)
        {
            if (oData == null) return;
            if (SearchedForDataConverter) return;
            SearchedForDataConverter = true;
            DataConverter = CreateDataConverter(oData.GetType());
        }

        private static Func<IEnumerable<T>, IQueryable<T>> CreateDataConverter(Type oqt)
        {
            Func<IEnumerable<T>, IQueryable<T>> NewDataConverter = null;
            var createSourceQueryMethod = oqt.GetMethod("CreateSourceQuery", new Type[] { });
            if (createSourceQueryMethod != null)
            {
                var sourceQueryParam = Expression.Parameter(typeof(IEnumerable<T>), "x");
                var callQueryMethod = Expression.Call(Expression.Convert(sourceQueryParam, oqt),
                    createSourceQueryMethod);
                var callQueryMethodConv = Expression.Convert(callQueryMethod, typeof(IQueryable<T>));
                var createSourceQueryAction = Expression.Lambda<Func<IEnumerable<T>, IQueryable<T>>>(callQueryMethodConv, sourceQueryParam).Compile();

                NewDataConverter = createSourceQueryAction;
            }
            return NewDataConverter;
        }

        private void FilterData()
        {
            bool Applied;
            EnsureFilterAction(out Applied);
            if (Applied) return;
            foreach (var FilterAction in FilterActions)
            {
                FData = FilterAction(FData);
            }
        }

        private void EnsureFilterAction(out bool Applied)
        {
            Applied = false;
            if (Filters == null || Filters.Length == 0)
            {
                FilterActions = new Func<IQueryable<T>, IQueryable<T>>[0];
                return;
            }

            FilterActions = new Func<IQueryable<T>, IQueryable<T>>[Filters.Length];
            for (int i = 0; i < Filters.Length; i++)
            {
                FilterActions[i] = GetFilterAction(Filters[i]);
                FData = FilterActions[i](FData);
            }
            Applied = true;
        }

        /// <summary>
        /// This method returns a function that can be used to filter the data. This implementation
        /// calls <see cref="SqlFilter"/> when rowFilter starts with "@", or calls <see cref="SimpleFilter"/> when
        /// rowFilter doesn't start with "@". You might want to replace this method
        /// by a different one that filters in other way.
        /// </summary>
        /// <param name="rowFilter">String with the filter as the user wrote it in the report.</param>
        /// <returns></returns>
        public virtual Func<IQueryable<T>, IQueryable<T>> GetFilterAction(string rowFilter)
        {
            if (rowFilter.StartsWith("@")) return SqlFilter(rowFilter.Substring(1));
            return SimpleFilter(rowFilter);
        }

        /// <summary>
        /// This method is called by <see cref="GetFilterAction"/> when the rowFilter starts with "@".
        /// It will do a simple parse of the rowFilter string, allowing "AND" "OR" "()" and equality comparisons.
        /// </summary>
        /// <param name="rowFilter"></param>
        /// <returns></returns>
        public Func<IQueryable<T>, IQueryable<T>> SimpleFilter(string rowFilter)
        {
            var sfp = new TSimpleFilterParser<T>(rowFilter, FieldsByName, TableName);
            return sfp.ToExpression();
        }

        /// <summary>
        /// This method is called by <see cref="GetFilterAction"/> when the rowFilter starts with "@". When overriding
        /// <see cref="GetFilterAction"/> you might want to call this method if rowfilter starts with "@".
        /// <br></br>
        ///  This implementation
        /// tries to find a "Where(string)" method in the data and call it.
        /// </summary>
        /// <param name="rowFilter"></param>
        /// <returns></returns>
        public Func<IQueryable<T>, IQueryable<T>> SqlFilter(string rowFilter)
        {
            Type oqt = FData.GetType();
            Type opt = Type.GetType("System.Data.Objects.ObjectParameter, " + oqt.Assembly.FullName);

            MethodInfo whereMethod = null;

            if (opt != null)
            {
                whereMethod = oqt.GetMethod("Where", new Type[] { typeof(string), opt.MakeArrayType() });
            }

            if (whereMethod == null) //try to find a simple "Where(string)"
            {
                whereMethod = oqt.GetMethod("Where", new Type[] { typeof(string) });
            }

            if (whereMethod == null)
            {
                FlxMessages.ThrowException(FlxErr.ErrDatasetDoesntSupportWhere, TableName);
            }

            object emptyParams = Activator.CreateInstance(opt.MakeArrayType(), 0);

            var sourceParam = Expression.Parameter(typeof(IQueryable<T>), "x");
            var callMethod = Expression.Call(Expression.Convert(sourceParam, oqt),
                whereMethod, Expression.Constant(rowFilter), Expression.Constant(emptyParams));

            return Expression.Lambda<Func<IQueryable<T>, IQueryable<T>>>(callMethod, sourceParam).Compile();
        }

        ///<inheritdoc />
        public override int FilteredRowCount(TMasterDetailLink[] masterDetailLinks)
        {
            EnsureCachedCount();
            return CachedCount;
        }

        #endregion

        #region IDisposable Members

        ///<inheritdoc />
        protected override void Dispose(bool disposing)
        {
            try
            {
                if (disposing)
                {
                    DisposeEnumerator();
                }
            }
            finally
            {
                //last call.
                base.Dispose(disposing);
            }
        }

        private void DisposeEnumerator()
        {
            EofReached = true;
            NextIsEof = true;
            IEnumerator<T> e = FilteredDataEnumerator;
            FilteredDataEnumerator = null;
            if (e != null) e.Dispose();
        }

        private void StartEnumerator()
        {
            if (FilteredDataEnumerator != null) return;
            FilteredDataEnumerator = Data.GetEnumerator();
            EofReached = !FilteredDataEnumerator.MoveNext(); //to the first position.
            if (!EofReached)
            {
                FilteredDataCurrent = FilteredDataEnumerator.Current;
                NextIsEof = !FilteredDataEnumerator.MoveNext();
            }
            else
            {
                NextIsEof = true;
            }

            AtBof = true;
        }


        #endregion

    }
    #endregion

    struct TColumnAndOrder
    {
        internal string Column;
        internal bool Descending;

        public static TColumnAndOrder[] GetSortFields(string SortStr)
        {
            if (SortStr == null) return null;
            string[] f = SortStr.Split(',');
            TColumnAndOrder[] Result = new TColumnAndOrder[f.Length];
            for (int i = 0; i < Result.Length; i++)
            {
                Result[i] = new TColumnAndOrder();
                string s = f[i].Trim();
                int OrdPos = s.LastIndexOf(' ');
                Result[i].Column = s;
                if (OrdPos > 0 && OrdPos < s.Length - 2)
                {
                    string ord = s.Substring(OrdPos + 1).Trim();
                    if (string.Equals(ord, "DESC", StringComparison.InvariantCultureIgnoreCase))
                    {
                        Result[i].Descending = true;
                        Result[i].Column = s.Substring(0, OrdPos).Trim();
                    }
                    else
                        if (string.Equals(ord, "ASC", StringComparison.InvariantCultureIgnoreCase))
                        {
                            Result[i].Column = s.Substring(0, OrdPos).Trim();
                        }
                }
            }

            return Result;
        }

    }

    internal class TSimpleFilterParser<T>
    {
        string Filter;
        int FilterPos;
        string CurrentToken;
        Dictionary<string, int> FieldsByName;
        ParameterExpression sourceParam;
        string TableName;

        public TSimpleFilterParser(string aFilter, Dictionary<string, int> aFieldsByName, string aTableName)
        {
            Filter = aFilter;
            FilterPos = 0;
            FieldsByName = aFieldsByName;
            TableName = aTableName;
        }

        #region Parse
        void NextWord()
        {
            while (!EOF() && Char.IsWhiteSpace(Filter[FilterPos])) FilterPos++;
            if (EOF())
            {
                CurrentToken = null;
                return;
            }

            StringBuilder sb = new StringBuilder();
            sb.Append(Filter[FilterPos]);
            if (IsChar(Filter[FilterPos]))
            {
                FilterPos++;
                while (!EOF() && IsChar(Filter[FilterPos]))
                {
                    sb.Append(Filter[FilterPos]);
                    FilterPos++;
                }
            }
            else
            {
                FilterPos++;
            }

            string s = sb.ToString();
            char c = PeekChar();

            if (
                (s == "<" && (c == '=' || c == '>'))
                || (s == ">" && (c == '='))
                )
            {
                s += c;
                FilterPos++;
            }

            CurrentToken = s;
        }

        private bool IsChar(char p)
        {
            if (Char.IsLetterOrDigit(p) || p == '_' || p == '.') return true;
            return false;
        }

        bool EOF()
        {
            return FilterPos >= Filter.Length;
        }

        void NextChar()
        {
            if (!EOF()) FilterPos++;
        }

        private char PeekChar()
        {
            if (EOF()) return (char)0; else return Filter[FilterPos];
        }
        #endregion

        public Func<IQueryable<T>, IQueryable<T>> ToExpression()
        {
            sourceParam = Expression.Parameter(typeof(T), "x");
            NextWord();
            TFilterExpression expr = Parse();
            if (CurrentToken != null) FlxMessages.ThrowException(FlxErr.ErrUnexpectedId, CurrentToken, Filter);
            if (!EOF()) FlxMessages.ThrowException(FlxErr.ErrFormulaInvalid, Filter);
            Expression<Func<T,bool>> filter = Expression.Lambda<Func<T, bool>>(expr.Expr, sourceParam);
            return (x) => { return x.Where(filter); }; 
        }

        private TFilterExpression Parse()
        {
            return DoAndOr();
        }


        private TFilterExpression DoBinary(Func<TFilterExpression> Next, TTokenAndExpr[] Tokens, bool MoreThanOne)
        {
            TFilterExpression expr = Next();
            if (CurrentToken == null) return expr;
            bool found;

            do
            {
                found = false;
                foreach (var t in Tokens)
                {
                    if (string.Equals(CurrentToken, t.Token, StringComparison.OrdinalIgnoreCase))
                    {
                        NextWord();
                        TFilterExpression expr2 = Next();
                        if (expr.Expr.Type != expr2.Expr.Type) MakeSimilarTypes(ref expr, ref expr2);
                        expr = new TFilterExpression(t.Exp(expr.Expr, expr2.Expr), expr.CustomType);
                        found = true;
                    }
                }
            }
            while (found && MoreThanOne);
            return expr; 
        }

        private void MakeSimilarTypes(ref TFilterExpression expr, ref TFilterExpression expr2)
        {
            if (expr2.CustomType)
            {
                expr2.Expr = Expression.Convert(expr2.Expr, expr.Expr.Type);
                expr2.CustomType = expr.CustomType;
            }
            else
            {
                expr.Expr = Expression.Convert(expr.Expr, expr2.Expr.Type);
                expr.CustomType = expr2.CustomType;
            }
        }

        static TTokenAndExpr[] AndOr = new TTokenAndExpr[]
                {
                    new TTokenAndExpr("AND", Expression.AndAlso),
                    new TTokenAndExpr("OR", Expression.OrElse)
                };

        private TFilterExpression DoAndOr()
        {
            return DoBinary(DoNot, AndOr, true);
        }

        static TTokenAndExpr[] Binary = new TTokenAndExpr[]
                {
                    new TTokenAndExpr("<", Expression.LessThan),
                    new TTokenAndExpr("<=", Expression.LessThanOrEqual),
                    new TTokenAndExpr(">", Expression.GreaterThan),
                    new TTokenAndExpr(">=", Expression.GreaterThanOrEqual),
                    new TTokenAndExpr("<>", Expression.NotEqual),
                    new TTokenAndExpr("=", Expression.Equal)
                };

        private TFilterExpression DoNot()
        {
            TFilterExpression expr;
            if (String.Equals(CurrentToken, "NOT", StringComparison.OrdinalIgnoreCase))
            {
                NextWord();
                expr = DoCompare();
                expr.Expr = Expression.Not(expr.Expr);
            }
            else
            {
                expr = DoCompare();
            }
            return expr;
        }

        private TFilterExpression DoCompare()
        {
            return DoBinary(DoAddSub, Binary, false);
        }

        static TTokenAndExpr[] AddSub = new TTokenAndExpr[]
                {
                    new TTokenAndExpr("+", Expression.Add),
                    new TTokenAndExpr("-", Expression.Subtract)
                };

        private TFilterExpression DoAddSub()
        {
            return DoBinary(DoMultDiv, AddSub, true);
        }

        static TTokenAndExpr[] MultDiv = new TTokenAndExpr[]
                {
                    new TTokenAndExpr("*", Expression.Multiply),
                    new TTokenAndExpr("/", Expression.Divide)
                };

        private TFilterExpression DoMultDiv()
        {
            return DoBinary(DoUnaryPlusMinus, MultDiv, true);
        }

        private TFilterExpression DoUnaryPlusMinus()
        {
            if (CurrentToken == "+" || CurrentToken == "-")
            {
                bool Negate = CurrentToken == "-";
                if (!Char.IsWhiteSpace(PeekChar()))
                {
                    NextWord();
                    TFilterExpression expr = DoParenthesis();
                    if (Negate) expr.Expr = Expression.Negate(expr.Expr);
                    return expr;
                }
            }
            
            return DoParenthesis();
        }

        private TFilterExpression GetExprInsideParenthesis()
        {
            TFilterExpression exp = Parse();
            if (CurrentToken != ")")
            {
                FlxMessages.ThrowException(FlxErr.ErrMissingParen, Filter);
            }
            NextWord();

            return exp;
        }

        private TFilterExpression DoParenthesis()
        {
            if (CurrentToken == "(")
            {
                NextWord();
                return GetExprInsideParenthesis(); 
            }
            return DoToken();
        }

        private TFilterExpression DoToken()
        {
            if (CurrentToken == null) FlxMessages.ThrowException(FlxErr.ErrUnexpectedEof, Filter);
            string s = CurrentToken;

            if (s == "\'")
            {
                return new TFilterExpression(Expression.Constant(ReadStr('\'')), false);
            }

            if (s == "[")
            {
                string FieldName = GetField(']');
                if (!FieldsByName.ContainsKey(FieldName))
                {
                    FlxMessages.ThrowException(FlxErr.ErrColumNotFound, FieldName, TableName);
                }
                return new TFilterExpression(Expression.Property(sourceParam, FieldName), false);
            }
            NextWord();

            TFilterExpression FuncExp;
            if (CurrentToken == "(" && TryParseFunction(s, out FuncExp)) return FuncExp;
            if (CurrentToken == "'" && TryParseDateLiteral(s, false, out FuncExp)) return FuncExp;
            if (s == "{" && TryParseOleDateLiteral(s, out FuncExp)) return FuncExp;

            if (FieldsByName.ContainsKey(s)) return new TFilterExpression(Expression.Property(sourceParam, s), false);

            Int32 i32;
            if (Int32.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out i32))
            {
                return new TFilterExpression(Expression.Constant(i32), true);
            }

            Int64 i64;
            if (Int64.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out i64))
            {
                return new TFilterExpression(Expression.Constant(i64), true);
            }

            double d;
            if (double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out d))
            {
                return new TFilterExpression(Expression.Constant(d), true);
            }

            bool b;
            if (bool.TryParse(s, out b))
            {
                return new TFilterExpression(Expression.Constant(b), true);            
            }

            if (string.Equals(s, "null", StringComparison.InvariantCultureIgnoreCase))
            {
                return new TFilterExpression(Expression.Constant(null), false);
            }

            FlxMessages.ThrowException(FlxErr.ErrUnexpectedId, s, Filter);
            return new TFilterExpression(); //just to compile
        }

        private bool TryParseFunction(string s, out TFilterExpression expr)
        {
            switch (s.ToUpperInvariant())
            {
                case "ISNULL":
                    NextWord();
                    TFilterExpression pexp = GetExprInsideParenthesis();
                    expr = new TFilterExpression(Expression.Equal(pexp.Expr, Expression.Constant(null)), false);
                    return true;

                case "YEAR":
                    expr = ReadDateFunction("Year");
                    return true;
                
                case "MONTH":
                    expr = ReadDateFunction("Month");
                    return true;

                case "DAY":
                    expr = ReadDateFunction("Day");
                    return true;

                case "STREQUALS":
                    NextWord();
                    TFilterExpression str1 = GetOneArg();
                    TFilterExpression str2 = GetOneArg();

                    StringComparison strc = StringComparison.InvariantCultureIgnoreCase;
                    switch (CurrentToken.ToUpperInvariant())
                    {
                        case "IGNORECASE":
                            strc = StringComparison.CurrentCultureIgnoreCase;
                            break;

                        case "SAMECASE":
                            strc = StringComparison.CurrentCulture;
                            break;

                        default: FlxMessages.ThrowException(FlxErr.ErrUnexpectedId, CurrentToken, Filter);
                            break;
                    }
                    NextWord();

                    if (CurrentToken != ")")
                    {
                        FlxMessages.ThrowException(FlxErr.ErrMissingParen, Filter);
                    }

                    //The static String.Equals fails in Entity Framework
                    var strEquals = typeof(string).GetMethod("Equals", new Type[] { typeof(string), typeof(StringComparison) });
                    Expression IsNull1 = Expression.Equal(str1.Expr, Expression.Constant(null));
                    Expression IsNull2 = Expression.Equal(str2.Expr, Expression.Constant(null));
                    Expression ExprEq = Expression.Call(str1.Expr, strEquals, str2.Expr, Expression.Constant(strc));
                    Expression ExprEqNull = Expression.Condition(IsNull1, IsNull2, ExprEq);
                   
                    expr = new TFilterExpression
                    ( ExprEqNull
                        , false);
                    NextWord();
                    return true;

            }

            expr = new TFilterExpression();
            return false;
        }

        private TFilterExpression ReadDateFunction(string method)
        {
            TFilterExpression expr;
            NextWord();
            TFilterExpression yexp = GetExprInsideParenthesis();
            
            var dtYear = typeof(DateTime).GetProperty(method).GetGetMethod();

            Type texp = yexp.Expr.Type;
            if (IsNullable(texp))
            {
                expr = new TFilterExpression(
                    Expression.Condition(Expression.Equal(yexp.Expr, Expression.Constant(null)), Expression.Constant(-1),
                    Expression.Call(Expression.Convert(yexp.Expr, typeof(DateTime)), dtYear)), true);
            }
            else
            {
                expr = new TFilterExpression(Expression.Call(yexp.Expr, dtYear), true);
            }
    
            return expr;
        }

        private bool IsNullable(Type texp)
        {
            return (texp.IsGenericType && texp.GetGenericTypeDefinition().Equals(typeof(Nullable<>)));
        }

        private TFilterExpression GetOneArg()
        {
            TFilterExpression str1 = Parse();
            if (CurrentToken != ",")
            {
                FlxMessages.ThrowException(FlxErr.ErrUnexpectedChar, CurrentToken, FilterPos, Filter);
            }
            NextWord();
            return str1;
        }

        private string GetField(char EndOfField)
        {
            StringBuilder s = new StringBuilder();
            while (!EOF() && PeekChar() != EndOfField)
            {
                s.Append(PeekChar());
                NextChar();
            }

            if (PeekChar() != EndOfField) FlxMessages.ThrowException(FlxErr.ErrMissingParen, Filter);
            NextChar();
            NextWord();

            return s.ToString();
        }

        private bool TryParseOleDateLiteral(string s, out TFilterExpression expr)
        {
            string Id = CurrentToken;
            NextWord();
            if (!TryParseDateLiteral(Id, true, out expr))
            {
                FlxMessages.ThrowException(FlxErr.ErrUnexpectedId, Id, Filter);
            }
            if (CurrentToken != "}")
            {
                FlxMessages.ThrowException(FlxErr.ErrMissingParen, Filter);
            }
            NextWord();
            return true;
        }

        private bool TryParseDateLiteral(string s, bool ole, out TFilterExpression expr)
        {
            int SaveParsePos = FilterPos;
            string dateTimeString = ReadStr('\'');

            int FmtId = ole ? GetOleDateLiteral(s) : GetSql92DateLiteral(s);
            if (FmtId < 0)
            {
                expr = new TFilterExpression();
                FilterPos = SaveParsePos;
                return false;
            }

            string[] Formats = new string[]{"yyyy-MM-dd", "HH:mm:ss", "yyyy-MM-dd HH:mm:ss"};
            string Fmt = Formats[FmtId];


            DateTime dt;
            if (DateTime.TryParseExact(dateTimeString, Fmt, CultureInfo.InvariantCulture, DateTimeStyles.NoCurrentDateDefault, out dt))
            {
                expr = new TFilterExpression(Expression.Constant(dt), true);
                return true;
            }

            FlxMessages.ThrowException(FlxErr.ErrInvalidFilterDateTime, dateTimeString, Filter, Fmt);
            expr = new TFilterExpression();
            return false;
            
        }

        private int GetSql92DateLiteral(string s)
        {
            switch (s.ToUpperInvariant())
            {
                case "DATE":
                    return 0;

                case "TIME":
                    return 1;

                case "TIMESTAMP":
                    return 2;
            }
            return -1;
        }

        private int GetOleDateLiteral(string s)
        {
            switch (s.ToUpperInvariant())
            {
                case "D":
                    return 0;

                case "T":
                    return 1;

                case "TS":
                    return 2;
            }
            return -1;
        }

        private string ReadStr(char strdelim)
        {
            StringBuilder s = new StringBuilder();
            bool StrOk = ReadStrChars(s, strdelim);

            if (!StrOk) FlxMessages.ThrowException(FlxErr.ErrUnterminatedString, Filter);
            NextWord();

            return s.ToString();
        }

        private bool ReadStrChars(StringBuilder s, char strdelim)
        {
            while (!EOF())
            {
                if (PeekChar() == strdelim)
                {
                    NextChar();
                    if (PeekChar() != strdelim) return true;
                }
                s.Append(PeekChar());
                NextChar();
            }

            return false;
        }
    }

    internal struct TFilterExpression
    {
        internal Expression Expr;
        internal bool CustomType;

        internal TFilterExpression(Expression aExpr, bool aCustomType)
        {
            Expr = aExpr;
            CustomType = aCustomType;
        }
    }
    internal struct TTokenAndExpr
    {
        internal string Token;
        internal Func<Expression, Expression, Expression> Exp;

        internal TTokenAndExpr(string aToken, Func<Expression, Expression, Expression> aExp)
        {
            Token = aToken;
            Exp = aExp;
        }
    }
}





#endif
