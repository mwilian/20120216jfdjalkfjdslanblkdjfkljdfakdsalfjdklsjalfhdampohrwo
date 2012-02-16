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
    /// <summary>
    /// An specialization of <see cref="TLinqDataTable{T}"/> that we will use as default implementation.
    /// </summary>
    internal class TEFDataTable<T>: TLinqDataTable<T>
    {
         public TEFDataTable(string aTableName, VirtualDataTable aCreatedBy, IQueryable<T> aData)
            : base(aTableName, aCreatedBy, aData)
         {
         }

         public TEFDataTable(string aTableName, VirtualDataTable aCreatedBy, IQueryable<T> aData, TLinqDataTable<T> aOrigTable, string[] aFilters) :
            base(aTableName, aCreatedBy, aData, aOrigTable, aFilters)
        {
        }

         protected override VirtualDataTable CreateNewDataTable(string aTableName, VirtualDataTable aCreatedBy, IQueryable<T> aData, TLinqDataTable<T> aOrigTable, string[] aFilters)
        {
            return new TEFDataTable<T>(aTableName, aCreatedBy, aData, aOrigTable, aFilters);
        }

        public override VirtualDataTableState CreateState(string sort, TMasterDetailLink[] masterDetailLinks, TSplitLink splitLink)
        {
            return new TEFDataTableState<T>(this, Data, Fields, FieldExpressions, FieldsByName, Filters, sort, masterDetailLinks, splitLink);
        }
    }

    internal class TEFDataTableState<T> : TLinqDataTableState<T>
    {
        public TEFDataTableState(VirtualDataTable aTableData, IQueryable<T> aData, PropertyInfo[] aFields,
    TFieldExpressions<T> aFieldExpressions, Dictionary<string, int> aFieldsByName,
    string[] aFilters,
    string sort, TMasterDetailLink[] masterDetailLinks, TSplitLink splitLink)
            : base(aTableData, aData, aFields, aFieldExpressions, aFieldsByName, aFilters, sort, masterDetailLinks, splitLink)
        { }


    }

}
#endif
