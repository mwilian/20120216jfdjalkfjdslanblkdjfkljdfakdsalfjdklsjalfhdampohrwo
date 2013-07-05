#if (FRAMEWORK30)
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using FlexCel.Core;
using System.Xml;
using System.Globalization;

namespace FlexCel.XlsAdapter
{
    #region Pivot
    class TXlsxPivotList<K, T> where T: TXlsxPivot<K> where K: TXlsxPivotRecord
    {
        private List<T> FList;

        internal TXlsxPivotList()
        {
            FList = new List<T>();
        }

        internal List<T> List { get { return FList; } }

        internal void Clear()
        {
            FList.Clear();
        }


        internal void ArrangeInsertRange(TXlsCellRange CellRange, int aRowCount, int aColCount, TSheetInfo SheetInfo)
        {
            for (int i = 0; i < List.Count; i++)
            {
                T it = List[i];
                if (it != null) it.ArrangeInsertRange(CellRange, aRowCount, aColCount, SheetInfo);
            }
        }

        internal void ArrangeMoveRange(TXlsCellRange CellRange, int NewRow, int NewCol, TSheetInfo SheetInfo)
        {
            for (int i = 0; i < List.Count; i++)
            {
                T it = List[i];
                if (it != null) it.ArrangeMoveRange(CellRange, NewRow, NewCol, SheetInfo);
            }
        }

        internal virtual void CopyTo(TXlsxPivotList<K, T> Target, TWorkbookGlobals Globals)
        {
            for (int i = 0; i < FList.Count; i++)
            {
               // Target.FList.Add(FList[i].Clone());
            }
        }

        internal void UpdateDeletedRanges(int SheetIndex, int SheetCount, TDeletedRanges DeletedRanges)
        {
            for (int i = 0; i < List.Count; i++)
            {
                T it = List[i];
                if (it != null) it.UpdateDeletedRanges(SheetIndex, SheetCount, DeletedRanges);
            }
        }

    }

    abstract class TXlsxPivot<T> where T: TXlsxPivotRecord
    {
        internal abstract long CacheId { get; }
        private TXlsxAttribute[] Attributes;

        protected List<T> FRecords;

        protected TXlsxPivot()
        {
            FRecords = new List<T>();
        }


        protected void LoadAtts(TOpenXmlReader DataStream, bool AddInvalid)
        {
            List<TXlsxAttribute> TmpAtt = new List<TXlsxAttribute>();
            if (AddInvalid) TmpAtt.Add(new TXlsxAttribute("", "invalid", "true")); //as we won't save id, this is the first one.

            TXlsxAttribute[] Atts = DataStream.GetAttributes();
            foreach (TXlsxAttribute att in Atts)
            {
                if (att.Name != "id" && att.Name != "invalid" && att.Name != "refreshOnLoad" && att.Name != "upgradeOnRefresh")
                {
                    TmpAtt.Add(att);
                }
            }

            Attributes = TmpAtt.ToArray();
        }

        internal void SaveToXlsx(TOpenXmlWriter DataStream, string DocName, bool IsCache)
        {
            if (Attributes == null) return;
            DataStream.WriteStartDocument(DocName, true);
            foreach (TXlsxAttribute att in Attributes)
            {
                if (att.Namespace != "http://www.w3.org/2000/xmlns/")
                {
                    if (att.Name == "cacheId") DataStream.WriteAtt("cacheId", CacheId);
                    else DataStream.WriteAttRaw(att.Namespace, att.Name, att.Value);
                }
            }

            if (IsCache)
            {
                DataStream.WriteAtt("refreshOnLoad", true);
                DataStream.WriteAtt("upgradeOnRefresh", true);
            }
            foreach (TXlsxPivotRecord rec in FRecords)
            {
                rec.SaveToXlsx(DataStream);
            }
            DataStream.WriteEndDocument();
        }


        internal void ArrangeInsertRange(TXlsCellRange CellRange, int aRowCount, int aColCount, TSheetInfo SheetInfo)
        {
            for (int i = 0; i < FRecords.Count; i++)
            {
                T it = FRecords[i];
                if (it != null) it.ArrangeInsertRange(CellRange, aRowCount, aColCount, SheetInfo);
            }
        }

        internal void ArrangeMoveRange(TXlsCellRange CellRange, int NewRow, int NewCol, TSheetInfo SheetInfo)
        {
            for (int i = 0; i < FRecords.Count; i++)
            {
                T it = FRecords[i];
                if (it != null) it.ArrangeMoveRange(CellRange, NewRow, NewCol, SheetInfo);
            }
        }

        internal void UpdateDeletedRanges(int SheetIndex, int SheetCount, TDeletedRanges DeletedRanges)
        {
            for (int i = 0; i < FRecords.Count; i++)
            {
                T it = FRecords[i];
                if (it != null) it.UpdateDeletedRanges(SheetIndex, SheetCount, DeletedRanges);
            }
        }
    }

    class TXlsxPivotRecord : TFutureStorageRecord
    {
        internal TXlsxPivotRecord(string aXml)
            : base(aXml)
        {
        }

        internal virtual void SaveToXlsx(TOpenXmlWriter DataStream)
        {
            DataStream.WriteRaw(Xml);
        }


        internal virtual void ArrangeInsertRange(TXlsCellRange CellRange, int aRowCount, int aColCount, TSheetInfo SheetInfo)
        {
        }

        internal virtual void ArrangeMoveRange(TXlsCellRange CellRange, int NewRow, int NewCol, TSheetInfo SheetInfo)
        {
        }

        internal virtual void UpdateDeletedRanges(int SheetIndex, int SheetCount, TDeletedRanges DeletedRanges)
        {
        }

    }


    #endregion

    #region PivotCache
    class TXlsxPivotCacheList : TXlsxPivotList<TXlsxPivotCacheRecord, TXlsxPivotCache>
    {
        private Dictionary<long, TXlsxPivotCache> FCacheIdsOnDisk;

        internal TXlsxPivotCacheList(): base()
        {
            FCacheIdsOnDisk = new Dictionary<long, TXlsxPivotCache>(); 
        }

        internal void Load(TOpenXmlReader DataStream)
        {
            TXlsxPivotCache Pivot = TXlsxPivotCache.Load(DataStream);
            List.Add(Pivot);
            FCacheIdsOnDisk[Pivot.CacheId] = Pivot;

        }

        internal void RemoveUnusedCaches()
        {
            /* No need to clean the flags, as we will cleant them at the end, so they always start clean.
            foreach (TXlsxPivotCache pc in PivotCacheList.List)
            {
                pc.HasRefs = false;

            }*/

            for (int i = List.Count - 1; i >= 0; i--)
            {
                if (List[i].HasRefs)
                {
                    List[i].HasRefs = false; //Keep flag always false.
                }
                else
                {
                    List.RemoveAt(i);
                }
            }
        }

        internal TXlsxPivotCache FindCacheFromDisk(long cid)
        {
            TXlsxPivotCache Result;
            if (!FCacheIdsOnDisk.TryGetValue(cid, out Result)) return null;
            return Result;
        }

    }

    class TXlsxPivotCache: TXlsxPivot<TXlsxPivotCacheRecord>
    {
        internal Uri LastSavedUri;
        private long FCacheId;
        public bool HasRefs;

        internal static TXlsxPivotCache Load(TOpenXmlReader DataStream)
        {
            TXlsxPivotCache Result = new TXlsxPivotCache();
            Result.FCacheId = DataStream.GetAttributeAsLong("cacheId", -1);
            Result.ReadPivotCache(DataStream, DataStream.GetRelationship("id"));
            DataStream.FinishTagAndIgnoreChildren();
            return Result;
        }

        internal override long CacheId
        {
            get { return FCacheId; }
        }

        internal void ReadPivotCache(TOpenXmlReader DataStream, string relId)
        {
            DataStream.SelectFromCurrentPartAndPush(relId, TOpenXmlManager.MainNamespace, false);

            while (DataStream.NextTag())
            {
                switch (DataStream.RecordName())
                {
                    case "pivotCacheDefinition":
                        ReadPivotCacheDefinition(DataStream);
                        break;

                    default: //shouldn't happen
                        DataStream.GetXml();
                        break;
                }
            }

            DataStream.PopPart();
        }

        private void ReadPivotCacheDefinition(TOpenXmlReader DataStream)
        {
            LoadAtts(DataStream, true);
            /* Fields that need updating:
             * cacheFields->formula
             * cacheSource.consolidation.rangeSets.rangeSet->ref
             * cacheSource.worksheetSource->ref
             * calculatedItems doesn't need it because formulas don't have cell refs.
             */

            if (DataStream.IsSimpleTag) { DataStream.NextTag(); return; }
            string StartElement = DataStream.RecordName();
            if (!DataStream.NextTag()) return;

            while (!DataStream.AtEndElement(StartElement))
            {
                switch (DataStream.RecordName())
                {
                    case "cacheSource":
                        ReadCacheSource(DataStream);
                        break;
                    case "cacheFields":
                        //break;
                    case "cacheHierarchies":
                    case "kpis":
                    case "tupleCache":
                    case "calculatedItems":
                    case "calculatedMembers":
                    case "dimensions":
                    case "measureGroups":
                    case "maps":
                    case "extLst":
                    default:
                        FRecords.Add(new TXlsxPivotCacheRecord(DataStream.GetXml()));
                        break;
                }
            }   
        }

        private void ReadCacheSource(TOpenXmlReader DataStream)
        {
            string SourceType = DataStream.GetAttribute("type");
            if (SourceType == "external")
            {
                FRecords.Add(new TXlsxPivotCacheRecord(DataStream.GetXml()));
                return;
            }

            if (DataStream.IsSimpleTag) { DataStream.NextTag(); return; }
            string StartElement = DataStream.RecordName();
            if (!DataStream.NextTag()) return;

            TCacheSourceRecord R = new TCacheSourceRecord(SourceType, DataStream.GetAttributeAsLong("connectionId", -1));
            FRecords.Add(R);

            while (!DataStream.AtEndElement(StartElement))
            {
                switch (DataStream.RecordName())
                {
                    case "consolidation":
                        ReadConsolidation(DataStream, R);
                        break;
                    case "worksheetSource":
                        ReadWorksheetSource(DataStream, R);
                        break;

                    case "extLst":
                    default:
                        TFutureStorage.Add(ref R.FutureStorage, new TFutureStorageRecord(DataStream.GetXml()));
                        break;
                }
            }   
  
        }

        private void ReadConsolidation(TOpenXmlReader DataStream, TCacheSourceRecord R)
        {
            R.Consolidation = new TConsolidation();
            R.Consolidation.LoadXlsx(DataStream);
        }

        private void ReadWorksheetSource(TOpenXmlReader DataStream, TCacheSourceRecord R)
        {
            R.WorksheetSource = new TWorksheetSource();
            R.WorksheetSource.LoadXlsx(DataStream);
        }
    }

    class TXlsxPivotCacheRecord : TXlsxPivotRecord
    {
        internal TXlsxPivotCacheRecord(string aXml) : base(aXml)
        {        
        }
    }

    class TCacheSourceRecord : TXlsxPivotCacheRecord
    {
        string SourceType;
        long ConnectionId;
        internal TWorksheetSource WorksheetSource;
        internal TConsolidation Consolidation;
        internal TFutureStorage FutureStorage;

        public TCacheSourceRecord(string aSourceType, long aConnectionId)
            : base(null)
        {
            SourceType = aSourceType;
            ConnectionId = aConnectionId;
        }

        internal override void SaveToXlsx(TOpenXmlWriter DataStream)
        {
            DataStream.WriteStartElement("cacheSource");
            DataStream.WriteAtt("type", SourceType);
            if (ConnectionId > 0) DataStream.WriteAtt("connectionId", ConnectionId, 0);
            if (WorksheetSource != null) WorksheetSource.SaveToXlsx(DataStream);
            if (Consolidation != null) Consolidation.SaveToXlsx(DataStream);
            DataStream.WriteFutureStorage(FutureStorage);
            DataStream.WriteEndElement();
        }

        internal override void ArrangeInsertRange(TXlsCellRange CellRange, int aRowCount, int aColCount, TSheetInfo SheetInfo)
        {
            if (WorksheetSource != null) WorksheetSource.ArrangeInsertRange(CellRange, aRowCount, aColCount, SheetInfo);
            if (Consolidation != null) Consolidation.ArrangeInsertRange(CellRange, aRowCount, aColCount, SheetInfo);
        }

        internal override void ArrangeMoveRange(TXlsCellRange CellRange, int NewRow, int NewCol, TSheetInfo SheetInfo)
        {
            if (WorksheetSource != null) WorksheetSource.ArrangeMoveRange(CellRange, NewRow, NewCol, SheetInfo);
            if (Consolidation != null) Consolidation.ArrangeMoveRange(CellRange, NewRow, NewCol, SheetInfo);
        }

        internal override void UpdateDeletedRanges(int SheetIndex, int SheetCount, TDeletedRanges DeletedRanges)
        {
            if (WorksheetSource != null) WorksheetSource.UpdateDeletedRanges(SheetIndex, SheetCount, DeletedRanges);
        }
    }

    class TWorksheetSource
    {
        TXlsCellRange Refe;
        string Name;
        string Sheet;

        public TWorksheetSource()
        {
        }


        internal void SaveToXlsx(TOpenXmlWriter DataStream)
        {
            DataStream.WriteStartElement("worksheetSource");
            SaveAttsToXlsx(DataStream);
            DataStream.WriteEndElement();
        }

        internal void SaveAttsToXlsx(TOpenXmlWriter DataStream)
        {
            DataStream.WriteAttAsRange("ref", Refe, true);
            DataStream.WriteAtt("name", Name);
            DataStream.WriteAtt("sheet", Sheet);
        }

        internal void LoadXlsx(TOpenXmlReader DataStream)
        {
            LoadAttsFromXlsx(DataStream);
            DataStream.FinishTag();
        }

        internal void LoadAttsFromXlsx(TOpenXmlReader DataStream)
        {
            Refe = DataStream.GetAttributeAsRange("ref", true);
            Name = DataStream.GetAttribute("name");
            Sheet = DataStream.GetAttribute("sheet");
        }


        internal void ArrangeInsertRange(TXlsCellRange CellRange, int aRowCount, int aColCount, TSheetInfo SheetInfo)
        {
            if (Refe == null) return;
            if (!String.Equals(SheetInfo.SourceGlobals.GetSheetName(SheetInfo.InsSheet), Sheet, StringComparison.InvariantCultureIgnoreCase)) return;

            TRangeManipulator.ArrangeInsertRange(Refe, CellRange, aRowCount, aColCount);
        }

        internal void ArrangeMoveRange(TXlsCellRange CellRange, int NewRow, int NewCol, TSheetInfo SheetInfo)
        {
            if (Refe == null) return;
            if (!String.Equals(SheetInfo.SourceGlobals.GetSheetName(SheetInfo.InsSheet), Sheet, StringComparison.InvariantCultureIgnoreCase)) return;
            TRangeManipulator.ArrangeMoveRange(Refe, CellRange, NewRow, NewCol);
        }

        internal void UpdateDeletedRanges(int SheetIndex, int SheetCount, TDeletedRanges DeletedRanges)
        {
            if (Name == null) return;
            int sheet = -1;
            int NameIndex = DeletedRanges.Names.GetNamePos(sheet, Name);
            if (NameIndex < 0) return;
            if (DeletedRanges.Update)
            {
                //Nothing to do here, as we save the name string, not position
            }
            else
            {
                TNameRecord NameRec = DeletedRanges.Names[NameIndex];
                DeletedRanges.Reference(NameIndex);
                NameRec.UpdateDeletedRanges(DeletedRanges);
            }

        }

    }

    class TConsolidation
    {
        string Pages;
        List<TRangeSet> RangeSets;
        bool AutoPage;

        public TConsolidation()
        {
            RangeSets = new List<TRangeSet>();
        }

        internal void SaveToXlsx(TOpenXmlWriter DataStream)
        {
            DataStream.WriteStartElement("consolidation");
            DataStream.WriteAtt("autoPage", AutoPage);
            if (!String.IsNullOrEmpty(Pages))
            {
                DataStream.WriteRaw(Pages);
            }

            if (RangeSets.Count > 0)
            {
                DataStream.WriteStartElement("rangeSets");
                DataStream.WriteAtt("count", RangeSets.Count);
                foreach (TRangeSet r in RangeSets)
                {
                    r.SaveToXlsx(DataStream);
                }
                DataStream.WriteEndElement();
            }
            DataStream.WriteEndElement();
        }

        internal void LoadXlsx(TOpenXmlReader DataStream)
        {
            AutoPage = DataStream.GetAttributeAsBool("autoPage", false);

            if (DataStream.IsSimpleTag) { DataStream.NextTag(); return; }
            string StartElement = DataStream.RecordName();
            if (!DataStream.NextTag()) return;

            while (!DataStream.AtEndElement(StartElement))
            {
                switch (DataStream.RecordName())
                {
                    case "pages":
                        Pages = DataStream.GetXml();
                        break;
                    case "rangeSets":
                        ReadRangeSets(DataStream);
                        break;

                    default:
                        DataStream.GetXml(); //Shouldn't happen.
                        break;
                }
            }      
        }

        private void ReadRangeSets(TOpenXmlReader DataStream)
        {
            if (DataStream.IsSimpleTag) { DataStream.NextTag(); return; }
            string StartElement = DataStream.RecordName();
            if (!DataStream.NextTag()) return;

            while (!DataStream.AtEndElement(StartElement))
            {
                switch (DataStream.RecordName())
                {
                    case "rangeSet":
                        RangeSets.Add(TRangeSet.LoadFromXlsx(DataStream));
                        break;

                    default:
                        DataStream.GetXml(); //Shouldn't happen.
                        break;
                }
            }
        }

        internal void ArrangeInsertRange(TXlsCellRange CellRange, int aRowCount, int aColCount, TSheetInfo SheetInfo)
        {
            foreach (TRangeSet rs in RangeSets)
            {
                rs.ArrangeInsertRange(CellRange, aRowCount, aColCount, SheetInfo);
            }
        }

        internal void ArrangeMoveRange(TXlsCellRange CellRange, int NewRow, int NewCol, TSheetInfo SheetInfo)
        {
            foreach (TRangeSet rs in RangeSets)
            {
                rs.ArrangeMoveRange(CellRange, NewRow, NewCol, SheetInfo);
            }
        }
    }

    class TRangeSet
    {
        int[] ix;
        TWorksheetSource Source;

        public TRangeSet()
        {
            ix = new int[4];
            for (int i = 0; i < ix.Length; i++)
            {
                ix[i] = -1;
            }
        }

        internal void SaveToXlsx(TOpenXmlWriter DataStream)
        {
            DataStream.WriteStartElement("rangeSet");
            for (int i = 0; i < ix.Length; i++)
            {
                if (ix[i] >= 0) DataStream.WriteAtt("i" + (i + 1).ToString(CultureInfo.InvariantCulture), ix[i]);
            }
            if (Source != null) Source.SaveAttsToXlsx(DataStream);
            DataStream.WriteEndElement();
        }

        internal static TRangeSet LoadFromXlsx(TOpenXmlReader DataStream)
        {
            TRangeSet Result = new TRangeSet();
            for (int i = 0; i < Result.ix.Length; i++)
            {
                Result.ix[i] = DataStream.GetAttributeAsInt("i" + (i + 1).ToString(CultureInfo.InvariantCulture), -1);
            }
            Result.Source = new TWorksheetSource();
            Result.Source.LoadAttsFromXlsx(DataStream);
            DataStream.FinishTag();
            return Result;
        }

        internal void ArrangeInsertRange(TXlsCellRange CellRange, int aRowCount, int aColCount, TSheetInfo SheetInfo)
        {
            if (Source != null) Source.ArrangeInsertRange(CellRange, aRowCount, aColCount, SheetInfo);

        }

        internal void ArrangeMoveRange(TXlsCellRange CellRange, int NewRow, int NewCol, TSheetInfo SheetInfo)
        {
            if (Source != null) Source.ArrangeMoveRange(CellRange, NewRow, NewCol, SheetInfo);
        }
    }
    #endregion

    #region PivotTable
    class TXlsxPivotTableList : TXlsxPivotList<TXlsxPivotTableRecord, TXlsxPivotTable>
    {
        internal TXlsxPivotTable Add()
        {
            List.Add(new TXlsxPivotTable());
            return List[List.Count - 1];
        }

        internal void MarkUsedCacheIds()
        {
            for (int i = List.Count - 1; i >= 0; i--)
            {
                if (List[i].Cache == null) List.RemoveAt(i); else List[i].Cache.HasRefs = true;
            }
         }

    }

    class TXlsxPivotTable : TXlsxPivot<TXlsxPivotTableRecord>
    {
        private TXlsxPivotCache FCache;

        internal TXlsxPivotCache Cache
        {
            get { return FCache; }
        }

        internal override long CacheId
        {
            get { if (FCache == null) return -1; else return FCache.CacheId; }
        }

        internal void ReadPivotTable(TOpenXmlReader DataStream, TXlsxPivotCacheList PivotCacheList)
        {
            /* Fields that need updating:
             * location
             * filter
             */
            LoadAtts(DataStream, false);
            long cid = DataStream.GetAttributeAsLong("cacheId", -1);
            if (cid >= 0) FCache = PivotCacheList.FindCacheFromDisk(cid);

            if (DataStream.IsSimpleTag) { DataStream.NextTag(); return; }
            string StartElement = DataStream.RecordName();
            if (!DataStream.NextTag()) return;

            while (!DataStream.AtEndElement(StartElement))
            {
                switch (DataStream.RecordName())
                {
                    case "location":
                        ReadLocation(DataStream);
                        break;

                    case "pivotFields":
                    case "rowFields":
                    case "rowItems":
                    case "colFields":
                    case "colItems":
                    case "pageFields":
                    case "dataFields":
                    case "formats":
                    case "conditionalFormats":
                    case "chartFormats":
                    case "pivotHierarchies":
                    case "pivotTableStyleInfo":
                    case "filters":
                    case "rowHierarchiesUsage":
                    case "colHierarchiesUsage":
                    case "extLst":
                    default:
                        FRecords.Add(new TXlsxPivotTableRecord(DataStream.GetXml()));
                        break;
                }
            }

        }

        private void ReadLocation(TOpenXmlReader DataStream)
        {
            FRecords.Add(TXlsxPivotLocationRecord.Load(DataStream));
        }

    }

    class TXlsxPivotTableRecord:  TXlsxPivotRecord
    {
        
        internal TXlsxPivotTableRecord(string aXml) : base(aXml)
        {        
        }
    }

    class TXlsxPivotLocationRecord: TXlsxPivotTableRecord
    {
        TXlsCellRange Refe;
        int firstHeaderRow;
        int firstDataRow;
        int firstDataCol;
        int rowPageCount;
        int colPageCount;

        private TXlsxPivotLocationRecord()
            : base(null)
        { 
        }

        internal static TXlsxPivotTableRecord Load(TOpenXmlReader DataStream)
        {
            TXlsxPivotLocationRecord Result = new TXlsxPivotLocationRecord();
            Result.Refe = DataStream.GetAttributeAsRange("ref", true);
            Result.firstHeaderRow = DataStream.GetAttributeAsInt("firstHeaderRow", 0);
            Result.firstDataRow = DataStream.GetAttributeAsInt("firstDataRow", 0);
            Result.firstDataCol = DataStream.GetAttributeAsInt("firstDataCol", 0);
            Result.rowPageCount = DataStream.GetAttributeAsInt("rowPageCount", 0);
            Result.colPageCount = DataStream.GetAttributeAsInt("colPageCount", 0);
            DataStream.FinishTag();
            return Result;
        }

        internal override void SaveToXlsx(TOpenXmlWriter DataStream)
        {
            DataStream.WriteStartElement("location");
            DataStream.WriteAttAsRange("ref", Refe, true);
            DataStream.WriteAtt("firstHeaderRow", firstHeaderRow);
            DataStream.WriteAtt("firstDataRow", firstDataRow);
            DataStream.WriteAtt("firstDataCol", firstDataCol);
            DataStream.WriteAtt("rowPageCount", rowPageCount, 0);
            DataStream.WriteAtt("colPageCount", colPageCount, 0);

            DataStream.WriteEndElement();
        }

        internal override void ArrangeInsertRange(TXlsCellRange CellRange, int aRowCount, int aColCount, TSheetInfo SheetInfo)
        {
            if (SheetInfo.InsSheet != SheetInfo.DestFormulaSheet) return;
            TRangeManipulator.ArrangeInsertRange(Refe, CellRange, aRowCount, aColCount);
        }

        internal override void ArrangeMoveRange(TXlsCellRange CellRange, int NewRow, int NewCol, TSheetInfo SheetInfo)
        {
            if (SheetInfo.InsSheet != SheetInfo.DestFormulaSheet) return;
            TRangeManipulator.ArrangeMoveRange(Refe, CellRange, NewRow, NewCol);
        }
    }

    #endregion
}
#endif


