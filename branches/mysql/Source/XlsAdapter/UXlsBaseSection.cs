using System;
using System.IO;
using FlexCel.Core;

namespace FlexCel.XlsAdapter
{
	/// <summary>
	/// Abstract class for sections on a workbook stream. Children might be sheets, globals, etc.
	/// </summary>
	internal abstract class TBaseSection
	{
        private TBOFRecord FBOF;
        private TEOFRecord FEOF;
        internal TFutureStorage FutureStorage;

		internal TBaseSection()
		{
            FBOF=null;
            FEOF=null;
		}

        internal TBOFRecord sBOF { get { return FBOF;} set {FBOF=value;}}
        internal TEOFRecord sEOF { get { return FEOF;} set {FEOF=value;}}

        internal virtual long TotalSize(TEncryptionData Encryption, bool Repeatable)
        {
            if ((sEOF==null)||(sBOF==null)) XlsMessages.ThrowException(XlsErr.ErrSectionNotLoaded);
            return sEOF.TotalSize()+ sBOF.TotalSize();
        }

        internal virtual long TotalRangeSize(int SheetIndex, TXlsCellRange CellRange, TEncryptionData Encryption, bool Repeatable)
        {
            if ((sEOF==null)||(sBOF==null)) XlsMessages.ThrowException(XlsErr.ErrSectionNotLoaded);
            return sEOF.TotalSize()+ sBOF.TotalSize();
        }

        internal abstract void LoadFromStream(TBaseRecordLoader RecordLoader, TBOFRecord First);
        internal abstract void SaveToStream(IDataStream DataStream, TSaveData SaveData);
        internal abstract void SaveRangeToStream(IDataStream DataStream, TSaveData SaveData, int SheetIndex, TXlsCellRange CellRange);

        internal abstract void SaveToPxl(TPxlStream PxlStream, TPxlSaveData SaveData);

        internal void AddFutureStorage(TFutureStorageRecord R)
        {
            TFutureStorage.Add(ref FutureStorage, R);
        }
    }
}
