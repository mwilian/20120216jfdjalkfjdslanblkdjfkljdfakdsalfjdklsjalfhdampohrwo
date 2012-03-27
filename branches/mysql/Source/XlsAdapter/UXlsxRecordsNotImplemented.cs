using System;
using System.IO;
using System.Text;

using FlexCel.Core;
using System.Globalization;

namespace FlexCel.XlsAdapter
{
	internal enum TXlsxSection
	{
		Workbook_Root,
		Workbook_Workbook,
		Workbook_Workbook_Sheets,
		Workbook_Workbook_Names,

		SST_Root,
		SST_SST,

		Styles_Root,
		Styles_Styles,

		Worksheet
	}

	internal class TXlsxBaseRecord : TBaseRecord
	{
		internal override int GetId
		{
			get
			{
				return -1;
			}
		}

		protected override TBaseRecord DoCopyTo(TSheetInfo SheetInfo)
		{
			return null;
		}
		internal override void SaveToStream(IDataStream Workbook, TSaveData SaveData, int Row)
		{

		}

		internal override int TotalSize()
		{
			return 0;
		}

		internal override int TotalSizeNoHeaders()
		{
			return 0;
		}


	}

	internal class TxXlsxBaseRecord : TXlsxBaseRecord
	{
		internal string Xml;
		internal TXlsxSection Section;

		internal TxXlsxBaseRecord(string aXml, TXlsxSection aSection)
		{
			Xml = aXml;
			Section = aSection;
		}
	}

	internal class TFutureStorageRecord : TXlsxBaseRecord
	{
		internal TFutureStorageRecord(string aXml)
		{
		}
	}

	internal class TFutureStorage
	{
		internal static void Add(ref TFutureStorage Fs, TFutureStorageRecord R)
		{
		}

		internal TFutureStorage Clone()
		{
			return this;
		}

        internal static TFutureStorage Clone(TFutureStorage a)
        {
            return a;
        }

#if (COMPACTFRAMEWORK && !FRAMEWORK20)
		public static bool Equals(Object o1, Object o2)
		{
			return (o1 != null && o1.Equals(o2)) || (o1 == null && o2 == null);
		}
#endif
    }



	internal class TRichTextRun
	{
	}

	internal class TxSSTRecord : TXlsxBaseRecord
	{
		internal string Text = null;
		internal TRTFRun[] RTFRuns = null;
	}
}
