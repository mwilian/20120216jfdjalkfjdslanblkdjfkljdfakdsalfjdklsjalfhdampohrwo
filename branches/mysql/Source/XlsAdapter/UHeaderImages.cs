using System;

namespace FlexCel.XlsAdapter
{
	/// <summary>
	/// Drawing Group record for header and footer images. There is only one of them on a xls file, and is on the global section.
	/// </summary>
	internal class THeaderImageGroupRecord: TBaseDrawingGroupRecord
	{
		internal THeaderImageGroupRecord(int aId, byte[] aData): base(aId, aData){}

        internal override void LoadIntoWorkbook(TWorkbookGlobals Globals, TWorkbookLoader WorkbookLoader)
        {
            Globals.HeaderImages.LoadFromStream(WorkbookLoader.RecordLoader, this, false);
        }
	}

	/// <summary>
	/// Drawing record for header and footer images. It appears on the sheets.
	/// </summary>
	internal class THeaderImageRecord: TBaseDrawingRecord
	{
		internal THeaderImageRecord(int aId, byte[] aData): base(aId, aData){}

		internal override void LoadIntoSheet(TSheet ws, int rRow, TBaseRecordLoader RecordLoader, ref TLoaderInfo Loader)
		{
			ws.HeaderImages.LoadFromStream(RecordLoader, ws.FWorkbookGlobals, this, ws.SheetType == FlexCel.Core.TSheetType.Chart);
		}

	}
}
