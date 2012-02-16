using System;
using FlexCel.Core;
using System.Diagnostics;
using System.Globalization;

namespace FlexCel.XlsAdapter
{

    /// <summary>
    /// Cache for fast-finding Sheet names. It has support for finding the next number of the sheet.
    /// For example, if you have 3 sheets on a workbook named "Sheet1, Sheet2, Sheet3", and you insert an empty new one,
    /// it should be named "Sheet4"
    /// </summary>
    internal class TSheetNameList
    {
        protected TCaseInsensitiveHashtableStrInt FList;
        internal TSheetNameList()
        {
            FList = new  TCaseInsensitiveHashtableStrInt();
        }

        #region Generics

        internal int this[string index]
        {
            get
            {
                int Result;
                if (!FList.TryGetValue(index, out Result)) return -1;
                return Result;
            }
        }

        internal void Clear()
        {
            FList.Clear();
        }
        
        internal int Count
        {
            get {return FList.Count;}
        }
        #endregion


        internal void Add(TBoundSheetRecordList BoundSheets, int SheetToInsert, string aName, int SheetPos) //Error if duplicated entry
        {
			if (FList.ContainsKey(aName)) XlsMessages.ThrowException(XlsErr.ErrDuplicatedSheetName, aName);
            FixSheetPos(BoundSheets, SheetToInsert, 1);

            FList.Add(aName, SheetPos);

        }

        internal static string MakeValidSheetName(string aName)
        {
            string Result=aName.Substring(0, Math.Min(31, aName.Length));
            Result = Result.Replace("/","_").Replace("\\","_").Replace("?","_").Replace("[","_").Replace("]","_").Replace("*","_").Replace(":",".");
			if (Result.Length == 0) Result = " ";
            if (Result.EndsWith("\'")) Result = Result.Substring(0, Result.Length - 1) + '_';
			return Result;
        }


        internal string AddUniqueName(TBoundSheetRecordList BoundSheets, int SheetToInsert, string aName, int SheetPos)
        {
			int n = FList.Count + 1;
			string NewName;
			do
			{
				NewName = aName + n.ToString(CultureInfo.InvariantCulture);
				n++;
			} while (FList.ContainsKey(NewName));

            Add(BoundSheets, SheetToInsert, NewName, SheetPos);
            return NewName;
        }

        internal void DeleteSheet(string SheetName, TBoundSheetRecordList BoundSheets, int SheetToDelete)
        {
            FList.Remove(SheetName);
            FixSheetPos(BoundSheets, SheetToDelete + 1, -1);
        }

        private void FixSheetPos(TBoundSheetRecordList BoundSheets, int FirstSheet, int p)
        {
            if (BoundSheets == null) return;
            for (int i = FirstSheet; i < BoundSheets.Count; i++)
            {
                string key = BoundSheets[i].SheetName;
                FList[key] += p;
            }
        }

        internal void Rename(string OldName, string NewName)
        {
            if (OldName==NewName) return;
            Debug.Assert(FList.ContainsKey(OldName));
            int SheetPos = FList[OldName];
            if (FList.ContainsKey(NewName)) XlsMessages.ThrowException(XlsErr.ErrDuplicatedSheetName, NewName);
            FList.Remove(OldName);
            Add(null, 0, NewName, SheetPos);
        }
    }
}
