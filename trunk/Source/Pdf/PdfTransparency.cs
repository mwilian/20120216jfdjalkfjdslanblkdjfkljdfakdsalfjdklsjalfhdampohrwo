using System;
using System.Globalization;
using FlexCel.Core;

namespace FlexCel.Pdf
{
	/// <summary>
	/// A class for storing ca/CA graphics states.
	/// </summary>
	internal class TPdfTransparency: IComparable
	{
		int GStateId;
		int GStateObjId;
        string SMask;
        string BBox;
        int SMaskObjId;
		int Alpha;
		TPdfToken Operator;

        public TPdfTransparency(int aGStateId, int aAlpha, TPdfToken aOperator, string aSMask, string aBBox)
        {
            GStateId = aGStateId;
            Alpha = aAlpha;
            Operator = aOperator;
            SMask = aSMask;
            BBox = aBBox;
        }

		public void Select(TPdfStream DataStream)
		{
			TPdfBaseRecord.WriteLine(DataStream,
				GStateName + " " +
				TPdfTokens.GetString(TPdfToken.Commandgs));
		}

		public void WriteGState(TPdfStream DataStream, TXRefSection XRef)
		{
			TPdfBaseRecord.Write(DataStream, GStateName + " ");
			GStateObjId = XRef.GetNewObject(DataStream);
			TIndirectRecord.CallObj(DataStream, GStateObjId);
		}

        public string GStateName
        {
            get
            {
                return TPdfTokens.GetString(TPdfToken.GStatePrefix) + GStateId.ToString(CultureInfo.InvariantCulture);
            }
        }

		public void WriteGStateObject(TPdfStream DataStream, TXRefSection XRef)
		{
			XRef.SetObjectOffset(GStateObjId, DataStream);
			TIndirectRecord.SaveHeader(DataStream, GStateObjId);
			TDictionaryRecord.BeginDictionary(DataStream);
			TDictionaryRecord.SaveKey(DataStream, TPdfToken.TypeName, TPdfTokens.GetString(TPdfToken.ExtGStateName));
			TDictionaryRecord.SaveKey(DataStream, Operator, PdfConv.CoordsToString(Alpha / 255.0));
			
            if (SMask != null)
            {
                SMaskObjId = XRef.GetNewObject(DataStream);
                TDictionaryRecord.WriteLine(DataStream, TPdfTokens.GetString(TPdfToken.SMaskName));
                TDictionaryRecord.BeginDictionary(DataStream);
                TDictionaryRecord.SaveKey(DataStream, TPdfToken.TypeName, TPdfTokens.GetString(TPdfToken.MaskName));
                TDictionaryRecord.SaveKey(DataStream, TPdfToken.SName, TPdfTokens.GetString(TPdfToken.AlphaName));
                TDictionaryRecord.Write(DataStream, TPdfTokens.GetString(TPdfToken.GName) + " ");
                TIndirectRecord.CallObj(DataStream, SMaskObjId);
                TDictionaryRecord.EndDictionary(DataStream);
            }

            TDictionaryRecord.EndDictionary(DataStream);

			TIndirectRecord.SaveTrailer(DataStream);

            if (SMask != null)
            {
                WriteSMaskObject(DataStream, XRef);
            }

		}

        public void WriteSMaskObject(TPdfStream DataStream, TXRefSection XRef)
        {
            //Actually write stream
            XRef.SetObjectOffset(SMaskObjId, DataStream);
            TIndirectRecord.SaveHeader(DataStream, SMaskObjId);
            TDictionaryRecord.BeginDictionary(DataStream);
            TDictionaryRecord.SaveKey(DataStream, TPdfToken.LengthName, SMask.Length);
            TDictionaryRecord.SaveKey(DataStream, TPdfToken.TypeName, TPdfTokens.GetString(TPdfToken.XObjectName));
            TDictionaryRecord.SaveKey(DataStream, TPdfToken.SubtypeName, TPdfTokens.GetString(TPdfToken.FormName));
            TDictionaryRecord.SaveKey(DataStream, TPdfToken.BBoxName, BBox);
            TDictionaryRecord.WriteLine(DataStream, TPdfTokens.GetString(TPdfToken.GroupName));
            TDictionaryRecord.BeginDictionary(DataStream);
            TDictionaryRecord.SaveKey(DataStream, TPdfToken.SName, TPdfTokens.GetString(TPdfToken.TransparencyName));

            TDictionaryRecord.EndDictionary(DataStream);

            TDictionaryRecord.EndDictionary(DataStream);
            
            TStreamRecord.BeginSave(DataStream);
            TStreamRecord.Write(DataStream, SMask);
            TStreamRecord.EndSave(DataStream);
            TIndirectRecord.SaveTrailer(DataStream);
        }

		#region IComparable Members
		public int CompareTo(object obj)
		{
			TPdfTransparency p2= obj as TPdfTransparency;
			if (p2==null)
				return -1;
			int Result = Alpha.CompareTo(p2.Alpha);
			if (Result != 0) return Result;
			
            Result = Operator.CompareTo(p2.Operator);
            if (Result != 0) return Result;

            Result = String.Compare(SMask, p2.SMask, StringComparison.InvariantCulture);
            if (Result != 0) return Result;

            return 0;

		}
		#endregion


	}
}
