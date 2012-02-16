using System;
using System.Text;
using System.Globalization;
using FlexCel.Core;
using System.Collections.Generic;

#if (WPF)
using RectangleF = System.Windows.Rect;
#else
using System.Drawing;
using System.IO;
#endif


namespace FlexCel.Pdf
{
	/// <summary>
	/// Represents an interactive form (AcroForm).
	/// </summary>
	internal class TAcroFormRecord: TDictionaryRecord
	{
		private int FId;
		private int SigFlags;
		private TAcroFormFieldList FieldList;
		private TPdfSignature Signature;

		public TAcroFormRecord(TPdfSignature aSignature)
		{
			FId = -1;
			SigFlags  = 0;
			FieldList = new TAcroFormFieldList();
			Signature = aSignature;
			if (Signature != null) 
			{
				FieldList.Add(new TAcroSigField(Signature));
				SigFlags = 3;
			}
		}

        internal int GetId(TPdfStream DataStream, TXRefSection XRef)
        {
			if (FieldList.Count <= 0) return -1; //We don't have any forms.
            if (FId < 0) FId = XRef.GetNewObject(DataStream); ;
            return FId;
        }

		internal void SaveToStream(TPdfStream DataStream, TXRefSection XRef, TPageInfo ParentPage)	
		{
			if (FieldList.Count <= 0) return;  //No AcroForm in this file.


			XRef.SetObjectOffset(GetId(DataStream, XRef), DataStream);
			TIndirectRecord.SaveHeader(DataStream, FId);
			BeginDictionary(DataStream);
			SaveKey(DataStream, TPdfToken.FieldsName, FieldList.GetCallArray(DataStream, XRef, false));

			SaveKey(DataStream, TPdfToken.SigFlagsName, SigFlags);
			EndDictionary(DataStream);
			TIndirectRecord.SaveTrailer(DataStream);

			FieldList.SaveToStream(DataStream, XRef, ParentPage);

		}

		internal static int GetSignaturePage(TPdfSignature Signature, int PageCount)
		{
			int Result = 0;

			TPdfVisibleSignature VSig = Signature as TPdfVisibleSignature;

			if (VSig != null)
			{
				Result = VSig.Page - 1; 
			}

			if (Result < 0 || Result >= PageCount) Result = PageCount - 1;

			return Result;
		}

		internal string GetSignatureFieldsCallArray(TPdfStream DataStream, TXRefSection XRef, int PageNumber, int PageCount)
		{
			if (GetSignaturePage(Signature, PageCount) != PageNumber) return String.Empty; // Annotation does not belong to this page.

			return FieldList.GetCallArray(DataStream, XRef, true);
		}

        internal bool HasSignatureFields(int PageNumber, int PageCount)
        {
            int SigPage = GetSignaturePage(Signature, PageCount);
            if (SigPage != PageNumber) return false;

            return FieldList.HasSigFields();
        }

        internal int SignatureFieldId(TPdfStream DataStream, TXRefSection XRef)
        {
            return FieldList.SigFieldId(DataStream, XRef);
        }

		internal void FinishSign(TPdfStream DataStream)
		{
			FieldList.FinishSign(DataStream);
		}

	}

	#region Field List
	/// <summary>
	/// Holds a list of fields in one form.
	/// </summary>
	internal class TAcroFormFieldList
	{
		private List<TAcroField> Fields;

		internal TAcroFormFieldList()
		{
			Fields = new List<TAcroField>();
		}

		internal int Count{ get{return Fields.Count;}}

		internal string GetCallArray(TPdfStream DataStream, TXRefSection XRef, bool OnlySignatures)
		{
			StringBuilder sb = new StringBuilder();
			int aCount = Count;
			for (int i = 0; i < aCount; i++)
			{
				int id =Fields[i].GetId(DataStream, XRef, OnlySignatures);
				if (id < 0) continue;
				if (sb.Length > 0) sb.Append(" ");
				sb.Append(TIndirectRecord.GetCallObj(id));
			}
			if (sb.Length==0) return String.Empty;

			if (OnlySignatures) return sb.ToString();

			return TPdfTokens.GetString(TPdfToken.OpenArray) + sb.ToString()
				+ TPdfTokens.GetString(TPdfToken.CloseArray);
		}

		internal void SaveToStream(TPdfStream DataStream, TXRefSection XRef, TPageInfo ParentPage)
		{
			int aCount = Count;
			for (int i = 0; i < aCount; i++)
			{
				Fields[i].SaveToStream(DataStream, XRef, ParentPage);
			}
		}

		internal void FinishSign(TPdfStream DataStream)
		{
			int aCount = Count;
			for (int i = 0; i < aCount; i++)
			{
				Fields[i].FinishSign(DataStream);
			}
		}

		internal void Add(TAcroField Field)
		{
			Fields.Add(Field);
		}

        internal bool HasSigFields()
        {
            foreach (TAcroField f in Fields)
            {
                if (f is TAcroSigField) return true;
            }

            return false;
        }

        internal int SigFieldId(TPdfStream DataStream, TXRefSection XRef)
        {
            foreach (TAcroField f in Fields)
            {
                TAcroSigField Sig = f as TAcroSigField;
                if (Sig != null) return Sig.GetSigDictionaryId(DataStream, XRef);
            }

            return -1;
        }

    }
	#endregion

	#region Fields
	internal abstract class TAcroField: TDictionaryRecord
	{
		private int FId = -1;

		internal TAcroField()
		{
		}

		internal virtual int GetId(TPdfStream DataStream, TXRefSection XRef, bool OnlySignatures)
		{
			if (OnlySignatures) return -1;
			if (FId < 0) FId = XRef.GetNewObject(DataStream);
			return FId;
		}

		internal abstract void SaveToStream(TPdfStream DataStream, TXRefSection XRef, TPageInfo ParentPage);

		protected void BeginSaveToStream(TPdfStream DataStream, TXRefSection XRef)
		{
			XRef.SetObjectOffset(FId, DataStream);
			TIndirectRecord.SaveHeader(DataStream, FId);
			BeginDictionary(DataStream);
		}

		protected static void EndSaveToStream(TPdfStream DataStream, TXRefSection XRef)
		{
			EndDictionary(DataStream);
			TIndirectRecord.SaveTrailer(DataStream);
		}

		internal virtual void FinishSign(TPdfStream DataStream)
		{
			//Nothing here.
		}

	}

	internal class TAcroSigField: TAcroField
	{
		private TPdfSignature Signature;
        private int PKCSSize;
        private int SigDictionaryId = -1;

		internal TAcroSigField(TPdfSignature aSignature): base()
		{
			Signature = aSignature;
		}

		internal override int GetId(TPdfStream DataStream, TXRefSection XRef, bool OnlySignatures)
		{
			return base.GetId (DataStream, XRef, false);
		}

        internal int GetSigDictionaryId(TPdfStream DataStream, TXRefSection XRef)
        {
            if (SigDictionaryId < 0) SigDictionaryId = XRef.GetNewObject(DataStream);
            return SigDictionaryId;
        }

		internal override void SaveToStream(TPdfStream DataStream, TXRefSection XRef, TPageInfo ParentPage)
		{
			BeginSaveToStream(DataStream, XRef);

			//Annot
			SaveKey(DataStream, TPdfToken.TypeName, TPdfTokens.GetString(TPdfToken.AnnotName)); //Annots and Fields mix their dictionaries. See "Digital Signature Appearances" to see why annots are required even in non visible sigs.
			SaveKey(DataStream, TPdfToken.SubtypeName, TPdfTokens.GetString(TPdfToken.WidgetName)); 
			
			RectangleF Rect = new RectangleF(0,0,0,0);
			TPdfVisibleSignature VSig = Signature as TPdfVisibleSignature;

			int APId = -1;
			if (VSig != null)
			{
				Rect = VSig.Rect;
				APId = SaveAPRef(DataStream, XRef);
			}
			
			SaveKey(DataStream, TPdfToken.RectName, 
				TPdfTokens.GetString(TPdfToken.OpenArray)+
				String.Format(CultureInfo.InvariantCulture, "{0} {1} {2} {3}", Rect.Left, Rect.Top, Rect.Right, Rect.Bottom)+ 
				TPdfTokens.GetString(TPdfToken.CloseArray) 
				);

			SaveKey(DataStream, TPdfToken.PName, TIndirectRecord.GetCallObj(ParentPage.Id));
			SaveKey(DataStream, TPdfToken.FName, 132);  //bits 3 and 8

			//Field
			SaveKey(DataStream, TPdfToken.FTName, TPdfTokens.GetString(TPdfToken.SigName));
			if (Signature.Name != null) SaveUnicodeKey(DataStream, TPdfToken.TName, Signature.Name);
			SaveKey(DataStream, TPdfToken.FfName, 1);

			XRef.SetObjectOffset(GetSigDictionaryId(DataStream, XRef), DataStream);
			SaveKey(DataStream, TPdfToken.VName, TIndirectRecord.GetCallObj(SigDictionaryId));
			
			EndSaveToStream(DataStream, XRef);

			if (VSig != null)
			{
				SaveAPObj(DataStream, XRef, APId, VSig);
			}

			SaveSigDictionary(DataStream, XRef, SigDictionaryId);

        }

        #region AP

        private static void WriteCommonXObject(TPdfStream DataStream, RectangleF Rect, string StreamContents)
        {
            SaveKey(DataStream, TPdfToken.LengthName, StreamContents.Length);
            SaveKey(DataStream, TPdfToken.TypeName, TPdfTokens.GetString(TPdfToken.XObjectName));
            SaveKey(DataStream, TPdfToken.SubtypeName, TPdfTokens.GetString(TPdfToken.FormName));
            SaveKey(DataStream, TPdfToken.FormTypeName, 1);
            SaveKey(DataStream, TPdfToken.MatrixName, "[1.0 0.0 0.0 1.0 0.0 0.0]"); //Standard drawing matrix
            SaveKey(DataStream, TPdfToken.BBoxName, PdfConv.ToRectangleWH(new RectangleF(0, 0, Rect.Width, Rect.Height), true));
        }

		private static int SaveAPRef(TPdfStream DataStream, TXRefSection XRef)
		{
			//AP call
			int APId = XRef.GetNewObject(DataStream);
			Write(DataStream, TPdfTokens.GetString(TPdfToken.APName));
			BeginDictionary(DataStream);
			SaveKey(DataStream, TPdfToken.NName, TIndirectRecord.GetCallObj(APId));
			EndDictionary(DataStream);
			return APId;
		}

		private static void SaveAPObj(TPdfStream DataStream, TXRefSection XRef, int APId, TPdfVisibleSignature VSig)
		{
            string StreamContents = "q 1 0 0 1 0 0 cm /FRM Do Q";

			//AP Object
			XRef.SetObjectOffset(APId, DataStream);
			TIndirectRecord.SaveHeader(DataStream, APId);
			BeginDictionary(DataStream);
            WriteCommonXObject(DataStream, VSig.Rect, StreamContents);
            int FRMId = XRef.GetNewObject(DataStream);
            SaveProcSet(DataStream, XRef, false);
            SaveResourcesFirstXObject(DataStream, XRef, FRMId);
            EndDictionary(DataStream);

            WriteStream(DataStream, StreamContents);

			TIndirectRecord.SaveTrailer(DataStream);

            SaveFRM(DataStream, XRef, VSig, FRMId);
		}

        private static void WriteStream(TPdfStream DataStream, string StreamContents)
        {
            TStreamRecord.BeginSave(DataStream);
            WriteLine(DataStream, StreamContents);
            TStreamRecord.EndSave(DataStream);
        }

        private static void SaveProcSet(TPdfStream DataStream, TXRefSection XRef, bool HasExtraProcs)
        {
            string ExtraProcs = HasExtraProcs ?
                TPdfTokens.NewLine + TPdfTokens.GetString(TPdfToken.TextName) + TPdfTokens.NewLine +
                TPdfTokens.GetString(TPdfToken.ImageCName)+
                TPdfTokens.GetString(TPdfToken.ImageIName)+
                TPdfTokens.GetString(TPdfToken.ImageBName)
                :
                String.Empty;

            Write(DataStream, TPdfTokens.GetString(TPdfToken.ResourcesName));
            BeginDictionary(DataStream);
            SaveKey(DataStream, TPdfToken.ProcSetName, 
                TPdfTokens.GetString(TPdfToken.OpenArray) +
                TPdfTokens.GetString(TPdfToken.PDFName) +  ExtraProcs +
                TPdfTokens.GetString(TPdfToken.CloseArray)
                );
        }
        
        private static void SaveResourcesFirstXObject(TPdfStream DataStream, TXRefSection XRef, int FRMId)
        {
            Write(DataStream, TPdfTokens.GetString(TPdfToken.XObjectName));
            BeginDictionary(DataStream);
            SaveKey(DataStream, TPdfToken.FRMName, TIndirectRecord.GetCallObj(FRMId));
            EndDictionary(DataStream);
            EndDictionary(DataStream);
        }

        private static void SaveFRM(TPdfStream DataStream, TXRefSection XRef, TPdfVisibleSignature VSig, int FRMId)
        {
            string StreamContents = "q 1 0 0 1 0 0 cm /n0 Do Q" + TPdfTokens.NewLine + "q 1 0 0 1 0 0 cm /n2 Do Q";

            XRef.SetObjectOffset(FRMId, DataStream);
            TIndirectRecord.SaveHeader(DataStream, FRMId);
            BeginDictionary(DataStream);
            WriteCommonXObject(DataStream, VSig.Rect, StreamContents);
            int n0Id = XRef.GetNewObject(DataStream);
            int n2Id = XRef.GetNewObject(DataStream);
            SaveProcSet(DataStream, XRef, false);
            SaveResourcesSecondXObject(DataStream, n0Id, n2Id);
            EndDictionary(DataStream);
            WriteStream(DataStream, StreamContents);

            TIndirectRecord.SaveTrailer(DataStream);

            Saven0(DataStream, XRef, VSig, n0Id);
            Saven2(DataStream, XRef, VSig, n2Id);
        }

        private static void SaveResourcesSecondXObject(TPdfStream DataStream, int n0Id, int n2Id)
        {
            Write(DataStream, TPdfTokens.GetString(TPdfToken.XObjectName));
            BeginDictionary(DataStream);
            SaveKey(DataStream, TPdfToken.n0Name, TIndirectRecord.GetCallObj(n0Id));
            SaveKey(DataStream, TPdfToken.n2Name, TIndirectRecord.GetCallObj(n2Id));
            EndDictionary(DataStream);
            EndDictionary(DataStream);
        }

        private static void Saven0(TPdfStream DataStream, TXRefSection XRef, TPdfVisibleSignature VSig, int n0Id)
        {
            string StreamContents = "% DSBlank";

            XRef.SetObjectOffset(n0Id, DataStream);
            TIndirectRecord.SaveHeader(DataStream, n0Id);
            BeginDictionary(DataStream);
            WriteCommonXObject(DataStream, new RectangleF(0, 0, 100, 100), StreamContents);
            SaveProcSet(DataStream, XRef, false);
            EndDictionary(DataStream);
            EndDictionary(DataStream);
            WriteStream(DataStream, StreamContents);

            TIndirectRecord.SaveTrailer(DataStream);
        }

        private static void Saven2(TPdfStream DataStream, TXRefSection XRef, TPdfVisibleSignature VSig, int n2Id)
        {
            string s = String.Format(CultureInfo.InvariantCulture, "{0} {1} {2} {3} {4} {5} cm ",
                            VSig.Rect.Width,
                            0, 0,
                            VSig.Rect.Height,
                            0, 0);

            string StreamContents = "q " + s +  TPdfTokens.GetString(TPdfToken.ImgPrefix) + "0 Do Q";

            XRef.SetObjectOffset(n2Id, DataStream);
            TIndirectRecord.SaveHeader(DataStream, n2Id);
            BeginDictionary(DataStream);
            WriteCommonXObject(DataStream, VSig.Rect, StreamContents);

            TPdfResources Resources = new TPdfResources(null, false, null);
            using (MemoryStream ImgStream = new MemoryStream(VSig.ImageData))
            {
                Resources.AddImage(null, ImgStream, FlxConsts.NoTransparentColor, false);
            }

            SaveProcSet(DataStream, XRef, true);
            SaveResourcesImgXObject(DataStream, XRef, Resources);

            EndDictionary(DataStream);
            WriteStream(DataStream, StreamContents);

            TIndirectRecord.SaveTrailer(DataStream);

            Resources.SaveObjects(DataStream, XRef);

        }

        private static void SaveResourcesImgXObject(TPdfStream DataStream, TXRefSection XRef, TPdfResources Resources)
        {
            Resources.SaveResourceDesc(DataStream, XRef, false);
            EndDictionary(DataStream);
        }

        #endregion

        private void SaveSigDictionary(TPdfStream DataStream, TXRefSection XRef, int SigDictionaryId)
		{
			XRef.SetObjectOffset(SigDictionaryId, DataStream);
			TIndirectRecord.SaveHeader(DataStream, SigDictionaryId);
			BeginDictionary(DataStream);

			SaveKey(DataStream, TPdfToken.TypeName, TPdfTokens.GetString(TPdfToken.SigName));
			SaveKey(DataStream, TPdfToken.FilterName, TPdfTokens.GetString(TPdfToken.Adobe_PPKLiteName));
            SaveKey(DataStream, TPdfToken.SubFilterName, TPdfTokens.GetString(TPdfToken.adbe_pkcs7_detachedName));

            DateTime dt = Signature.SignDate;
            if (dt == DateTime.MinValue) dt = DateTime.Now;
            SaveKey(DataStream, TPdfToken.MName, TDateRecord.GetDate(dt));

			if (Signature.Location != null) SaveUnicodeKey(DataStream, TPdfToken.LocationName, Signature.Location);
			if (Signature.Reason != null) SaveUnicodeKey(DataStream, TPdfToken.ReasonName, Signature.Reason);
			if (Signature.ContactInfo != null) SaveUnicodeKey(DataStream, TPdfToken.ContactInfoName, Signature.ContactInfo);

			Write(DataStream, TPdfTokens.GetString(TPdfToken.ContentsName)); //After writing "contents" we need to write the pkcs data, but it can not be written yet since we didn't calculate it.

            PKCSSize = DataStream.GetEstimatedLength() * 2 + 2;  //*2 +2 is because the real string will be hexa, and each byte is 2 bytes, plus "<" and ">" to mark an hexa string.

            int SignOffset = 50 + PKCSSize;  //The 50 is for /Contents[a b c d] where a, b, c and d are of variable length.
            DataStream.GoVirtual(SignOffset);

            SaveReferenceDict(DataStream, XRef);

            EndSaveToStream(DataStream, XRef);            
		}

        private void SaveReferenceDict(TPdfStream DataStream, TXRefSection XRef)
        {
            Write(DataStream, TPdfTokens.GetString(TPdfToken.ReferenceName));
            Write(DataStream, TPdfTokens.GetString(TPdfToken.OpenArray));
            SaveDocMDP(DataStream, XRef);
            SaveFieldMDP(DataStream, XRef);
            Write(DataStream, TPdfTokens.GetString(TPdfToken.CloseArray));
        }

        private void SaveDocMDP(TPdfStream DataStream, TXRefSection XRef)
        {
            BeginDictionary(DataStream);
            SaveKey(DataStream, TPdfToken.TypeName, TPdfTokens.GetString(TPdfToken.SigRefName));
            SaveKey(DataStream, TPdfToken.TransformMethodName, TPdfTokens.GetString(TPdfToken.DocMDPName));
            SaveTransformParams(DataStream);
            SaveKey(DataStream, TPdfToken.DigestMethodName, TPdfTokens.GetString(TPdfToken.MD5Name));
            
            //This is a small hack. We don't have an object model of the pdf file able to compute the object hash,
            //and it is not needed anyway for Acrobat >=7 (and <6 does not allow MDP anyway). 
            //But we need to write something here, or acrobat will complain the document is not PDF/SigQ complaint.
            byte[] MDPHash = new byte[16];
            string MDPHashStr = PdfConv.ToHexString(MDPHash, true);
            Write(DataStream, TPdfTokens.GetString(TPdfToken.DigestValueName));
            long StartMDPHash = DataStream.Position;
            Write(DataStream, MDPHashStr);
            string LocationString =
                TPdfTokens.GetString(TPdfToken.OpenArray) +
                String.Format(CultureInfo.InvariantCulture, "{0} {1}", StartMDPHash, MDPHashStr.Length) +
                TPdfTokens.GetString(TPdfToken.CloseArray);

            SaveKey(DataStream, TPdfToken.DigestLocationName, LocationString);
            
            EndDictionary(DataStream);
        }

        private void SaveTransformParams(TPdfStream DataStream)
        {
            Write(DataStream, TPdfTokens.GetString(TPdfToken.TransformParamsName));
            BeginDictionary(DataStream);
            SaveKey(DataStream, TPdfToken.TypeName, TPdfTokens.GetString(TPdfToken.TransformParamsName));
            SaveKey(DataStream, TPdfToken.PName, Signature.AllowedChangesValue);
            SaveKey(DataStream, TPdfToken.VName, TPdfTokens.GetString(TPdfToken.V1_2Name));
            EndDictionary(DataStream);
        }

        private static void SaveFieldMDP(TPdfStream DataStream, TXRefSection XRef)
        {
           //Not in this implementation.
        }
 

        internal byte[] GetByteRange(TPdfStream DataStream)
        {
            long cut = DataStream.SignPosition;
            long endcut = cut + PKCSSize;

            string s = TPdfTokens.GetString(TPdfToken.ByteRangeName) +
                TPdfTokens.GetString(TPdfToken.OpenArray) +
                String.Format(CultureInfo.InvariantCulture, "{0} {1} {2} {3}", 0, cut, endcut, DataStream.Length - endcut) +
                TPdfTokens.GetString(TPdfToken.CloseArray);

            return Coder.GetBytes(s);
        }

		internal override void FinishSign(TPdfStream DataStream)
		{
			DataStream.GoReal(GetByteRange(DataStream), PKCSSize);
		}

	}
	#endregion
}
