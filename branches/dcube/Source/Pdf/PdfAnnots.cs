using System;
using System.Text;
using System.Globalization;
using System.Collections.Generic;

#if (WPF)
using RectangleF = System.Windows.Rect;
using real = System.Double;
#else
using real = System.Single;
using System.Drawing;
#endif

namespace FlexCel.Pdf
{
	#region Annotation Classes
	/// <summary>
	/// Holds one annotation on a page.
	/// </summary>
	internal abstract class TAnnot: TDictionaryRecord
	{
		internal int FId;
		internal real x1;
		internal real y1;
		internal real Width;
		internal real Height;

		protected TAnnot(real ax1, real ay1, real aWidth, real aHeight)
		{
			x1=ax1;
			y1=ay1;
			Width = aWidth;
			Height = aHeight;
		}

		internal virtual int GetId(TPdfStream DataStream, TXRefSection XRef)
		{
			FId = XRef.GetNewObject(DataStream);
			return FId;
		}

		internal abstract void SaveToStream(TPdfStream DataStream, TXRefSection XRef);

		protected void BeginSaveToStream(TPdfStream DataStream, TXRefSection XRef)
		{
			XRef.SetObjectOffset(FId, DataStream);
			TIndirectRecord.SaveHeader(DataStream, FId);
			BeginDictionary(DataStream);
			SaveKey(DataStream, TPdfToken.TypeName, TPdfTokens.GetString(TPdfToken.AnnotName));
			SaveKey(DataStream, TPdfToken.RectName, 
				TPdfTokens.GetString(TPdfToken.OpenArray)+
				String.Format(CultureInfo.InvariantCulture, "{0} {1} {2} {3}", x1, y1, x1+Width, y1+Height)+ 
				TPdfTokens.GetString(TPdfToken.CloseArray)
				);

			DateTime dt= DateTime.Now;
			SaveKey(DataStream, TPdfToken.MName, TDateRecord.GetDate(dt));
		}

		protected static void EndSaveToStream(TPdfStream DataStream, TXRefSection XRef)
		{
			EndDictionary(DataStream);
			TIndirectRecord.SaveTrailer(DataStream);
		}

	}

	internal class TLinkAnnot: TAnnot
	{
		string FURL;
        int ActionId;

		internal TLinkAnnot(real ax1, real ay1, real aWidth, real aHeight, string URL):
			base(ax1, ay1, aWidth, aHeight)
		{
			//only 7-bits allowed.

            byte[] bt = TPdfStringRecord.EscapeString(Encoding.ASCII.GetBytes(URL));
			FURL = TPdfTokens.GetString(TPdfToken.OpenString)+
				Encoding.ASCII.GetString(bt, 0, bt.Length)+
				TPdfTokens.GetString(TPdfToken.CloseString);
		}

		internal override int GetId(TPdfStream DataStream, TXRefSection XRef)
		{
			base.GetId (DataStream, XRef);
			ActionId = XRef.GetNewObject(DataStream);
			return FId;
		}


		internal override void SaveToStream(TPdfStream DataStream, TXRefSection XRef)
		{
			BeginSaveToStream(DataStream, XRef);
			SaveKey(DataStream, TPdfToken.SubtypeName, TPdfTokens.GetString(TPdfToken.LinkName));
			SaveKey(DataStream, TPdfToken.AName, TIndirectRecord.GetCallObj(ActionId));
			SaveKey(DataStream, TPdfToken.BorderName, TPdfTokens.GetString(TPdfToken.Border0));
			EndSaveToStream(DataStream, XRef);

			SaveAction(DataStream, XRef);
		}

		private void SaveAction(TPdfStream DataStream, TXRefSection XRef)
		{
			XRef.SetObjectOffset(ActionId, DataStream);
			TIndirectRecord.SaveHeader(DataStream, ActionId);
			BeginDictionary(DataStream);

			TPdfBaseRecord.Write(DataStream,
				String.Format(CultureInfo.InvariantCulture, TPdfTokens.GetString(TPdfToken.LinkData), FURL));

			EndDictionary(DataStream);
			TIndirectRecord.SaveTrailer(DataStream);
		}
	}

	internal class TCommentAnnot: TAnnot
	{
		string FComment;
		TPdfCommentProperties FCommentProperties;

		internal TCommentAnnot(real ax1, real ay1, real aWidth, real aHeight, string aComment, TPdfCommentProperties aCommentProperties):
			base(ax1, ay1, aWidth, aHeight)
		{
			FComment = aComment;
			FCommentProperties = aCommentProperties;
		}

		internal override int GetId(TPdfStream DataStream, TXRefSection XRef)
		{
			base.GetId (DataStream, XRef);
			return FId;
		}

		internal override void SaveToStream(TPdfStream DataStream, TXRefSection XRef)
		{
			BeginSaveToStream(DataStream, XRef);
			SaveKey(DataStream, TPdfToken.SubtypeName, "/"+FCommentProperties.CommentType.ToString());
			SaveUnicodeKey(DataStream, TPdfToken.ContentsName, FComment);
			SaveKey(DataStream, TPdfToken.NameName, "/"+FCommentProperties.Icon.ToString());
			SaveKey(DataStream, TPdfToken.CAName, PdfConv.DoubleToString(FCommentProperties.Opacity));
			if (FCommentProperties.CommentType != TPdfCommentType.Text)
			{
				Color aColor = FCommentProperties.BackgroundColor;
				string BgColor = TPdfTokens.GetString(TPdfToken.OpenArray)+
					PdfConv.CoordsToString(aColor.R / 255.0) + " " +
					PdfConv.CoordsToString(aColor.G / 255.0) + " " +
					PdfConv.CoordsToString(aColor.B / 255.0) +
					TPdfTokens.GetString(TPdfToken.CloseArray);

				SaveKey(DataStream, TPdfToken.ICName, BgColor);

			}
			EndSaveToStream(DataStream, XRef);
		}
	}

	#endregion

	#region Annotation List
	/// <summary>
	/// Holds a list of annotation for one page.
	/// </summary>
	internal class TAnnotList
	{
		private List<TAnnot> Annots;

		internal TAnnotList()
		{
			Annots = new List<TAnnot>();
		}

		private int Count{ get{return Annots.Count;}}

        internal bool HasAnnots(TAcroFormRecord AcroForm, int PageNumber, int PageCount)
        {
            if (Count > 0) return true;
            return AcroForm.HasSignatureFields(PageNumber, PageCount);
        }

		internal string GetCallArray(TPdfStream DataStream, TXRefSection XRef, TAcroFormRecord AcroForm, int PageNumber, int PageCount)
		{
			StringBuilder sb = new StringBuilder();
			int aCount = Count;
			for (int i = 0; i < aCount; i++)
			{
                int id =Annots[i].GetId(DataStream, XRef);
				if (i>0) sb.Append(" ");
				sb.Append(TIndirectRecord.GetCallObj(id));
			}

			string SignatureFields = AcroForm.GetSignatureFieldsCallArray(DataStream, XRef, PageNumber, PageCount); //Signature Fields are also widget annotations.
			if (SignatureFields.Length > 0) 
			{
				sb.Append(" ");
				sb.Append(SignatureFields);
			}

			if (sb.Length==0) return String.Empty;
			return TPdfTokens.GetString(TPdfToken.OpenArray) + sb.ToString()
				+ TPdfTokens.GetString(TPdfToken.CloseArray);
		}

		internal void SaveToStream(TPdfStream DataStream, TXRefSection XRef)
		{
			int aCount = Count;
			for (int i = 0; i < aCount; i++)
			{
				Annots[i].SaveToStream(DataStream, XRef);
			}
		}

		internal void Add(TAnnot Annot)
		{
			Annots.Add(Annot);
		}
	}
	#endregion
}
