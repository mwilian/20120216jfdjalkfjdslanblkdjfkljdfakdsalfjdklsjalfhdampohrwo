#region Using directives

using System;
using System.Text;
using System.IO;
using System.Globalization;
using FlexCel.Core;


#if (WPF)
using RectangleF = System.Windows.Rect;
using real = System.Double;
using System.Windows.Media;
#else
using real = System.Single;
using System.Drawing;
using System.Drawing.Drawing2D;
#endif


#endregion

namespace FlexCel.Pdf
{
    /// <summary>
    /// Base section for all PDF sections on a file.
    /// </summary>
    internal class TPdfBaseSection
    {
        public TPdfBaseSection()
        {
        }
    }

    /// <summary>
    /// Header. It should only have a header and a comment record. The Comment record
    /// should have at lest 4 non ASCII 128 characters.
    /// </summary>
    internal sealed class THeaderSection : TPdfBaseSection
    {
		private THeaderSection(){}
        internal static void SaveToStream(TPdfStream DataStream)
        {
			if (DataStream.HasSignature)  //Our sign algorithm requires Reader 7 or newer.
			{
				TPdfHeaderRecord.SaveToStream(DataStream, TPdfTokens.GetString(TPdfToken.Header16));
			}
			else
			{
				TPdfHeaderRecord.SaveToStream(DataStream, TPdfTokens.GetString(TPdfToken.Header14));
			}

            TPdfBaseRecord.WriteLine(DataStream, TPdfTokens.GetString(TPdfToken.HeaderComment));
        }

    }

    /// <summary>
    /// Where the data goes.
    /// </summary>
    internal class TBodySection : TPdfBaseSection
    {
        private int LengthId;
        private bool Compress;
        private long StartStream, EndStream;
        private TPageTreeRecord PageTree;
		internal TBookmarkList Bookmarks;
		internal TPageLayout PageLayout;
        internal int CatalogId;
        private TInfoRecord Info;
        private TPermsRecord Perms;

		private TAcroFormRecord AcroForm;

        internal TBodySection(bool aCompress)
        {
            Compress = aCompress;
			Bookmarks = new TBookmarkList();
        }

		internal static TBodySection CreateTempBody(string aFallbackFontList, TFontEvents FontEvents)
		{
			TBodySection Result = new TBodySection(false);
			Result.PageTree = new TPageTreeRecord(0, aFallbackFontList, false, FontEvents);
			return Result;
		}
		
		internal int InfoId { get { return Info.InfoId; } }

        internal void BeginSave(TPdfStream DataStream, TXRefSection XRef, TPaperDimensions PageSize, TPdfSignature Signature, string aFallbackFontList, TFontEvents FontEvents)
        {
			CatalogId = XRef.GetNewObject(DataStream);
            int PageTreeId = XRef.GetNewObject(DataStream);
			PageTree = new TPageTreeRecord(PageTreeId, aFallbackFontList, Compress, FontEvents);
			Info = new TInfoRecord();
			AcroForm = new TAcroFormRecord(Signature);
            Perms = new TPermsRecord();

            CreateNewPage(DataStream, XRef, PageSize);
        }

        internal void EndSave(TPdfStream DataStream, TXRefSection XRef, TPdfProperties Properties, TPdfSignature Signature)
        {
            FinishPage(DataStream, XRef);

			TPdfToken P = TPdfToken.None;
			switch (PageLayout)
			{
				case TPageLayout.Outlines: P = TPdfToken.UseOutlinesName;break;
				case TPageLayout.Thumbs: P = TPdfToken.UseThumbsName;break;
				case TPageLayout.FullScreen: P = TPdfToken.FullScreenName;break;
			}

			int OutlinesId = -1;
			if (Bookmarks.Count > 0)
			{
				OutlinesId = XRef.GetNewObject(DataStream);
			}

			TDocumentCatalogRecord.SaveToStream(DataStream, XRef, CatalogId, PageTree.Id, OutlinesId, 
                AcroForm.GetId(DataStream, XRef), Perms.GetId(DataStream, XRef, AcroForm.SignatureFieldId(DataStream, XRef)), P);
			PageTree.SaveToStream(DataStream, XRef, AcroForm);
			SaveBookmarks(DataStream, XRef, OutlinesId);
		
            Info.SaveToStream(DataStream, XRef, Properties);

            Perms.SaveToStream(DataStream, XRef, AcroForm.SignatureFieldId(DataStream, XRef));

			AcroForm.SaveToStream(DataStream, XRef, PageTree.GetSigPage(Signature)); //We save it last because it needs to keep everything after it in memory for signing it.
        }

		internal void FinishSign(TPdfStream DataStream)
		{
			AcroForm.FinishSign(DataStream);
		}

        internal void FinishPage(TPdfStream DataStream, TXRefSection XRef)
        {
			DataStream.FlushEndText();
            DataStream.Compress = false;
            EndStream = DataStream.Position;
            TStreamRecord.EndSave(DataStream);
            TIndirectRecord.SaveTrailer(DataStream);

            XRef.SetObjectOffset(LengthId, DataStream);
            TIndirectRecord.SaveHeader(DataStream, LengthId);
            TPdfBaseRecord.WriteLine(DataStream, (EndStream - StartStream).ToString(CultureInfo.InvariantCulture));
            TIndirectRecord.SaveTrailer(DataStream);
        }

        private void CreateNewPage(TPdfStream DataStream, TXRefSection XRef, TPaperDimensions PageSize)
        {
			int PageId = XRef.GetNewObject(DataStream);
			int ContentId = XRef.GetNewObject(DataStream);
            PageTree.AddPage(PageId, ContentId, PageSize);
            TIndirectRecord.SaveHeader(DataStream, ContentId);
            LengthId = XRef.GetNewObject(DataStream);
            TStreamRecord.BeginSave(DataStream, LengthId, Compress);
            DataStream.Compress = Compress;
            StartStream = DataStream.Position;
        }

        internal void NewPage(TPdfStream DataStream, TXRefSection XRef, TPaperDimensions PageSize)
        {
            FinishPage(DataStream, XRef);
            CreateNewPage(DataStream, XRef, PageSize);
        }

		internal string GetDestStr(TPdfDestination Dest)
		{
            int PageId = PageTree.GetPageId(Dest.PageNumber);

			string PageStr = TIndirectRecord.GetCallObj(PageId);
				
			switch (Dest.ZoomOptions)
			{
				case TZoomOptions.Fit: 
					return 
						TPdfTokens.GetString(TPdfToken.OpenArray) +
						PageStr + TPdfTokens.GetString(TPdfToken.FitName) +
						TPdfTokens.GetString(TPdfToken.CloseArray);
				case TZoomOptions.FitH: 
					return 
						TPdfTokens.GetString(TPdfToken.OpenArray) +
						PageStr + TPdfTokens.GetString(TPdfToken.FitHName) +
						TPdfTokens.GetString(TPdfToken.CloseArray);
				case TZoomOptions.FitV: 
					return 
						TPdfTokens.GetString(TPdfToken.OpenArray) +
						PageStr + TPdfTokens.GetString(TPdfToken.FitVName) +
						TPdfTokens.GetString(TPdfToken.CloseArray);
			}
			return 
				TPdfTokens.GetString(TPdfToken.OpenArray) +
				PageStr + TPdfTokens.GetString(TPdfToken.XYZName) +
				TPdfTokens.GetString(TPdfToken.CloseArray);

		}
		internal void SaveBookmarkObjects(TPdfStream DataStream, TXRefSection XRef, TBookmarkList bmks, int ParentId, int ObjectId, ref int FirstId, ref int LastId, ref int AllOpenCount)
		{
			int PreviousId = -1;
			for (int i = 0; i < bmks.Count; i++)
			{
				TBookmark b = bmks[i];
				AllOpenCount++;

				int NextId = -1;
				if (i < bmks.Count - 1) NextId = XRef.GetNewObject(DataStream);
				LastId = ObjectId;
				if (FirstId == -1)
				{
					FirstId = ObjectId;
				}

				int FirstChildId = -1;
				int LastChildId = -1;
				int ChildOpenCount = 0;
				if (b.FChildren.Count > 0)
				{
					FirstChildId = XRef.GetNewObject(DataStream);
					int ChildLastId = -1;
					SaveBookmarkObjects(DataStream, XRef, b.FChildren, ObjectId, FirstChildId, ref FirstId, ref ChildLastId, ref ChildOpenCount);
					if (!b.ChildrenCollapsed) AllOpenCount += ChildOpenCount;
					LastChildId = ChildLastId;
				}


				XRef.SetObjectOffset(ObjectId, DataStream);
				TIndirectRecord.SaveHeader(DataStream, ObjectId);
				TDictionaryRecord.BeginDictionary(DataStream);
				TDictionaryRecord.SaveUnicodeKey(DataStream, TPdfToken.TitleName, b.Title);
				TDictionaryRecord.SaveKey(DataStream, TPdfToken.ParentName, TIndirectRecord.GetCallObj(ParentId));
				if (PreviousId >= 0)
				{
					TDictionaryRecord.SaveKey(DataStream, TPdfToken.PrevName, TIndirectRecord.GetCallObj(PreviousId));
				}

				if (NextId >= 0)
				{
					TDictionaryRecord.SaveKey(DataStream, TPdfToken.NextName, TIndirectRecord.GetCallObj(NextId));
				}

				if (FirstChildId >= 0)
				{
					TDictionaryRecord.SaveKey(DataStream, TPdfToken.FirstName, TIndirectRecord.GetCallObj(FirstChildId));
				}

				if (LastChildId >= 0)
				{
					TDictionaryRecord.SaveKey(DataStream, TPdfToken.LastName, TIndirectRecord.GetCallObj(LastChildId));
				}
				if (ChildOpenCount > 0)
				{
					if (b.ChildrenCollapsed)
					{
						TDictionaryRecord.SaveKey(DataStream, TPdfToken.CountName, -ChildOpenCount);
					}
					else
					{
						TDictionaryRecord.SaveKey(DataStream, TPdfToken.CountName, ChildOpenCount);
					}
				}


				if (b.Destination != null)
				{
					TDictionaryRecord.SaveKey(DataStream, TPdfToken.DestName, GetDestStr(b.Destination));
				}

				if (b.TextColor.R != 0 || b.TextColor.G != 0 || b.TextColor.B != 0)
				{
					TDictionaryRecord.SaveKey(DataStream, TPdfToken.CName, TPdfTokens.GetString(TPdfToken.OpenArray) + PdfConv.ToString(b.TextColor) + TPdfTokens.GetString(TPdfToken.CloseArray));
				}

				if (b.TextStyle != TBookmarkStyle.None)
				{
					int k = 0;
					if ((b.TextStyle & TBookmarkStyle.Italic) != 0) k |= 1;
					if ((b.TextStyle & TBookmarkStyle.Bold) != 0) k |= 2;
					TDictionaryRecord.SaveKey(DataStream, TPdfToken.FName, k);			
				}

				TDictionaryRecord.EndDictionary(DataStream);
				TIndirectRecord.SaveTrailer(DataStream);

				PreviousId = ObjectId;
				ObjectId = NextId;
			}
		}

		internal void SaveBookmarks(TPdfStream DataStream, TXRefSection XRef, int OutlinesId)
		{
			if (OutlinesId < 0) return;

			int FirstId = -1;
			int LastId = -1;

			int NextId = XRef.GetNewObject(DataStream);
			int AllOpenCount = 0;
			SaveBookmarkObjects(DataStream, XRef, Bookmarks, OutlinesId, NextId, ref FirstId, ref LastId, ref AllOpenCount);

			XRef.SetObjectOffset(OutlinesId, DataStream);
			TIndirectRecord.SaveHeader(DataStream, OutlinesId);
			TDictionaryRecord.BeginDictionary(DataStream);

			if (FirstId >=0)
			{
				TDictionaryRecord.SaveKey(DataStream, TPdfToken.FirstName, TIndirectRecord.GetCallObj(FirstId));
				TDictionaryRecord.SaveKey(DataStream, TPdfToken.LastName, TIndirectRecord.GetCallObj(LastId));  //LastId must be positive if FirstId is positive.
			}

			if (AllOpenCount > 0)
			{
				TDictionaryRecord.SaveKey(DataStream, TPdfToken.CountName, AllOpenCount);
			}


			TDictionaryRecord.EndDictionary(DataStream);
			TIndirectRecord.SaveTrailer(DataStream);


		}


		internal TPdfFont SelectFont(TPdfStream DataStream, TFontMapping Mapping, Font aFont, string s, TFontEmbed aEmbed, TFontSubset aSubset, bool aUseKerning,
			ref string LastFont)
		{
			return PageTree.SelectFont(DataStream, Mapping, aFont, s, aEmbed, aSubset, aUseKerning, ref LastFont);
		}

		internal TPdfFont GetFont(TFontMapping Mapping, Font aFont, string s, TFontEmbed aEmbed, TFontSubset aSubset, bool aUseKerning)
		{
			return PageTree.GetFont(Mapping, aFont, s, aEmbed, aSubset, aUseKerning);
		}

		internal void SelectImage(TPdfStream DataStream, Image aImage, Stream ImageData, long transparentColor, bool defaultToJpg)
		{
			PageTree.SelectImage(DataStream, aImage, ImageData, transparentColor, defaultToJpg);
		}

        internal void SelectBrush(TPdfStream DataStream, HatchBrush aBrush)
        {
            PageTree.SelectBrush(DataStream, aBrush);
        }

		internal void SelectBrush(TPdfStream DataStream, TextureBrush aBrush, real[] aPatternMatrix)
		{
			PageTree.SelectBrush(DataStream, aBrush, aPatternMatrix);
		}
		
		internal void SelectBrush(TPdfStream DataStream, LinearGradientBrush aBrush, RectangleF Rect, RectangleF RotatedCoords, string DrawingMatrix)
        {
            PageTree.SelectBrush(DataStream, aBrush, Rect, RotatedCoords, DrawingMatrix);
        }

        internal void SelectBrush(TPdfStream DataStream, PathGradientBrush aBrush, RectangleF Rect, RectangleF RotatedCoords, PointF CenterPoint, string DrawingMatrix)
        {
            PageTree.SelectBrush(DataStream, aBrush, Rect, RotatedCoords, CenterPoint, DrawingMatrix);
        }
        
        internal void SelectTransparency(TPdfStream DataStream, int Alpha, TPdfToken aOperator, string aSMask, string aBBox)
        {
            PageTree.SelectTransparency(DataStream, Alpha, aOperator, aSMask, aBBox);
        }

        internal void SelectTransparency(TPdfStream DataStream, int Alpha, TPdfToken aOperator)
        {
            SelectTransparency(DataStream, Alpha, aOperator, null, null);
        }

		internal void Hyperlink(real x, real y, real width, real height, string Url)
		{
			PageTree.Hyperlink(x, y, width, height, Url);
		}
	
		internal void Comment(real x, real y, real width, real height, string comment, TPdfCommentProperties commentProperties)
		{
			PageTree.Comment(x, y, width, height, comment, commentProperties);
		}
	}

    /// <summary>
    /// Fixed form section.
    /// xref
    /// n l
    /// 0000000000 00000 n
    /// </summary>
    internal class TXRefSection : TPdfBaseSection
    {
#if (FRAMEWORK20)
        private System.Collections.Generic.List<long> FList;
#else
        private ArrayList FList;
#endif
        private long FStartPosition;

        internal TXRefSection()
        {
#if (FRAMEWORK20)
        FList = new System.Collections.Generic.List<long>();
#else
            FList = new ArrayList();
#endif
        }

        internal int GetNewObject(TPdfStream DataStream)
        {
#if (FRAMEWORK20)
            FList.Add(DataStream.Position);
            return FList.Count;
#else
            return FList.Add(DataStream.Position)+1;
#endif
        }

        internal void SetObjectOffset(int ObjId, TPdfStream DataStream)
        {
            FList[ObjId-1] = DataStream.Position;
        }

        public long StartPosition { get { return FStartPosition; } }
        public int Count { get { return FList.Count+1; } }

        internal void SaveToStream(TPdfStream DataStream)
        {
            FStartPosition = DataStream.Position;
            string Generation = " 00000 n ";

            TPdfBaseRecord.WriteLine(DataStream, TPdfTokens.GetString(TPdfToken.XRef));
            TPdfBaseRecord.WriteLine(DataStream, "0 " + (FList.Count + 1).ToString(CultureInfo.InvariantCulture));
            int aCount = FList.Count;

            TPdfBaseRecord.WriteLine(DataStream, "0000000000 65535 f ");

            for (int i = 0; i < aCount; i++)
            {
                TPdfBaseRecord.WriteLine(DataStream,(
#if(!FRAMEWORK20)
                    (long)
#endif 
                    FList[i]).ToString("0000000000", CultureInfo.InvariantCulture) +
                    Generation);
            }

        }
    }

    /// <summary>
    /// The trailer with XRefs offsets
    /// </summary>
    internal sealed class TTrailerSection : TPdfBaseSection
    {
		private TTrailerSection(){}

        internal static void SaveToStream(TPdfStream DataStream, TXRefSection XRef, int CatalogId, int InfoId)
        {
            TPdfBaseRecord.WriteLine(DataStream, TPdfTokens.GetString(TPdfToken.Trailer));

            TDictionaryRecord.BeginDictionary(DataStream);
            TDictionaryRecord.SaveKey(DataStream, TPdfToken.SizeName, XRef.Count);
            TDictionaryRecord.SaveKey(DataStream, TPdfToken.RootName, TIndirectRecord.GetCallObj(CatalogId));
            TDictionaryRecord.SaveKey(DataStream, TPdfToken.InfoName, TIndirectRecord.GetCallObj(InfoId));

            Byte[] FileIdBytes = Guid.NewGuid().ToByteArray();
            string FileId = PdfConv.ToHexString(FileIdBytes, true);
            TDictionaryRecord.SaveKey(DataStream, TPdfToken.IDName,
                TPdfTokens.GetString(TPdfToken.OpenArray) +
                FileId + " " + FileId +
                TPdfTokens.GetString(TPdfToken.CloseArray));

            TDictionaryRecord.EndDictionary(DataStream);

            TPdfBaseRecord.WriteLine(DataStream, TPdfTokens.GetString(TPdfToken.StartXRef));
            TPdfBaseRecord.WriteLine(DataStream, XRef.StartPosition.ToString(CultureInfo.InvariantCulture));
            TPdfBaseRecord.WriteLine(DataStream, TPdfTokens.GetString(TPdfToken.Eof));
        }
    }

	

}
