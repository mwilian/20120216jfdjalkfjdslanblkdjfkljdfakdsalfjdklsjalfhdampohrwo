#region Using directives

using System;
using System.Text;
using System.IO;
using System.Globalization;
using System.Reflection;
using FlexCel.Core;
using System.Collections.Generic;

#if (WPF)
using RectangleF = System.Windows.Rect;
using real = System.Double;
using ColorBlend = System.Windows.Media.GradientStopCollection;
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
    /// Container for the whole document.
    /// </summary>
    internal sealed class TDocumentCatalogRecord: TDictionaryRecord
    {
		private TDocumentCatalogRecord(){}

        internal static void SaveToStream(TPdfStream DataStream, TXRefSection XRef, int CatalogId, int PagesId, int BookmarksId, int AcroFormId, int PermsId, TPdfToken PageLayout)
        {
			XRef.SetObjectOffset(CatalogId, DataStream);
            TIndirectRecord.SaveHeader(DataStream, CatalogId);
            TDictionaryRecord.BeginDictionary(DataStream);
            TDictionaryRecord.SaveKey(DataStream, TPdfToken.TypeName, TPdfTokens.GetString(TPdfToken.CatalogName));
			SaveKey(DataStream, TPdfToken.PagesName, TIndirectRecord.GetCallObj(PagesId));
			if (BookmarksId >=0)
			{
				SaveKey(DataStream, TPdfToken.OutlinesName, TIndirectRecord.GetCallObj(BookmarksId));
			}
			if (PageLayout != TPdfToken.None)
			{
				SaveKey(DataStream, TPdfToken.PageModeName, TPdfTokens.GetString(PageLayout));
			}
			if (AcroFormId >= 0)
			{
				SaveKey(DataStream, TPdfToken.AcroFormName, TIndirectRecord.GetCallObj(AcroFormId));
			}
            if (PermsId >= 0)
            {
                SaveKey(DataStream, TPdfToken.PermsName, TIndirectRecord.GetCallObj(PermsId));
            }
            EndDictionary(DataStream);
            TIndirectRecord.SaveTrailer(DataStream);
        }
    }

	internal class TPageInfo
	{
		internal TPageContents Contents;
		internal TAnnotList Annots;
		private int FId;

		internal TPageInfo(int aId, TPageContents aContents, TAnnotList aAnnots)
		{
			FId = aId;
			Contents = aContents;
			Annots = aAnnots;
		}

		internal int Id
		{
			get
			{
				return FId;
			}
		}
	}

    internal class TPageContents
    {
        internal int Id;
        internal TPaperDimensions PageSize;

        internal TPageContents(int aId, TPaperDimensions aPageSize)
        {
            Id=aId;
            PageSize=aPageSize;
        }
    }

    /// <summary>
    /// Contains a list of pages or other PagesTrees.
    /// </summary>
    internal class TPageTreeRecord: TDictionaryRecord
    {
        private List<TPageInfo> FList;
        internal int Id;
        private TPdfResources Resources;

        internal TPageTreeRecord(int aId, string aFallbackFontList, bool aCompress, TFontEvents FontEvents)
        {
            FList = new List<TPageInfo>();
            Resources = new TPdfResources(aFallbackFontList, aCompress, FontEvents);
            Id = aId;
        }

        internal void AddPage(int PageId, int ContentId, TPaperDimensions PageSize)
        {
            FList.Add(new TPageInfo(PageId, new TPageContents(ContentId, PageSize), new TAnnotList()));
        }

		internal int GetPageId(int page)
		{
			if (page < 1 || page > FList.Count) PdfMessages.ThrowException(PdfErr.ErrInvalidPageNumber, page, FList.Count);
			return FList[page - 1].Id;
		}

        internal void SaveToStream(TPdfStream DataStream, TXRefSection XRef, TAcroFormRecord AcroForm)
        {
            XRef.SetObjectOffset(Id, DataStream);
            TIndirectRecord.SaveHeader(DataStream, Id);
            BeginDictionary(DataStream);
            SaveKey(DataStream, TPdfToken.TypeName, TPdfTokens.GetString(TPdfToken.PagesName));

            SaveKids(DataStream, XRef);

            SaveKey(DataStream, TPdfToken.CountName, FList.Count);
            WriteLine(DataStream, TPdfTokens.GetString(TPdfToken.ResourcesName));
            Resources.SaveResourceDesc(DataStream, XRef, true);
            EndDictionary(DataStream);
            TIndirectRecord.SaveTrailer(DataStream);

            int aCount = FList.Count;
            for (int i = 0; i < aCount; i++)
            {
                TPageInfo PInfo = FList[i];
                TPageRecord.SaveToStream(DataStream, Id, PInfo, AcroForm, XRef, i, aCount);
                PInfo.Annots.SaveToStream(DataStream, XRef);

            }

            Resources.SaveObjects(DataStream, XRef);
        }
        private void SaveKids(TPdfStream DataStream, TXRefSection XRef)
        {
            WriteLine(DataStream, TPdfTokens.GetString(TPdfToken.KidsName) + " " + TPdfTokens.GetString(TPdfToken.OpenArray));
            for (int i = 0; i < FList.Count; i++)
            {
                TIndirectRecord.CallObj(DataStream, FList[i].Id);
            }
            WriteLine(DataStream, TPdfTokens.GetString(TPdfToken.CloseArray));
        }

        internal TPdfFont SelectFont(TPdfStream DataStream, TFontMapping Mapping, Font aFont, string s,
            TFontEmbed aEmbed, TFontSubset aSubSet, bool aUseKerning, ref string LastFont)
        {
            return Resources.SelectFont(DataStream, Mapping, aFont, s, aEmbed, aSubSet, aUseKerning, ref LastFont);
        }

        internal TPdfFont GetFont(TFontMapping Mapping, Font aFont, string s,
            TFontEmbed aEmbed, TFontSubset aSubset, bool aUseKerning)
        {
            return Resources.GetFont(Mapping, aFont, s, aEmbed, aSubset, aUseKerning);
        }

        internal void SelectImage(TPdfStream DataStream, Image aImage, Stream ImageData, long transparentColor, bool defaultToJpg)
        {
            Resources.SelectImage(DataStream, aImage, ImageData, transparentColor, defaultToJpg);
        }

		internal void SelectBrush(TPdfStream DataStream, HatchBrush aBrush)
		{
			Resources.SelectPattern(DataStream, aBrush.HatchStyle, aBrush.ForegroundColor);
		}

		internal void SelectBrush(TPdfStream DataStream, TextureBrush aBrush, real[] aPatternMatrix)
		{
			Resources.SelectPattern(DataStream, aBrush.Image, aPatternMatrix);
		}

        private Color InterpolateColorRGB(Color c0, Color c1, real factor)
        {
            //We will interpolate in RGB mode, since we are defining gradients with DeviceRGB color space.
            return ColorUtil.FromArgb((byte)(c0.R + (c1.R - c0.R) * factor),
                                  (byte)(c0.G + (c1.G - c0.G) * factor),
                                  (byte)(c0.B + (c1.B - c0.B) * factor));
        }

        private ColorBlend GetColorBlendWithoutInterpolationColors(Blend bl, Color Color1, Color Color2)
        {
            //We will only use Blend when it has been set by SetSigmaBellShape. Since we only use SetSigmaBellShape in gradients without InterpolationColors, we can test this only here.    
            ColorBlend Result = new ColorBlend(Math.Max(bl.Factors.Length, 2));

            real factor = 0;
            real position = 0;
            for (int i = 0; i < Result.Colors.Length; i++)
            {
                if (bl.Factors.Length > 1 && i < bl.Factors.Length) factor = bl.Factors[i];
                if (bl.Positions.Length > 1 && i < bl.Positions.Length) position = bl.Positions[i];

                Result.Colors[i] = InterpolateColorRGB(Color1, Color2, factor);
                Result.Positions[i] = position;

                factor = 1;
                position = 1;
            }

            return Result;
        }

        internal void SelectBrush(TPdfStream DataStream, LinearGradientBrush aBrush, RectangleF Rect, RectangleF RotatedCoords, string DrawingMatrix)
        {
#if(WPF)
            ColorBlend cb = aBrush.GradientStops;

#else
            ColorBlend cb = null;
            try
            {
                if (aBrush.Blend == null) cb = aBrush.InterpolationColors; //if it has interpolationcolors, blend must be null.
            }
            catch (ArgumentException) //Awful way to tell if it has Interpolationcolors, but the framework does not give us a choice.
            {
            }

            if (cb == null)
            {
                cb = GetColorBlendWithoutInterpolationColors(aBrush.Blend, aBrush.LinearColors[0], aBrush.LinearColors[1]);
            }
#endif

            int n = cb.Colors.Length -1;
            if (cb.Colors[0].A != 255 || cb.Colors[n].A != 255)
            {
                ColorBlend TransparencyBlend = new ColorBlend(2);
                TransparencyBlend.Colors[0] = ColorUtil.FromArgb(cb.Colors[0].A, cb.Colors[0].A, cb.Colors[0].A);
                TransparencyBlend.Positions[0] = 0;
                TransparencyBlend.Colors[1] = ColorUtil.FromArgb(cb.Colors[n].A, cb.Colors[n].A, cb.Colors[n].A);
                TransparencyBlend.Positions[1] = 1;
                TPdfGradient SMask = Resources.GetGradient(TGradientType.Axial, TransparencyBlend, Rect, Rect.Location, RotatedCoords, DrawingMatrix);
                SelectTransparency(DataStream, 255, TPdfToken.CommandSetAlphaBrush, SMask.GetSMask(), PdfConv.ToRectangleXY(Rect, true));
            }
            Resources.SelectGradient(DataStream, TGradientType.Axial, cb, Rect, Rect.Location, RotatedCoords, DrawingMatrix);
        }

        internal void SelectBrush(TPdfStream DataStream, PathGradientBrush aBrush, RectangleF Rect, RectangleF RotatedCoords, PointF CenterPoint, string DrawingMatrix)
        {
            ColorBlend cb = null;
            try
            {
                cb = aBrush.InterpolationColors;
            }
            catch (ArgumentException)//Awful way to tell if it has InterpolationColors, but the framework does not give us a choice.
            {
            }

            if (cb == null || cb.Colors.Length == 1)
            {
                cb = GetColorBlendWithoutInterpolationColors(aBrush.Blend, aBrush.SurroundColors[0], aBrush.CenterColor);
            }

            int n = cb.Colors.Length -1;
            if (cb.Colors[0].A != 255 || cb.Colors[n].A != 255)
            {
                ColorBlend TransparencyBlend = new ColorBlend(2);
                TransparencyBlend.Colors[0] = ColorUtil.FromArgb(cb.Colors[0].A, cb.Colors[0].A, cb.Colors[0].A);
                TransparencyBlend.Positions[0] = 0;
                TransparencyBlend.Colors[1] = ColorUtil.FromArgb(cb.Colors[n].A, cb.Colors[n].A, cb.Colors[n].A);
                TransparencyBlend.Positions[1] = 1;
                TPdfGradient SMask = Resources.GetGradient(TGradientType.Radial, TransparencyBlend, Rect, aBrush.CenterPoint, RotatedCoords, DrawingMatrix);
                SelectTransparency(DataStream, 255, TPdfToken.CommandSetAlphaBrush, SMask.GetSMask(), PdfConv.ToRectangleXY(Rect, true));
            }

            Resources.SelectGradient(DataStream, TGradientType.Radial, cb, Rect, CenterPoint, RotatedCoords, DrawingMatrix);
        }

        internal void SelectTransparency(TPdfStream DataStream, int Alpha, TPdfToken aOperator)
        {
            Resources.SelectTransparency(DataStream, Alpha, aOperator);
        }

        internal void SelectTransparency(TPdfStream DataStream, int Alpha, TPdfToken aOperator, string aSMask, string aBBox)
        {
            Resources.SelectTransparency(DataStream, Alpha, aOperator, aSMask, aBBox);
        }


		internal void Hyperlink(real x, real y, real width, real height, string Url)
		{
			if (FList.Count<=0) return;
			FList[FList.Count-1].Annots.Add(new TLinkAnnot(x, y, width, height, Url));
		}

		internal void Comment(real x, real y, real width, real height, string comment, TPdfCommentProperties commentProperties)
		{
			if (FList.Count<=0) return;
			FList[FList.Count-1].Annots.Add(new TCommentAnnot(x, y, width, height, comment, commentProperties));
		}

		internal TPageInfo GetSigPage(TPdfSignature Signature)
		{
			int CurrentPage = TAcroFormRecord.GetSignaturePage(Signature, FList.Count);
			return FList[CurrentPage];
		}

    }

    /// <summary>
    /// The real page.
    /// </summary>
    internal sealed class TPageRecord : TDictionaryRecord
    {
		private TPageRecord(){}

        internal static void SaveToStream(TPdfStream DataStream, int PageListId, TPageInfo PageInfo, TAcroFormRecord AcroForm, TXRefSection XRef, int PageNumber, int PageCount)
        {
			XRef.SetObjectOffset(PageInfo.Id, DataStream);
            TIndirectRecord.SaveHeader(DataStream, PageInfo.Id);
            BeginDictionary(DataStream);
            SaveKey(DataStream, TPdfToken.TypeName, TPdfTokens.GetString(TPdfToken.PageName));
            SaveKey(DataStream, TPdfToken.ParentName, TIndirectRecord.GetCallObj(PageListId));
            SaveKey(DataStream, TPdfToken.ContentsName, TIndirectRecord.GetCallObj(PageInfo.Contents.Id));
            SaveKey(DataStream, TPdfToken.MediaBoxName, 
                TPdfTokens.GetString(TPdfToken.OpenArray)+
                "0 0 "+
                PdfConv.CoordsToString(PageInfo.Contents.PageSize.Width*72/100)+" "+
                PdfConv.CoordsToString(PageInfo.Contents.PageSize.Height*72/100)+
                TPdfTokens.GetString(TPdfToken.CloseArray)
                );
			if (PageInfo.Annots.HasAnnots(AcroForm, PageNumber, PageCount))
				SaveKey(DataStream, TPdfToken.AnnotsName, PageInfo.Annots.GetCallArray(DataStream, XRef, AcroForm, PageNumber, PageCount));
			EndDictionary(DataStream);
            TIndirectRecord.SaveTrailer(DataStream);
        }
    }

    internal class TInfoRecord : TDictionaryRecord
    {
        internal int InfoId;

        internal void SaveToStream(TPdfStream DataStream, TXRefSection XRef, TPdfProperties p)
        {
            InfoId = XRef.GetNewObject(DataStream);
            TIndirectRecord.SaveHeader(DataStream, InfoId);
            BeginDictionary(DataStream);
            if (p.Title != null && p.Title.Length > 0) SaveUnicodeKey(DataStream, TPdfToken.TitleName, p.Title);
            if (p.Author != null && p.Author.Length > 0) SaveUnicodeKey(DataStream, TPdfToken.AuthorName, p.Author);
            if (p.Subject != null && p.Subject.Length > 0) SaveUnicodeKey(DataStream, TPdfToken.SubjectName, p.Subject);
            if (p.Keywords != null && p.Keywords.Length > 0) SaveUnicodeKey(DataStream, TPdfToken.KeywordsName, p.Keywords);
            if (p.Creator != null && p.Creator.Length > 0) SaveUnicodeKey(DataStream, TPdfToken.CreatorName, p.Creator);

			string Producer = TPdfTokens.GetString(TPdfToken.Producer);
			if (!PdfWriter.FTesting)
				Producer += Assembly.GetExecutingAssembly().GetName().Version.ToString();

            if (Producer != null && Producer.Length > 0) SaveUnicodeKey(DataStream, TPdfToken.ProducerName, Producer);

			DateTime dt = DateTime.Now;
            SaveKey(DataStream, TPdfToken.CreationDateName, TDateRecord.GetDate(dt));
            EndDictionary(DataStream);
            TIndirectRecord.SaveTrailer(DataStream);
        }
    }

    internal class TPermsRecord : TDictionaryRecord
    {
        int PermsId = -1;

        internal void SaveToStream(TPdfStream DataStream, TXRefSection XRef, int MDPId)
        {
            if (GetId(DataStream, XRef, MDPId) < 0) return;
            if (MDPId < 0) return;
            XRef.SetObjectOffset(PermsId, DataStream);
            TIndirectRecord.SaveHeader(DataStream, PermsId);
            BeginDictionary(DataStream);
            SaveKey(DataStream, TPdfToken.DocMDPName, TIndirectRecord.GetCallObj(MDPId));
            EndDictionary(DataStream);
            TIndirectRecord.SaveTrailer(DataStream);
        }

        internal int GetId(TPdfStream DataStream, TXRefSection XRef, int MDPId)
        {
            if (MDPId < 0) return -1;
            if (PermsId < 0) PermsId = XRef.GetNewObject(DataStream);
            return PermsId;
        }
    }
}
