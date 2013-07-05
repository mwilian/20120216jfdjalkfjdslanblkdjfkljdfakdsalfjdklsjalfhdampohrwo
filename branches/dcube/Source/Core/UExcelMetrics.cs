using System;
using System.Runtime.CompilerServices;
using System.Globalization;

#if(MONOTOUCH)
using Font = MonoTouch.UIKit.UIFont;
#endif

#if (WPF)
using RectangleF = System.Windows.Rect;
using real = System.Double;
using System.Windows.Controls;
using System.Windows.Media;
#else
using real = System.Single;
using System.Drawing;
#endif

namespace FlexCel.Core
{
	/// <summary>
	/// Returns Information to convert between standard units and Excel units.
	/// </summary>
	public sealed class ExcelMetrics
	{
		private ExcelMetrics(){}
#if(!COMPACTFRAMEWORK)
        [ThreadStatic]
        private static real CachedFontWidth; //STATIC*
        [ThreadStatic]
        private static string CachedFontName; //STATIC*
        [ThreadStatic]
        private static int CachedFontSize; //STATIC*
        [ThreadStatic]
        private static TFlxFontStyles CachedFontStyle; //STATIC*
#endif

		#region Font0Width
		/// <summary>
		/// Returns the width of the 0 font on Excel. Normally this is 7, but can change depending on
		/// the file. The user might modify it by changing Format->Style->Normal.
		/// </summary>
		/// <param name="Workbook"></param>
		/// <returns></returns>
		public static real GetFont0Width(IRowColSize Workbook)
		{
			TFlxFont Fx = Workbook.GetDefaultFont;
			return GetFont0Width(Fx);
		}

		internal static real GetFont0Width(TFlxFont Fx)
		{
			try
			{
				return FullGetFont0Width(Fx);
			}
			catch (TypeLoadException)
			{
				return 7;
			}
        }


#if(WPF)
        private static real DoFullGetFont0Width(TFlxFont Fx)
        {
            TextBlock tb = new TextBlock();
            ExcelFont.SetFont(tb, Fx);
            tb.Text = "0";
            return (real)tb.ActualWidth; //Result is in silverlight pixels. Note that no transformation applies to TextBlock
        }

        private static real FullGetFont0Width(TFlxFont Fx)
        {
            if (String.Equals(Fx.Name, "ARIAL", StringComparison.InvariantCultureIgnoreCase) && Fx.Size20 == 200 && Fx.Style == TFlxFontStyles.None)
                return 7; //Most usual case.

            if (String.Equals(Fx.Name, CachedFontName, StringComparison.InvariantCultureIgnoreCase) && Fx.Size20 == CachedFontSize && Fx.Style == CachedFontStyle)
                return CachedFontWidth; //Most usual case.

            real Result = DoFullGetFont0Width(Fx);
            CachedFontWidth = Result;
            CachedFontName = Fx.Name;
            CachedFontSize = Fx.Size20;
            CachedFontStyle = Fx.Style;
            return Result; 
        }
#else
#if(!COMPACTFRAMEWORK)
        [MethodImpl(MethodImplOptions.NoInlining)]
		private static real DoFullGetFont0Width(TFlxFont Fx)
		{
#if(MONOTOUCH)
			using (Font MyFont = ExcelFont.CreateFont(Fx.Name, (Fx.Size20 / 20F), ExcelFont.ConvertFontStyle(Fx)))
			{
				using (MonoTouch.Foundation.NSString o = new MonoTouch.Foundation.NSString("0"))
				{
				    return o.StringSize(MyFont).Width;
				}
			}
#else

			using (Font MyFont = ExcelFont.CreateFont(Fx.Name, (Fx.Size20 / 20F), ExcelFont.ConvertFontStyle(Fx)))
			{
				using (Bitmap bm = new Bitmap(1,1))
				{
					using(Graphics gr = Graphics.FromImage(bm))
					using (StringFormat sfTemplate = StringFormat.GenericTypographic) //GenericTypographic returns a NEW instance. It has to be disposed.
					{
						using (StringFormat sf = (StringFormat) sfTemplate.Clone()) //Even when sfTemplate is a new instance, changing directly on it will change the standard generic typographic :(
						{
							//sf.SetMeasurableCharacterRanges was causing a deadlock here.
							//DONT DO!!
							/*CharacterRange[] r = {new CharacterRange(0,1)};
							sf.SetMeasurableCharacterRanges(r);*/

                            sf.Alignment = StringAlignment.Near; //this should be set, but just in case someone changed it.
                            sf.LineAlignment = StringAlignment.Far; //this should be set, but just in case someone changed it.
							sf.FormatFlags = 0;
							gr.PageUnit = GraphicsUnit.Pixel;
							SizeF sz = gr.MeasureString("0", MyFont, 1000, sf);
							return (real)Math.Round(sz.Width);	
						}
					}
				}
			}
#endif
		}

		/// <summary>
		/// No need for threadstatic.
		/// </summary>
		private static bool MissingFrameworkFont;

		private static real FullGetFont0Width(TFlxFont Fx)
		{
			if (String.Equals(Fx.Name, "ARIAL", StringComparison.InvariantCultureIgnoreCase) && Fx.Size20 == 200 && Fx.Style == TFlxFontStyles.None)
				return 7; //Most usual case.

            if (MissingFrameworkFont) return 7;

            if (String.Equals(Fx.Name, CachedFontName, StringComparison.InvariantCultureIgnoreCase) && Fx.Size20 == CachedFontSize && Fx.Style == CachedFontStyle)
                return CachedFontWidth; //Most usual case.
            
            try
            {
                real Result= DoFullGetFont0Width(Fx);
                CachedFontWidth = Result;
                CachedFontName = Fx.Name;
                CachedFontSize = Fx.Size20;
                CachedFontStyle = Fx.Style;
                return Result;
            }
            catch (MissingMethodException)
            {
                MissingFrameworkFont=true;
                return 7;
            }
		}

#else
		private static real FullGetFont0Width(TFlxFont Fx)
		{
			return 7;
		}
#endif
#endif
        #endregion

        #region RowHeight
        /// <summary>
		/// Returns the Height of an XF format Excel. 
		/// </summary>
		/// <param name="Fx"></param>
		/// <returns></returns>
		internal static int GetRowHeightInPixels(TFlxFont Fx)
		{
			try
			{
				return FullGetRowHeight(Fx);
			}
			catch (TypeLoadException)
			{
				return 0xFF;
			}
		}

#if(!COMPACTFRAMEWORK)
		[MethodImpl(MethodImplOptions.NoInlining)]
		private static int DoFullGetRowHeight(TFlxFont Fx)
		{
			using (Font MyFont = ExcelFont.CreateFont(Fx.Name, (Fx.Size20 / 20F), ExcelFont.ConvertFontStyle(Fx)))
			{
#if (MONOTOUCH)
				real h = MyFont.LineHeight;
#else
				real h = MyFont.GetHeight(75);
#endif
				return (int)(h * 20.87 + 5);	
			}
		}

		private static int FullGetRowHeight(TFlxFont Fx)
		{
			if (Fx.Size20<8*20) //Fonts less than 8 points are all the same
			{
				 int[] heights = {105, 105, 105, 120, 135, 165, 165, 180, 180};
				return heights[(int) Math.Round(Fx.Size20 / 20f)];
			}

			if (
				(
				String.Equals(Fx.Name, "ARIAL", StringComparison.InvariantCultureIgnoreCase) 
				||
				String.Equals(Fx.Name, "TIMES NEW ROMAN", StringComparison.InvariantCultureIgnoreCase) 
				)
				&& Fx.Size20 == 200 && Fx.Style == TFlxFontStyles.None)
				return 0xFF; //Most usual case.

			if (MissingFrameworkFont) return 0xFF;
			try
			{
				int Result= DoFullGetRowHeight(Fx);
				return Result;
			}
			catch (MissingMethodException)
			{
				MissingFrameworkFont=true;
				return 7;
			}
		}

#else
		private static int FullGetRowHeight(TFlxFont Fx)
		{
			return 0xFF;
		}
#endif
		#endregion

		#region FmlaMult
		/// <summary>
		/// When showing/printing the sheet and "Show formula" check box is on, column widths are double of the normal ones.
		/// This method returns 0.5 when "Show formulas" is turned on, and 1 if it is not.
		/// </summary>
		public static real FmlaMult(IRowColSize Workbook)
		{
			return Workbook.ShowFormulaText? 0.5f: 1f;
		}

		#endregion

		#region MultDisplay

		/// <summary>
		/// Multiply by this number to convert the width of a column from GraphicsUnit.Display units (1/100 inch) 
		/// to Excel internal units. Note that the default column width is different, you need to multiply by <see cref="DefColWidthAdapt(int, ExcelFile)"/>
		/// </summary>
		public static real ColMultDisplay(IRowColSize Workbook)
		{
			return (real)(FmlaMult(Workbook) * 33.358*7F / GetFont0Width(Workbook) * Workbook.WidthCorrection);
			//33.34F return 256f/GetFont0Width(Workbook);
			
		}

		/// <summary>
		/// Multiply by this number to convert the height of a row from GraphicsUnit.Display units (1/100 inch) 
		/// to Excel internal units.
		/// </summary>
		/// <remarks>
		/// 1 Height unit=1/20 pt. 1pt=1/72 inch. -> 1 Height unit=1/(72*20) inch. -> 1inch/100= 72*20/100= 14.4
		/// PrintPreview on Excel uses different coordinates than the screen.
		/// </remarks>
		//public static readonly real RowMultDisplay= 14.64F;  //14.72F;
		public static real RowMultDisplay(IRowColSize Workbook)
		{
			return 14.83F * Workbook.HeightCorrection;  //14.72F; 83
		}

		/// <summary>
		/// Multiply by this number to convert the width of a column from pixels to excel internal units. 
		/// Note that the default column width is different, you need to multiply by <see cref="DefColWidthAdapt(int, ExcelFile)"/>
		/// </summary>
		public static real ColMult(IRowColSize Workbook)
		{
			return FmlaMult(Workbook) * 256F / GetFont0Width(Workbook);
		}

		/// <summary>
		/// Convert the DEFAULT column width to pixels. This is different from <see cref="ColMult"/>, that goes in a column by column basis.
		/// </summary>
		public static int DefColWidthAdapt(int width, ExcelFile workbook)
		{
			return Convert.ToInt32(Math.Round((width+0.43) * 256 *8f/7f));
		}

		internal static int DefColWidthAdapt(int width)
		{
			//Test for font != null when used.
			//It is difficult to predict, but we are going to handle both most used cases (arial, 8 and arial, 10 that are the defaults on xls (97,2000) and (xp,2003))
			if (width == 8) return 0x924;
			if (width == 10) return 0xb6d;
			return Convert.ToInt32(Math.Round((width+0.43) * 256f * 1.09f));
		}

		internal static int InverseDefColWidthAdapt(int w)
		{
			//Test for font != null when used.
			//It is difficult to predict, but we are going to handle both most used cases (arial, 8 and arial, 10 that are the defaults on xls (97,2000) and (xp,2003))
			if (w == 0x924) return 8;
			if (w == 0xb6d) return 10;
			return Convert.ToInt32(Math.Round((w / 256f / 1.09f) - 0.43));
		}

		#endregion
	}
}
