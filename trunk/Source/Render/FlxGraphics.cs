using System;
using System.IO;

using FlexCel.Core;
#if (MONOTOUCH)
using System.Drawing;
using real = System.Single;
using Color = MonoTouch.UIKit.UIColor;
using Font = MonoTouch.UIKit.UIFont;
#else
	#if (WPF)
	using RectangleF = System.Windows.Rect;
	using SizeF = System.Windows.Size;
	using System.Windows.Media;
	using real = System.Double;
	#else
	using System.Drawing;
	using real = System.Single;
	#endif
#endif

namespace FlexCel.Render
{
	#region IFlxGraphics
	/// <summary>
	/// An interface to abstract all graphics operations we need to do.
	/// </summary>
	internal interface IFlxGraphics
	{
		RectangleF ClipBounds{get;}

		void CreateSFormat();
		void DestroySFormat();

		void Rotate(real x, real y, real Alpha);
		void Scale(real xScale, real yScale);
		TPointF Transform (TPointF p);
		void SaveTransform();
		void ResetTransform();

		void DrawString(string Text, Font aFont, Brush aBrush, real x, real y);
		void DrawString(string Text, Font aFont, Pen aPen, Brush aBrush, real x, real y);
		SizeF MeasureString(string Text, Font aFont, TPointF p); 
		SizeF MeasureString(string Text, Font aFont); 
		SizeF MeasureStringEmptyHasHeight(string Text, Font aFont);

		real FontDescent(Font aFont);
		real FontLinespacing(Font aFont);


		void DrawLines(Pen aPen, TPointF[] points);
		void DrawLine(Pen aPen, real x1, real y1, real x2, real y2);
		void FillRectangle (Brush b, RectangleF rect);
		void FillRectangle (Brush b, RectangleF rect, TClippingStyle clippingStyle);
		void FillRectangle (Brush b, real x1, real y1, real width, real height);
		void DrawRectangle(Pen pen, real x, real y, real width, real height);
		void DrawAndFillRectangle(Pen pen, Brush b, real x, real y, real width, real height);
		void DrawImage(Image image, RectangleF destRect, RectangleF srcRect, long transparentColor, int brightness, int contrast, int gamma, Color shadowColor, Stream imgData);

		void DrawAndFillBeziers(Pen pen, Brush brush, TPointF[] points, TClippingStyle clippingStyle);
		void DrawAndFillBeziers(Pen pen, Brush brush, TPointF[] points);
		void DrawAndFillPolygon(Pen pen, Brush brush, TPointF[] points, TClippingStyle clippingStyle);
		void DrawAndFillPolygon(Pen pen, Brush brush, TPointF[] points);

		void SetClipIntersect(RectangleF rect);
		void SetClipReplace(RectangleF rect);

		void ResetClip();

		void SaveState(); //always follow with restorestate
		void RestoreState();


		void AddHyperlink(real x1, real y1, real width, real height, string Url);

		void AddComment(real x1, real y1, real width, real height, string comment);

	}

	/// <summary>
	/// Can be used to clip into a region instead of drawing on it.
	/// </summary>
	public enum TClippingStyle
	{
		/// <summary>
		/// Normal draw, this will draw the image.
		/// </summary>
		None,

		/// <summary>
		/// This will include the region into the clipping area, no image will be drawn.
		/// </summary>
		Include,

		/// <summary>
		/// This will exclude the region from the clipping area, no image will be drawn.
		/// </summary>
		Exclude
	}

	#endregion

	//Generic classes that can be used with both System.Drawing and WPF classes.
	#region Supporting Classes
	internal interface IFont
	{
	}

	#region PointOutsides
	internal class PointOutside
	{
		private const int MaxSize = 4000000; //To avoid overflow errors. This is more than 1 km in points.
		private const int MinSize = -MaxSize;   //actually it is 0, but just to be sure when discarding text that starts at negative coords.

		private PointOutside(){}

		public static bool Check(ref real x, ref real y)
		{
			bool Result = false;
			if (x > MaxSize) {x = MaxSize; Result = true;}
			if (y > MaxSize) {y = MaxSize; Result = true;}
			if (x < MinSize) {x = MinSize; Result = true;}
			if (y < MinSize) {y = MinSize; Result = true;}
			
			return Result;
		}

		public static bool Check(ref TPointF p)
		{
			bool Result = false;
			if (p.X > MaxSize) {p.X = MaxSize; Result = true;}
			if (p.Y > MaxSize) {p.Y = MaxSize; Result = true;}
			if (p.X < MinSize) {p.X = MinSize; Result = true;}
			if (p.Y < MinSize) {p.Y = MinSize; Result = true;}
			
			return Result;
		}

		public static bool Check(ref RectangleF r)
		{
			if (r.Right < MinSize || r.Bottom < MinSize || r.Left > MaxSize || r.Top > MaxSize) return true;

			real x1 =r.Left < MinSize? MinSize: r.Left;
			real x2 =r.Right > MaxSize? MaxSize: r.Right;
			real y1 =r.Top < MinSize? MinSize: r.Top;
			real y2 =r.Bottom > MaxSize? MaxSize: r.Bottom;

			r = FlexCelRender.RectangleXY(x1, y1, x2, y2);
			return false;
		}

		public static bool Check(ref real x, ref real y, ref real w, ref real h)
		{
			real x2 = x + w;
			real y2 = y + h;

			if (x2 < MinSize || y2 < MinSize || x > MaxSize || y > MaxSize) return true;


			if (x < MinSize) {x = MinSize; w = x2 - x;}
			if (y < MinSize) {y = MinSize; h = y2 - y;}

			if (x2 > MaxSize) w = MaxSize - x;
			if (y2 > MaxSize) h = MaxSize - y;

			return false;
		}
	}
	#endregion


	#endregion
}
