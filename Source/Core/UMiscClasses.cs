using System;
using System.Reflection;
using System.Globalization;
using System.Text;
using System.Collections.Generic;

#if (MONOTOUCH)
    using Color = MonoTouch.UIKit.UIColor;
#endif

#if (WPF)
using System.Windows.Media;
using real = System.Double;
#else
using System.Drawing;
using real = System.Single;
#endif

namespace FlexCel.Core
{
    #region Anchors
    /// <summary>
	/// A rectangle holding image position in Excel internal units. Not for normal use, it can be used to do relative changes
	/// (for example, reducing the x to the the half).
	/// </summary>
	internal class TAbsoluteAnchorRect: ICloneable
	{
		#region Privates
		private TFlxAnchorType FAnchorType;
		private long Fx1;
		private long Fx2;
		private long Fy1;
		private long Fy2;
		#endregion

		/// <summary>
		/// Creates a new TAbsoluteAnchorRect instance.
		/// </summary>
		/// <param name="aAnchorType">Anchor type.</param>
		/// <param name="ax1">First x coord.</param>
		/// <param name="ay1">First y coord.</param>
		/// <param name="ax2">Second x coord.</param>
		/// <param name="ay2">Second y coord.</param>
		public TAbsoluteAnchorRect(TFlxAnchorType aAnchorType, long ax1, long ay1, long ax2, long ay2)
		{
			FAnchorType = aAnchorType;
			Fx1 = ax1;
			Fy1 = ay1;
			Fx2 = ax2;
			Fy2 = ay2;
		}

		/// <summary>
		/// Anchor type.
		/// </summary>
		public TFlxAnchorType AnchorType {get{return FAnchorType;} set{FAnchorType=value;}}

		/// <summary>
		/// X1 of the image.
		/// </summary>
		public long x1 {get{return Fx1;} set{Fx1=value;}}

		/// <summary>
		/// Y1 of the image.
		/// </summary>
		public long y1 {get{return Fy1;} set{Fy1=value;}}

		/// <summary>
		/// X2 of the image.
		/// </summary>
		public long x2 {get{return Fx2;} set{Fx2=value;}}

		/// <summary>
		/// Y2 of the image.
		/// </summary>
		public long y2 {get{return Fy2;} set{Fy2=value;}}

		#region ICloneable Members

		public object Clone()
		{
			return MemberwiseClone();
		}

		#endregion
	}

	/// <summary>
	/// A class to hold the offsets relative to the parent on grouped shapes.
	/// </summary>
	public class TChildAnchor: ICloneable
	{
		#region Privates
		private double FDx1;
		private double FDy1;
		private double FDx2;
		private double FDy2;
		#endregion

		/// <summary>
		/// Creates a new empty instance.
		/// </summary>
		public TChildAnchor(){}

		/// <summary>
		/// Creates an instance with the defined values.
		/// </summary>
		/// <param name="aDx1"></param>
		/// <param name="aDy1"></param>
		/// <param name="aDx2"></param>
		/// <param name="aDy2"></param>
		public TChildAnchor(double aDx1, double aDy1, double aDx2, double aDy2)
		{
			FDx1 = aDx1;
			FDy1 = aDy1;
			FDx2 = aDx2;
			FDy2 = aDy2;
		}

		/// <summary>
		/// Offset from the left on the parent, on percent of the total width of the parent.
		/// </summary>
		public double Dx1 {get {return FDx1;} set {FDx1 = value;}}

		/// <summary>
		/// Offset from the top on the parent, on percent of the total height of the parent.
		/// </summary>
		public double Dy1 {get {return FDy1;} set {FDy1 = value;}}
		
		/// <summary>
		/// Right coordinate of the box, in percent of the total width of the parent.
		/// </summary>
		public double Dx2 {get {return FDx2;} set {FDx2 = value;}}

        /// <summary>
        /// Bottom coordinate of the box, in percent of the total height of the parent.
        /// </summary>
        public double Dy2 { get { return FDy2; } set { FDy2 = value; } }

		#region ICloneable Members

		/// <summary>
		/// Returns a deep copy of the Anchor.
		/// </summary>
		/// <returns></returns>
		public object Clone()
		{
			return MemberwiseClone();
		}

		#endregion

		/// <summary>
		/// Returns true if both instances of the objects contain the same values. Instances might be different, this method will return
		/// if their values are equal. Instances can be null.
		/// </summary>
		/// <param name="a1">First instance to compare.</param>
		/// <param name="a2">Second instance to compare.</param>
		/// <returns></returns>
		public static bool EqualValues (TChildAnchor a1, TChildAnchor a2)
		{
			if (a1 == null) return a2 == null;
			if (a2 == null) return false;

			return
				a1.FDx1 == a2.FDx1 &&
				a1.FDy1 == a2.FDy1 &&
				a1.FDx2 == a2.FDx2 &&
				a1.FDy2 == a2.FDy2;
		}

        /// <summary>
        /// Returns true if both objects have the same data.
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public override bool Equals(object obj)
        {
            return EqualValues(this, obj as TChildAnchor);
        }

        /// <summary>
        /// Returns the hashcode for the object.
        /// </summary>
        /// <returns></returns>
        public override int GetHashCode()
        {
            return HashCoder.GetHashObj(
                FDx1,
                FDy1,
                FDx2,
                FDy2);
        }
	}

	/// <summary>
	/// Image Anchor information.
	/// </summary>
	public class TClientAnchor: ICloneable
	{
		#region Privates
		private bool FChartCoords;
		private TFlxAnchorType FAnchorType;
		private int FCol1;
		private int FDx1;
		private int FRow1;
		private int FDy1;
		private int FCol2;
		private int FDx2;
		private int FRow2;
		private int FDy2;

		private TChildAnchor FChildAnchor;
		#endregion

		/// <summary>
		/// If true, this object is inside a chart, and columns and rows range from 0 to 4000.
		/// </summary>
		public bool ChartCoords {get {return FChartCoords;} set{FChartCoords = value;}}

		/// <summary>
		/// How the image behaves when copying/inserting cells.
		/// </summary>
		public TFlxAnchorType AnchorType {get {return FAnchorType;} set {FAnchorType=value;}}
		/// <summary>
		/// First column of object
		/// </summary>
		public int Col1 {get {return FCol1;} set {FCol1=value;}}
		/// <summary>
		/// Delta x of image, on 1/1024 of a cell.  0 means totally at the left, 512 on half of the cell, 1024 means at the left of next cell.
		/// </summary>
		public int Dx1 {get {return FDx1<0? 0: FDx1>1024? 1024: FDx1;} set {FDx1 = value<0? 0: value>1024? 1024: value;}}
		/// <summary>
		/// First Row of object.
		/// </summary>
		public int Row1 {get {return FRow1;} set {FRow1=value;}}
		/// <summary>
		/// Delta y of image on 1/255 of a cell. 0 means totally at the top, 128 on half of the cell, 255 means at the top of next cell.
		/// </summary>
        public int Dy1 { get { return FDy1 < 0 ? 0 : FDy1 > 255 ? 255 : FDy1; } set { FDy1 = value < 0 ? 0 : value > 255 ? 255 : value; } }
		/// <summary>
		/// Last column of object.
		/// </summary>
		public int Col2 {get {return FCol2;} set {FCol2=value;}}
		/// <summary>
		/// Delta x of image, on 1/1024 of a cell.  0 means totally at the left, 512 on half of the cell, 1024 means at the left of next cell.
		/// </summary>
		public int Dx2 {get {return FDx2<0? 0: FDx2>1024? 1024: FDx2;} set {FDx2 = value<0? 0: value>1024? 1024: value;}}
		/// <summary>
		/// Last row of object.
		/// </summary>
		public int Row2 {get {return FRow2;} set {FRow2=value;}}
		/// <summary>
		/// Delta y of image on 1/255 of a cell. 0 means totally at the top, 128 on half of the cell, 255 means at the top of next cell.
		/// </summary>
		public int Dy2 {get {return FDy2<0? 0: FDy2>255? 255: FDy2;} set {FDy2=value<0? 0: value>255? 255: value;}}

		/// <summary>
		/// Returns the offset on the parent system for the image, when it is grouped.
		/// For example, if the parent shape on the group is 100 px wide, and ChildAnchor has a dx
		/// of 0.5, the image starts 50px to the right of the parent. If the shape is not grouped 
		/// or it is the shape on top of the group, ChildAnchor is null. When this member is
		/// not null, other values on ClientAnchor have no meaning.
		/// </summary>
		/// <remarks>
		/// This offset is relative to the parent shape. If there are 2 parent shapes before on the hierarchy,
		/// this object applies to the parent, not the grandparent.
		/// </remarks>
        public TChildAnchor ChildAnchor { get { return FChildAnchor; } set { FChildAnchor = value; } }
        
		/// <summary>
		/// Creates an Empty ClientAnchor.
		/// </summary>
		public TClientAnchor(){}

		/// <summary>
		/// Creates a new ClientAnchor object, based on cell coords. Does not take in count actual image size.
		/// </summary>
		/// <param name="aAnchorType">How the image behaves when copying/inserting cells.</param>
		/// <param name="aCol1">First column of object</param>
		/// <param name="aDx1">Delta x of image, on 1/1024 of a cell.  0 means totally at the left, 512 on half of the cell, 1024 means at the left of next cell.</param>
		/// <param name="aRow1">First Row of object.</param>
		/// <param name="aDy1">Delta y of image on 1/255 of a cell. 0 means totally at the top, 128 on half of the cell, 255 means at the top of next cell.</param>
		/// <param name="aCol2">Last column of object.</param>
		/// <param name="aDx2">Delta x of image, on 1/1024 of a cell.  0 means totally at the left, 512 on half of the cell, 1024 means at the left of next cell.</param>
		/// <param name="aRow2">Last row of object.</param>
		/// <param name="aDy2">Delta y of image on 1/255 of a cell. 0 means totally at the top, 128 on half of the cell, 255 means at the top of next cell.</param>
		public TClientAnchor(TFlxAnchorType aAnchorType, int aRow1, int aDy1, int aCol1, int aDx1, int aRow2, int aDy2, int aCol2, int aDx2):
			this(false, aAnchorType, aRow1, aDy1, aCol1, aDx1, aRow2, aDy2, aCol2, aDx2)
		{
		}

		/// <summary>
		/// Creates a new ClientAnchor object, based on cell coords. Ignores actual image size.
		/// </summary>
		/// <param name="aAnchorType">How the image behaves when copying/inserting cells.</param>
		/// <param name="aCol1">First column of object.</param>
		/// <param name="aDx1">Delta x of image, on 1/1024 of a cell.  0 means totally at the left, 512 on half of the cell, 1024 means at the left of next cell.</param>
		/// <param name="aRow1">First Row of object.</param>
		/// <param name="aDy1">Delta y of image on 1/255 of a cell. 0 means totally at the top, 128 on half of the cell, 255 means at the top of next cell.</param>
		/// <param name="aCol2">Last column of object.</param>
		/// <param name="aDx2">Delta x of image, on 1/1024 of a cell.  0 means totally at the left, 512 on half of the cell, 1024 means at the left of next cell.</param>
		/// <param name="aRow2">Last row of object.</param>
		/// <param name="aDy2">Delta y of image on 1/255 of a cell. 0 means totally at the top, 128 on half of the cell, 255 means at the top of next cell.</param>
		/// <param name="aChartCoords">If true, the object is inside a chart and rows and columns range from 0 to 4000.</param>
		public TClientAnchor(bool aChartCoords, TFlxAnchorType aAnchorType, int aRow1, int aDy1, int aCol1, int aDx1, int aRow2, int aDy2, int aCol2, int aDx2)
		{
			//Anchors in objects inside charts might have row and col up to 4000.
			if (!aChartCoords)
			{
				if ((aRow1<0)||(aRow1>FlxConsts.Max_Rows+1)) throw new ArgumentOutOfRangeException(FlxMessages.GetString(FlxErr.ErrInvalidRow, aRow1)); //2 Args constructor not compatible with cf!
				if ((aRow2<0)||(aRow2>FlxConsts.Max_Rows+1)) throw new ArgumentOutOfRangeException(FlxMessages.GetString(FlxErr.ErrInvalidRow, aRow2)); 
				if ((aCol1<0)||(aCol1>FlxConsts.Max_Columns+1)) throw new ArgumentOutOfRangeException(FlxMessages.GetString(FlxErr.ErrInvalidColumn, aCol1)); 
				if ((aCol2<0)||(aCol2>FlxConsts.Max_Columns+1)) throw new ArgumentOutOfRangeException(FlxMessages.GetString(FlxErr.ErrInvalidColumn, aCol2)); 
			}
			FChartCoords = aChartCoords;
			
			AnchorType=aAnchorType;
			Col1=aCol1;
			Dx1=aDx1;
			Row1=aRow1;
			Dy1=aDy1;
			Col2=aCol2;
			Dx2=aDx2;
			Row2=aRow2;
			Dy2=aDy2;
		}

        private static double Rh(IRowColSize Workbook, int Row)
        {
            if ((Row > 0) && (Row <= FlxConsts.Max_Rows + 1) && (!Workbook.IsEmptyRow(Row))) return Workbook.GetRowHeight(Row, true) / FlxConsts.RowMult;
            else
                return Workbook.DefaultRowHeight / FlxConsts.RowMult;
        }

        private static double Rhi(IRowColSize Workbook, int Row)
        {
            if ((Row > 0) && (Row <= FlxConsts.Max_Rows + 1) && (!Workbook.IsEmptyRow(Row))) return Workbook.GetRowHeight(Row, true);
            else
                return Workbook.DefaultRowHeight;
        }

        private static double Cw(IRowColSize Workbook, int Col)
        {
            if ((Col > 0) && (Col <= FlxConsts.Max_Columns + 1))
                return Workbook.GetColWidth(Col, true) / ExcelMetrics.ColMult(Workbook);
            else
                return Workbook.DefaultColWidth / ExcelMetrics.ColMult(Workbook);
        }

        private static double Rhp(IRowColSize Workbook, int Row)
        {
            double RowMultDisplay = ExcelMetrics.RowMultDisplay(Workbook) * 100F / FlxConsts.DispMul;
            if ((Row > 0) && (Row <= FlxConsts.Max_Rows + 1) && (!Workbook.IsEmptyRow(Row))) return Workbook.GetRowHeight(Row, true) / RowMultDisplay;
            else
                return Workbook.DefaultRowHeight / RowMultDisplay;
        }

        private static double Cwp(IRowColSize Workbook, int Col)
		{
			double ColMultDisplay = ExcelMetrics.ColMultDisplay(Workbook) * 100F / FlxConsts.DispMul;
			if ((Col>0)&&(Col<=FlxConsts.Max_Columns+1))
				return Workbook.GetColWidth(Col, true)/ColMultDisplay;
			else
				return Workbook.DefaultColWidth/ColMultDisplay;
		}

		/// <summary>
		/// Creates a new image based on the image size.
		/// </summary>
		/// <param name="aAnchorType">How the image behaves when copying/inserting cells.</param>
		/// <param name="aRow1">Row where to insert the image.</param>
		/// <param name="aCol1">Column where to insert the image.</param>
		/// <param name="aPixDy1">Delta in pixels that the image is moved from aRow1.</param>
		/// <param name="aPixDx1">Delta in pixels that the image is moved from aCol1.</param>
		/// <param name="height">Height in pixels.</param>
		/// <param name="width">Width in pixels.</param>
		/// <param name="Workbook">ExcelFile with the workbook, used to calculate the cells.</param>
        public TClientAnchor(TFlxAnchorType aAnchorType, int aRow1, int aPixDy1, int aCol1, int aPixDx1, int height, int width, IRowColSize Workbook)
		{

			AnchorType=aAnchorType;
			if (Workbook==null) FlxMessages.ThrowException(FlxErr.ErrNotConnected);

            double ddx1;
            double ddy1;
            CalcFromPixels(aRow1, aPixDy1, aCol1, aPixDx1, Workbook, out ddx1, out ddy1, out FRow1, out FCol1, out FDy1, out FDx1, true);

			int r=Row1; double h=Rh(Workbook, Row1)-ddy1;
			while ((int)Math.Round(h)<height && r <= FlxConsts.Max_Rows + 2) 
			{
				r++;
				h+= Rh(Workbook, r);
			}
			Row2=r;
			double lrh2 = Rh(Workbook, r);
            if (lrh2 == 0) Dy2 = 0; else Dy2 = (int)Math.Round((Rh(Workbook, r) - (h - height)) / lrh2 * 255);

			int c=Col1; double w=Cw(Workbook, Col1)-ddx1;
			while ((int)Math.Round(w)<width && c <= FlxConsts.Max_Columns + 2)
			{
				c++;
				w+=Cw(Workbook, c);
			}
			Col2=c;
			double lcw2 = Cw(Workbook, c);
			if (lcw2 == 0) Dx2 = 0; else Dx2=(int)Math.Round((Cw(Workbook, c)-(w-width))/lcw2*1024);

			if (Row2>FlxConsts.Max_Rows+1)
			{
				Row1=FlxConsts.Max_Rows+1;
				
				ddy1=height;
				while (Row1 > 0 && ddy1 > 0)
				{
					ddy1 -= Rh(Workbook, Row1);
					Row1--;
				}
				Row1++;
				double lrh1 = Rh(Workbook, Row1);
                if (lrh1 == 0 || ddy1 > 0) Dy1 = 0; else Dy1 = (int)Math.Round(-255.0 * ddy1 / lrh1);

				Row2=FlxConsts.Max_Rows+1;
				Dy2 = 255 - 1;
			}

			if (Col2>FlxConsts.Max_Columns+1)
			{
				Col1=FlxConsts.Max_Columns+1;
				ddx1=width;
				while (Col1 > 0 && ddx1 > 0)
				{
					ddx1 -= Cw(Workbook, Col1);
					Col1--;
				}
				Col1++;
				double lcw1 = Cw(Workbook, Col1);
				if (lcw1 == 0 | ddx1 > 0) Dx1 = 0; else Dx1=(int)Math.Round(-1024.0*ddx1/lcw1);

				Col2=FlxConsts.Max_Columns+1;
				Dx2 = 1024 - 1;
			}

			//Just in case of an image bigger than the spreadsheet...
			if (Col1<1) Col1=1;
			if (Row1<1) Row1=1;
		}

        /// <summary>
        /// Creates a new image based on the image size.
        /// </summary>
        /// <param name="aAnchorType">How the image behaves when copying/inserting cells.</param>
        /// <param name="aRow1">Row where to insert the image.</param>
        /// <param name="aCol1">Column where to insert the image.</param>
        /// <param name="aPixDy1">Delta in pixels that the image is moved from aRow1.</param>
        /// <param name="aPixDx1">Delta in pixels that the image is moved from aCol1.</param>
        /// <param name="aRow2">Row where to insert the image.</param>
        /// <param name="aCol2">Column where to insert the image.</param>
        /// <param name="aPixDy2">Delta in pixels that the image is moved from aRow1.</param>
        /// <param name="aPixDx2">Delta in pixels that the image is moved from aCol1.</param>
        /// <param name="Workbook">ExcelFile with the workbook, used to calculate the cells.</param>
        public TClientAnchor(TFlxAnchorType aAnchorType, int aRow1, int aPixDy1, int aCol1, int aPixDx1, int aRow2, int aPixDy2, int aCol2, int aPixDx2, IRowColSize Workbook):
            this (aAnchorType, aRow1, aPixDy1, aCol1, aPixDx1, aRow2, aPixDy2, aCol2, aPixDx2, Workbook, true)
        {}

        internal TClientAnchor(TFlxAnchorType aAnchorType, int aRow1, int aPixDy1, int aCol1, int aPixDx1, int aRow2, int aPixDy2, int aCol2, int aPixDx2, IRowColSize Workbook, bool SpanCells)
        {
            AnchorType = aAnchorType;
            if (Workbook == null) FlxMessages.ThrowException(FlxErr.ErrNotConnected);

            double ddx1;
            double ddy1;
            CalcFromPixels(aRow1, aPixDy1, aCol1, aPixDx1, Workbook, out ddx1, out ddy1, out FRow1, out FCol1, out FDy1, out FDx1, SpanCells);
            CalcFromPixels(aRow2, aPixDy2, aCol2, aPixDx2, Workbook, out ddx1, out ddy1, out FRow2, out FCol2, out FDy2, out FDx2, SpanCells);

            if (Col1 < 1) Col1 = 1;
            if (Row1 < 1) Row1 = 1;

            if (Row2 > FlxConsts.Max_Rows + 1)
            {
                Row2 = FlxConsts.Max_Rows + 1;
                Dy2 = 255 - 1;
            }

            if (Col2 > FlxConsts.Max_Columns + 1)
            {
                Col2 = FlxConsts.Max_Columns + 1;
                Dx2 = 1024 - 1;
            }
        }

        private static void CalcFromPixels(int aRow1, int aPixDy1, int aCol1, int aPixDx1,
            IRowColSize Workbook, out double ddx1, out double ddy1,
            out int Row1, out int Col1, out int Dy1, out int Dx1, bool SpanCells)
        {
            Row1 = aRow1; Col1 = aCol1; ddx1 = aPixDx1; ddy1 = aPixDy1;
            //If delta spans more than one cell, advance the cells.
            if (SpanCells)
            {
                SpanCellsToNext(Workbook, ref ddx1, ref ddy1, ref Row1, ref Col1);
            }

            if (Row1 < 1 || ddy1 < 0) { Row1 = 1; ddy1 = 0; }
            if (Col1 < 1 || ddx1 < 0) { Col1 = 1; ddx1 = 0; }

            //Convert from pixels to percent of cells.
            double lcw1 = Cw(Workbook, Col1);
            if (ddx1 > lcw1) ddx1 = lcw1;
            if (lcw1 == 0) Dx1 = 0; else Dx1 = (int)Math.Round(1024.0 * ddx1 / lcw1);
            double lrh1 = Rh(Workbook, Row1);
            if (ddy1 > lrh1) ddy1 = lrh1;
            if (lrh1 == 0) Dy1 = 0; else Dy1 = (int)Math.Round(255.0 * ddy1 / lrh1);

        }

        private static void SpanCellsToNext(IRowColSize Workbook, ref double ddx1, ref double ddy1, ref int Row1, ref int Col1)
        {
            while (ddx1 > Cw(Workbook, Col1) && Col1 <= FlxConsts.Max_Columns + 2)
            {
                ddx1 -= Cw(Workbook, Col1);
                Col1++;
            }
            while (Col1 > 1 && ddx1 < 0)
            {
                Col1--;
                ddx1 += Cw(Workbook, Col1);
            }

            while (ddy1 > Rh(Workbook, Row1) && Row1 <= FlxConsts.Max_Rows + 2)
            {
                ddy1 -= Rh(Workbook, Row1);
                Row1++;
            }

            while (Row1 > 1 && ddy1 < 0)
            {
                Row1--;
                ddy1 += Rh(Workbook, Row1);
            }
        }
		
		/// <summary>
		/// Returns a COPY of the anchor with rows and cols incremented by one
		/// </summary>
		/// <returns></returns>
		public TClientAnchor Inc()
		{
			TClientAnchor Result= (TClientAnchor) Clone();
			Result.Row1++;
			Result.Col1++;
			Result.Row2++;
			Result.Col2++;
			return Result;
		}

		/// <summary>
		/// Returns a COPY of the anchor with rows and cols decremented by one
		/// </summary>
		/// <returns></returns>
		public TClientAnchor Dec()
		{
			TClientAnchor Result= (TClientAnchor) Clone();
			Result.Row1--;
			Result.Col1--;
			Result.Row2--;
			Result.Col2--;
			return Result;
		}

		/// <summary>
		/// Length of the Serialized array.
        /// This serialized array is in biff8 format, so it doesn't allow more than 65536 rows.
        /// </summary>
        internal static byte Biff8Length { get { return 18; } }

		/// <summary>
        /// Calculates the width and height of the image in pixels. MAKE SURE THE ACTIVESHEET IN WORKBOOK IS THE CORRECT ONE.
		/// </summary>
		/// <param name="height">Will return the height of the object.</param>
		/// <param name="width">Will return the width of the object.</param>
		/// <param name="Workbook">Workbook used to know the column widths and row heights.</param>
        public void CalcImageCoords(ref double height, ref double width, IRowColSize Workbook)
        {
            height = 0;
            for (int r = Row1 + 1; r < Row2; r++)
                height += Rh(Workbook, r);
            height += Rh(Workbook, Row1) * (255 - Dy1) / 255.0;
            if (Row2 == Row1)
                height -= Rh(Workbook, Row2) * (255 - Dy2) / 255.0;
            else
                height += Rh(Workbook, Row2) * (Dy2) / 255.0;


            width = 0;
            for (int c = Col1 + 1; c < Col2; c++)
                width += Cw(Workbook, c);
            width += Cw(Workbook, Col1) * (1024 - Dx1) / 1024.0;
            if (Col2 == Col1)
                width -= Cw(Workbook, Col2) * (1024 - Dx2) / 1024.0;
            else
                width += Cw(Workbook, Col2) * (Dx2) / 1024.0;

        }

		/// <summary>
		/// Calculates the width and height of the image in Points.
		/// </summary>
		/// <param name="height">Will return the height of the object.</param>
		/// <param name="width">Will return the width of the object.</param>
		/// <param name="Workbook">Workbook used to know the column widths and row heights.</param>
        public void CalcImageCoordsInPoints(ref double height, ref double width, IRowColSize Workbook)
        {
            height = 0;
            for (int r = Row1 + 1; r < Row2; r++)
                height += Rhp(Workbook, r);
            height += Rhp(Workbook, Row1) * (255 - Dy1) / 255.0;
            if (Row2 == Row1)
                height -= Rhp(Workbook, Row2) * (255 - Dy2) / 255.0;
            else
                height += Rhp(Workbook, Row2) * (Dy2) / 255.0;


            width = 0;
            for (int c = Col1 + 1; c < Col2; c++)
                width += Cwp(Workbook, c);
            width += Cwp(Workbook, Col1) * (1024 - Dx1) / 1024.0;
            if (Col2 == Col1)
                width -= Cwp(Workbook, Col2) * (1024 - Dx2) / 1024.0;
            else
                width += Cwp(Workbook, Col2) * (Dx2) / 1024.0;

        }

        /// <summary>
        /// This might be called with a workbook or a sheet. If with a workbook, verify the sheet is the correct one.
        /// </summary>
        /// <param name="Workbook"></param>
        /// <returns></returns>
        internal double CalcImageHeightInternal(IRowColSize Workbook)
        {
            double Result = 0;
            for (int r = Row1 + 1; r < Row2; r++)
                Result += Rhi(Workbook, r);
            Result += Rhi(Workbook, Row1) * (255 - Dy1) / 255.0;
            if (Row2 == Row1)
                Result -= Rhi(Workbook, Row2) * (255 - Dy2) / 255.0;
            else
                Result += Rhi(Workbook, Row2) * (Dy2) / 255.0;

            return Result;
        }


        /// <summary>
		/// Returns the offset of the object in pixels from the left of the cell.
		/// </summary>
		/// <param name="Workbook">ExcelFile where this object is placed</param>
		/// <returns>Offset in pixels.</returns>
        public int Dx1Pix(IRowColSize Workbook)
		{
			return (int)Math.Round(Dx1*Cw(Workbook, Col1)/1024.0);
		}

		/// <summary>
		/// Returns the offset of the object in pixels from the top of the cell.
		/// </summary>
		/// <param name="Workbook">ExcelFile where this object is placed</param>
		/// <returns>Offset in pixels.</returns>
        public int Dy1Pix(IRowColSize Workbook)
        {
            return (int)Math.Round(Dy1 * Rh(Workbook, Row1) / 255.0);
        }

        /// <summary>
        /// Returns the offset of the object in pixels from the left of the cell.
        /// </summary>
        /// <param name="Workbook">ExcelFile where this object is placed</param>
        /// <returns>Offset in pixels.</returns>
        public int Dx2Pix(IRowColSize Workbook)
        {
            return (int)Math.Round(Dx2 * Cw(Workbook, Col2) / 1024.0);
        }

        /// <summary>
        /// Returns the offset of the object in pixels from the top of the cell.
        /// </summary>
        /// <param name="Workbook">ExcelFile where this object is placed</param>
        /// <returns>Offset in pixels.</returns>
        public int Dy2Pix(IRowColSize Workbook)
        {
            return (int)Math.Round(Dy2 * Rh(Workbook, Row2) / 255.0);
        }

        /// <summary>
		/// Returns the offset of the object in points from the left of the cell. This is used for display, and is not the exact conversion from Dx1Pix.
		/// </summary>
		/// <param name="Workbook">ExcelFile where this object is placed</param>
		/// <returns>Offset in points (1/72 of an inch).</returns>
        public float Dx1Points(IRowColSize Workbook)
		{
			return (float)(Dx1*Cwp(Workbook, Col1)/1024f);
		}

		/// <summary>
		/// Returns the offset of the object in points from the top of the cell.
		/// </summary>
		/// <param name="Workbook">ExcelFile where this object is placed</param>
		/// <returns>Offset in points (1/72 of an inch).</returns>
        public float Dy1Points(IRowColSize Workbook)
		{
			return (float)(Dy1*Rhp(Workbook, Row1)/255f);
		}

        /// <summary>
        /// Returns the offset of the object in points from the left of the cell.
        /// </summary>
        /// <param name="Workbook">ExcelFile where this object is placed</param>
        /// <returns>Offset in points (1/72 of an inch).</returns>
        public float Dx2Points(IRowColSize Workbook)
        {
            return (float)(Dx2 * Cwp(Workbook, Col2) / 1024f);
        }

        /// <summary>
        /// Returns the offset of the object in points from the top of the cell.
        /// </summary>
        /// <param name="Workbook">ExcelFile where this object is placed</param>
        /// <returns>Offset in points (1/72 of an inch).</returns>
        public float Dy2Points(IRowColSize Workbook)
        {
            return (float)(Dy2 * Rhp(Workbook, Row2) / 255f);
        }

		#region ICloneable Members

		/// <summary>
		/// Creates a copy of the Anchor
		/// </summary>
		/// <returns>Anchor copy.</returns>
		public object Clone()
		{
			TClientAnchor Result = (TClientAnchor)MemberwiseClone();
			if (ChildAnchor != null) Result.ChildAnchor = (TChildAnchor)ChildAnchor.Clone();
			return Result;
		}

		#endregion

		/// <summary>
		/// Returns true if both instances of the objects contain the same values. Instances might be different, this method will return
		/// if their values are equal. Instances can be null. This method is equivalent to TClientAnchor.Equals(a,b), and kept for backwards compatibility.
		/// </summary>
		/// <param name="a1">First instance to compare.</param>
		/// <param name="a2">Second instance to compare.</param>
		/// <returns></returns>
		public static bool EqualValues (TClientAnchor a1, TClientAnchor a2)
		{
			if (a1 == null) return a2 == null;
			if (a2 == null) return false;

			return
				a1.FChartCoords == a2.FChartCoords &&
				a1.FAnchorType == a2.FAnchorType &&
				a1.FCol1 == a2.FCol1 &&
				a1.FDx1 == a2.FDx1 &&
				a1.FRow1 == a2.FRow1 &&
				a1.FDy1 == a2.FDy1 &&
				a1.FCol2 == a2.FCol2 &&
				a1.FDx2 == a2.FDx2 &&
				a1.FRow2 == a2.FRow2 &&
				a1.FDy2 == a2.FDy2 &&

				TChildAnchor.EqualValues(a1.ChildAnchor, a2.ChildAnchor);
		}

        /// <summary>
        /// Returns true if both objects have the same data.
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public override bool Equals(object obj)
        {
            return EqualValues(this, obj as TClientAnchor);
        }

        /// <summary>
        /// Returns the hash code for the object.
        /// </summary>
        /// <returns></returns>
        public override int GetHashCode()
        {
            return HashCoder.GetHashObj(
                FChartCoords,
                FAnchorType,
                FCol1,
                FDx1,
                FRow1,
                FDy1,
                FCol2,
                FDx2,
                FRow2,
                FDy2,
                ChildAnchor);
        }

        /// <summary>
        /// Returns true if the target anchor is inside or equal to this one.
        /// </summary>
        /// <param name="targetAnchor">Anchor to test.</param>
        /// <returns></returns>
        public bool Contains(TClientAnchor targetAnchor)
        {
            if (targetAnchor.Row1 < Row1 || targetAnchor.Col1 < Col1) return false;
            if (targetAnchor.Row2 > Row2 || targetAnchor.Col2 > Col2) return false;
            if (targetAnchor.Row1 == Row1 && targetAnchor.Dy1 < Dy1) return false;
            if (targetAnchor.Row2 == Row2 && targetAnchor.Dy2 > Dy2) return false;
            if (targetAnchor.Col1 == Col1 && targetAnchor.Dx1 < Dx1) return false;
            if (targetAnchor.Col2 == Col2 && targetAnchor.Dx2 > Dx2) return false;
            return true;
        }

    }

	/// <summary>
	/// Image information for an image inside a header or footer.
	/// </summary>
	public class THeaderOrFooterAnchor: ICloneable
	{
		private long FWidth;
		private long FHeight;

		/// <summary>
		/// Creates a new Anchor for a Header or footer image.
		/// </summary>
		/// <param name="aWidth">Width of the image in pixels.</param>
		/// <param name="aHeight">Height of the image in pixels.</param>
		public THeaderOrFooterAnchor(long aWidth, long aHeight)
		{
			Width = aWidth;
			Height = aHeight;
		}
		/// <summary>
		/// Width of the image in pixels.
		/// </summary>
		public long Width {get{return FWidth;} set{FWidth=value;}}

		/// <summary>
		/// Height of the image in pixels.
		/// </summary>
		public long Height {get{return FHeight;} set{FHeight=value;}}

		/// <summary>
		/// Length of the Serialized array.
		/// </summary>
		public byte Length {get {return 8;}}

		/// <summary>
		/// All the data as a byte array.
		/// </summary>
		public byte[] GetData()
		{
			byte[] Result=new byte[Length];
			BitConverter.GetBytes((UInt32)Width).CopyTo(Result,0);
			BitConverter.GetBytes((UInt32)Height).CopyTo(Result,4);
			return Result;
		}


		#region ICloneable Members

		/// <summary>
		/// Returns a clone of the anchor.
		/// </summary>
		/// <returns></returns>
		public object Clone()
		{
			return new THeaderOrFooterAnchor(Width, Height);
		}

		#endregion

		/// <summary>
		/// Returns true if both instances of the objects contain the same values. Instances might be different, this method will return
		/// if their values are equal. Instances can be null.
		/// </summary>
		/// <param name="a1">First instance to compare.</param>
		/// <param name="a2">Second instance to compare.</param>
		/// <returns></returns>
		public static bool EqualValues (THeaderOrFooterAnchor a1, THeaderOrFooterAnchor a2)
		{
			if (a1 == null) return a2 == null;
			if (a2 == null) return false;

			return
				a1.FHeight == a2.FHeight &&
				a1.FWidth == a2.FWidth;
		}

    }
    #endregion

    #region Margins and ranges

    /// <summary>
	/// Sheet margin for printing, in inches.
	/// </summary>
	public class TXlsMargins
	{
		#region Privates
		private double FLeft;
		private double FTop;
		private double FRight;
		private double FBottom;
		private double FHeader;
		private double FFooter;
		#endregion

		/// <summary>
		/// Left Margin in inches.
		/// </summary>
		public double Left {get {return FLeft;} set {FLeft=value;}}


		/// <summary>
		/// Top Margin in inches.
		/// </summary>
		public double Top {get {return FTop;} set {FTop=value;}}

		/// <summary>
		/// Right Margin in inches.
		/// </summary>
		public double Right {get {return FRight;} set {FRight=value;}}

		/// <summary>
		/// Bottom Margin in inches.
		/// </summary>
		public double Bottom {get {return FBottom;} set {FBottom=value;}}
        
		/// <summary>
		/// Header Margin in inches. Space for the header at top of page, it is taken from Top margin.
		/// </summary>
		public double Header {get {return FHeader;} set {FHeader=value;}}

		/// <summary>
		/// Footer Margin in inches. Space for the footer at bottom of page, it is taken from Bottom margin.
		/// </summary>
		public double Footer {get {return FFooter;} set {FFooter=value;}}

		/// <summary>
		/// Creates default Margins
		/// </summary>
		public TXlsMargins(){}

		/// <summary>
		/// Creates Personalized Margins. All units are in inches.
		/// </summary>
		/// <param name="aLeft">Left margin in inches.</param>
		/// <param name="aTop">Top margin in inches.</param>
		/// <param name="aRight">Right margin in inches.</param>
		/// <param name="aBottom">Bottom margin in inches.</param>
		/// <param name="aHeader">Header margin in inches.</param>
		/// <param name="aFooter">Footer margin in inches.</param>
		public TXlsMargins (double aLeft, double aTop, double aRight, double aBottom, double aHeader, double aFooter)
		{
			Left=aLeft;
			Top=aTop;
			Right=aRight;
			Bottom=aBottom;
			Header=aHeader;
			Footer=aFooter;
		}
	}

	/// <summary>
	/// An Excel Cell range, 1-based.
	/// </summary>
	public class TXlsCellRange: ICloneable
	{
		/// <summary>
		/// First column on range.
		/// </summary>
		public int Left;

		/// <summary>
		/// First row on range.
		/// </summary>
		public int Top;

		/// <summary>
		/// Last column on range.
		/// </summary>
		public int Right;
		
		/// <summary>
		/// Last row on range.
		/// </summary>
		public int Bottom;

		/// <summary>
		/// Creates a new TXlsCellRange class.
		/// </summary>
		/// <param name="aFirstRow">First row on range.</param>
		/// <param name="aFirstCol">First column on range.</param>
		/// <param name="aLastRow">Last row on range.</param>
		/// <param name="aLastCol">Last column on range.</param>
		public TXlsCellRange(int aFirstRow, int aFirstCol, int aLastRow, int aLastCol)
		{
			Top=aFirstRow;
			Left=aFirstCol;
			Bottom=aLastRow;
			Right=aLastCol;
		}

        /// <summary>
        /// Creates a cell range based in an Excel range string, like "A1:A10"
        /// </summary>
        /// <param name="rangeDef">Definition for the range, in Excel A1 notation. For example A1:B3</param>
        public TXlsCellRange(string rangeDef)
        {
            if (rangeDef == null) FlxMessages.ThrowException(FlxErr.ErrInvalidRange, rangeDef);
            ExcelFile LocalXls; //not needed, there are no external refs here.
            int sheet1, sheet2;
            TCellAddress.ParseAddress(null, rangeDef.Trim(), -1, out LocalXls, out sheet1, out sheet2, out Top, out Left, out Bottom, out Right);
            if (sheet1 != -1 || sheet2 != -1) FlxMessages.ThrowException(FlxErr.ErrInvalidRange, rangeDef);
        }

		/// <summary>
		/// Creates an empty Cell range.
		/// </summary>
		public TXlsCellRange()
		{
			Left=0;
			Top=0;
			Right=0;
			Bottom=0;
		}

		/// <summary>
		/// Creates a range with all cells on the sheet (65536 rows x 256 columns in Excel 97-2003)
		/// </summary>
		public static TXlsCellRange FullRange()
		{
			return new TXlsCellRange(1, 1, FlxConsts.Max_Rows+1, FlxConsts.Max_Columns+1);
		}

		/// <summary>
		/// Returns true if the range has only one cell.
		/// </summary>
		public virtual bool IsOneCell
		{
			get
			{
				return Top == Bottom && Left == Right;
			}
		}

		/// <summary>
		/// Returns the range transposed, rows by columns.
		/// </summary>
		/// <returns></returns>
		public TXlsCellRange Transpose()
		{
			return new TXlsCellRange(Left, Top, Right, Bottom);
		}
		
        #region ICloneable Members

		/// <summary>
		/// Returns a copy of the original range.
		/// </summary>
		/// <returns></returns>
		public object Clone()
		{
			return MemberwiseClone();
		}

		/// <summary>
		/// Creates a new range with the start at newTopRow, newLeftCol.
		/// </summary>
		/// <param name="newTopRow"></param>
		/// <param name="newLeftCol"></param>
		/// <returns></returns>
		public TXlsCellRange Offset(int newTopRow, int newLeftCol)
		{
			TXlsCellRange Result = (TXlsCellRange)Clone();
			Result.Top=newTopRow;
			Result.Bottom=Bottom-Top+newTopRow;
			Result.Left=newLeftCol;
			Result.Right=Right-Left+newLeftCol;
			return Result;
		}

		/// <summary>
		/// Number of rows on the range.
		/// </summary>
		public int RowCount
		{
			get
			{
				return Bottom-Top+1;
			}
		}

		/// <summary>
		/// Number of columns on the range.
		/// </summary>
		public int ColCount
		{
			get
			{
				return Right-Left+1;
			}
		}

		/// <summary>
		/// True if the specified row is in the range
		/// </summary>
		/// <param name="row">Row we want to know if is on the range.</param>
		/// <returns></returns>
		public bool HasRow(int row)
		{
			return (row>=Top)&&(row<=Bottom);
		}

		/// <summary>
		/// True if the specified column is in the range
		/// </summary>
		/// <param name="col">Column we want to know if is on the range.</param>
		/// <returns></returns>
		public bool HasCol(int col)
		{
			return (col>=Left)&&(col<=Right);
		}

		/// <summary>
		/// Returns a COPY of the range decremented by one.
		/// </summary>
		/// <returns></returns>
		public TXlsCellRange Dec()
		{
			TXlsCellRange Result= (TXlsCellRange)Clone();
			Result.Left--;
			Result.Right--;
			Result.Top--;
			Result.Bottom--;
			return Result;
		}

		/// <summary>
		/// Returns a COPY of the range incremented by one.
		/// </summary>
		/// <returns></returns>
		public TXlsCellRange Inc()
		{
			TXlsCellRange Result= (TXlsCellRange)Clone();
			Result.Left++;
			Result.Right++;
			Result.Top++;
			Result.Bottom++;
			return Result;
		}

		#region Compare
		/// <summary>
		/// Returns true if both objects are equal.
		/// </summary>
		/// <param name="obj"></param>
		/// <returns></returns>
		public override bool Equals(object obj)
		{
			TXlsCellRange o2 = obj as TXlsCellRange;
			if (o2 == null) return false;
			return  Left == o2.Left && Top == o2.Top && Right == o2.Right && Bottom == o2.Bottom;
		}

        /// <summary>
        /// Returns true if both objects are equal.
        /// Note this is for backwards compatibility, this is a class and not immutable,
        /// so this method should return true if references are different. But that would break old code.
        /// </summary>
        /// <param name="o1"></param>
        /// <param name="o2"></param>
        /// <returns></returns>
        public static bool operator==(TXlsCellRange o1, TXlsCellRange o2)
        {
            if ((object)o1 == null) return (object)o2 == null;
            return o1.Equals(o2);
        }

        /// <summary>
        /// Returns true if both objects do not have the same value.
        /// Note this is for backwards compatibility, this is a class and not immutable,
        /// so this method should return true if references are different. But that would break old code.
        /// </summary>
        /// <param name="s1"></param>
        /// <param name="s2"></param>
        /// <returns></returns>
        public static bool operator!=(TXlsCellRange s1, TXlsCellRange s2)
        {
            if ((object)s1 == null) return (object)s2 != null;
            return !(s1.Equals(s2));
        }

        /// <summary>
        /// Returns the hashcode of the object.
        /// </summary>
        /// <returns></returns>
        public override int GetHashCode()
        {
            return HashCoder.GetHash(Left, Top, Right, Bottom);
        }
		
		#endregion

		#endregion

        internal TXlsCellRange OffsetForIns(int DestRow, int DestCol, TFlxInsertMode InsertMode)
        {
            int dr = DestRow;
            int dc = DestCol;
            if (InsertMode == TFlxInsertMode.ShiftColRight) dr = 0;
            if (InsertMode == TFlxInsertMode.ShiftRowDown) dc = 0;

            return Offset(dr, dc);
        }
    }

	/// <summary>
	/// A 3d Excel range.
	/// </summary>
	public class TXls3DRange : TXlsCellRange
	{
		#region Privates
		private int FSheet1;
		private int FSheet2;
		#endregion

		/// <summary>
		/// Creates a new TXls3DRange class.
		/// </summary>
		/// <param name="aSheet1">First sheet of the range.</param>
		/// <param name="aSheet2">Second sheet of the range.</param>
		/// <param name="aFirstRow">First row on range.</param>
		/// <param name="aFirstCol">First column on range.</param>
		/// <param name="aLastRow">Last row on range.</param>
		/// <param name="aLastCol">Last column on range.</param>
		public TXls3DRange(int aSheet1, int aSheet2, int aFirstRow, int aFirstCol, int aLastRow, int aLastCol):
			base(aFirstRow, aFirstCol, aLastRow, aLastCol)
		{
			Sheet1 = aSheet1;
			Sheet2 = aSheet2;
		}

		/// <summary>
		/// Creates an empty 3d range.
		/// </summary>
		public TXls3DRange():
			base()
		{
			Sheet1 = 0;
			Sheet2 = 0;
		}


		/// <summary>
		/// Returns true if the range has only one cell.
		/// </summary>
		public override bool IsOneCell
		{
			get
			{
				if (!base.IsOneCell) return false;
				return Sheet1 == Sheet2;
			}
		}
 

		/// <summary>
		/// First sheet of the range.
		/// </summary>
		public int Sheet1 {get {return FSheet1;} set{FSheet1 = value;}}

		/// <summary>
		/// Second sheet of the range.
		/// </summary>
		public int Sheet2 {get {return FSheet2;} set{FSheet2 = value;}}

        #region Compare
        /// <summary>
        /// Returns true if both objects are equal.
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public override bool Equals(object obj)
        {
            if (!base.Equals(obj)) return false;
            TXls3DRange o2 = obj as TXls3DRange;
            if (o2 == null) return false;
            return FSheet1 == o2.FSheet1 && FSheet2 == o2.FSheet2;
        }

        /// <summary>
        /// Returns true if both objects are equal.
        /// Note this is for backwards compatibility, this is a class and not immutable,
        /// so this method should return true if references are different. But that would break old code.
        /// </summary>
        /// <param name="o1"></param>
        /// <param name="o2"></param>
        /// <returns></returns>
        public static bool operator==(TXls3DRange o1, TXls3DRange o2)
        {
            if ((object)o1 == null) return (object)o2 == null;
            return o1.Equals(o2);
        }

        /// <summary>
        /// Returns true if both objects do not have the same value.
        /// Note this is for backwards compatibility, this is a class and not immutable,
        /// so this method should return true if references are different. But that would break old code.
        /// </summary>
        /// <param name="s1"></param>
        /// <param name="s2"></param>
        /// <returns></returns>
        public static bool operator!=(TXls3DRange s1, TXls3DRange s2)
        {
            if ((object)s1 == null) return (object)s2 != null;
            return !(s1.Equals(s2));
        }

        /// <summary>
        /// Returns the hashcode of the object.
        /// </summary>
        public override int GetHashCode()
        {
            return HashCoder.GetHash(base.GetHashCode(), FSheet1, FSheet2);
        }

        #endregion
	}

	/// <summary>
	/// An Excel named range.
	/// </summary>
	public class TXlsNamedRange: TXlsCellRange
	{
		#region Privates
		private string FName;
		private int FSheetIndex;
		private int FNameSheetIndex;
		private int FOptionFlags;
		private TParsedTokenList FFormulaData;
		private string FRangeFormula;
        private string FComment;

        static readonly char[] InvalidChars = NameInvalidChars();


		#endregion

        #region Constructors
        /// <summary>
        /// Creates a new Named range with the given values.
        /// </summary>
        /// <param name="aName">Name of the range.</param>
        /// <param name="aSheetIndex">Sheet index where the range is. (1-based)</param>
        /// <param name="aFirstRow">First row on range.</param>
        /// <param name="aFirstCol">First column on range.</param>
        /// <param name="aLastRow">Last row on range.</param>
        /// <param name="aLastCol">Last column on range.</param>
        /// <param name="aOptionFlags">Options of this Range.</param>
        /// <param name="aFormulaData">Formula data expressed as RPN array. Set it to null when creating a new range.</param>
        /// <param name="aRangeFormula">The formula for the range, expressed as text. Use it if the range is complex and cannot be expressed with aSheetIndex, aFirstRow... etc
        /// When you specify this parameter, all SheetIndex, aFirstRow, etc. lose their meaning.</param>
        internal TXlsNamedRange(string aName, int aSheetIndex, int aFirstRow, int aFirstCol, int aLastRow, int aLastCol, int aOptionFlags, TParsedTokenList aFormulaData, string aRangeFormula)
            : base(aFirstRow, aFirstCol, aLastRow, aLastCol)
        {
            Name = aName;
            SheetIndex = aSheetIndex;
            NameSheetIndex = 0;
            OptionFlags = aOptionFlags;
            if (aFormulaData == null)
                FFormulaData = null;
            else
            {
                FFormulaData = aFormulaData.Clone();
            }
            RangeFormula = aRangeFormula;
        }

        /// <summary>
        /// Creates a new Named range with the given values. Use this overload to create a simple range.
        /// </summary>
        /// <param name="aName">Name of the range.</param>
        /// <param name="aNameSheetIndex">Sheet index for the sheet that holds the range. 0 means a global range (default on Excel)</param>
        /// <param name="aSheetIndex">Sheet index where the range apply. This is where row and col properties apply, not where the range is stored. (1-based)</param>
        /// <param name="aFirstRow">First row on range.</param>
        /// <param name="aFirstCol">First column on range.</param>
        /// <param name="aLastRow">Last row on range.</param>
        /// <param name="aLastCol">Last column on range.</param>
        /// <param name="aOptionFlags">Options of this Range.</param>
        /// <param name="aRangeFormula">The formula for the range, expressed as text. Use it if the range is complex and cannot be expressed with aSheetIndex, aFirstRow... etc
        /// When you specify this parameter, all SheetIndex, aFirstRow, etc. lose their meaning.</param>
        public TXlsNamedRange(string aName, int aNameSheetIndex, int aSheetIndex, int aFirstRow, int aFirstCol, int aLastRow, int aLastCol, int aOptionFlags, string aRangeFormula)
            : this(aName, aSheetIndex, aFirstRow, aFirstCol, aLastRow, aLastCol, aOptionFlags, null, aRangeFormula)
        {
            NameSheetIndex = aNameSheetIndex;
        }

        /// <summary>
        /// Creates a new Named range with the given values. This is the most complete overload, you normally don't need to call it.
        /// </summary>
        /// <param name="aName">Name of the range.</param>
        /// <param name="aNameSheetIndex">Sheet index for the sheet that holds the range. 0 means a global range (default on Excel)</param>
        /// <param name="aSheetIndex">Sheet index where the range apply. This is where row and col properties apply, not where the range is stored. (1-based)</param>
        /// <param name="aFirstRow">First row on range.</param>
        /// <param name="aFirstCol">First column on range.</param>
        /// <param name="aLastRow">Last row on range.</param>
        /// <param name="aLastCol">Last column on range.</param>
        /// <param name="aOptionFlags">Options of this Range.</param>
        /// <param name="aFormulaData"></param>
        /// <param name="aRangeFormula">The formula for the range, expressed as text. Use it if the range is complex and cannot be expressed with aSheetIndex, aFirstRow... etc
        /// When you specify this parameter, all SheetIndex, aFirstRow, etc. lose their meaning.
        /// <br/>Note that with <b>relative</b> references, we always consider "A1" to be the cell where the name is. This means that the formula:
        /// "=$A$1 + A1" when evaluated in Cell B8, will read "=$A$1 + B8". To provide a negative offset, you need to wrap the formula.
        /// For example "=A1048575" will evaluate to B7 when evaluated in B8.
        /// </param>
        /// <param name="aComment">Comment for the named range. Note that this will only show in Excel 2007 or newer.</param>
        internal TXlsNamedRange(string aName, int aNameSheetIndex, int aSheetIndex, int aFirstRow, int aFirstCol, int aLastRow, int aLastCol, int aOptionFlags, TParsedTokenList aFormulaData, string aRangeFormula, string aComment)
            : this(aName, aSheetIndex, aFirstRow, aFirstCol, aLastRow, aLastCol, aOptionFlags, aFormulaData, aRangeFormula)
        {
            NameSheetIndex = aNameSheetIndex;
            Comment = aComment;
        }

        /// <summary>
        /// Creates a new Named range with the given values. Use this overload to create a simple range.
        /// </summary>
        /// <param name="aName">Name of the range.</param>
        /// <param name="aNameSheetIndex">Sheet index for the sheet that holds the range. 0 means a global range (default on Excel)</param>
        /// <param name="aSheetIndex">Sheet index where the range apply. This is where row and col properties apply, not where the range is stored. (1-based)</param>
        /// <param name="aFirstRow">First row on range.</param>
        /// <param name="aFirstCol">First column on range.</param>
        /// <param name="aLastRow">Last row on range.</param>
        /// <param name="aLastCol">Last column on range.</param>
        /// <param name="aOptionFlags">Options of this Range.</param>
        public TXlsNamedRange(string aName, int aNameSheetIndex, int aSheetIndex, int aFirstRow, int aFirstCol, int aLastRow, int aLastCol, int aOptionFlags)
            : this(aName, aSheetIndex, aFirstRow, aFirstCol, aLastRow, aLastCol, aOptionFlags, null, null)
        {
            NameSheetIndex = aNameSheetIndex;
        }


        /// <summary>
        /// Creates a complex Named range, with a formula.
        /// </summary>
        /// <param name="aName">Name of the range.</param>
        /// <param name="aNameSheetIndex">Sheet index for the sheet that holds the range. 0 means a global range (default on Excel)</param>
        /// <param name="aOptionFlags">Options of this Range.</param>
        /// <param name="aRangeFormula">The formula for the range, expressed as text. For example: "A1:B2,C3:C7"
        /// <br/>Note that with <b>relative</b> references, we always consider "A1" to be the cell where the name is. This means that the formula:
        /// "=$A$1 + A1" when evaluated in Cell B8, will read "=$A$1 + B8". To provide a negative offset, you need to wrap the formula.
        /// For example "=A1048575" will evaluate to B7 when evaluated in B8.
        /// </param>
        public TXlsNamedRange(string aName, int aNameSheetIndex, int aOptionFlags, string aRangeFormula)
            : this(aName, 1, 1, 1, 1, 1, aOptionFlags, null, aRangeFormula)
        {
            NameSheetIndex = aNameSheetIndex;
        }

        /// <summary>
        /// Creates an empty NamedRange.
        /// </summary>
        public TXlsNamedRange()
            : base()
        {
            Name = String.Empty;
            SheetIndex = 0;
            OptionFlags = 0;
        }
        #endregion

        #region Properties
        /// <summary>
		/// The name of the range.
		/// </summary>
        public string Name { get { return FName; } set { FName = value; } }
		/// <summary>
		/// The Sheet index where the row and col properties apply. 1-based. When RangeFormula is set, it is not used.
		/// </summary>
        public int SheetIndex { get { return FSheetIndex; } set { FSheetIndex = value; } }

		/// <summary>
		/// The sheet index for the name (1 based). A named range can have the same name than other
		/// as long as they are on different sheets. The default value(0) means a global named range, not tied to
		/// any specific sheet.
		/// </summary>
        public int NameSheetIndex { get { return FNameSheetIndex; } set { FNameSheetIndex = value; } }

		/// <summary>
		/// Options of the range. You can access the options by using the corresponding properties. (Hidden, BuiltIn, etc)
		/// </summary>
        public int OptionFlags { get { return FOptionFlags; } set { FOptionFlags = value; } }

		/// <summary>
		/// This is a formula defining the range. It can be used to define complex ranges.
		/// For example you can use "=Sheet1!$A:$B,Sheet1!$1:$2". 
		/// When this parameter is set, SheetIndex, Left, Top, Right and Bottom properties have no meaning.
        /// <br/>Note that with <b>relative</b> references, we always consider "A1" to be the cell where the name is. This means that the formula:
        /// "=$A$1 + A1" when evaluated in Cell B8, will read "=$A$1 + B8". To provide a negative offset, you need to wrap the formula.
        /// For example "=A1048575" will evaluate to B7 when evaluated in B8.
        /// </summary>
		/// <remarks>
		/// In latest FlexCel versions, you *can* use ranges like "A:B" instead (A1:B65536). Using "A:B" form is prefered, since it will work also in Excel 2007.
		/// </remarks>
        public string RangeFormula { get { return FRangeFormula; } set { FRangeFormula = value; } }

		/// <summary>
		/// Formula data as RPN array.
		/// </summary>
		internal TParsedTokenList FormulaData {get {return FFormulaData;} set {FFormulaData=value;} }

        /// <summary>
		/// True if the range is hidden.
		/// </summary>
		public bool Hidden {get {return (OptionFlags & 0x0001)!=0;} set {if (value) OptionFlags|= 0x0001; else OptionFlags &= ~(int)0x0001; }}
		
        /// <summary>
		/// True if the range is a function.
		/// </summary>
		public bool Function {get {return (OptionFlags & 0x0002)!=0;} set {if (value) OptionFlags|= 0x0002; else OptionFlags &= ~(int)0x0002; }}
		
        /// <summary>
		/// True if the range is a Visual Basic Procedure
		/// </summary>
		public bool VisualBasicProc {get {return (OptionFlags & 0x0004)!=0;} set {if (value) OptionFlags|= 0x0004; else OptionFlags &= ~(int)0x0004; }}
		
        /// <summary>
		/// True if the range is a function on a macro sheet.
		/// </summary>
		public bool Proc {get {return (OptionFlags & 0x0008)!=0;} set {if (value) OptionFlags|= 0x0008; else OptionFlags &= ~(int)0x0008; }}
		
        /// <summary>
		/// True if the range contains a complex function.
		/// </summary>
		public bool CalcExp {get {return (OptionFlags & 0x0010)!=0;} set {if (value) OptionFlags|= 0x0010; else OptionFlags &= ~(int)0x0010; }}

        /// <summary>
        /// True if the range is a built in name. Built in names are 1 char long.
        /// </summary>
        public bool BuiltIn { get { return (OptionFlags & 0x0020) != 0; } set { if (value) OptionFlags |= 0x0020; else OptionFlags &= ~(int)0x0020; } }

        /// <summary>
        /// Specifies the function group index if the defined name refers to a function. The function 
        /// group defines the general category for the function. This attribute is used when there is 
        /// an add-in or other code project associated with the file.
        /// </summary>
        public TFunctionGroup FunctionGroup
        {
            get { return (TFunctionGroup)((OptionFlags >> 6) & 0x003F); }
            set
            {
                OptionFlags &= ~0x0FC0; //clear existing
                OptionFlags |= ((int)value & 0x003F) << 6;
            }
        }

        /// <summary>
        /// Indicates whether the defined name is included in the 
        /// version of the workbook that is published to or rendered on a Web or application server. This is new to Excel 2007.
        /// </summary>
        public bool PublishToServer { get { return (OptionFlags & 0x2000) != 0; } set { if (value) OptionFlags |= 0x2000; else OptionFlags &= ~(int)0x2000; } }

        /// <summary>
        /// indicates that the name is used as a workbook parameter 
        /// on a version of the workbook that is published to or rendered on a Web or application server. This is new to Excel 2007.
        /// </summary>
        public bool WorkbookParameter { get { return (OptionFlags & 0x4000) != 0; } set { if (value) OptionFlags |= 0x4000; else OptionFlags &= ~(int)0x4000; } }

        /// <summary>
        /// Returns the comment associated with the name, if there is one. Comments are only available in Excel 2007, but they are saved too in xls file format.
        /// </summary>
        public string Comment { get { return FComment; } set { FComment = value; } }
        #endregion

        #region Compare
        /// <summary>
        /// Returns true if both objects are equal.
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public override bool Equals(object obj)
        {
            if (!base.Equals(obj)) return false;
            TXlsNamedRange o2 = obj as TXlsNamedRange;
            if (o2 == null) return false;
            return
                    FName == o2.FName &&
                    FSheetIndex == o2.FSheetIndex &&
                    FNameSheetIndex == o2.FNameSheetIndex &&
                    FOptionFlags == o2.FOptionFlags &&
                    //FFormulaData == o2.FFormulaData &&  We will not test for data.
                    FRangeFormula == o2.FRangeFormula &&
                    FComment == o2.FComment;
        }

        /// <summary>
        /// Returns true if both objects are equal.
        /// Note this is for backwards compatibility, this is a class and not immutable,
        /// so this method should return true if references are different. But that would break old code.
        /// </summary>
        /// <param name="o1"></param>
        /// <param name="o2"></param>
        /// <returns></returns>
        public static bool operator==(TXlsNamedRange o1, TXlsNamedRange o2)
        {
            if ((object)o1 == null) return (object)o2 == null;
            return o1.Equals(o2);
        }

        /// <summary>
        /// Returns true if both objects do not have the same value.
        /// Note this is for backwards compatibility, this is a class and not immutable,
        /// so this method should return true if references are different. But that would break old code.
        /// </summary>
        /// <param name="s1"></param>
        /// <param name="s2"></param>
        /// <returns></returns>
        public static bool operator!=(TXlsNamedRange s1, TXlsNamedRange s2)
        {
            if ((object)s1 == null) return (object)s2 != null;
            return !(s1.Equals(s2));
        }


        /// <summary>
        /// Returns the hashcode of the object.
        /// </summary>
        public override int GetHashCode()
        {

            return HashCoder.GetHash(
                base.GetHashCode(),
                FName != null ? FName.GetHashCode() : 0,
                FRangeFormula != null ? FRangeFormula.GetHashCode() : 0  //no need to check for everything, if name, formula and coords are the same there are good probabilities it is the same.
                );
        }

        #endregion

		/// <summary>
		/// Returns the string that corresponds to an internal name.
		/// </summary>
		/// <param name="name">Internal name we want to find.</param>
		/// <returns>The one-char string that represents that internal name.</returns>
		public static string GetInternalName(InternalNameRange name)
		{
			return ((char)name).ToString(CultureInfo.InvariantCulture);
		}


        internal static string GetInternal(string defname, out bool IsInternal)
        {
            IsInternal = true;
            switch (defname)
            {
                //Some of this names are not mentioned, but if you open a file with them, Excel understands them.
                case "_xlnm.Consolidate_Area": return TXlsNamedRange.GetInternalName(InternalNameRange.Consolidate_Area);
                case "_xlnm.Auto_Open": return TXlsNamedRange.GetInternalName(InternalNameRange.Auto_Open); //Not mentioned in docs, but exists
                case "_xlnm.Auto_Close": return TXlsNamedRange.GetInternalName(InternalNameRange.Auto_Close);
                case "_xlnm.Extract": return TXlsNamedRange.GetInternalName(InternalNameRange.Extract);
                case "_xlnm.Database": return TXlsNamedRange.GetInternalName(InternalNameRange.Database);
                case "_xlnm.Criteria": return TXlsNamedRange.GetInternalName(InternalNameRange.Criteria);
                case "_xlnm.Print_Area": return TXlsNamedRange.GetInternalName(InternalNameRange.Print_Area);
                case "_xlnm.Print_Titles": return TXlsNamedRange.GetInternalName(InternalNameRange.Print_Titles);
                case "_xlnm.Recorder": return TXlsNamedRange.GetInternalName(InternalNameRange.Recorder);
                case "_xlnm.Data_Form": return TXlsNamedRange.GetInternalName(InternalNameRange.Data_Form);
                case "_xlnm.Auto_Activate": return TXlsNamedRange.GetInternalName(InternalNameRange.Auto_Activate);
                case "_xlnm.Auto_Deactivate": return TXlsNamedRange.GetInternalName(InternalNameRange.Auto_Deactivate);
                case "_xlnm.Sheet_Title": return TXlsNamedRange.GetInternalName(InternalNameRange.Sheet_Title);
                case "_xlnm._FilterDatabase": return TXlsNamedRange.GetInternalName(InternalNameRange.Filter_DataBase);
            }

            IsInternal = false;
            return defname;
        }

        internal static string GetXlsxInternal(string biff8Name)
        {
            if (biff8Name == null || biff8Name.Length == 0) return biff8Name;

            if (biff8Name.Length == 1 && (int)biff8Name[0] <= 0x0D) //Internal name.
            {
                switch ((InternalNameRange)biff8Name[0])
                {
                    case InternalNameRange.Consolidate_Area: return "_xlnm.Consolidate_Area";
                    case InternalNameRange.Auto_Open: return "_xlnm.Auto_Open";
                    case InternalNameRange.Auto_Close: return "_xlnm.Auto_Close";
                    case InternalNameRange.Extract: return "_xlnm.Extract";
                    case InternalNameRange.Database: return "_xlnm.Database";
                    case InternalNameRange.Criteria: return "_xlnm.Criteria";
                    case InternalNameRange.Print_Area: return "_xlnm.Print_Area";
                    case InternalNameRange.Print_Titles: return "_xlnm.Print_Titles";
                    case InternalNameRange.Recorder: return "_xlnm.Recorder";
                    case InternalNameRange.Data_Form: return "_xlnm.Data_Form";
                    case InternalNameRange.Auto_Activate: return "_xlnm.Auto_Activate";
                    case InternalNameRange.Auto_Deactivate: return "_xlnm.Auto_Deactivate";
                    case InternalNameRange.Sheet_Title: return "_xlnm.Sheet_Title";
                    case InternalNameRange.Filter_DataBase: return "_xlnm._FilterDatabase";
                }
            }
            return biff8Name;
        }


        /// <summary>
        /// Returns true if the string is a valid name for a named range. Valid names must start with 
        /// a letter or an underscore
        /// </summary>
        /// <param name="Name">String we want to check.</param>
        /// <param name="IsInternal">Returns true if this is an internal name, like Print_Range. Internal names have only one character.</param>
        /// <returns></returns>
        public static bool IsValidRangeName(string Name, out bool IsInternal)
        {
            IsInternal = false;
            if (Name == null || Name.Length < 1 || Name.Length > 254) return false;
            if (Name == "R" || Name == "r") return false;
            if (Name == "C" || Name == "c") return false;
            if (string.Equals(Name, "true", StringComparison.InvariantCultureIgnoreCase)) return false;
            if (string.Equals(Name, "false", StringComparison.InvariantCultureIgnoreCase)) return false;

            if (Name.Length == 1 && (int)Name[0] <= 0x0D) //Internal name.
            {
                IsInternal = true;
                return true;
            }

            if (Name.IndexOfAny(InvalidChars) >= 0) return false;
            if (Name[0] < 'A') return false;

            //Check it is not a valid cell reference.
            TCellAddress a = new TCellAddress();
            if (a.TrySetCellRef(Name)) return false;

            return true;

        }

        private static char[] NameInvalidChars()
        {
            char[] InvalidChars = new char[0x41 + 0xC0 - 0x7F];
            for (int i = 0; i < 0x41; i++) InvalidChars[i] = (char)i;
            InvalidChars[0x30] = '{'; //replace '0' with another invalid char.
            InvalidChars[0x31] = '/'; //replace '1' with another invalid char.
            InvalidChars[0x32] = '}'; //replace '2' with another invalid char.
            InvalidChars[0x33] = '['; //replace '3' with another invalid char.
            InvalidChars[0x34] = ']'; //replace '4' with another invalid char.
            InvalidChars[0x35] = '~'; //replace '5' with another invalid char.
            InvalidChars[0x36] = (char)0xA0; //replace '6' with another invalid char.
            InvalidChars[0x37] = '{'; //replace '7' with another invalid char.
            InvalidChars[0x38] = '{'; //replace '8' with another invalid char.
            InvalidChars[0x39] = '{'; //replace '9' with another invalid char.

            InvalidChars[0x3F] = '{'; //replace '?' with another invalid char.

            InvalidChars[(int)'.'] = '{'; //replace '.' with another invalid char.

            for (int i = 0x7F; i < 0xC0; i++)
                InvalidChars[0x41 + i - 0x7F] = (char)i;

            InvalidChars[0x41 + 0xB5 - 0x7F] = '{'; //replace 'u'(micro) with another invalid char.
            return InvalidChars;
        }

        internal TXlsCellRange[] GetRanges()
        {
            if (FormulaData == null || FormulaData.Count < 0) return null;
            FormulaData.ResetPositionToLast();
            List<TXlsCellRange> Ranges = new List<TXlsCellRange>();

            while (!FormulaData.Bof())
            {
                TBaseParsedToken basetoken = FormulaData.LightPop();
                TArea3dToken at = basetoken as TArea3dToken;
                if (at != null && !at.IsErr())
                {
                    Ranges.Add(new TXlsCellRange(at.Row1 + 1, at.Col1 + 1, at.Row2 + 1, at.Col2 + 1));
                }

                TRef3dToken rf = basetoken as TRef3dToken;
                if (rf != null && !rf.IsErr())
                {
                    Ranges.Add(new TXlsCellRange(rf.Row + 1, rf.Col + 1, rf.Row + 1, rf.Col + 1));
                }
            }

            if (Ranges.Count == 0) return null;
            return Ranges.ToArray();
        }
    }

    /// <summary>
    /// Small class that can convert between a string reference ("A1") into row and col integers (1,1).
    /// </summary>
    public class TCellAddress
    {
        private int FRow;
        private int FCol;
        private bool FRowAbsolute;
        private bool FColAbsolute;
        private string FSheet;

        /// <summary>
        /// Creates Cell Address pointing to A1.
        /// </summary>
        public TCellAddress()
        {
            FCol = 1;
            FRow = 1;
            FSheet = String.Empty;
        }

        /// <summary>
        /// Creates Cell Address pointing to (aRow, aCol).
        /// </summary>
        /// <param name="aRow">Row index of the reference (1-based).</param>
        /// <param name="aCol">Column index of the reference (1-based).</param>
        public TCellAddress(int aRow, int aCol)
        {
            FRow = aRow;
            FCol = aCol;
            FSheet = String.Empty;
        }

        /// <summary>
        /// Creates Cell Address pointing to (aRow, aCol) with the corresponding absolute values.
        /// </summary>
        /// <param name="aRow">Row index of the reference (1-based).</param>
        /// <param name="aCol">Column index of the reference (1-based).</param>
        /// <param name="aRowAbsolute">If true row will be an absolute reference. (As in A$5).</param>
        /// <param name="aColAbsolute">If true col will be an absolute reference. (As in $A5).</param>
        public TCellAddress(int aRow, int aCol, bool aRowAbsolute, bool aColAbsolute)
        {
            FRow = aRow;
            FCol = aCol;
            FRowAbsolute = aRowAbsolute;
            FColAbsolute = aColAbsolute;
            FSheet = String.Empty;
        }

        /// <summary>
        /// Creates Cell Address pointing to (aRow, aCol) with the corresponding absolute values.
        /// </summary>
        /// <param name="aSheet">Sheet name of the reference.</param>
        /// <param name="aRow">Row index of the reference (1-based).</param>
        /// <param name="aCol">Column index of the reference (1-based).</param>
        /// <param name="aRowAbsolute">If true row will be an absolute reference. (As in A$5).</param>
        /// <param name="aColAbsolute">If true col will be an absolute reference. (As in $A5).</param>
        public TCellAddress(string aSheet, int aRow, int aCol, bool aRowAbsolute, bool aColAbsolute)
        {
            FRow = aRow;
            FCol = aCol;
            FRowAbsolute = aRowAbsolute;
            FColAbsolute = aColAbsolute;
            if (aSheet == null) FSheet = String.Empty;
            else
                FSheet = aSheet;
        }


        /// <summary>
        /// Creates a Cell Address pointing to (aCellRef).
        /// </summary>
        /// <param name="aCellRef">
        /// String containing the cell address in Excel notation (for example "A5").
        /// Absolute references ($A$5) work too.
        ///</param>
        public TCellAddress(string aCellRef)
        {
            CellRef = aCellRef;
        }


        /// <summary>
        /// Returns a deep copy of the cell address.
        /// </summary>
        /// <returns></returns>
        public TCellAddress Clone()
        {
            return (TCellAddress)MemberwiseClone();
        }

        /// <summary>
        /// Sheet name of the reference.
        /// </summary>
        public string Sheet { get { return FSheet; } set { FSheet = value; } }

        /// <summary>
        /// Row index for this reference (1-based).
        /// </summary>
        public int Row { get { return FRow; } set { if (value < 1 || value > FlxConsts.Max_Rows + 1) FlxMessages.ThrowException(FlxErr.ErrInvalidRow, value); FRow = value; } }
        /// <summary>
        /// Column index for this reference (1-based).
        /// </summary>
        public int Col { get { return FCol; } set { if (value < 1 || value > FlxConsts.Max_Columns + 1) FlxMessages.ThrowException(FlxErr.ErrInvalidColumn, value); FCol = value; } }

        /// <summary>
        /// True if the row is an absolute reference (as in A$5)
        /// </summary>
        public bool RowAbsolute { get { return FRowAbsolute; } set { FRowAbsolute = value; } }

        /// <summary>
        /// True if the column is an absolute reference (as in $A5)
        /// </summary>
        public bool ColAbsolute { get { return FColAbsolute; } set { FColAbsolute = value; } }

        /// <summary>
        /// Quotes a sheet name if it is needed. For example, Sheet 1 should be quoted as 'Sheet 1'
        /// </summary>
        /// <param name="SheetName">Sheet to quote</param>
        /// <returns>Quoted sheet.</returns>
        public static string QuoteSheet(string SheetName)
        {
            bool NeedsQuote = false;
            bool StartsWithBraket = false;
            bool BraketEnded = false;

            for (int i = 0; i < SheetName.Length; i++)
            {
                char c = SheetName[i];
                bool IsAlpha = (c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z') || (c == '_') || (c == '\\');
                if (IsAlpha) continue;

                if (i == 0)
                {
                    if (c == '[')
                    {
                        StartsWithBraket = true;
                        continue;
                    }
                }
                else
                {
                    bool IsNumber = (c >= '0' && c <= '9');
                    if (IsNumber) continue;
                    bool IsExtra = (c == '?' || c == '.');
                    if (IsExtra) continue;

                    if (StartsWithBraket && !BraketEnded && c == ']')
                    {
                        BraketEnded = true;
                        continue;
                    }
                }

                NeedsQuote = true;
                break;
            }

            if (StartsWithBraket && !BraketEnded) NeedsQuote = true;

            if (!NeedsQuote) return SheetName;
            char quote = TFormulaMessages.TokenChar(TFormulaToken.fmSingleQuote);
            string squote = quote.ToString();
            return quote + SheetName.Replace(squote, squote + squote) + quote;
        }

        /// <summary>
        /// Returns "A" for column 1, "B"  for 2 and so on.
        /// </summary>
        /// <param name="C">Index of the column, 1 based.</param>
        /// <returns></returns>
        public static string EncodeColumn(int C)
        {
            const int Delta = 'Z' - 'A' + 1;
            if (C <= Delta) return ((char)('A' + C - 1)).ToString();
            else
                return EncodeColumn(((C - 1) / Delta)) + ((char)('A' + ((C - 1) % Delta))).ToString();
        }

        /// <summary>
        /// Cell address in Excel A1 notation. (For example "A5").
        /// Absolute references ($A$5) will work too.
        /// </summary>
        public string CellRef
        {
            get
            {
                StringBuilder sb = GetSheetRef();
                if (FColAbsolute) sb.Append(TFormulaMessages.TokenChar(TFormulaToken.fmAbsoluteRef));
                sb.Append(EncodeColumn(Col));
                if (FRowAbsolute) sb.Append(TFormulaMessages.TokenChar(TFormulaToken.fmAbsoluteRef));
                sb.Append(Row.ToString(CultureInfo.InvariantCulture));
                return sb.ToString();
            }
            set
            {
                if (!TrySetCellRef(value)) FlxMessages.ThrowException(FlxErr.ErrInvalidRef, value);
            }
        }

        private StringBuilder GetSheetRef()
        {
            StringBuilder sb = new StringBuilder();
            if (FSheet != null && FSheet.Length > 0)
            {
                sb.Append(QuoteSheet(FSheet));
                sb.Append(TFormulaMessages.TokenString(TFormulaToken.fmExternalRef));
            }
            return sb;
        }

        /// <summary>
        /// Returns the cell reference in the objects in R1C1 notation.
        /// </summary>
        /// <param name="cellRow">Row where the cell that has the formula is. This is used in relative R1C1 references,
        /// so for example, if the cell with the formula is B7, and this object holds a relative C9, the result would br R[2]C[1]</param>
        /// <param name="cellCol">Column where the cell that has the formula is. This is used in relative R1C1 references,
        /// so for example, if the cell with the formula is B7, and this object holds a relative C9, the result would br R[2]C[1]</param>
        /// <returns>The cell reference in R1C1 notation.</returns>
        internal string CellRefR1C1(int cellRow, int cellCol)
        {
            StringBuilder sb = GetSheetRef();
            return sb.ToString() + GetR1C1Ref(Row, Col, cellRow, cellCol, RowAbsolute, ColAbsolute);
        }

        /// <summary>
        /// Tries to set a cell reference, and returns true if the cell is a correct A1 reference. This is similar to setting <see cref="CellRef"/> to a 
        /// string, but it will not raise an exception.
        /// </summary>
        /// <param name="value">String with the cell reference, in A1 format. To use R1C1 notation, see <see cref="TrySetCellRef(string, TReferenceStyle, int, int)"/></param>
        /// <returns>True if value was a correct cell reference, false otherwise.</returns>
        public bool TrySetCellRef(string value)
        {
            return TrySetCellRef(value, TReferenceStyle.A1, 0, 0);
        }

        private bool ParseSheet(string value, out string v)
        {
            FSheet = String.Empty;
            v = value.ToUpper(CultureInfo.InvariantCulture);

            string SheetSep = TFormulaMessages.TokenString(TFormulaToken.fmExternalRef);
            int SheetIndex = v.IndexOf(SheetSep, StringComparison.Ordinal);
            if (SheetIndex >= 0)
            {
                FSheet = v.Substring(0, SheetIndex);
                int dummy = 0;
                if (FSheet.Length > 2)
                {
                    if (!TryUnquote(FSheet, out FSheet, out dummy)) return false;
                }
                v = v.Substring(SheetIndex + 1);
            }
            return true;
        }

        private bool ParseA1Address(string v)
        {
            char AbsRef = TFormulaMessages.TokenChar(TFormulaToken.fmAbsoluteRef);
            FColAbsolute = v.Length > 0 && v[0] == AbsRef;
            int k = 0; if (FColAbsolute) k++;

            int aCol = 0;
            k = ReadSimpleCol(v, k, out aCol);
            if (k < 0) return false;
            Col = aCol;

            FRowAbsolute = v.Length > k && v[k] == AbsRef;
            if (FRowAbsolute) k++;

            int aRow = 0;
            k = ReadSimpleRow(v, k, out aRow);
            if (k < 0) return false;
            Row = aRow;

            return true;
        }

        private bool ParseR1C1Address(string v, int cellRow, int cellCol, out bool IsFullRowRange, out bool IsFullColRange)
        {
            TFormulaConvertTextToInternal ps = new TFormulaConvertTextToInternal(null, -1, false, v, false);
            ps.SetStartForRelativeRefs(cellRow - 1, cellCol - 1);
            IsFullRowRange = false; IsFullColRange = false;
            bool ok = ps.ReadR1C1Ref(false, ref FRowAbsolute, ref FColAbsolute, out FRow, out FCol, ref IsFullRowRange, ref IsFullColRange);
            if (!ok || ps.RemainingFormula.Trim().Length != 0) return false;

            return true;
        }


        /// <summary>
        /// Tries to set a cell reference, and returns true if the cell is a correct A1 or R1C1 reference. 
        /// </summary>
        /// <param name="value">String with the cell reference, in A1 or R1C1 format.</param>
        /// <param name="referenceStyle">Style the reference is in (A1 or R1C1).</param>
        /// <param name="cellRow">Row where the cell is. This is only used for R1C1 relative references, to know where the
        /// cell is. For example, the reference R[1]C[1] when cellRow is 5 will reference row 6. A1 references or absolute R1C1 references ignore this parameter.</param>
        /// <param name="cellCol">Column where the cell is. This is only used for R1C1 relative references, to know where the
        /// cell is. For example, the reference R[1]C[1] when cellCol is 5 will reference column 6. A1 references or absolute R1C1 references ignore this parameter.</param>
        /// <returns>True if value was a correct cell reference, false otherwise.</returns>
        public bool TrySetCellRef(string value, TReferenceStyle referenceStyle, int cellRow, int cellCol)
        {
            bool frr, fcr;
            return TrySetCellRef(value, referenceStyle, cellRow, cellCol, out frr, out fcr);
        }

        internal bool TrySetCellRef(string value, TReferenceStyle referenceStyle, int cellRow, int cellCol, out bool IsFullRowRange, out bool IsFullColRange)
        {
            IsFullRowRange = false;
            IsFullColRange = false;
            string v;
            if (!ParseSheet(value, out v)) return false;
            if (referenceStyle == TReferenceStyle.R1C1)
            {
                return ParseR1C1Address(v, cellRow, cellCol, out IsFullRowRange, out IsFullColRange);
            }
            return ParseA1Address(v);
        }

        internal static bool TryUnquote(string sheetName, out string Sheets, out int i)
        {
            char quote = TFormulaMessages.TokenChar(TFormulaToken.fmSingleQuote);
            i = 0;
            Sheets = null;

            if (sheetName.Length < 1) return false;


            if (sheetName[0] == quote)
            {
                if (sheetName.Length < 2) return false;
                StringBuilder SheetBuilder = new StringBuilder();
                i = 1;
                while (i < sheetName.Length - 1)
                {
                    if (sheetName[i] == quote) //Get rid of double quotes.
                        if (sheetName[i + 1] == quote)
                        {
                            i++;
                        }
                        else break;

                    SheetBuilder.Append(sheetName[i]);
                    i++;
                }
                i++;
                Sheets = SheetBuilder.ToString();
            }
            else
            {
                i = sheetName.LastIndexOf(TFormulaMessages.TokenChar(TFormulaToken.fmExternalRef));
                if (i < 1)
                    i = sheetName.Length;
                Sheets = sheetName.Substring(0, i);
            }

            return true;
        }

        internal static string UnQuote(string QuotedString)
        {
            char quote = TFormulaMessages.TokenChar(TFormulaToken.fmSingleQuote);
            if (QuotedString.Length < 1) return QuotedString;


            if (QuotedString[0] == quote)
            {
                if (QuotedString.Length < 2) return QuotedString;
                StringBuilder SheetBuilder = new StringBuilder();
                int i = 1;
                while (i < QuotedString.Length - 1)
                {
                    if (QuotedString[i] == quote) //Get rid of double quotes.
                        if (QuotedString[i + 1] == quote)
                        {
                            i++;
                        }
                        else break;

                    SheetBuilder.Append(QuotedString[i]);
                    i++;
                }
                if (i + 1 < QuotedString.Length) SheetBuilder.Append(QuotedString.Substring(i + 1));
                return SheetBuilder.ToString();
            }
            else
            {
                return QuotedString;
            }
        }

        internal static void SplitFileName(string SheetPlusFilename, out string FileName, out string Sheets)
        {
            string UnquotedSheetPlusFilename = UnQuote(SheetPlusFilename);

            FileName = null;
            Sheets = UnquotedSheetPlusFilename;
            int w = UnquotedSheetPlusFilename.LastIndexOf(TBaseFormulaParser.fts(TFormulaToken.fmWorkbookClose));
            if (w > 0)
            {
                FileName = UnquotedSheetPlusFilename.Substring(0, w);
                int ws = FileName.IndexOf(TBaseFormulaParser.fts(TFormulaToken.fmWorkbookOpen));
                if (ws >= 0) FileName = FileName.Substring(ws + 1);
                Sheets = UnquotedSheetPlusFilename.Substring(w + 1);
            }
        }

        /// <summary>
        /// Returns both sheets from a string. If the string has only one sheet, sheet1==sheet2.
        /// Note that this routine uses real names and not quoted ones. for example, the string "sheet'1" but not "'sheet''1'"
        /// </summary>
        /// <param name="sheetAndFileName">Expression with a sheet name or a sheet range. Might be for example "Sheet1", "Sheet 1" or "Sheet1:Sheet2"</param>
        /// <param name="LocalXls">XlsFile where the sheets are located.</param>
        /// <param name="sheet1">First sheet of the range.</param>
        /// <param name="sheet2">Second sheet of the range. If there is no second sheet, this value is equal to sheet1.</param>
        /// <param name="Xls">Excel file.</param>
        private static bool TryParseUnquotedSheet(ExcelFile Xls, string sheetAndFileName, out ExcelFile LocalXls, out int sheet1, out int sheet2)
        {
            LocalXls = Xls; sheet1 = 0; sheet2 = 0;

            string sheetName; string fileName;
            SplitFileName(sheetAndFileName, out fileName, out sheetName);

            string[] Sheets = sheetName.Split(TFormulaMessages.TokenChar(TFormulaToken.fmRangeSep));
            if (Sheets.Length > 2 || Sheets.Length < 1)
                return false;

            if (fileName != null)
            {
                LocalXls = Xls.GetSupportingFile(fileName);
                if (LocalXls == null) return false;
            }

            sheet1 = LocalXls.GetSheetIndex(Sheets[0], false);
            if (sheet1 < 1) return false;

            sheet2 = sheet1;
            if (Sheets.Length > 1)
            {
                sheet2 = LocalXls.GetSheetIndex(Sheets[1], false);
                if (sheet2 < 1) return false;
            }

            return true;
        }

        /// <summary>
        /// Parses an string with an address. Quotes are allowed here.
        /// </summary>
        internal static void ParseAddress(ExcelFile Xls, string s, int activeSheet, out ExcelFile LocalXls, out int sheet1, out int sheet2, out int row1, out int col1, out int row2, out int col2)
        {
            if (!TryParseAddress(Xls, s, activeSheet, out LocalXls, out sheet1, out sheet2, out row1, out col1, out row2, out col2))
                FlxMessages.ThrowException(FlxErr.ErrInvalidRef, s);
        }

        /// <summary>
        /// Parses an string with an address. Quotes are allowed here. Xls might be null, in which case external refs are not allowed.
        /// </summary>
        internal static bool TryParseAddress(ExcelFile Xls, string s, int activeSheet, out ExcelFile LocalXls, out int sheet1, out int sheet2, out int row1, out int col1, out int row2, out int col2)
        {
            return TryParseAddress(Xls, s, activeSheet, out LocalXls, out sheet1, out sheet2, out row1, out col1, out row2, out col2, TReferenceStyle.A1, 0, 0); 
        }

        /// <summary>
        /// Parses an string with an address. Quotes are allowed here. Xls might be null, in which case external refs are not allowed.
        /// </summary>
        internal static bool TryParseAddress(ExcelFile Xls, string s, int activeSheet, out ExcelFile LocalXls, 
            out int sheet1, out int sheet2, out int row1, out int col1, out int row2, out int col2,
            TReferenceStyle RefStyle, int CellRow, int CellCol)
        {
            int i = 0; row1 = 0; col1 = 0; row2 = 0; col2 = 0; LocalXls = Xls; sheet1 = 0; sheet2 = 0;
            if (s.IndexOf(TFormulaMessages.TokenChar(TFormulaToken.fmExternalRef)) >= 0)
            {
                if (Xls == null) return false;
                string Sheets;
                if (!TryUnquote(s, out Sheets, out i)) return false;

                if (!TryParseUnquotedSheet(Xls, Sheets, out LocalXls, out sheet1, out sheet2)) return false;

                if (i < 1 || i + 1 >= s.Length || s[i] != TFormulaMessages.TokenChar(TFormulaToken.fmExternalRef))
                    return false;
                i++; //skip the "!"
            }
            else
            {
                LocalXls = Xls;
                sheet1 = activeSheet;
                sheet2 = activeSheet;
            }

            string[] Cells = s.Substring(i).Split(TFormulaMessages.TokenChar(TFormulaToken.fmRangeSep));
            if (Cells.Length > 2 || Cells.Length < 1)
                return false;

            //Support for ranges like 1:3 or a:b
            if (RefStyle == TReferenceStyle.A1 && Cells.Length == 2)
            {
                if (TryA1FullReferences(Cells, out row1, out row2, out col1, out col2)) return true;
            }


            TCellAddress Addr = new TCellAddress();
            bool IsFullRowRange1, IsFullColRange1;
            if (!Addr.TrySetCellRef(Cells[0], RefStyle, CellRow, CellCol, out IsFullRowRange1, out IsFullColRange1)) return false;

            row1 = Addr.Row;
            col1 = Addr.Col;

            bool IsFullRowRange2 = true; bool IsFullColRange2 = true;
            if (Cells.Length == 2)
            {
                if (!Addr.TrySetCellRef(Cells[1], RefStyle, CellRow, CellCol, out IsFullRowRange2, out IsFullColRange2)) return false;
            }
            row2 = Addr.Row;
            col2 = Addr.Col;

            if (IsFullRowRange1 && IsFullRowRange2)
            {
                col1 = 1;
                col2 = FlxConsts.Max_Columns + 1;
            }

            if (IsFullColRange1 && IsFullColRange2)
            {
                row1 = 1;
                row2 = FlxConsts.Max_Rows + 1;
            }

            return true;
        }

        private static bool TryA1FullReferences(string[] Cells, out int row1, out int row2, out int col1, out int col2)
        {
            col2 = 0;
            row2 = 0;

            string v1 = Cells[0].ToUpper(CultureInfo.InvariantCulture);
            string v2 = Cells[1].ToUpper(CultureInfo.InvariantCulture);

            if (GetFullCol(v1, out col1) && GetFullCol(v2, out col2))
            {
                if (col1 > col2)
                {
                    int tmp = col1;
                    col1 = col2;
                    col2 = tmp;
                }

                row1 = 1;
                row2 = FlxConsts.Max_Rows + 1;
                return true;
            }

            if (GetFullRow(v1, out row1) && GetFullRow(v2, out row2))
            {
                if (row1 > row2)
                {
                    int tmp = row1;
                    row1 = row2;
                    row2 = tmp;
                }

                col1 = 1;
                col2 = FlxConsts.Max_Columns + 1;
                return true;
            }

            return false;
        }

        private static bool GetFullCol(string v, out int Col)
        {
            bool ColAbsolute = v.Length > 0 && v[0] == TFormulaMessages.TokenChar(TFormulaToken.fmAbsoluteRef);
            int k = 0; if (ColAbsolute) k++;

            k = ReadSimpleCol(v, k, out Col);
            return k == v.Length;
        }

        private static bool GetFullRow(string v, out int Row)
        {
            bool RowAbsolute = v.Length > 0 && v[0] == TFormulaMessages.TokenChar(TFormulaToken.fmAbsoluteRef);
            int k = 0; if (RowAbsolute) k++;

            k = ReadSimpleRow(v, k, out Row);
            return k == v.Length;
        }

        /// <summary>
        /// An optimized method to read cell references in xlsx files. It won't allow absolute or sheets.
        /// </summary>
        internal static int ReadSimpleCol(string CellRef, int StrStart, out int aCol)
        {
            aCol = 0;
            const int Delta = 'Z' - 'A' + 1;

            int i = StrStart;
            while (i < CellRef.Length)
            {
                char c = CellRef[i];
                if (c >= 'a' && c <= 'z') c -= (char)('a' - 'A');
                if (c < 'A' || c > 'Z') break;
                if (i - StrStart > FlxConsts.Max_LettersInColumnName) return -1;

                aCol *= Delta;
                aCol += c - 'A' + 1;
                i++;
            }

            if (aCol < 1 || aCol > FlxConsts.Max_Columns + 1) return -1;
            return i;
        }

        internal static int ReadSimpleRow(string CellRef, int StrStart, out int aRow)
        {
            aRow = 0;

            int i = StrStart;
            while (i < CellRef.Length)
            {
                char c = CellRef[i];
                if (c < '0' || c > '9') return -1; //not allowed anything after the row.

                aRow *= 10;
                aRow += c - '0';
                if (aRow > FlxConsts.Max_Rows + 1) return -1;
                i++;
            }

            if (aRow < 1 || aRow > FlxConsts.Max_Rows + 1) return -1;
            return i;
        }


        internal static string GetR1C1Ref(int Row, int Col, int CellRow, int CellCol, bool RowAbs, bool ColAbs)
        {
            string Result = TFormulaConvertTextToInternal.fts(TFormulaToken.fmR1C1_R);
            Result += GetR1SimpleRef(Row, CellRow, RowAbs);
            Result += TFormulaConvertTextToInternal.fts(TFormulaToken.fmR1C1_C);
            Result += GetR1SimpleRef(Col, CellCol, ColAbs);
            return Result;
        }

        internal static string GetR1SimpleRef(int RowCol, int CellRowCol, bool Abs)
        {
            if (Abs)
            {
                return RowCol.ToString();
            }
            else
            {
                int Delta = RowCol - 1 - CellRowCol;
                if (Delta == 0) return String.Empty;
                return TFormulaConvertTextToInternal.fts(TFormulaToken.fmR1C1RelativeRefStart) + (Delta).ToString() + TFormulaConvertTextToInternal.fts(TFormulaToken.fmR1C1RelativeRefEnd);
            }
        }
        
        /// <summary>
        /// Returns true if both objects have the same data.
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public override bool Equals(object obj)
        {
            TCellAddress o2 = obj as TCellAddress;
            if (o2 == null) return false;
            return o2.FRow == FRow && o2.FCol == FCol && o2.FRowAbsolute == FRowAbsolute && o2.FColAbsolute == FColAbsolute
                && o2.FSheet == FSheet;
        }

        /// <summary>
        /// Returns the hashcode of this object.
        /// </summary>
        /// <returns></returns>
        public override int GetHashCode()
        {
            return HashCoder.GetHashObj(FRow, FCol, FRowAbsolute, FColAbsolute, FSheet);
        }
    }
    
    /// <summary>
    /// A class with 2 TCellAddress objects marking the start and end of a cell range.
    /// </summary>
    public class TCellAddressRange
    {
        TCellAddress FTopLeft;
        TCellAddress FBottomRight;

        /// <summary>
        /// Creates a new instance. Addresses can't be null.
        /// </summary>
        /// <param name="aTopLeft"></param>
        /// <param name="aBottomRight"></param>
        public TCellAddressRange(TCellAddress aTopLeft, TCellAddress aBottomRight)
        {
            TopLeft = aTopLeft;
            BottomRight = aBottomRight;
        }

        /// <summary>
        /// The cell at the top left position in the range. It can't be null.
        /// </summary>
        public TCellAddress TopLeft { get { return FTopLeft; }
            set
            {
                if (value == null) FlxMessages.ThrowException(FlxErr.ErrNullParameter, "TopLeft");
                FTopLeft = value;
            }
        }


        /// <summary>
        /// The cell at the bottom right position in the range. It can't be null.
        /// </summary>
        public TCellAddress BottomRight { get { return FBottomRight; }
            set 
            { 
                if (value == null) FlxMessages.ThrowException(FlxErr.ErrNullParameter, "BottomRight"); 
                FBottomRight = value; 
            }
        }

        /// <summary>
        /// Returns true if both objects have the same data.
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public override bool Equals(object obj)
        {
            TCellAddressRange o2 = obj as TCellAddressRange;
            if (o2 == null) return false;

            return Object.Equals(TopLeft, o2.TopLeft) && Object.Equals(BottomRight, o2.BottomRight);
        }

        /// <summary>
        /// HashCode for the object
        /// </summary>
        /// <returns></returns>
        public override int GetHashCode()
        {
            return HashCoder.GetHashObj(TopLeft, BottomRight);
        }

    }
    #endregion

    #region Image Properties
    /// <summary>
	/// Defines a cropping area for an image. If the values are not zero, only a part of the image will display on Excel.
	/// </summary>
	public class TCropArea: ICloneable
	{
		private int FCropFromTop;
		private int FCropFromBottom;
		private int FCropFromLeft;
		private int FCropFromRight;

		/// <summary>
		/// Creates a new crop area with no crop.
		/// </summary>
		public TCropArea()
		{
		}

		/// <summary>
		/// Creates a new crop area with the indicated coordinates.
		/// </summary>
		/// <param name="aCropFromTop"></param>
		/// <param name="aCropFromBottom"></param>
		/// <param name="aCropFromLeft"></param>
		/// <param name="aCropFromRight"></param>
		public TCropArea(int aCropFromTop, int aCropFromBottom, int aCropFromLeft, int aCropFromRight)
		{
			FCropFromTop = aCropFromTop;
			FCropFromBottom = aCropFromBottom;
			FCropFromLeft = aCropFromLeft;
			FCropFromRight = aCropFromRight;
		}

		/// <summary>
		/// How much to crop the image, in fractions of 65536 of the total image height.
		/// </summary>
		public int CropFromTop {get {return FCropFromTop;} set {FCropFromTop=value;}}

		/// <summary>
		/// How much to crop the image, in fractions of 65536 of the total image height.
		/// </summary>
		public int CropFromBottom {get {return FCropFromBottom;} set {FCropFromBottom=value;}}

		/// <summary>
		/// How much to crop the image, in fractions of 65536 of the total image width.
		/// </summary>
		public int CropFromLeft {get {return FCropFromLeft;} set {FCropFromLeft=value;}}

		/// <summary>
		/// How much to crop the image, in fractions of 65536 of the total image width.
		/// </summary>
		public int CropFromRight {get {return FCropFromRight;} set {FCropFromRight=value;}}

		#region ICloneable Members

		/// <summary>
		/// Returns a deep copy of this object
		/// </summary>
		/// <returns>A deep copy of this object.</returns>
		public object Clone()
		{
			return MemberwiseClone();
		}

		#endregion

		/// <summary>
		/// Returns true if all the coordinates are 0.
		/// </summary>
		/// <returns>True if all the coordinates are 0.</returns>
		public bool IsEmpty()
		{
			return CropFromLeft == 0 && CropFromTop == 0 && CropFromRight == 0 && CropFromBottom == 0;
		}

		/// <summary>
		/// Returns true if both instances of the objects contain the same values. Instances might be different, this method will return
		/// if their values are equal. Instances can be null.
		/// </summary>
		/// <param name="a1">First instance to compare.</param>
		/// <param name="a2">Second instance to compare.</param>
		/// <returns></returns>
		public static bool EqualValues (TCropArea a1, TCropArea a2)
		{
			if (a1 == null) return a2 == null;
			if (a2 == null) return false;

			return
				a1.FCropFromTop == a2.FCropFromTop &&
				a1.FCropFromBottom == a2.FCropFromBottom &&
				a1.FCropFromLeft == a2.FCropFromLeft &&
				a1.FCropFromRight == a2.FCropFromRight;
		}


	}

    /// <summary>
    /// Specifies properties for the text in an autoshape or object.
    /// </summary>
    public class TObjectTextProperties : ICloneable
    {
        private bool FLockText;
        private THFlxAlignment FHAlignment;
        private TVFlxAlignment FVAlignment;
        private TTextRotation FTextRotation;

        /// <summary>
        /// Creates a new instance and sets its values to defaults.
        /// </summary>
        public TObjectTextProperties()
            : this(true, THFlxAlignment.left, TVFlxAlignment.top, TTextRotation.Normal)
        {
        }

        /// <summary>
        /// Creates a new instance and sets its values.
        /// </summary>
        /// <param name="aLockText"></param>
        /// <param name="aHAlignment"></param>
        /// <param name="aVAlignment"></param>
        /// <param name="aTextRotation"></param>
        public TObjectTextProperties(bool aLockText, THFlxAlignment aHAlignment, TVFlxAlignment aVAlignment, TTextRotation aTextRotation)
        {
            FLockText = aLockText;
            FHAlignment = aHAlignment;
            FVAlignment = aVAlignment;
            FTextRotation = aTextRotation;
        }

        /// <summary>
        /// Specifies if the text of the object is locked.
        /// </summary>
        public bool LockText { get { return FLockText; } set { FLockText = value; } }

        /// <summary>
        /// Horizontal alignment for the text in the object.
        /// </summary>
        public THFlxAlignment HAlignment { get { return FHAlignment; } set { FHAlignment = value; } }

        /// <summary>
        /// Horizontal alignment for the text in the object.
        /// </summary>
        public TVFlxAlignment VAlignment { get { return FVAlignment; } set { FVAlignment = value; } }

        /// <summary>
        /// Determines how the text is orientated in the object.
        /// </summary>
        public TTextRotation TextRotation { get { return FTextRotation; } set { FTextRotation = value; } }

        /// <summary>
        /// Returns true if both instances of the objects contain the same values. Instances might be different, this method will return
        /// if their values are equal. Instances can be null.
        /// </summary>
        /// <param name="a1">First instance to compare.</param>
        /// <param name="a2">Second instance to compare.</param>
        /// <returns></returns>
        public static bool EqualValues(TObjectTextProperties a1, TObjectTextProperties a2)
        {
            if (a1 == null) return a2 == null;
            if (a2 == null) return false;

            return
                a1.FLockText == a2.FLockText &&
                a1.FHAlignment == a2.FHAlignment &&
                a1.FVAlignment == a2.FVAlignment &&
                a1.FTextRotation == a2.FTextRotation;
        }

        #region ICloneable Members

        /// <summary>
        ///  Creates a new object that is a copy of the current instance.
        /// </summary>
        /// <returns></returns>
        public object Clone()
        {
            return MemberwiseClone();
        }

        #endregion
    }

    /// <summary>
	/// Image information, for headers and footers, normal images or objects in general.
	/// </summary>
	public abstract class TBaseImageProperties
	{
		private string FFileName;
		private TCropArea FCropArea;
		private long FTransparentColor = FlxConsts.NoTransparentColor;
		private int FContrast= FlxConsts.DefaultContrast;
		private int FBrightness = FlxConsts.DefaultBrightness;
		private int FGamma = FlxConsts.DefaultGamma;
        private bool FLock = true;
        private bool FPrint = true;
        private bool FPublished = false;
        private bool FDisabled = false;
        private bool FDefaultSize = false;
        private bool FAutoFill = false;
        private bool FAutoLine = false;
        private string FMacro;

        private bool FPreferRelativeSize = true;
        private bool FLockAspectRatio = true;
        private bool FGrayscale;
        private bool FBiLevel;

        internal TBlipFill BlipFill = null;


        /// <summary>
        /// This variable must be set by the class inheriting this one.
        /// </summary>
        protected bool FDefaultsToLockedAspectRatio = false;


        /// <summary>
        /// This property returns true if the shape by default locks its aspect ratio. Images do it, comments don't.
        /// You will normally not need to use this value.
        /// </summary>
        public bool DefaultsToLockedAspectRatio { get { return FDefaultsToLockedAspectRatio; } }


		/// <summary>
		/// FileName of the image. It sets/gets the original filename of the image before it was inserted.
		/// (For example: c:\image.jpg) It is not necessary to set this field, and when the image is not inserted
		/// from a file but pasted, Excel does not set it either.
		/// </summary>
		public string FileName{ get {return FFileName;} set {FFileName=value;}}
	
		/// <summary>
		/// Cropping coordinates for the Image.
		/// </summary>
		public TCropArea CropArea {get{return FCropArea;} 
			set 
			{
				if (value==null) 
					FCropArea = new TCropArea();
				else
					FCropArea=(TCropArea)value.Clone();	
			}
		}

		/// <summary>
		/// This method will NOT copy the area. Only internal use.
		/// </summary>
		/// <param name="aCropArea"></param>
		internal void SetCropArea(TCropArea aCropArea)
		{
			if (aCropArea == null)
				FCropArea = new TCropArea();
			else
				FCropArea=aCropArea;
		}

		/// <summary>
		/// Transparent Color. <see cref="FlxConsts.NoTransparentColor"/> (~0L) means no transparent color.
		/// </summary>
		public long TransparentColor  {get {return FTransparentColor ;} set {FTransparentColor =value;}}

		/// <summary>
		/// Contrast of the image. <see cref="FlxConsts.DefaultContrast"/> is the default Contrast.
		/// </summary>
		public int Contrast {get {return FContrast;} set {FContrast=value;}}

		/// <summary>
		/// Brightness of the image. <see cref="FlxConsts.DefaultBrightness"/> is the default Brightness.
		/// </summary>
		public int Brightness {get {return FBrightness;} set {FBrightness=value;}}

		/// <summary>
		/// Gamma of the image. <see cref="FlxConsts.DefaultGamma"/> is the default Gamma.
		/// </summary>
		public int Gamma {get {return FGamma;} set {FGamma=value;}}

        /// <summary>
        /// True if this image can't be selected when the sheet is protected.
        /// </summary>
        public bool Lock { get { return FLock; } set { FLock = value; } }

        /// <summary>
        /// If false, the image won't be printed.
        /// </summary>
        public bool Print { get { return FPrint; } set { FPrint = value; } }

        /// <summary>
        /// Determines if the image should be published when sent to a server. This only applies to charts.
        /// </summary>
        public bool Published { get { return FPublished; } set { FPublished = value; } }

        /// <summary>
        /// If true, the object is disabled.
        /// </summary>
        public bool Disabled { get { return FDisabled; } set { FDisabled = value; } }

        /// <summary>
        /// If true, the application is expected to choose the default size of the object.
        /// </summary>
        public bool DefaultSize { get { return FDefaultSize; } set { FDefaultSize = value; } }

        /// <summary>
        /// If true, the object uses automatic fill style.
        /// </summary>
        internal bool AutoFill { get { return FAutoFill; } set { FAutoFill = value; } }

        /// <summary>
        /// If true, the object uses automatic line style.
        /// </summary>
        internal bool AutoLine { get { return FAutoLine; } set { FAutoLine = value; } }

        /// <summary>
        /// Macro attached to the image.
        /// </summary>
        public string Macro { get { return FMacro; } set { FMacro = value; } }

        /// <summary>
        /// Specifies whether the original size of an object is saved after reformatting. 
        /// If true, the original size of the object is stored and all resizing is based on a 
        /// percentage of that original size.  Otherwise, each resizing resets the scale to 100%.
        /// </summary>
        public bool PreferRelativeSize { get { return FPreferRelativeSize; } set { FPreferRelativeSize = value; } }

        /// <summary>
        /// Specifies whether the aspect ratio of a shape is locked from being edited.
        /// </summary>
        public bool LockAspectRatio { get { return FLockAspectRatio; } set { FLockAspectRatio = value; } }

        /// <summary>
        /// Image should be displayed in grayscale.
        /// </summary>
        public bool Grayscale { get { return FGrayscale; } set { FGrayscale = value; } }

        /// <summary>
        /// If true, the image will display in 2 color black and white.
        /// </summary>
        public bool BiLevel { get { return FBiLevel; } set { FBiLevel = value; } }


		/// <summary>
		/// Returns true if both instances of the objects contain the same values. Instances might be different, this method will return
		/// if their values are equal. Instances can be null.
		/// </summary>
		/// <param name="a1">First instance to compare.</param>
		/// <param name="a2">Second instance to compare.</param>
		/// <returns></returns>
		public static bool EqualValues (TBaseImageProperties a1, TBaseImageProperties a2)
		{
			if (a1 == null) return a2 == null;
			if (a2 == null) return false;

            return
                a1.FFileName == a2.FFileName &&
                TCropArea.EqualValues(a1.CropArea, a2.CropArea) &&
                a1.FTransparentColor == a2.FTransparentColor &&
                a1.FContrast == a2.FContrast &&
                a1.FBrightness == a2.FBrightness &&
                a1.FGamma == a2.FGamma &&
                a1.FLock == a2.FLock &&
                a1.FPrint == a2.FPrint &&
                a1.FPublished == a2.FPublished &&
                a1.FMacro == a2.FMacro &&
                a1.FPreferRelativeSize == a2.FPreferRelativeSize &&
                a1.FLockAspectRatio == a2.FLockAspectRatio &&
                a1.FBiLevel == a2.FBiLevel &&
                a1.FGrayscale == a2.FGrayscale;
		}

	}

	/// <summary>
	/// Image information for a normal image.
	/// </summary>
	public class TImageProperties: TBaseImageProperties, ICloneable
	{
		private TClientAnchor FAnchor;
		private string FShapeName;
        private string FAltText;
        internal TShapeProperties ShapeOptions = null;

		/// <summary>
		/// Creates a new empty TImageProperties instance.
		/// </summary>
		public TImageProperties(): base()
		{
			CropArea = null;
		}

		/// <summary>
		/// This method will create a new copy of aAnchor, so you can modify it later.
		/// </summary>
		/// <param name="aAnchor">Anchor. It will be copied here.</param>
		/// <param name="aFileName">Filename of the image, for saving the info on the sheet. Mostly unused.</param>
		public TImageProperties(TClientAnchor aAnchor, string aFileName)
		{
			FAnchor= (TClientAnchor)aAnchor.Clone();
			FileName=aFileName;   
			CropArea = null;
            FDefaultsToLockedAspectRatio = true;
		}

		/// <summary>
		/// This method will create a new copy of aAnchor, so you can modify it later.
		/// </summary>
		/// <param name="aAnchor">Anchor. It will be copied here.</param>
		/// <param name="aFileName">Filename of the image, for saving the info on the sheet. Mostly unused.</param>
		/// <param name="aShapeName">Shape name as it will appear on the names combo box.</param>
		public TImageProperties(TClientAnchor aAnchor, string aFileName, string aShapeName)
		{
			FAnchor= (TClientAnchor)aAnchor.Clone();
			FileName=aFileName;   
			FShapeName = aShapeName;
			CropArea = null;
            FDefaultsToLockedAspectRatio = true;
        }

		/// <summary>
		/// This method will create a new copy of aAnchor and aCropArea, so you can modify it later.
		/// </summary>
		/// <param name="aAnchor">Anchor. It will be copied here.</param>
		/// <param name="aFileName">Filename of the image, for saving the info on the sheet. Mostly unused.</param>
		/// <param name="aShapeName">Name of the shape, as it appears on the names combobox. Set it to null to not add an imagename.</param>
		/// <param name="aCropArea">Crop area for the image.</param>
		public TImageProperties(TClientAnchor aAnchor, string aFileName, string aShapeName, TCropArea aCropArea)
		{
			FAnchor= (TClientAnchor)aAnchor.Clone();
			FileName=aFileName;   
			FShapeName=aShapeName;
			CropArea = aCropArea;
            FDefaultsToLockedAspectRatio = true;
        }

		/// <summary>
		/// This method WILL NOT copy the anchor, to avoid the overhead. It should be used with care, that's why it isn't public.
		/// </summary>
		internal TImageProperties(TClientAnchor aAnchor, string aFileName, string aShapeName, TCropArea aCropArea, long aTransparentColor, 
            int aBrightness, int aContrast, int aGamma, bool aLock, bool aPrint, bool aPublished, bool aDisabled, bool aDefaultSize,
            bool aAutoFill, bool aAutoLine, string aAltText, string aMacro, bool aPreferRelativeSize, bool aLockAspectRatio, bool aBiLevel, bool aGrayscale, bool DontCopyAnchor)
		{
            FDefaultsToLockedAspectRatio = true;
            if (DontCopyAnchor)
			{
				FAnchor=aAnchor;
				SetCropArea(aCropArea);
			}
			else 
			{
				FAnchor= (TClientAnchor)aAnchor.Clone();
				CropArea = aCropArea;
			}
			FileName=aFileName;   
			FShapeName = aShapeName;
			TransparentColor = aTransparentColor;
			Brightness = aBrightness;
			Contrast = aContrast;
			Gamma = aGamma;
            Lock = aLock;
            Print = aPrint;
            Published = aPublished;
            Disabled = aDisabled;
            DefaultSize = aDefaultSize;
            AutoFill = aAutoFill;
            AutoLine = aAutoLine;
            AltText = aAltText;
            Macro = aMacro;

            PreferRelativeSize = aPreferRelativeSize;
            LockAspectRatio = aLockAspectRatio;
            BiLevel = aBiLevel;
            Grayscale = aGrayscale;
		}

		/// <summary>
		/// Image position
		/// </summary>
		public TClientAnchor Anchor {get{return FAnchor;} set {FAnchor=(TClientAnchor)value.Clone();}}

		/// <summary>
		/// Name of the image. It sets/gets the name of the shape for the image as you can see it on the names combobox.
		/// If you set it to null Excel will show a generic name, like "Picture 31"
		/// </summary>
		public string ShapeName{ get {return FShapeName;} set {FShapeName=value;}}

        /// <summary>
        /// Alternative Text. This is the same as the "Alt Text" tab in the properties dialog for the image, and is used when exporting to HTML.
        /// </summary>
        public string AltText { get { return FAltText; } set { FAltText = value; } }

		/// <summary>
		/// Returns a COPY of the class with its coords incremented by 1.
		/// </summary>
		public virtual TImageProperties Inc()
		{
			TImageProperties Result= (TImageProperties) Clone();
			Result.Anchor.Row1++;
			Result.Anchor.Col1++;
			Result.Anchor.Row2++;
			Result.Anchor.Col2++;
			return Result;
		}

		/// <summary>
		/// Returns a COPY of the class with its coords decremented by 1.
		/// </summary>
		public virtual TImageProperties Dec()
		{
			TImageProperties Result= (TImageProperties) Clone();
			Result.Anchor.Row1--;
			Result.Anchor.Col1--;
			Result.Anchor.Row2--;
			Result.Anchor.Col2--;
			return Result;
		}
        
		#region ICloneable Members

		/// <summary>
		/// Performs a deep copy of the object.
		/// </summary>
		/// <returns>An TImageProperties class that is a copy of this instance, but has a different
		/// ClientAnchor. So you can change the new ClientAnchor without modifying this one.</returns>
		public object Clone()
		{
			TImageProperties Result= (TImageProperties)MemberwiseClone();
			Result.FAnchor= (TClientAnchor) FAnchor.Clone();
			Result.CropArea = CropArea;
			return Result;
		}

		#endregion

		/// <summary>
		/// Returns true if both instances of the objects contain the same values. Instances might be different, this method will return
		/// if their values are equal. Instances can be null.
		/// </summary>
		/// <param name="a1">First instance to compare.</param>
		/// <param name="a2">Second instance to compare.</param>
		/// <returns></returns>
		public static bool EqualValues (TImageProperties a1, TImageProperties a2)
		{
			if (a1 == null) return a2 == null;
			if (a2 == null) return false;

			return
				TBaseImageProperties.EqualValues(a1, a2) &&
				TClientAnchor.EqualValues(a1.Anchor, a2.Anchor) &&
				a1.FShapeName == a2.FShapeName &&
                a1.FAltText == a2.FAltText;
		}


	}
	
	/// <summary>
	/// Image information for an image embedded inside a header or footer.
	/// </summary>
	public class THeaderOrFooterImageProperties: TBaseImageProperties, ICloneable
	{
		private THeaderOrFooterAnchor FAnchor;

		/// <summary>
		/// Creates a new empty THeaderOrFooterImageProperties instance.
		/// </summary>
		public THeaderOrFooterImageProperties()
		{
			CropArea = null;
            FDefaultsToLockedAspectRatio = true;
        }

		/// <summary>
		/// This method will create a new copy of aAnchor and aCropArea, so you can modify it later.
		/// </summary>
		/// <param name="aAnchor">Anchor. It will be copied here.</param>
		/// <param name="aFileName">Filename of the image, for saving the info on the sheet. Mostly unused.</param>
		/// <param name="aCropArea">Crop area for the image.</param>
		public THeaderOrFooterImageProperties(THeaderOrFooterAnchor aAnchor, string aFileName, TCropArea aCropArea)
		{
			FAnchor= (THeaderOrFooterAnchor)aAnchor.Clone();
			FileName=aFileName;   
			CropArea = aCropArea;
            FDefaultsToLockedAspectRatio = true;
		}

		/// <summary>
		/// This method WONT copy the anchor, to avoid the overhead. It should be used with care, that's why it isn't public.
		/// </summary>
		internal THeaderOrFooterImageProperties(THeaderOrFooterAnchor aAnchor, string aFileName, TCropArea aCropArea, long aTransparentColor, 
            int aBrightness, int aContrast, int aGamma, bool aLock, bool aPrint, bool aPublished, bool aDisabled, bool aDefaultSize,
            bool aAutoFill, bool aAutoLine, string aMacro, 
            bool aPreferRelativeSize, bool aLockAspectRatio, bool aBiLevel, bool aGrayscale,
            bool DontCopyAnchor)
		{
            FDefaultsToLockedAspectRatio = true;
            
            if (DontCopyAnchor)
			{
				FAnchor=aAnchor;
				SetCropArea (aCropArea);
			}
			else 
			{
				FAnchor= (THeaderOrFooterAnchor)aAnchor.Clone();
				CropArea = aCropArea;
			}
			FileName=aFileName;   
			TransparentColor = aTransparentColor;
			Brightness = aBrightness;
			Contrast = aContrast;
			Gamma = aGamma;
            Lock = aLock;
            Print = aPrint;
            Published = aPublished;
            Disabled = aDisabled;
            DefaultSize = aDefaultSize;
            AutoFill = aAutoFill;
            AutoLine = aAutoLine;

            Macro = aMacro;

            PreferRelativeSize = aPreferRelativeSize;
            LockAspectRatio = aLockAspectRatio;
            BiLevel = aBiLevel;
            Grayscale = aGrayscale;

		}

		/// <summary>
		/// Image position
		/// </summary>
		public THeaderOrFooterAnchor Anchor {get{return FAnchor;} set {FAnchor=(THeaderOrFooterAnchor)value.Clone();}}
        
		#region ICloneable Members

		/// <summary>
		/// Performs a deep copy of the object.
		/// </summary>
		/// <returns>An TImageProperties class that is a copy of this instance, but has a different
		/// Anchor. So you can change the new Anchor without modifying this one.</returns>
		public object Clone()
		{
			THeaderOrFooterImageProperties Result= (THeaderOrFooterImageProperties)MemberwiseClone();
			Result.FAnchor= (THeaderOrFooterAnchor) FAnchor.Clone();
			Result.CropArea = CropArea;
			return Result;
		}
		#endregion

		/// <summary>
		/// Returns true if both instances of the objects contain the same values. Instances might be different, this method will return
		/// if their values are equal. Instances can be null.
		/// </summary>
		/// <param name="a1">First instance to compare.</param>
		/// <param name="a2">Second instance to compare.</param>
		/// <returns></returns>
		public static bool EqualValues (THeaderOrFooterImageProperties a1, THeaderOrFooterImageProperties a2)
		{
			if (a1 == null) return a2 == null;
			if (a2 == null) return false;

			return
				TBaseImageProperties.EqualValues(a1, a2) &&
				THeaderOrFooterAnchor.EqualValues(a1.Anchor, a2.Anchor);
		}

	}

    class TBodyPr
    {
        internal TDrawingCoordinate l;
        internal TDrawingCoordinate t;
        internal TDrawingCoordinate r;
        internal TDrawingCoordinate b;

        internal string xml;

        internal TBodyPr(TDrawingCoordinate al, TDrawingCoordinate at, TDrawingCoordinate ar, TDrawingCoordinate ab, string axml)
        {
            l = al;
            t = at;
            r = ar;
            b = ab;
            xml = axml;
        }

        public override bool Equals(object obj)
        {
            TBodyPr o2 = obj as TBodyPr;
            if (o2 == null) return false;
            return o2.xml == xml;
        }

        public override int GetHashCode()
        {
            if (xml == null) return 0;
            return xml.GetHashCode();
        }
    }

    /// <summary>
    /// Holds the properties for an object.
    /// </summary>
    public class TObjectProperties : TImageProperties
    {
        /// <summary>
        /// Internal use.
        /// </summary>
        protected internal TObjectTextProperties FTextProperties;

        /// <summary>
        /// Internal use.
        /// </summary>
        protected internal bool FAutoSize;

        /// <summary>
        /// Internal use.
        /// </summary>
        protected internal bool FHidden;

        /// <summary>
        /// Internal use.
        /// </summary>
        internal TShapeFill FShapeFill;

        /// <summary>
        /// Internal use.
        /// </summary>
        internal TShapeLine FShapeLine;

        /// <summary>
        /// Internal use.
        /// </summary>
        protected internal TRichString FText;

#if (FRAMEWORK30 && !COMPACTFRAMEWORK)
        internal TDrawingRichString FTextExt;

        internal TDrawingHyperlink HLinkClick;
        internal TDrawingHyperlink HLinkHover;
        internal TEffectProperties FEffectProperties;

        internal TShapeEffects FShapeEffects;
        internal TBodyPr BodyPr;
        internal string LstStyle;
#endif

        /// <summary>
        /// Internal use.
        /// </summary>
        protected internal TCheckboxState FCheckboxState;

        /// <summary>
        /// Internal use.
        /// </summary>
        protected internal string FLinkedFmla;

        /// <summary>
        /// Internal use.
        /// </summary>
        protected internal bool FIs3D;
    
#if (MONOTOUCH || (FRAMEWORK30 && !COMPACTFRAMEWORK))
        /// <summary>
        /// Internal use.
        /// </summary>
        internal TComboBoxProperties FComboBoxProperties;

        /// <summary>
        /// Internal use.
        /// </summary>
        internal TSpinProperties FSpinProperties;

        internal List<TObjectProperties> FGroupedShapes;
        internal TDrawingPoint Offs;
        internal Size Ext;
        internal TDrawingPoint? ChOffs;
        internal Size? ChExt;
#endif

        /// <summary>
        /// Creates a new empty instance.
        /// </summary>
        public TObjectProperties()
            : base()
        {
            FDefaultsToLockedAspectRatio = false;            
            FTextProperties = new TObjectTextProperties();
#if (FRAMEWORK30 && !COMPACTFRAMEWORK)
            FGroupedShapes = new List<TObjectProperties>();
#endif
            FAutoSize = false;
            FShapeFill = null;
        }

        /// <summary>
        /// This method will create a new copy of aAnchor, so you can modify it later.
        /// </summary>
        /// <param name="aAnchor">Anchor. It will be copied here.</param>
        public TObjectProperties(TClientAnchor aAnchor)
            : base(aAnchor, null)
        {
            FDefaultsToLockedAspectRatio = false;
            FTextProperties = new TObjectTextProperties();
#if (FRAMEWORK30 && !COMPACTFRAMEWORK)
            FGroupedShapes = new List<TObjectProperties>();
#endif
            FAutoSize = false;
            FShapeFill = null;
        }

        /// <summary>
        /// This method will create a new copy of aAnchor, so you can modify it later.
        /// </summary>
        /// <param name="aAnchor">Anchor. It will be copied here.</param>
        /// <param name="aShapeName">Shape name for the object, as it will appear on the names combo box.</param>
        public TObjectProperties(TClientAnchor aAnchor, string aShapeName)
            : base(aAnchor, null, aShapeName)
        {
            FDefaultsToLockedAspectRatio = false;
            FAutoSize = false;
            FShapeFill = null;
        }

        /// <summary>
        /// This method will create a new copy of aAnchor, so you can modify it later.
        /// </summary>
        /// <param name="aAnchor">Anchor. It will be copied here.</param>
        /// <param name="aShapeName">Name of the shape, as it appears on the names combobox. Set it to null to not add an imagename.</param>
        /// <param name="aTextProperties">Propertied of the text in the object.</param>
        /// <param name="aLock">Specifies if the object is locked.</param>
        /// <param name="aAutoSize">If true, the object will autosize to hold the text.</param>
        public TObjectProperties(TClientAnchor aAnchor, string aShapeName, TObjectTextProperties aTextProperties, bool aLock, bool aAutoSize)
            : base(aAnchor, null, aShapeName)
        {
            FDefaultsToLockedAspectRatio = false;
            Lock = aLock;
            if (aTextProperties != null) FTextProperties = (TObjectTextProperties)aTextProperties.Clone(); else FTextProperties = new TObjectTextProperties();
#if (FRAMEWORK30 && !COMPACTFRAMEWORK)
            FGroupedShapes = new List<TObjectProperties>();
#endif
            FAutoSize = aAutoSize;
            FShapeFill = null;
        }

        /// <summary>
        /// This method WILL NOT copy the anchor, to avoid the overhead. It should be used with care, that's why it isn't public.
        /// </summary>
        internal TObjectProperties(TClientAnchor aAnchor, string aFileName, string aShapeName, TCropArea aCropArea, long aTransparentColor, int aBrightness, 
            int aContrast, int aGamma, bool aLock, bool aPrint, bool aPublished, bool aDisabled, bool aDefaultSize, 
            bool aAutoFill, bool aAutoLine, string aAltText, string aMacro, bool aLockAspectRatio, 
            TObjectTextProperties aTextProperties, bool aAutoSize, TShapeFill aShapeFill, TShapeLine aShapeLine, 
            bool aHidden, bool aIs3D, bool DontCopyAnchor)

            : base(aAnchor, aFileName, aShapeName, aCropArea, aTransparentColor, aBrightness,
            aContrast, aGamma, aLock, aPrint, aPublished, aDisabled, aDefaultSize, 
            aAutoFill, aAutoLine, aAltText, aMacro, false, aLockAspectRatio, false, false, DontCopyAnchor)
        {
            FDefaultsToLockedAspectRatio = false;
            if (aTextProperties == null) FTextProperties = new TObjectTextProperties(); else FTextProperties = aTextProperties;
#if (FRAMEWORK30 && !COMPACTFRAMEWORK)
            FGroupedShapes = new List<TObjectProperties>();
#endif
            FAutoSize = aAutoSize;
            FHidden = aHidden;
            FShapeFill = aShapeFill;
            FShapeLine = aShapeLine;
            FIs3D = aIs3D;
        }

        /// <summary>
        /// Returns true if both instances of the objects contain the same values. Instances might be different, this method will return
        /// if their values are equal. Instances can be null.
        /// </summary>
        /// <param name="a1">First instance to compare.</param>
        /// <param name="a2">Second instance to compare.</param>
        /// <returns></returns>
        public static bool EqualValues(TObjectProperties a1, TObjectProperties a2)
        {
            if (a1 == null) return a2 == null;
            if (a2 == null) return false;

            return
                TImageProperties.EqualValues(a1, a2) &&
                TObjectTextProperties.EqualValues(a1.FTextProperties, a2.FTextProperties) &&
                a1.FAutoSize == a2.FAutoSize &&
                object.Equals(a1.FShapeFill, a2.FShapeFill) &&
                object.Equals(a1.FShapeLine, a2.FShapeLine) &&
#if (FRAMEWORK30 && !COMPACTFRAMEWORK)
                object.Equals(a1.FShapeEffects, a2.FShapeEffects) &&
                object.Equals(a1.FEffectProperties, a2.FEffectProperties) &&
                object.Equals(a1.HLinkClick, a2.HLinkClick) &&
                object.Equals(a1.HLinkHover, a2.HLinkHover) &&
                object.Equals(a1.FTextExt, a2.FTextExt) &&
                object.Equals(a1.BodyPr, a2.BodyPr) &&
                object.Equals(a1.LstStyle, a2.LstStyle) &&
#endif

                a1.FHidden == a2.FHidden &&
                a1.FIs3D == a2.FIs3D;
        }

        /// <summary>
        /// Fill style used to fill the background of the comment. If you are using a solid color, only Indexed colors or RGB are allowed here, if you specify something else,
        /// the color will be converted to RGB. It might be a gradient fill or a texture too. If null, default fill style will be used.
        /// </summary>
        public TShapeFill ShapeFill { get { return FShapeFill; } set { FShapeFill = value; } }

        /// <summary>
        /// Linestyle for the comment.
        /// </summary>
        public TShapeLine ShapeBorder { get { return FShapeLine; } set { FShapeLine = value; } }

#if (MONOTOUCH || (FRAMEWORK30 && !COMPACTFRAMEWORK))
        internal TDrawingPoint GetChOffs()
        {
            if (ChOffs.HasValue) return ChOffs.Value;
            return Offs;
        }

        internal Size GetChExt()
        {
            if (ChExt.HasValue) return ChExt.Value;
            return Ext;
        }
#endif
    }

    /// <summary>
    /// Holds the properties for a comment. This class is a descendant of <see cref="TObjectProperties"/>, and it adds specific behavior for a comment.
    /// </summary>
    public class TCommentProperties: TObjectProperties
    {
        /// <summary>
        /// Default color used in the comments.
        /// </summary>
        public const int DefaultFillColorRGB = unchecked((int)0xffffffe1);

        /// <summary>
        /// Default color used in the comments, as a system color.
        /// </summary>
        public const TSystemColor DefaultFillColorSystem = TSystemColor.InfoBk;

        /// <summary>
        /// Default color used in the comments, as a system color.
        /// </summary>
        public const TSystemColor DefaultLineColorSystem = TSystemColor.None;

		/// <summary>
		/// Creates a new empty TCommentProperties instance.
		/// </summary>
		public TCommentProperties(): base()
        {
            FShapeLine = new TShapeLine(true, null);
            FHidden = true;
            LockAspectRatio = false;
            PreferRelativeSize = false;
        }

		/// <summary>
		/// This method will create a new copy of aAnchor, so you can modify it later.
		/// </summary>
		/// <param name="aAnchor">Anchor. It will be copied here.</param>
		public TCommentProperties(TClientAnchor aAnchor) : base(aAnchor)
        {
            FShapeLine = new TShapeLine(true, null);
            FHidden = true;
            LockAspectRatio = false;
            PreferRelativeSize = false;
        }

		/// <summary>
		/// This method will create a new copy of aAnchor, so you can modify it later.
		/// </summary>
		/// <param name="aAnchor">Anchor. It will be copied here.</param>
		/// <param name="aShapeName">Shape name for the comment box, as it will appear on the names combo box.</param>
		public TCommentProperties(TClientAnchor aAnchor, string aShapeName): base(aAnchor, aShapeName)
        {
            FShapeLine = new TShapeLine(true, null);
            FHidden = true;
            LockAspectRatio = false;
            PreferRelativeSize = false;
        }


		/// <summary>
		/// This method will create a new copy of aAnchor, so you can modify it later.
		/// </summary>
		/// <param name="aAnchor">Anchor. It will be copied here.</param>
		/// <param name="aShapeName">Name of the shape, as it appears on the names combobox. Set it to null to not add an imagename.</param>
        /// <param name="aTextProperties">Propertied of the text in the comment.</param>
        /// <param name="aLock">Specifies if the comment is locked.</param>
        /// <param name="aAutoSize">If true, the comment box will autosize to hold the text.</param>
		public TCommentProperties(TClientAnchor aAnchor, string aShapeName, TObjectTextProperties aTextProperties, bool aLock, bool aAutoSize): 
            base(aAnchor, aShapeName, aTextProperties, aLock, aAutoSize)
        {
            FShapeLine = new TShapeLine(true, null);
            FHidden = true;
            LockAspectRatio = false;
            PreferRelativeSize = false;
        }


		/// <summary>
		/// This method WILL NOT copy the anchor, to avoid the overhead. It should be used with care, that's why it isn't public.
		/// </summary>
        internal TCommentProperties(TClientAnchor aAnchor, string aFileName, string aShapeName, TCropArea aCropArea, 
            long aTransparentColor, int aBrightness, int aContrast, int aGamma, bool aLock, 
            bool aPrint, bool aPublished, bool aDisabled, bool aDefaultSize, bool aAutoFill, bool aAutoLine, 
            string aAltText, string aMacro, bool aLockAspectRatio, TObjectTextProperties aTextProperties, bool aAutoSize, 
            TShapeFill aShapeFill, TShapeLine aShapeLine, bool aHidden, bool aIs3D, bool DontCopyAnchor)

            : base(aAnchor, aFileName, aShapeName, aCropArea, aTransparentColor, aBrightness, aContrast, aGamma, aLock, 
            aPrint, aPublished, aDisabled, aDefaultSize, aAutoFill, aAutoLine, 
            aAltText, aMacro, aLockAspectRatio, aTextProperties, aAutoSize, 
            aShapeFill, aShapeLine, aHidden, aIs3D, DontCopyAnchor) { }

        internal static TCommentProperties GetDefaultProps(int rowBase1, int colBase1, ExcelFile Workbook)
        {
            TXlsCellRange Cr = Workbook.CellMergedBounds(rowBase1, colBase1);
            return new TCommentProperties(new TClientAnchor(TFlxAnchorType.DontMoveAndDontResize, rowBase1 - 1, 8, colBase1 + 
                Cr.Right - Cr.Left + 1, 14, 75, 130, Workbook),
                String.Empty, null,
                new TCropArea(),
                FlxConsts.NoTransparentColor, FlxConsts.DefaultBrightness, FlxConsts.DefaultContrast, FlxConsts.DefaultGamma,
                true, true, false, false, false, false, true, null, null, //here commentproperties is null, so we use the default
                false, null, false, new TShapeFill(true, null), new TShapeLine(true, null),
                true, false,
                true);
        }

        /// <summary>
        /// Properties of the text in the object.
        /// </summary>
        public TObjectTextProperties TextProperties { get { return FTextProperties; } set {if (value == null) FTextProperties = new TObjectTextProperties(); else FTextProperties = value; } }

        /// <summary>
        /// If true, the comment box will adapt its size to the size of the text.
        /// </summary>
        public bool AutoSize { get { return FAutoSize; } set { FAutoSize = value; } }

        /// <summary>
        /// If true, the comment box will be hidden (this is the default).
        /// </summary>
        public bool Hidden { get { return FHidden; } set { FHidden = value; } }
    }

    /// <summary>
    /// Spin properties of a scrollbar, spinner, listbox or combobox.
    /// </summary>
    public class TSpinProperties
    {
        private int FMin;
        private int FMax;
        private int FIncr;
        private int FPage;
        int FDx = 16;
#if (FRAMEWORK30 && !COMPACTFRAMEWORK)
        internal int FVal; //Internal use.
#endif

        /// <summary>
        /// Creates an empty TSpinProperties instance.
        /// </summary>
        public TSpinProperties()
        {

        }

        /// <summary>
        /// Creates a new instance with data and a default dx of 16.
        /// </summary>
        /// <param name="aMin">Minimum value for the spin control.</param>
        /// <param name="aMax">Maximum value for the spin control.</param>
        /// <param name="aIncr">Small increment.</param>
        /// <param name="aPage">Big increment.</param>
        public TSpinProperties(int aMin, int aMax, int aIncr, int aPage)
            : this(aMin, aMax, aIncr, aPage, 16)
        { 
        }

        /// <summary>
        /// Creates a new instance with data.
        /// </summary>
        /// <param name="aMin">Minimum value for the spin control.</param>
        /// <param name="aMax">Maximum value for the spin control.</param>
        /// <param name="aIncr">Small increment.</param>
        /// <param name="aPage">Big increment.</param>
        /// <param name="aDx">Width of the scrollbar. It should normally be 16.</param>
        public TSpinProperties(int aMin, int aMax, int aIncr, int aPage, int aDx)
        {
            FMin = aMin;
            FMax = aMax;
            FIncr = aIncr;
            FPage = aPage;
            FDx = aDx;
        }
        /// <summary>
        /// Minimum value for the spinner/scrollbar.
        /// </summary>
        public int Min { get { return FMin; } set { FMin = value; } }

        /// Maximum value for the spinner/scrollbar.
        public int Max { get { return FMax; } set { FMax = value; } }

        /// <summary>
        /// How much the scrollbar moves when you press the up or down arrow. You will probably want to keep this at 1.
        /// </summary>
        public int Incr { get { return FIncr; } set { FIncr = value; } }

        /// <summary>
        /// How much the scrollbar moves when you press pgup/down.
        /// </summary>
        public int Page { get { return FPage; } set { FPage = value; } }

        /// <summary>
        /// Width of the scrollbar. It should normally be 16.
        /// </summary>
        internal int Dx { get { return FDx; } set { FDx = value; } }

        /// <summary>
        /// Returns true if both objects have the same data.
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public override bool Equals(object obj)
        {
            TSpinProperties o2 = obj as TSpinProperties;
            if (o2 == null) return false;

            return o2.Min == Min && o2.Max == Max && o2.Incr == Incr && o2.Page == Page && o2.Dx == Dx; 
        }


        /// <summary>
        /// Hashcode for the object.
        /// </summary>
        /// <returns></returns>
        public override int GetHashCode()
        {
            return HashCoder.GetHash(Min, Max, Incr, Page, Dx);
        }
    }

    /// <summary>
    /// Properties for a combobox or a listbox.
    /// </summary>
    internal class TComboBoxProperties
    {
        int FDropLines = 8;
        string FFormulaRange;
        int FSel;
        TListBoxSelectionType FSelectionType;

        public TComboBoxProperties()
        {
        }

        /// <summary>
        /// In comboboxes, how many lines will be shown in the drop down list.
        /// </summary>
        public int DropLines { get { return FDropLines; } set { FDropLines = value; } }

        /// <summary>
        /// Internal use.
        /// </summary>
        internal string FormulaRange { get { return FFormulaRange; } set { FFormulaRange = value; } }
        
        /// <summary>
        /// Internal use.
        /// </summary>
        internal int Sel { get { return FSel; } set { FSel = value; } }

        /// <summary>
        /// Internal use.
        /// </summary>
        internal TListBoxSelectionType SelectionType { get { return FSelectionType; } set { FSelectionType = value; } }
    }
    #endregion

    #region Formulas
    /// <summary>
    /// This structure is used in formulas that span more than one cell, like some array formulas, or "what-if" table formulas.
    /// </summary>
    public struct TFormulaSpan
    {
        #region Privates
        private int FRowSpan;
        private int FColSpan;
        private bool FIsNotTopLeft;
        #endregion

        #region Constructor
        /// <summary>
        /// Creates a new TFormulaSpan and sets its values.
        /// </summary>
        /// <param name="aRowSpan">How many rows the formula will use.</param>
        /// <param name="aColSpan">How many columns the formula will use.</param>
        /// <param name="aIsTopLeft">Indicates if this is the first formula of the array. Only the formula that is at the top
        /// left cell of the group will be used when setting a formula. Other formulas will be ignored, so you can copy 
        /// formulas in a loop from one place to the other without worring if the cell is at the top left or not.</param>
        public TFormulaSpan(int aRowSpan, int aColSpan, bool aIsTopLeft)
        {
            //The private storage is offset -1, so a default struct is 1 based.
            FRowSpan = aRowSpan - 1;
            FColSpan = aColSpan - 1;
            FIsNotTopLeft = !aIsTopLeft;
        }
        #endregion

        #region Properties
        /// <summary>
        /// How many rows the formula will use.
        /// </summary>
        public int RowSpan { get {return FRowSpan + 1; } set { FRowSpan = value - 1; } }

        /// <summary>
        /// How many columns the formula will use.
        /// </summary>
        public int ColSpan { get {return FColSpan + 1; } set { FColSpan = value - 1; } }
        
        /// <summary>
        /// Indicates if this is the first formula of the array. Only the formula that is at the top
        /// left cell of the group will be used when setting a formula. Other formulas will be ignored, so you can copy 
        /// formulas in a loop from one place to the other without worring if the cell is at the top left or not.
        /// </summary>
        public bool IsTopLeft { get {return !FIsNotTopLeft; } set { FIsNotTopLeft = !value; } }

        /// <summary>
        /// Returns true if this formula spans over a single cell. (the most usual case)
        /// </summary>
        public bool IsOneCell
        {
            get
            {
                return RowSpan <= 1 && ColSpan <= 1;
            }
        }

        #endregion

        #region Compare

        /// <summary>
        /// Returns true of the 2 FormulaSpans are the same.
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public override bool Equals(object obj)
        {
            if (!(obj is TFormulaSpan)) return false;
            TFormulaSpan o1 = (TFormulaSpan)obj;

            return RowSpan == o1.RowSpan && ColSpan == o1.ColSpan && IsTopLeft == o1.IsTopLeft;
        }

        /// <summary>
        /// Returns the hashcode for this object.
        /// </summary>
        /// <returns></returns>
        public override int GetHashCode()
        {
            return HashCoder.GetHash(RowSpan, ColSpan, IsTopLeft.GetHashCode());
        }

        /// <summary>
        /// This operator is overriden so you can compare this structure directly with ==
        /// </summary>
        /// <param name="o1"></param>
        /// <param name="o2"></param>
        /// <returns></returns>
        public static bool operator==(TFormulaSpan o1, TFormulaSpan o2)
        {
            return o1.Equals(o2);
        }

        /// <summary>
        /// This operator is overriden so you can compare this structure directly with !=
        /// </summary>
        public static bool operator !=(TFormulaSpan o1, TFormulaSpan o2)
        {
            return !o1.Equals(o2);
        }

        #endregion

    }



	/// <summary>
	/// An Excel formula. Use this class to pass a formula to an Excel sheet.
	/// </summary>
	/// <remarks>
	/// There are 4 public members here:
	///    1) The formula text
	///    2) The internal formula data. This member is not public.
	///    3) The formula result. If RecalcMode is not manual, this value will be ignored.
    ///    4) The formula span. Some formulas, like some array formulas, can span over more than one cell.
	/// Not all of them need to be set, in fact, you won't probably never set internal formula data.
	/// Internal formula data is used so we can make Flexcel1.SetValue(FlexCel2.GetValue(1,1)); and not have to convert 
	/// data->text->back to data.
	/// </remarks>
	public class TFormula : ICloneable, IConvertible
	{
		#region private members
		/// <summary>
		/// The formula data on excel RPN notation. This is used internally by FlexCel.
		/// </summary>
        private TParsedTokenList FData;
		
        /// <summary>
		/// The formula data for an array on excel RPN notation. This is used internally by FlexCel.
		/// </summary>
		private TParsedTokenList FArrayData;

        /// <summary>
		/// The formula result.
		/// </summary>
		private object FResult;

		/// <summary>
		/// The formula Text. 
		/// </summary>
		private string FText;

        /// <summary>
        /// Only used in array formulas /what if tables.
        /// </summary>
        private TFormulaSpan FSpan;
		#endregion

		#region public properties
		/// <summary>
		/// The formula text, as it is written on Excel. It must begin with "=" or "{" for array formulas.
		/// </summary>
		public string Text
		{
			get
			{
				return FText;
			}
			set
			{
				FText=value;
				FData=null;  //We clear the formula data when setting the text.
				FArrayData= null;
			}
		}
    
		/// <summary>
		/// The formula result.
		/// </summary>
		public object Result
		{
			get {return FResult;} set {FResult=value;}
		}

        /// <summary>
        /// For multicell formulas (like an array formula entered over more than one cell) this property says how many rows and columns the 
        /// formula uses. Normal formulas will span one single cell
        /// </summary>
        public TFormulaSpan Span
        {
            get { return FSpan; }
        }

        /*
        /// <summary>
        /// Returns the formula as an array of tokens in RPN (Reverse Polish Notation). You can use this method to replace 
        /// references in a formula, for example to replace the reference A1 by B2.
        /// </summary>
        /// <example>
        /// You could use the following code to replace all references to cell A1 to cell B2:
        /// </example>
        public TToken[] AsRpn
        {
            get
            {
                if (FData == null)
                {
                    TFormulaConvertTextToInternal Ps = new TFormulaConvertTextToInternal(Workbook, Workbook.ActiveSheet, true, Fmla.Text, true);
                    Ps.SetStartForRelativeRefs(Row, Col);
                    Ps.Parse();
                    FmlaData = Ps.GetTokens();
                    FmlaArrayData = null;  
                }

                if (FArrayData != null) return ParseData(FArrayData);
                return ParseData(FData);
            }
            set
            {
            }
        }

        private TToken[] ParseData(TParsedTokenList FArrayData)
        {
            throw new NotImplementedException();
        }
*/

		#endregion

		/// <summary>
		/// Internal formula representation. Do not modify.
        /// </summary>
        internal TParsedTokenList Data
		{
			get
			{
				return FData;
			}
			set
			{
				FData=value;
			}
		}

		/// <summary>
		/// Internal formula representation. Do not modify.
		/// </summary>
        internal TParsedTokenList ArrayData
		{
			get
			{
				return FArrayData;
			}
			set
			{
				FArrayData=value;
			}
		}

		#region Constructors
		/// <summary>
		/// Creates an empty Excel formula
		/// </summary>
		public TFormula()
		{
			FData=null;
			FArrayData=null;
			FResult=null;
			FText=String.Empty;
            FSpan = new TFormulaSpan();
		}

		/// <summary>
		/// Creates a formula with the corresponding text and result=null.
		/// </summary>
		/// <param name="aText">Formula Text</param>
		public TFormula(string aText): this(aText, null)
		{
        }

		/// <summary>
		/// Creates a formula with the corresponding text and result.
		/// </summary>
		/// <param name="aText">Formula Text</param>
		/// <param name="aResult">Formula Result</param>
		public TFormula(string aText, object aResult)
		{
			Text=aText; //this sets data=null.
			FResult=aResult;
            FSpan = new TFormulaSpan();
		}

  		/// <summary>
		/// Creates a formula that spans to more than one row or column. Use it to create multicell array formulas or what-if tables.
		/// </summary>
		/// <param name="aText">Formula Text</param>
		/// <param name="aResult">Formula Result. You will normally want to set this to null, as it will be recalculated by FlexCel.</param>
        /// <param name="aSpan">How many rows and columns this formula will span, and in which position for the array the formula is.</param>
		public TFormula(string aText, object aResult, TFormulaSpan aSpan)
		{
			Text=aText; //this sets data=null.
			FResult=aResult;
            FSpan = aSpan;
		}

		/// <summary>
		/// Creates a formula with all fields. Internal use.
        /// </summary>
        internal TFormula(string aText, object aResult, TParsedTokenList aData, TParsedTokenList aArrayData, bool AddText, TFormulaSpan aFormulaSpan)
		{
            FSpan = aFormulaSpan;

			if (!AddText && aText != null && aText.Length > 0)
			{
				Text = aText;  //This sets data = null
			}
			else
			{
				FText = aText;
				if (aData==null)
				{
					Data= null;
				}
				else
				{
                    Data = aData.Clone();
				}

				if (aArrayData==null)
				{
					ArrayData= null;
				}
				else
				{
                    ArrayData = aArrayData.Clone();
				}
			}
			Result=aResult;
		}
        #endregion
        
        internal void SetFText(string value)
		{
			FText=value;
		}

        #region ICloneable Members
        /// <summary>
		/// Returns a Deep copy of the formula.
		/// </summary>
		/// <returns></returns>
		public object Clone()
		{
            return new TFormula(Text, Result, Data, ArrayData, false, FSpan);
		}

		#endregion

		#region IConvertible Members

		/// <summary></summary>
#if (!MONOTOUCH)
		[CLSCompliant(false)]
#endif
		public ulong ToUInt64(IFormatProvider provider)
		{
			if (Result==null) return 0; else
				return Convert.ToUInt64(Result); //CF!, provider);
		}

		/// <summary></summary>
#if (!MONOTOUCH)
		[CLSCompliant(false)]
#endif
		public sbyte ToSByte(IFormatProvider provider)
		{
			if (Result==null) return 0; else
				return Convert.ToSByte(Result, provider);
		}

		/// <summary></summary>
		public double ToDouble(IFormatProvider provider)
		{
			if (Result==null) return 0; else
				return Convert.ToDouble(Result); //CF, provider);
		}

		/// <summary></summary>
		public DateTime ToDateTime(IFormatProvider provider)
		{
			if (Result==null) return DateTime.MinValue; else
				return Convert.ToDateTime(Result, provider);
		}

		/// <summary></summary>
		public float ToSingle(IFormatProvider provider)
		{
			if (Result==null) return 0; else
				return Convert.ToSingle(Result); //CF!, provider);
		}

		/// <summary></summary>
		public bool ToBoolean(IFormatProvider provider)
		{
			if (Result==null) return false; else
				return Convert.ToBoolean(Result, provider);
		}

		/// <summary></summary>
		public int ToInt32(IFormatProvider provider)
		{
			if (Result==null) return 0; else
				return Convert.ToInt32(Result); //CF!, provider);
		}

		/// <summary></summary>
#if (!MONOTOUCH)
		[CLSCompliant(false)]
#endif
		public ushort ToUInt16(IFormatProvider provider)
		{
			if (Result==null) return 0; else
				return Convert.ToUInt16(Result); //CF!, provider);
		}

		/// <summary></summary>
		public short ToInt16(IFormatProvider provider)
		{
			if (Result==null) return 0; else
				return Convert.ToInt16(Result, provider);
		}

		/// <summary></summary>
		public string ToString(IFormatProvider provider)
		{
			if (Result==null) return String.Empty; else
				return Convert.ToString(Result); //CF!, provider);
		}

		/// <summary></summary>
		public byte ToByte(IFormatProvider provider)
		{
			if (Result==null) return 0; else
				return Convert.ToByte(Result, provider);
		}

		/// <summary></summary>
		public char ToChar(IFormatProvider provider)
		{
			if (Result==null) return ' '; else
				return Convert.ToChar(Result, provider);
		}

		/// <summary></summary>
		public long ToInt64(IFormatProvider provider)
		{
			if (Result==null) return 0; else
				return Convert.ToInt64(Result); //CF!, provider);
		}

		/// <summary></summary>
		public TypeCode GetTypeCode()
		{
			return TypeCode.Object;
		}

		/// <summary></summary>
		public decimal ToDecimal(IFormatProvider provider)
		{
			if (Result==null) return 0; else
				return Convert.ToDecimal(Result, provider);
		}

		/// <summary></summary>
		public object ToType(Type conversionType, IFormatProvider provider)
		{
			return ((IConvertible)Result).ToType(conversionType, provider);
		}

		/// <summary></summary>
#if (!MONOTOUCH)
		[CLSCompliant(false)]
#endif
		public uint ToUInt32(IFormatProvider provider)
		{
			if (Result==null) return 0; else
				return Convert.ToUInt32(Result); //CF! , provider);
		}

		#endregion
		/// <summary>Returns the formula result as a string.</summary>
		public override string ToString()
		{
			if (Result==null) return String.Empty; else
				return Convert.ToString (Result);
		}

		/// <summary>
		/// Returns true if obj is equal to this instance.
		/// </summary>
		/// <param name="obj"></param>
		/// <returns></returns>
		public override bool Equals(object obj)
		{
			TFormula fm2 = obj as TFormula;
			if (fm2==null) return false;

            if (FSpan != fm2.FSpan) return false;

			if (this.Text == null)
				if (fm2.Text == null) return true;
				else
					return false;
			return Text.Equals(fm2.Text);
		}

		/// <summary>
		/// Hashcode of the formula.
		/// </summary>
		/// <returns></returns>
		public override int GetHashCode()
		{
			if (Text==null) return 0; 
			else return Text.GetHashCode();
		}
    }
    #endregion

    #region RTF
    /// <summary>
	/// A string cell value with its rich text information.
	/// RTFRuns is an array of TRTFRun structures, where each struct identifies a font style for a portion of text.
	/// For example, if you have:
	///    Value = "Hello"
	///    RTFRuns = {FirstChar:1 FontIndex=1, FirstChar=3, FontIndex=2}
	///       "H" (char 0) will be formatted with the specific cell format.
	///       "el" (chars 1 and 2) will have font number 1
	///       "lo" (chars 3 and 4) will have font number 2
	///       
	/// </summary>
	public class TRichString: ICloneable, IConvertible
	{
		#region Private Members
		/// <summary>
		/// Cell text.
		/// </summary>
		private string FValue;
		/// <summary>
		/// Rich text info.
		/// </summary>
		private TRTFRun[] FRTFRuns;

		/// <summary>
		/// List with fonts. We will copy them here to make it independent of the current Workbook.
		/// Having a reference to the workbook could lead to a memory leak, if we null the workbook but not the RichString.
		/// </summary>
		private TFlxFont[] FontList;

		#endregion

		#region Constructors
		/// <summary>
		/// Constructs a default RichString with text and RTF info.
		/// </summary>
		/// <param name="aValue">Cell text.</param>
		/// <param name="aRTFRuns">Rich text info</param>
		/// <param name="aFontList">List with the fonts to convert.</param>
		internal TRichString(String aValue, byte[]aRTFRuns, IFlexCelFontList aFontList)
		{
			Value=aValue;
			FRTFRuns=TRTFRun.ToRTFRunArray(aRTFRuns);
			FillFontList(aFontList);
		}

		/// <summary>
		/// Constructs an empty RichString.
		/// </summary>
		public TRichString()
		{
			Value=string.Empty;
			FRTFRuns=new TRTFRun[0];
			FontList= new TFlxFont[0];
		}

		/// <summary>
		/// Constructs a RichString without formatting.
		/// </summary>
		public TRichString(string aValue)
		{
			Value=aValue;
			FRTFRuns=new TRTFRun[0];
			FontList= new TFlxFont[0];
		}

		/// <summary>
		/// Constructs a default RichString with text and RTF info.
		/// </summary>
		/// <param name="aValue">Cell Text</param>
		/// <param name="aRTFRuns">Array of TRTFRuns structs. This value will be COPIED, so old reference is not used</param>
		/// <param name="aWorkbook">Workbook containing the fonts.</param>
		public TRichString(String aValue, TRTFRun[] aRTFRuns, ExcelFile aWorkbook)
		{
			Value=aValue;
			FRTFRuns=new TRTFRun[aRTFRuns.Length];
			Array.Copy(aRTFRuns,0,FRTFRuns,0,aRTFRuns.Length);
			FillFontList(aWorkbook);
		}

		/// <summary>
		/// Constructs a default RichString with text and RTF info.
		/// </summary>
		/// <param name="aValue">Cell Text</param>
		/// <param name="aRTFRuns">Array of TRTFRuns structs. This value will be COPIED, so old reference is not used</param>
		/// <param name="aFontList">List with the fonts on the workbook.</param>
		internal TRichString(String aValue, TRTFRun[] aRTFRuns, TFlxFont[] aFontList)
		{
			Value=aValue;
			FRTFRuns=new TRTFRun[aRTFRuns.Length];
			Array.Copy(aRTFRuns,0,FRTFRuns,0,aRTFRuns.Length);
			FontList=new TFlxFont[aFontList.Length];
			//Deep copy. Array.Clone and Array.CopyTo won't call Clone()
			for (int i = 0; i < aFontList.Length; i++)
			{
				if (aFontList[i]!=null) FontList[i]=(TFlxFont)aFontList[i].Clone();
			}
		}

        /// <summary>
        /// This won't copy the values. Make sure the values are not reused.
        /// </summary>
        internal TRichString(String aValue, TRTFRun[] aRTFRuns, IFlexCelFontList aFontList, bool dummy)
        {
            Value = aValue;
            FRTFRuns = aRTFRuns;
            FillFontList(aFontList);
        }

		/// <summary>
		/// Constructs a default RichString with text and RTF info. 
		/// </summary>
		/// <param name="aValue">Cell Text</param>
		/// <param name="aRichString">Rich string with the RTF values to copy. This value will be COPIED, so old reference is not used</param>
		/// <param name="offset">How many characters the RTFRun must be moved. For example: RichString(s.SubString(3), RTFRuns, 3) will adapt the RTFRun for s to the new substring.</param>
		public TRichString(String aValue, TRichString aRichString, int offset)
		{
			Value=aValue;
			int NeededRuns=0;
			TRTFRun PrevRun;
			PrevRun.FirstChar=-1;
			PrevRun.FontIndex=-1;
			foreach (TRTFRun R in aRichString.FRTFRuns)
			{
				if ((R.FirstChar>=offset)&&(R.FirstChar<offset+aValue.Length)) NeededRuns++;
				if ((R.FirstChar<=offset)&&(R.FirstChar>PrevRun.FirstChar)) PrevRun=R;
			}

			if ((PrevRun.FirstChar>=0)&&(PrevRun.FirstChar<offset)) NeededRuns++; //We need to add a new firstchar with the latest format.
            
                     
			FRTFRuns= new TRTFRun[NeededRuns];
			int i=0;
			if ((PrevRun.FirstChar>=0)&&(PrevRun.FirstChar<offset))
			{
				FRTFRuns[i]=PrevRun;
				FRTFRuns[i].FirstChar=0;
				i++;
			}
			foreach (TRTFRun R in aRichString.FRTFRuns)
				if ((R.FirstChar>=offset)&&(R.FirstChar<offset+aValue.Length))
				{
					FRTFRuns[i]=R;  //struct copy
					FRTFRuns[i].FirstChar-=offset;
					i++;
				}

			FontList=new TFlxFont[aRichString.FontList.Length];
			//Deep copy. Array.Clone and Array.CopyTo won't call Clone()
			for (int k = 0; k < aRichString.FontList.Length; k++)
			{
				if (aRichString.FontList[k]!=null) FontList[k]=(TFlxFont)aRichString.FontList[k].Clone();
			}
		}

		/// <summary>
		/// Constructs a RichString with text and RTF info, using an RTF ArrayList. 
		/// </summary>
		/// <param name="aValue">Cell Text</param>
		/// <param name="RTFRuns">ArrayList with RTFRuns.</param>
		/// <param name="aFontList">List of fonts.</param>
#if(FRAMEWORK20)
        internal TRichString(String aValue, List<TRTFRun> RTFRuns, IFlexCelFontList aFontList)
#else
		internal TRichString(String aValue, ArrayList RTFRuns, IFlexCelFontList aFontList)
#endif
		{
			Value=aValue;                           
			FRTFRuns= new TRTFRun[RTFRuns.Count];
			for (int i=0;i<FRTFRuns.Length;i++)
			{
				FRTFRuns[i].FirstChar=((TRTFRun)RTFRuns[i]).FirstChar;
				FRTFRuns[i].FontIndex=((TRTFRun)RTFRuns[i]).FontIndex;
			}

			FillFontList(aFontList);
		}

		/// <summary>
		/// Returns a new TRichString from an HTML text. Note that only some tags from
		/// HTML are converted, the ones that do not have correspondence on Excel rich text 
		/// will be discarded.<b>Note: This method is for advanced
		/// uses only. Normally you would just use <see cref="ExcelFile.SetCellFromHtml(int, int, string, int)"/></b>
		/// </summary>
		/// <param name="HtmlString">Html string we want to convert.</param>
		/// <param name="aCellFormat">Original format of the cell where we want to enter the string. Note that depending on the starting cell, the Rich string will be different.
		/// For example, if you have a cell with Red text and add a "hello &lt;b&gt; world" string, then resulting
		/// RichString will include a RED bold "world" string</param>
		/// <param name="aWorkbook">File where this string will be added.</param>
		/// <returns>A TRichString containing the converted Html.</returns>
		public static TRichString FromHtml(string HtmlString, TFlxFormat aCellFormat, ExcelFile aWorkbook)
		{
			TRichString Result = new TRichString();
			Result.SetFromHtml(HtmlString, aCellFormat, aWorkbook);
			return Result;
		}
		#endregion

		#region Private Utilities
        private void FillFontList(IFlexCelFontList aFontList)
        {
            int MaxFontIndex = 0;
            for (int i = 0; i < RTFRunCount; i++)
            {
                if (RTFRun(i).FontIndex > MaxFontIndex) MaxFontIndex = RTFRun(i).FontIndex;
            }
            FontList = new TFlxFont[MaxFontIndex + 1];
            for (int i = 0; i < RTFRunCount; i++)
            {
                if (FontList[RTFRun(i).FontIndex] == null) FontList[RTFRun(i).FontIndex] = (TFlxFont)aFontList.GetFont(RTFRun(i).FontIndex).Clone();
            }
        }
		#endregion

		#region Public Methods
		/// <summary>
		/// Text of the string without formatting. Might be null.
		/// </summary>
		public string Value
		{
			get
			{
				return FValue;
			}
			set
			{
				FValue= value;
			}
		}

		/// <summary>
		/// A run of RTF.
		/// </summary>
		/// <param name="index">Index on the list. 0-Based.</param>
		/// <returns></returns>
		public TRTFRun RTFRun(int index)
		{
			return FRTFRuns[index];
		}

        /// <summary>
        /// Internal as it doesn't clone the struct.
        /// </summary>
        /// <returns></returns>
        internal TRTFRun[] GetRuns()
        {
            return FRTFRuns;
        }

		/// <summary>
		/// Number of RTF runs.
		/// </summary>
		public int RTFRunCount
		{
			get
			{
				return FRTFRuns.Length;
			}
		}

		/// <summary>
		/// Return the font for character i.
		/// </summary>
		/// <param name="i">index of the font.</param>
		/// <returns></returns>
		public TFlxFont GetFont(int i)
		{
			return FontList[i];
		}

		/// <summary>
		/// The count of Fonts on the richstring.
		/// </summary>
		public int MaxFontIndex
		{
			get
			{
				return FontList.Length;
			}
		}


		/// <summary>
		/// Returns the string without Rich text info.
		/// </summary>
		/// <returns></returns>
		public override string ToString()
		{
			if (Value==null) return String.Empty; else
				return Value;
		}

		/// <summary>
		/// Retrieves a substring from this instance. The substring starts at a specified character position and has a specified length.
		/// </summary>
		/// <param name="index">Start of the substring (0 based)</param>
		/// <param name="count">Number of characters to copy.</param>
		public TRichString Substring(int index, int count)
		{
			return new TRichString(Value.Substring(index, count), this, index);
		}

		/// <summary>
		/// Retrieves a substring from this instance. The substring starts at a specified character position and ends at the end of the string.
		/// </summary>
		/// <param name="index">Start of the substring (0 based)</param>
		public TRichString Substring(int index)
		{
			return Substring(index, Length-index);
		}

		/// <summary>
		/// Returns the character at position index.
		/// </summary>
		public char this[int index]
		{
			get
			{
				return Value[index];
			}
		}

		/// <summary>
		/// Length of the RichString.
		/// </summary>
		public int Length
		{
			get
			{
				if (Value==null) return 0; else return Value.Length;
			}
		}


		/// <summary>
		/// Concatenates two TRichString objects. Be careful that formats will be preserved from s1 to s2.
		/// If s1 ends up in red, s2 will start with red too. If you want to avoid this, make sure s2
		/// has a font definition at character 0. ALSO, MAKE SURE YOU ARE CONCATENATING RICH STRINGS FROM THE SAME WORKBOOK, SO FONT INDEXES ARE SIMILAR.
		/// </summary>
		/// <param name="s1">First string to concatenate.</param>
		/// <param name="s2">Second string to concatenate.</param>
		/// <returns></returns>
		public static TRichString operator +(TRichString s1, TRichString s2)
		{
			if (s1 == null || s1.Value == null) return s2;
			if (s2 == null || s2.Value == null) return s1;

			int MaxFontIndex=0;			
			TRTFRun[] aRTFRuns = new TRTFRun[s1.RTFRunCount + s2.RTFRunCount];
			for (int i = 0; i < s1.RTFRunCount; i++)
			{
				aRTFRuns[i] = s1.RTFRun(i);
				if (s1.RTFRun(i).FontIndex>MaxFontIndex) MaxFontIndex=s1.RTFRun(i).FontIndex;
			}

			for (int i = 0; i < s2.RTFRunCount; i++)
			{
				aRTFRuns[i + s1.RTFRunCount] = s2.RTFRun(i);
				if (s2.RTFRun(i).FontIndex>MaxFontIndex) MaxFontIndex=s2.RTFRun(i).FontIndex;
			}

			TFlxFont[] aFontList = new TFlxFont[MaxFontIndex+1];
			for (int i = 0; i < aRTFRuns.Length; i++)
			{
				int fi = aRTFRuns[i].FontIndex;
				TRichString sn = i < s1.RTFRunCount? s1: s2;
				if (aFontList[fi]==null) aFontList[fi]=(TFlxFont) sn.FontList[fi].Clone();
			}

			return new TRichString(s1.Value + s2.Value, aRTFRuns, aFontList);
		}

		/// <summary>
		/// Adds two richstrings together. If using C#, you can just use the overloaded "+" operator to contactenate rich strings.
		/// </summary>
		/// <param name="s1"></param>
		/// <returns></returns>
		public TRichString Add(TRichString s1)
		{
			return this + s1;
		} 

		/// <summary>
		/// Trims all the whitespace at the beginning and end of the string.
		/// </summary>
		/// <returns>The trimmed string.</returns>
		public TRichString Trim()
		{
			if (Value==null) return new TRichString();
			int i=0;
			while (i<Value.Length && Value[i]==' ') i++;
			int k=Value.Length-1;
			while (k>=0 && Value[k]==' ') k--;
			if (i<=k) return Substring(i, k-i+1); else return new TRichString();
		}

		/// <summary>
		/// Trims all the whitespace at the end of the string.
		/// </summary>
		/// <returns>The trimmed string.</returns>
		public TRichString RightTrim()
		{
			if (Value==null) return new TRichString();
			int i=0;
			int k=Value.Length-1;
			while (k>=0 && Value[k]==' ') k--;
			if (i<=k) return Substring(i, k-i+1); else return new TRichString();
		}

		/// <summary>
		/// Replaces all oldValue strings with newValue strings inside the RichString. (case sensitive) 
		/// </summary>
		/// <seealso cref="System.String.Replace(string, string)"/>
		/// <param name="oldValue">String to replace.</param>
		/// <param name="newValue">String that will replace oldValue</param>
		/// <returns>A new TRichString with all oldValues replaced with newValues.</returns>
		public TRichString Replace(string oldValue, string newValue)
		{
			return Replace(oldValue, newValue, false);
		}

		/// <summary>
		/// Replaces all oldValue strings with newValue strings inside the RichString. 
		/// </summary>
		/// <seealso cref="System.String.Replace(string, string)"/>
		/// <param name="oldValue">String to replace.</param>
		/// <param name="newValue">String that will replace oldValue</param>
		/// <param name="CaseInsensitive">If true, it will not take car of case for the search.</param>
		/// <returns>A new TRichString with all oldValues replaced with newValues.</returns>
		public TRichString Replace(string oldValue, string newValue, bool CaseInsensitive)
		{
			if (Value==null) return new TRichString();
			string SearchValue = Value;
			if (CaseInsensitive)
			{
				if (oldValue != null) oldValue = oldValue.ToUpper(CultureInfo.CurrentCulture);
				SearchValue = SearchValue.ToUpper(CultureInfo.CurrentCulture);
			}
			StringBuilder sb= new StringBuilder();
			TRTFRun[] NewRTFRuns= (TRTFRun[])FRTFRuns.Clone();
			int iPos=0; int newiPos=0;
			int ofs= newValue.Length-oldValue.Length;
			while ((newiPos=SearchValue.IndexOf(oldValue,iPos))>=0)
			{
				sb.Append(Value.Substring(iPos, newiPos-iPos));
				sb.Append(newValue);
				if (ofs!=0)
					for (int i=NewRTFRuns.Length-1; i>=0;i--)
						if (FRTFRuns[i].FirstChar>=newiPos)
						{
							if (FRTFRuns[i].FirstChar>=newiPos+oldValue.Length) 
								NewRTFRuns[i].FirstChar+=ofs;
							else if (FRTFRuns[i].FirstChar>= newiPos+newValue.Length) 
								NewRTFRuns[i].FirstChar= sb.Length - 1;

						}

				iPos=newiPos+oldValue.Length;
			}
			if (Value.Length-iPos>0)sb.Append(Value.Substring(iPos, Value.Length-iPos));
			return new TRichString(sb.ToString(), NewRTFRuns, FontList);
		}

		private static string GetTagName(string Tag)
		{
			for (int i = 0; i <Tag.Length; i++)
				if (Char.IsWhiteSpace(Tag[i]))
				{
					return Tag.Substring(0, i);
				}
			return Tag;
		}

		private static bool IsAttrValue(char c)
		{
			//HTML 4.0 reference http://www.w3.org/TR/REC-html40/intro/sgmltut.html 3.2.2
			return Char.IsLetterOrDigit(c) || c == '_' || c == '-' || c == '.' || c == ':' || c == '#' || c == '+'; // # is not mentioned, but is needed for colors.
		}

		private static bool GetAttr(string Tag, ref int i, out string Attr, out string Value)
		{
			Attr = null;
			Value = null;

			while (i < Tag.Length && Char.IsWhiteSpace(Tag[i])) i++;
			int j = i;
			while (j < Tag.Length && (Char.IsLetterOrDigit(Tag[j]) || Tag[j] == '-' || Tag[j] == '_')) j++;
			int k = j;
			while (k < Tag.Length && Char.IsWhiteSpace(Tag[k])) k++;
			if (k >= Tag.Length - 1 || Tag[k] != '=') return false; //Malformed attribute.
			k++;
			while (k < Tag.Length && Char.IsWhiteSpace(Tag[k])) k++;

			int z = k;
			if (Tag[k] == '\"')
			{
				z++;
				while (z < Tag.Length && !(Tag[z]== '\"')) z++;
				if (z >= Tag.Length) return false; //Malformed attribute.
				k++;
			}
			else
				if (Tag[k] == '\'')
			{
				z++;
				while (z < Tag.Length && !(Tag[z]== '\'')) z++;
				if (z >= Tag.Length) return false; //Malformed attribute.
				k++;
			}
			else
			{
				while (z < Tag.Length && IsAttrValue(Tag[z])) z++;
			}

			Attr = Tag.Substring(i, j-i);
			Value = Tag.Substring(k,z-k);
			i = z + 1; //when quoted attrs, we need to add one more char.
			return true;
		}

    

		private static void DoFont(ExcelFile aWorkbook, TFontState FontState, TFlxFont CellFont, string Tag, string TagName, bool MsFormat)
		{
			string Attr;
			string Value;

			int i =TagName.Length;
			while (i < Tag.Length)
			{
				if (!GetAttr(Tag, ref i, out Attr, out Value)) return; //Wrong attribute

				switch (Attr.ToLower(CultureInfo.InvariantCulture))
				{
					case "color":
                        Color FontColor = GetMsHtmlColor(aWorkbook, Value);
#if (COMPACTFRAMEWORK || (!FRAMEWORK30 && !MONOTOUCH))
                        if (FontColor.IsEmpty) return;
#else
						if (FontColor.IsEmpty()) return;
#endif
                        CellFont.Color = FontColor;
 
						break;

					case "face":
						int Comma = Value.IndexOf(",");
						if (Comma < 0) 
							CellFont.Name = Value;
						else
							CellFont.Name = Value.Substring(0, Comma);
						break;

					case "point-size":
						double Result=0;
                        if (TCompactFramework.ConvertToNumber(Value, CultureInfo.InvariantCulture, out Result))
						{
							CellFont.Size20 = (int)Math.Round(Result * 20);
						}
						break;

					case "size":
                        if (MsFormat)
                        {
                            double MsSize = 0;
                            if (!TCompactFramework.ConvertToNumber(Value, CultureInfo.InvariantCulture, out MsSize)) return;
                            CellFont.Size20 = (int)Math.Round(MsSize);
                            break;
                        }

						int[] SizesInPoints = {8, 9, 12, 14, 18, 24, 34};
						string Val = Value.Trim();
						if (Val.Length < 1) return;
						int Sz = SizesInPoints.Length - 1;

						if (Val[0] == '+' || Val[0] == '-')
						{
							//Calculate Actual Size.
							double SizePt = CellFont.Size20 / 20;
							for (int a = 0; a < SizesInPoints.Length - 1; a++)
							{
								if ((SizesInPoints[a] + SizesInPoints[a + 1]) / 2 >= SizePt) 
								{
									Sz = a;
									break;
								}
							}
							double Offset = 0;
                            if (!TCompactFramework.ConvertToNumber(Val, CultureInfo.InvariantCulture, out Offset)) return;

							Sz += (int)(Math.Round(Offset));
						}
						else
						{
							double Offset = 0;
                            if (!TCompactFramework.ConvertToNumber(Val, CultureInfo.InvariantCulture, out Offset)) return;
							Sz = (int)(Math.Round(Offset)) - 1;
						}

						if (Sz < 0) Sz = 0;
						if (Sz >= SizesInPoints.Length) Sz = SizesInPoints.Length - 1;
						CellFont.Size20 = SizesInPoints[Sz] * 20;

						break;

				}
			}
		}

        private static Color GetMsHtmlColor(ExcelFile aWorkbook, string Value)
        {
#if (!COMPACTFRAMEWORK)
            //stupid ms "html", can include a "colorindex" here
            int IntColor;
            if (int.TryParse(Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out IntColor))
            {
                if (IntColor > 7) IntColor -= 7;
                if (IntColor < 56) return TExcelColor.FromIndex(IntColor).ToColor(aWorkbook);

                TSystemColor syscol = ColorUtil.GetSystemColor(IntColor - 56);
                if (syscol == TSystemColor.None)
                {
                    return ColorUtil.Empty;
                }
                return TDrawingColor.FromSystem(syscol).ToColor(aWorkbook);

            }
#endif
            return THtmlColors.GetColor(Value);
        }

		private static int GetFontIndex(TFontState FontState, TFlxFont CellFont, ExcelFile aWorkbook, THtmlTag[] Tags, int first, int last, bool MsFormat)
		{
			for (int i = first; i < last; i++)
			{
				string Tag = Tags[i].Text.ToLower(CultureInfo.InvariantCulture);
				string TagName = GetTagName(Tag);

				switch (TagName)
				{
					case "b":
					case "strong":
						CellFont.Style |= TFlxFontStyles.Bold; break;
					case "/b": 
					case "/strong": 
						CellFont.Style &= ~TFlxFontStyles.Bold; break;
					case "i": 
					case "em": 
						CellFont.Style |= TFlxFontStyles.Italic; break;
					case "/i": 
					case "/em": 
						CellFont.Style &= ~TFlxFontStyles.Italic; break;

					case "s": 
					case "strike": 
						CellFont.Style |= TFlxFontStyles.StrikeOut; break;
					case "/s": 
					case "/strike": 
						CellFont.Style &= ~TFlxFontStyles.StrikeOut; break;

					case "sub": CellFont.Style |= TFlxFontStyles.Subscript; break;
					case "/sub": CellFont.Style &= ~TFlxFontStyles.Subscript; break;
					case "sup": CellFont.Style |= TFlxFontStyles.Superscript; break;
					case "/sup": CellFont.Style &= ~TFlxFontStyles.Superscript; break;

					case "u": CellFont.Underline = TFlxUnderline.Single; break;
					case "/u": CellFont.Underline = TFlxUnderline.None; break;

					case "tt": 
						FontState.Info.Push((TFlxFont)CellFont.Clone()); CellFont.Name = "Courier New"; break;
					case "/tt": 
						FontState.Pop(CellFont); break;

					case "font":
                        FontState.Info.Push((TFlxFont)CellFont.Clone());
						DoFont(aWorkbook, FontState, CellFont, Tag, TagName, MsFormat);
						break;

					case "big":
                        FontState.Info.Push((TFlxFont)CellFont.Clone());
                        DoFont(aWorkbook, FontState, CellFont, "font size = \"+1\"", "font", MsFormat);
						break;

					case "small":
                        FontState.Info.Push((TFlxFont)CellFont.Clone());
                        DoFont(aWorkbook, FontState, CellFont, "font size = \"-1\"", "font", MsFormat);
						break;


					case "/h1":
					case "/h2":
					case "/h3":
					case "/h4":
					case "/h5":
					case "/h6":
					case "/font":
					case "/big":
					case "/small":
						FontState.Pop(CellFont);
						break;

					case "h1":
					case "h2":
					case "h3":
					case "h4":
					case "h5":
					case "h6":
						FontState.Info.Push((TFlxFont)CellFont.Clone());
						int[] HeaderSizes = {24 * 20, 18 * 20, 270, 12*20, 10 * 20, 150};
						CellFont.Size20 = HeaderSizes[((int)TagName[1] - '1')];
						CellFont.Style |= TFlxFontStyles.Bold; 
						break;


				}
			}
			return aWorkbook.AddFont(CellFont);
		}

            

		/// <summary>
		/// Sets the rich string content from an HTML Formatted string. <b>Note: This method is for advanced
		/// uses only. Normally you would just use <see cref="ExcelFile.SetCellFromHtml(int, int, string, int)"/></b>
		/// </summary>
		/// <param name="HtmlString">String with the HTML data.</param>
        /// <param name="aCellFormat">Initial format of the cell where we want to enter the html string.</param>
		/// <param name="aWorkbook">ExcelFile where the cell is.</param>
        public void SetFromHtml(string HtmlString, TFlxFormat aCellFormat, ExcelFile aWorkbook)
        {
            SetFromHtml(HtmlString, aCellFormat, aWorkbook, false);
        }    

		/// <summary>
		/// Sets the rich string content from an HTML Formatted string. <b>Note: This method is for advanced
		/// uses only. Normally you would just use <see cref="ExcelFile.SetCellFromHtml(int, int, string, int)"/></b>
		/// </summary>
		/// <param name="HtmlString">String with the HTML data.</param>
        /// <param name="aCellFormat">Initial format of the cell where we want to enter the html string.</param>
		/// <param name="aWorkbook">ExcelFile where the cell is.</param>
        /// <param name="aMsFormat">If true, we are reading a legacy object.</param>
		internal void SetFromHtml(string HtmlString, TFlxFormat aCellFormat, ExcelFile aWorkbook, bool aMsFormat)
		{
			THtmlParsedString ParsedString = new THtmlParsedString(HtmlString);
			FValue = ParsedString.Text;
			TFontState FontState = new TFontState();

			TFlxFormat fmt = aCellFormat;
			TFlxFont CellFont = fmt.Font;
			int LastFontIndex = aWorkbook.AddFont(CellFont);

			THtmlTag[] Tags = ParsedString.Tags;
			TRTFRun[] Runs = new TRTFRun[Tags.Length]; //Maximum possible, to avoid overhead of arraylist.
			int RunPos = 0;
			int i = 0;
			while (i < Tags.Length)
			{
				int k = i+1;

				while (k < Tags.Length && Tags[i].Position == Tags[k].Position)
					k++;

				int FontIndex = GetFontIndex(FontState, CellFont, aWorkbook, Tags, i, k, aMsFormat);
				if (FontIndex != LastFontIndex && FValue != null && Tags[i].Position < FValue.Length)
				{
					Runs[RunPos].FirstChar = Tags[i].Position;
					Runs[RunPos].FontIndex = FontIndex;
					LastFontIndex = FontIndex;
					RunPos++;
				}

				i = k;
			}

			FRTFRuns = new TRTFRun[RunPos];
			Array.Copy(Runs, 0, FRTFRuns, 0, RunPos);

			FillFontList(aWorkbook);
		}

		/// <summary>
		/// Returns the rich string content as an HTML Formatted string. <b>Note: This method is for advanced
		/// uses only. Normally you would just use <see cref="ExcelFile.GetHtmlFromCell"/></b>
		/// </summary>
		/// <param name="aCellFormat">Format of the cell where this string is.</param>
		/// <param name="aWorkbook">ExcelFile where the cell is.</param>
		/// <returns>The string formatted as an HTML string.</returns>
		/// <param name="htmlStyle">Specifies whether to use CSS or not.</param>
		/// <param name="htmlVersion">Version of the html returned. In XHTML, single tags have a "/" at the end, while in 4.0 they don't.</param>
		/// <param name="aEncoding">Encoder used to return the string. Use UTF-8 for notmal cases.</param>
		public string ToHtml(ExcelFile aWorkbook, TFlxFormat aCellFormat, THtmlVersion htmlVersion, THtmlStyle htmlStyle, Encoding aEncoding)
		{
			return ToHtml(aWorkbook, aCellFormat, htmlVersion, htmlStyle, aEncoding, null);
		}

		/// <summary>
		/// Returns the rich string content as an HTML Formatted string. <b>Note: This method is for advanced
		/// uses only. Normally you would just use <see cref="ExcelFile.GetHtmlFromCell"/></b>
		/// </summary>
		/// <param name="aCellFormat">Format of the cell where this string is.</param>
		/// <param name="aWorkbook">ExcelFile where the cell is.</param>
		/// <returns>The string formatted as an HTML string.</returns>
		/// <param name="htmlStyle">Specifies whether to use CSS or not.</param>
		/// <param name="htmlVersion">Version of the html returned. In XHTML, single tags have a "/" at the end, while in 4.0 they don't.</param>
		/// <param name="aEncoding">Encoder used to return the string. Use UTF-8 for notmal cases.</param>
		/// <param name="OnHtmlFont">Provide this parameter to customize what to do when different fonts are found in the string.</param>
		public string ToHtml(ExcelFile aWorkbook, TFlxFormat aCellFormat, THtmlVersion htmlVersion, THtmlStyle htmlStyle, Encoding aEncoding, IHtmlFontEvent OnHtmlFont)
		{
            return ToHtml(aWorkbook, aCellFormat, htmlVersion, htmlStyle, aEncoding, OnHtmlFont, false);
        }

        internal string ToHtml(ExcelFile aWorkbook, TFlxFormat aCellFormat, THtmlVersion htmlVersion, THtmlStyle htmlStyle, Encoding aEncoding, IHtmlFontEvent OnHtmlFont, bool MsFormat)
        {
			if (Value == null) return String.Empty;
			if (FRTFRuns == null || FRTFRuns.Length == 0) return THtmlEntities.EncodeAsHtml(Value, htmlVersion, aEncoding); //normal case.

			StringBuilder Result = new StringBuilder();
			StringBuilder TagsToClose = new StringBuilder();
			int LastStart = 0;
			TFlxFont LastFont = aCellFormat.Font;

			if (htmlStyle == THtmlStyle.Css && FRTFRuns.Length > 0 && FRTFRuns[0].FirstChar != 0)
			{
				if (aCellFormat.Font.Underline != TFlxUnderline.None || (aCellFormat.Font.Style & TFlxFontStyles.StrikeOut) != 0)
				{
					Result.Append("<span style = '");
					string TextDeco = String.Empty;
					if (aCellFormat.Font.Underline != TFlxUnderline.None) TextDeco = "underline";
					if ((aCellFormat.Font.Style & TFlxFontStyles.StrikeOut) != 0) TextDeco += " line-through";
					Result.Append("text-decoration: ");
					Result.Append(TextDeco);
					Result.Append(";'>");
					TagsToClose.Append("</span>");
				}
			}

			for (int i = 0; i < FRTFRuns.Length; i++)
			{
                if ((FRTFRuns[i].FirstChar - LastStart) < 0) continue;
				TFlxFont NextFont = aWorkbook.GetFont(FRTFRuns[i].FontIndex);
				Result.Append(THtmlEntities.EncodeAsHtml(Value.Substring(LastStart, FRTFRuns[i].FirstChar - LastStart), htmlVersion, aEncoding));
				Result.Append(THtmlTagCreator.DiffFont(aWorkbook, aCellFormat.Font, LastFont, NextFont, htmlVersion, htmlStyle, ref TagsToClose, OnHtmlFont, MsFormat));

				LastStart = FRTFRuns[i].FirstChar;
				LastFont = NextFont;
			}
			Result.Append(THtmlEntities.EncodeAsHtml(Value.Substring(LastStart, Value.Length - LastStart), htmlVersion, aEncoding));
			Result.Append(THtmlTagCreator.CloseDiffFont(TagsToClose));
			return Result.ToString();
		}

		#endregion

		#region ICloneable Members

		/// <summary>
		/// Returns a Deep copy of the Rich string.
		/// </summary>
		/// <returns></returns>
		public object Clone()
		{
			return new TRichString(Value, FRTFRuns, FontList);
		}

		#endregion

		#region IConvertible Members

		/// <summary></summary>
#if (!MONOTOUCH)
		[CLSCompliant(false)]
#endif
		public ulong ToUInt64(IFormatProvider provider)
		{
			return Convert.ToUInt64(Value, provider);
		}

		/// <summary></summary>
#if (!MONOTOUCH)
		[CLSCompliant(false)]
#endif
		public sbyte ToSByte(IFormatProvider provider)
		{
			return Convert.ToSByte(Value, provider);
		}

		/// <summary></summary>
		public double ToDouble(IFormatProvider provider)
		{
			return Convert.ToDouble(Value, provider);
		}

		/// <summary></summary>
		public DateTime ToDateTime(IFormatProvider provider)
		{
			return Convert.ToDateTime(Value, provider);
		}

		/// <summary></summary>
		public float ToSingle(IFormatProvider provider)
		{
			return Convert.ToSingle(Value, provider);
		}

		/// <summary></summary>
		public bool ToBoolean(IFormatProvider provider)
		{
			return Convert.ToBoolean(Value, provider);
		}

		/// <summary></summary>
		public int ToInt32(IFormatProvider provider)
		{
			return Convert.ToInt32(Value, provider);
		}

		/// <summary></summary>
#if (!MONOTOUCH)
		[CLSCompliant(false)]
#endif
		public ushort ToUInt16(IFormatProvider provider)
		{
			return Convert.ToUInt16(Value, provider);
		}

		/// <summary></summary>
		public short ToInt16(IFormatProvider provider)
		{
			return Convert.ToInt16(Value, provider);
		}

		/// <summary></summary>
		public string ToString(IFormatProvider provider)
		{
			return ToString();
		}

		/// <summary></summary>
		public byte ToByte(IFormatProvider provider)
		{
			return Convert.ToByte(Value, provider);
		}

		/// <summary></summary>
		public char ToChar(IFormatProvider provider)
		{
			return Convert.ToChar(Value, provider);
		}

		/// <summary></summary>
		public long ToInt64(IFormatProvider provider)
		{
			return Convert.ToInt64(Value, provider);
		}


		/// <summary></summary>
		public decimal ToDecimal(IFormatProvider provider)
		{
			return Convert.ToDecimal(Value, provider);
		}

		/// <summary></summary>
		public TypeCode GetTypeCode()
		{
			return TypeCode.Object;
		}

		/// <summary></summary>
		public object ToType(Type conversionType, IFormatProvider provider)
		{
			return ((IConvertible)Value).ToType(conversionType, provider);
		}

		/// <summary></summary>
#if (!MONOTOUCH)
		[CLSCompliant(false)]
#endif
		public uint ToUInt32(IFormatProvider provider)
		{
			return Convert.ToUInt32(Value, provider);
		}

		#endregion

        #region Implicits
        /// <summary>
        /// Converts a string to a TRichstring.
        /// </summary>
        /// <param name="s"></param>
        /// <returns></returns>
        public static implicit operator TRichString(string s)
        {
            return new TRichString(s);
        }

        /// <summary>
        /// Converts a TRichstring to a string.
        /// </summary>
        /// <param name="r"></param>
        /// <returns></returns>
        public static implicit operator String(TRichString r)
        {
            return r == null? null: r.Value;
        }
        #endregion

        #region Basic overrides
        /// <summary>
		/// Returns true when both richstrings are equal.
		/// </summary>
		/// <param name="obj"></param>
		/// <returns></returns>
		public override bool Equals(object obj)
		{
			TRichString rs2= obj as TRichString;
			if (rs2==null) return false;
			if (FValue != rs2.FValue) return false;

			if (RTFRunCount != rs2.RTFRunCount) return false;

			TRTFRun[] rt2 = rs2.FRTFRuns;

			for (int i= RTFRunCount-1; i>=0; i--)
				if (FRTFRuns[i] != rt2[i]) return false;

			return true;

		}

        /// <summary>
        /// Returns true if both RichStrings are equal.
        /// Note this is for backwards compatibility, this is a class and not immutable,
        /// so this method should return true if references are different. But that would break old code.
        /// </summary>
        /// <param name="s1"></param>
        /// <param name="s2"></param>
        /// <returns></returns>
        public static bool operator==(TRichString s1, TRichString s2)
        {
            if ((object)s1 == null) return (object)s2 == null;
            return s1.Equals(s2);
        }

        /// <summary>
        /// Returns true if both RichStrings do not have the same value.
        /// Note this is for backwards compatibility, this is a class and not immutable,
        /// so this method should return true if references are different. But that would break old code.
        /// </summary>
        /// <param name="s1"></param>
        /// <param name="s2"></param>
        /// <returns></returns>
        public static bool operator!=(TRichString s1, TRichString s2)
        {
            if ((object)s1 == null) return (object)s2 != null;
            return !(s1.Equals(s2));
        }


		/// <summary>
		/// Hashcode for this richstring.
		/// </summary>
		/// <returns></returns>
		public override int GetHashCode()
		{
			if (Value != null) return Value.GetHashCode(); //Not perfect, but good enough.
			return 0;
		}
		#endregion

	}

	internal class TFontState
	{
#if(FRAMEWORK20)
        internal Stack<TFlxFont> Info;
#else
		internal Stack Info;
#endif

		internal TFontState()
		{
#if(FRAMEWORK20)
            Info = new Stack<TFlxFont>();
#else
			Info = new Stack();
#endif
        }


		internal void Pop(TFlxFont CellFont)
		{
			if (Info.Count > 0) 
			{
				TFlxFont fnt = ((TFlxFont)Info.Pop());
				fnt.CopyTo(CellFont);
			}

		}

	}

	internal class RTFFirstCharComparer: IComparer<TRTFRun>
	{
        #region IComparer<TRTFRun> Members

        public int Compare(TRTFRun x, TRTFRun y)
        {
            return x.FirstChar.CompareTo(y.FirstChar);
        }

        #endregion
    }

	/// <summary>
	/// One RTF run for the text in a cell. FirstChar is the first (base 0) character to apply the format, and FontIndex is the font index for the text
	/// </summary>
	public struct TRTFRun
	{
        private static readonly RTFFirstCharComparer RTFFirstCharComparerMethod = new RTFFirstCharComparer();

		/// <summary>
		/// First character on the string where we will apply the font. (0 based)
		/// </summary>
		public int FirstChar; //This is really UINT16, but to keep it CLS compliant, we define it as int.

		/// <summary>
		/// Font index for this string part.
		/// </summary>
		public int FontIndex;

		/// <summary>
		/// Converts a TRTFRun array into a byte array for serialization.
		/// </summary>
		/// <param name="runs">TRTFRun array</param>
		/// <returns>Serialized byte array</returns>
        public static byte[] ToByteArray(TRTFRun[] runs)
        {
            List<TRTFRun> SortedRuns = new List<TRTFRun>(runs);
            SortedRuns.Sort(RTFFirstCharComparerMethod);
            for (int i = SortedRuns.Count - 2; i >= 0; i--)
            {
                if (SortedRuns[i].FirstChar == SortedRuns[i + 1].FirstChar) SortedRuns.RemoveAt(i);
            }

            byte[] Result = new byte[SortedRuns.Count * 4];
            unchecked
            {
                for (int i = 0; i < SortedRuns.Count; i++)
                {
                    TRTFRun run = SortedRuns[i];
                    Result[(i << 2)] = (byte)run.FirstChar;
                    Result[(i << 2) + 1] = (byte)(run.FirstChar >> 8);
                    Result[(i << 2) + 2] = (byte)run.FontIndex;
                    Result[(i << 2) + 3] = (byte)(run.FontIndex >> 8);

                }
            }
            return Result;
        }

		/// <summary>
		/// Converts a byte array into a TRTFRun array for serialization.
		/// </summary>
		/// <param name="runs">Serialized byte array</param>
		/// <returns>TRTFRun array</returns>
		public static TRTFRun[] ToRTFRunArray(byte[] runs)
		{
			TRTFRun[] Result= new TRTFRun[ runs.Length/4];
			unchecked
			{
				for (int i=0;i<Result.Length;i++)
				{
					Result[i].FirstChar= runs[(i<<2)]+ (runs[(i<<2)+1]<<8);
					Result[i].FontIndex= runs[(i<<2)+2]+ (runs[(i<<2)+3]<<8);
				}
			}
			return Result;
		}

		/// <summary>
		/// Determines whether two TRTFRun instances are equal.
		/// </summary>
		/// <param name="obj">The Object to compare with the current TRTFRun.</param>
		/// <returns>true if the specified Object is equal to the current TRTFRun; otherwise, false.</returns>
		public override bool Equals(object obj)
		{
			if (obj == null || GetType() != obj.GetType()) return false;
			TRTFRun o2=(TRTFRun)obj;
			return (FirstChar==o2.FirstChar) && (FontIndex==o2.FontIndex);
		}

		/// <summary>
		/// Determines whether two TRTFRun instances are equal. To be considered equal, they must have the same text <b>and</b>
		/// the same formatting. That is, fontindex and firstchar must be equal.
		/// </summary>
		/// <remarks>Note that 2 TRTFRuns of different files might be equal but refer to different formatting, because 
		/// FontIndex might point to different fonts.</remarks>
		/// <param name="b1">First TRTFRun instance to compare.</param>
		/// <param name="b2">Second TRTFRun instance to compare.</param>
		/// <returns>true if both objects have the same text and the same formatting.</returns>
		public static bool operator== (TRTFRun b1, TRTFRun b2)
		{
			return (b1.FirstChar==b2.FirstChar) && (b1.FontIndex==b2.FontIndex);
		}

		/// <summary>
		/// Determines whether two TRTFRun instances are different. To be considered equal, they must have the same text <b>and</b>
		/// the same formatting. That is, fontindex and firstchar must be equal.
		/// </summary>
		/// <remarks>Note that 2 TRTFRuns of different files might be equal but refer to different formatting, because 
		/// FontIndex might point to different fonts.</remarks>
		/// <param name="b1">First TRTFRun instance to compare.</param>
		/// <param name="b2">Second TRTFRun instance to compare.</param>
		/// <returns></returns>
		public static bool operator!= (TRTFRun b1, TRTFRun b2)
		{
			return !(b1 == b2);
		}

		/// <summary>
		/// Gets a hashcode for the TRTFRun instance.
		/// </summary>
		/// <returns>hashcode.</returns>
		public override int GetHashCode()
		{
			return HashCoder.GetHash(FirstChar, FontIndex);
		}
    }


    internal class TRTFList : List<TRTFRun>
    {
    }

    #endregion

    #region FlxConvert
    /// <summary>
	/// Convert converting nulls to String.Empty.
	/// </summary>
	internal sealed class FlxConvert
	{
		private FlxConvert(){}

		/// <summary>
		/// Converts a string, if null to String.Empty.
		/// </summary>
		/// <param name="v">Object to convert.</param>
		/// <returns></returns>
		public static string ToString(object v)
		{
			if (v==null) return String.Empty; else return Convert.ToString(v); //This is not CF compatible..., CultureInfo.CurrentCulture);
		}

		/// <summary>
		/// Includes pretty-printing arrays.
		/// </summary>
		/// <param name="obj"></param>
		/// <returns></returns>
		internal static string ToStringWithArrays(object obj)
		{
			object[] arr = obj as object[];
			if (arr != null)
			{
				StringBuilder sb = new StringBuilder();
				sb.Append(TFormulaMessages.TokenChar(TFormulaToken.fmOpenArray));
				string sep = String.Empty;
				foreach (object o in arr)
				{
					sb.Append(sep);
					sb.Append(ToStringWithArrays(o));
					sep = TFormulaMessages.TokenString(TFormulaToken.fmArrayColSep);
				}
				sb.Append(TFormulaMessages.TokenChar(TFormulaToken.fmCloseArray));

				return sb.ToString();
			}

			return obj.ToString();
		}

	
		public static bool TryToDouble(object v, out double d)
		{
            d = 0;
            try
            {
                string s = null;
				if (v is TRichString) s = v.ToString(); else s = v as string;
                if (s != null)
                {
                    return TCompactFramework.ConvertToNumber(s, CultureInfo.CurrentCulture, out d);
                }
                if (v is object[,]) return false;

                d = Convert.ToDouble(v, CultureInfo.CurrentCulture);
            }
            catch (InvalidCastException)
            {
                return false;
            }
            return true;

		}

		public static bool TryStringToInt(string s, out int ResultValue)
		{
			ResultValue = 0;
			if (s == null || s.Length <= 0) return false;

			for (int i = 0; i < s.Length; i++)
			{
				if (s[i] < '0' || s[i] > '9') return false;
				ResultValue = ResultValue * 10 + (int)s[i] - (int)'0';
			}

			return true;
		}

        internal static bool ToXlsxBoolean(string CellValue)
        {
            switch (CellValue)
            {
                case "1":
                case "on":
                case "true":
                    return true;

                case "0":
                case "off":
                case "false":
                    return false;
            }

            FlxMessages.ThrowException(FlxErr.ErrInvalidCellValue, CellValue);
            return false; //just to compile
        }
    }
    #endregion

    #region Clipboard
    /// <summary>
    /// Excel formats to copy/paste to/from the clipboard
    /// </summary>
    public sealed class FlexCelDataFormats
    {
        private FlexCelDataFormats(){}

        /// <summary>
        /// Native Excel Format.
        /// </summary>
        public static string Excel97
        {
            get
            {
                return "Biff8";
            }
        }
    }
    #endregion

    #region Printer
    /// <summary>
    /// Printer specific settings. It is a byte array with a Win32 DEVMODE struct.
    /// </summary>
    public class TPrinterDriverSettings
    {
        private byte[] FData;

        /// <summary>
        /// Creates a new instance of a TPrinterDriverSettings class, with a COPY of aData
        /// </summary>
        /// <param name="aData"></param>
        public TPrinterDriverSettings(byte[] aData)
        {
            if (aData!=null)
            {
                FData= new byte[aData.Length];
                aData.CopyTo(FData,0);
            }
            else
                FData=null;
        }

        /// <summary>
        /// The current printer data as a byte stream. The first 2 bytes are the operating system (0=windows) and
        /// the rest is a Win32 DEVMODE struct.
        /// </summary>
        public byte[] GetData ()
        {
            return FData;
        }

        /// <summary>
        /// Returns true if two instances have the same data.
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
		public override bool Equals(object obj)
		{
			TPrinterDriverSettings o2 = obj as TPrinterDriverSettings;
			if (o2 == null) return false;
			return FlxUtils.CompareMem(FData, o2.FData);
		}

        /// <summary>
        /// Returns true if both objects are equal.
        /// Note this is for backwards compatibility, this is a class and not immutable,
        /// so this method should return true if references are different. But that would break old code.
        /// </summary>
        /// <param name="o1"></param>
        /// <param name="o2"></param>
        /// <returns></returns>
        public static bool operator==(TPrinterDriverSettings o1, TPrinterDriverSettings o2)
        {
            if ((object)o1 == null) return (object)o2 == null;
            return o1.Equals(o2);
        }

        /// <summary>
        /// Returns true if both objects do not have the same value.
        /// Note this is for backwards compatibility, this is a class and not immutable,
        /// so this method should return true if references are different. But that would break old code.
        /// </summary>
        /// <param name="s1"></param>
        /// <param name="s2"></param>
        /// <returns></returns>
        public static bool operator!=(TPrinterDriverSettings s1, TPrinterDriverSettings s2)
        {
            if ((object)s1 == null) return (object)s2 != null;
            return !(s1.Equals(s2));
        }

        /// <summary>
        /// Returns the hashcode for this instance.
        /// </summary>
        /// <returns></returns>
		public override int GetHashCode()
		{
			if (FData == null) return 0;
			return FData.GetHashCode ();
		}
		
    }


    /// <summary>
    /// Dimensions of an Excel paper
    /// </summary>
    public class TPaperDimensions 
    {
        private string FPaperName;
        private float FWidth;
        private float FHeight;

        /// <summary>
        /// Creates a new TPaperDimensions instance.
        /// </summary>
        /// <param name="aPaperName">A string identifying the paper name</param>
        /// <param name="aWidth">Width in inches/100</param>
        /// <param name="aHeight">Height in inches/100</param>
        public TPaperDimensions(string aPaperName, float aWidth, float aHeight)
        {
            PaperName = aPaperName;
            Width = aWidth;
            Height = aHeight;
        }

        /// <summary>
        /// Creates a new TPaperDimensions instance.
        /// </summary>
        /// <param name="PaperSize">Excel standard papersize.</param>
        public TPaperDimensions(TPaperSize PaperSize)
        {
            if (Enum.IsDefined(typeof(TPaperSize), PaperSize))
                PaperName = PaperSize.ToString();
            else
                PaperName = TPaperSize.Undefined.ToString();

            Width=0;
            Height=0;

            switch (PaperSize)
            {
                    //<summary>Letter - 81/2"" x 11""</summary>
                case TPaperSize.Letter: Width=850; Height=1100; break;
                    //<summary>Letter small - 81/2"" x 11""</summary>
                case TPaperSize.Lettersmall: Width=850; Height=1100; break;
                    //<summary>Tabloid - 11"" x 17""</summary>
                case TPaperSize.Tabloid: Width=1100; Height=1700; break;
                    //<summary>Ledger - 17"" x 11""</summary>
                case TPaperSize.Ledger: Width=1700; Height=1100; break;
                    //<summary>Legal - 81/2"" x 14""</summary>
                case TPaperSize.Legal: Width=850; Height=1400; break;
                    //<summary>Statement - 51/2"" x 81/2""</summary>
                case TPaperSize.Statement: Width=550; Height=850; break;
                    //<summary>Executive - 71/4"" x 101/2""</summary>
                case TPaperSize.Executive: Width=725; Height=1050; break;
                    //<summary>A3 - 297mm x 420mm</summary>
                case TPaperSize.A3: Width=mm(297); Height=mm(420); break;
                    //<summary>A4 - 210mm x 297mm</summary>
                case TPaperSize.A4: Width=mm(210); Height=mm(297); break;
                    //<summary>A4 small - 210mm x 297mm</summary>
                case TPaperSize.A4small: Width=mm(210); Height=mm(297); break;
                    //<summary>A5 - 148mm x 210mm</summary>
                case TPaperSize.A5: Width=mm(148); Height=mm(210); break;
                    //<summary>B4 (JIS) - 257mm x 364mm</summary>
                case TPaperSize.B4_JIS: Width=mm(257); Height=mm(364); break;
                    //<summary>B5 (JIS) - 182mm x 257mm</summary>
                case TPaperSize.B5_JIS: Width=mm(182); Height=mm(257); break;
                    //<summary>Folio - 81/2"" x 13""</summary>
                case TPaperSize.Folio: Width=850; Height=1300; break;
                    //<summary>Quarto - 215mm x 275mm</summary>
                case TPaperSize.Quarto: Width=mm(215); Height=mm(275); break;
                    //<summary>10x14 - 10"" x 14""</summary>
                case TPaperSize.s10x14: Width=1000; Height=1400; break;
                    //<summary>11x17 - 11"" x 17""</summary>
                case TPaperSize.s11x17: Width=1000; Height=1700; break;
                    //<summary>Note - 81/2"" x 11""</summary>
                case TPaperSize.Note: Width=850; Height=1100; break;
                    //<summary>Envelope #9 - 37/8"" x 87/8""</summary>
                case TPaperSize.Envelope9: Width=387.5F; Height=887.5F; break;
                    //<summary>Envelope #10 - 41/8"" x 91/2""</summary>
                case TPaperSize.Envelope10: Width=412.5F; Height=950; break;
                    //<summary>Envelope #11 - 41/2"" x 103/8""</summary>
                case TPaperSize.Envelope11: Width=450; Height=1037.5F; break;
                    //<summary>Envelope #12 - 43/4"" x 11""</summary>
                case TPaperSize.Envelope12: Width=475; Height=1100; break;
                    //<summary>Envelope #14 - 5"" x 111/2""</summary>
                case TPaperSize.Envelope14: Width=500; Height=1150; break;
                    //<summary>C - 17"" x 22""</summary>
                case TPaperSize.C: Width=1700; Height=2200; break;
                    //<summary>D - 22"" x 34""</summary>
                case TPaperSize.D: Width=2200; Height=3400; break;
                    //<summary>E - 34"" x 44""</summary>
                case TPaperSize.E: Width=3400; Height=4400; break;
                    //<summary>Envelope DL - 110mm x 220mm</summary>
                case TPaperSize.EnvelopeDL: Width=mm(110); Height=mm(220); break;
                    //<summary>Envelope C5 - 162mm x 229mm</summary>
                case TPaperSize.EnvelopeC5: Width=mm(162); Height=mm(229); break;
                    //<summary>Envelope C3 - 324mm x 458mm</summary>
                case TPaperSize.EnvelopeC3: Width=mm(324); Height=mm(458); break;
                    //<summary>Envelope C4 - 229mm x 324mm</summary>
                case TPaperSize.EnvelopeC4: Width=mm(229); Height=mm(324); break;
                    //<summary>Envelope C6 - 114mm x 162mm</summary>
                case TPaperSize.EnvelopeC6: Width=mm(114); Height=mm(162); break;
                    //<summary>Envelope C6/C5 - 114mm x 229mm</summary>
                case TPaperSize.EnvelopeC6_C5: Width=mm(114); Height=mm(229); break;
                    //<summary>B4 (ISO) - 250mm x 353mm</summary>
                case TPaperSize.B4_ISO: Width=mm(250); Height=mm(353); break;
                    //<summary>B5 (ISO) - 176mm x 250mm</summary>
                case TPaperSize.B5_ISO: Width=mm(176); Height=mm(250); break;
                    //<summary>B6 (ISO) - 125mm x 176mm</summary>
                case TPaperSize.B6_ISO: Width=mm(125); Height=mm(176); break;
                    //<summary>Envelope Italy - 110mm x 230mm</summary>
                case TPaperSize.EnvelopeItaly: Width=mm(110); Height=mm(230); break;
                    //<summary>Envelope Monarch - 37/8"" x 71/2""</summary>
                case TPaperSize.EnvelopeMonarch: Width=387.5F; Height=750; break;
                    //<summary>63/4 Envelope - 35/8"" x 61/2""</summary>
                case TPaperSize.s63_4Envelope: Width=3500F/8F; Height=650; break;
                    //<summary>US Standard Fanfold - 147/8"" x 11""</summary>
                case TPaperSize.USStandardFanfold: Width=1487.5F; Height=1100; break;
                    //<summary>German Std. Fanfold - 81/2"" x 12""</summary>
                case TPaperSize.GermanStdFanfold: Width=850; Height=1200; break;
                    //<summary>German Legal Fanfold - 81/2"" x 13""</summary>
                case TPaperSize.GermanLegalFanfold: Width=850; Height=1300; break;
                    //<summary>B4 (ISO) - 250mm x 353mm</summary>
                case TPaperSize.B4_ISO_2: Width=mm(250); Height=mm(353); break;
                    //<summary>Japanese Postcard - 100mm x 148mm</summary>
                case TPaperSize.JapanesePostcard: Width=mm(100); Height=mm(148); break;
                    //<summary>9x11 - 9"" x 11""</summary>
                case TPaperSize.s9x11: Width=900; Height=1100; break;
                    //<summary>10x11 - 10"" x 11""</summary>
                case TPaperSize.s10x11: Width=1000; Height=1100; break;
                    //<summary>15x11 - 15"" x 11""</summary>
                case TPaperSize.s15x11: Width=1500; Height=1100; break;
                    //<summary>Envelope Invite - 220mm x 220mm</summary>
                case TPaperSize.EnvelopeInvite: Width=mm(220); Height=mm(220); break;
                    //<summary>Letter Extra - 91/2"" x 12""</summary>
                case TPaperSize.LetterExtra: Width=950; Height=1200; break;
                    //<summary>Legal Extra - 91/2"" x 15""</summary>
                case TPaperSize.LegalExtra: Width=950; Height=1500; break;
                    //<summary>Tabloid Extra - 1111/16"" x 18""</summary>
                case TPaperSize.TabloidExtra: Width=1168.75F; Height=1800; break;
                    //<summary>A4 Extra - 235mm x 322mm</summary>
                case TPaperSize.A4Extra: Width=mm(235); Height=mm(322); break;
                    //<summary>Letter Transverse - 81/2"" x 11""</summary>
                case TPaperSize.LetterTransverse: Width=850; Height=1100; break;
                    //<summary>A4 Transverse - 210mm x 297mm</summary>
                case TPaperSize.A4Transverse: Width=mm(210); Height=mm(297); break;
                    //<summary>Letter Extra Transv. - 91/2"" x 12""</summary>
                case TPaperSize.LetterExtraTransv: Width=950; Height=1200; break;
                    //<summary>Super A/A4 - 227mm x 356mm</summary>
                case TPaperSize.SuperA_A4: Width=mm(227); Height=mm(356); break;
                    //<summary>Super B/A3 - 305mm x 487mm</summary>
                case TPaperSize.SuperB_A3: Width=mm(305); Height=mm(487); break;
                    //<summary>Letter Plus - 812"" x 1211/16""</summary>
                case TPaperSize.LetterPlus: Width=81200; Height=1268.75F; break;
                    //<summary>A4 Plus - 210mm x 330mm</summary>
                case TPaperSize.A4Plus: Width=mm(210); Height=mm(230); break;
                    //<summary>A5 Transverse - 148mm x 210mm</summary>
                case TPaperSize.A5Transverse: Width=mm(148); Height=mm(210); break;
                    //<summary>B5 (JIS) Transverse - 182mm x 257mm</summary>
                case TPaperSize.B5_JIS_Transverse: Width=mm(182); Height=mm(257); break;
                    //<summary>A3 Extra - 322mm x 445mm</summary>
                case TPaperSize.A3Extra: Width=mm(322); Height=mm(445); break;
                    //<summary>A5 Extra - 174mm x 235mm</summary>
                case TPaperSize.A5Extra: Width=mm(174); Height=mm(235); break;
                    //<summary>B5 (ISO) Extra - 201mm x 276mm</summary>
                case TPaperSize.B5_ISO_Extra: Width=mm(201); Height=mm(276); break;
                    //<summary>A2 - 420mm x 594mm</summary>
                case TPaperSize.A2: Width=mm(420); Height=mm(594); break;
                    //<summary>A3 Transverse - 297mm x 420mm</summary>
                case TPaperSize.A3Transverse: Width=mm(297); Height=mm(420); break;
                    //<summary>A3 Extra Transverse - 322mm x 445mm</summary>
                case TPaperSize.A3ExtraTransverse: Width=mm(322); Height=mm(445); break;
                    //<summary>Dbl. Japanese Postcard - 200mm x 148mm</summary>
                case TPaperSize.DblJapanesePostcard: Width=mm(200); Height=mm(148); break;
                    //<summary>A6 - 105mm x 148mm</summary>
                case TPaperSize.A6: Width=mm(105); Height=mm(148); break;
                    //<summary>Letter Rotated - 11"" x 81/2""</summary>
                case TPaperSize.LetterRotated: Width=1100; Height=850; break;
                    //<summary>A3 Rotated - 420mm x 297mm</summary>
                case TPaperSize.A3Rotated: Width=mm(420); Height=mm(297); break;
                    //<summary>A4 Rotated - 297mm x 210mm</summary>
                case TPaperSize.A4Rotated: Width=mm(297); Height=mm(210); break;
                    //<summary>A5 Rotated - 210mm x 148mm</summary>
                case TPaperSize.A5Rotated: Width=mm(210); Height=mm(148); break;
                    //<summary>B4 (JIS) Rotated - 364mm x 257mm</summary>
                case TPaperSize.B4_JIS_Rotated: Width=mm(364); Height=mm(257); break;
                    //<summary>B5 (JIS) Rotated - 257mm x 182mm</summary>
                case TPaperSize.B5_JIS_Rotated: Width=mm(257); Height=mm(182); break;
                    //<summary>Japanese Postcard Rot. - 148mm x 100mm</summary>
                case TPaperSize.JapanesePostcardRot: Width=mm(148); Height=mm(100); break;
                    //<summary>Dbl. Jap. Postcard Rot. - 148mm x 200mm</summary>
                case TPaperSize.DblJapPostcardRot: Width=mm(148); Height=mm(200); break;
                    //<summary>A6 Rotated - 148mm x 105mm</summary>
                case TPaperSize.A6Rotated: Width=mm(148); Height=mm(105); break;
                    //<summary>B6 (JIS) - 128mm x 182mm</summary>
                case TPaperSize.B6_JIS: Width=mm(128); Height=mm(182); break;
                    //<summary>B6 (JIS) Rotated - 182mm x 128mm</summary>
                case TPaperSize.B6_JIS_Rotated: Width=mm(182); Height=mm(128); break;
                    //<summary>12x11 - 12"" x 11""</summary>
                case TPaperSize.s12x11: Width=1200; Height=1100; break;

            }
        }

        /// <summary>
        /// Converts millimeters to inches/100
        /// </summary>
        /// <param name="v">Value in millimeters</param>
        /// <returns>Value in inches/100</returns>
        public static float mm(float v)
        {
            return (float)(100F*v/25.4);
        }

        /// <summary>
        /// Converts inches/100 to millimeters
        /// </summary>
        /// <param name="v">Value in inches/100</param>
        /// <returns>Value in millimeters</returns>
        public static float in100(float v)
        {
            return (float)(25.4*v/100F);
        }

        /// <summary>
        /// Paper Name.
        /// </summary>
        public string PaperName {get{return FPaperName;} set{FPaperName=value;}}

        /// <summary>
        /// Paper width in inches/100
        /// </summary>
        public float Width {get{return FWidth;} set{FWidth=value;}}

        /// <summary>
        /// Paper height in inches/100
        /// </summary>
        public float Height {get{return FHeight;} set{FHeight=value;}}

    }
    #endregion

    #region Hash
    internal static class FlxHash
    {
        internal static long MakeHash(int row, int col)
        {
            return ((long)row << 16) + col; //column is still 16 bit in xls 2007
        }

        internal static void UnHash(long rowcol, out int row, out int col)
        {
            col = (int)(rowcol & 0xFFFF);
            row = (int)(rowcol >> 16);
        }
    }
    #endregion
}
