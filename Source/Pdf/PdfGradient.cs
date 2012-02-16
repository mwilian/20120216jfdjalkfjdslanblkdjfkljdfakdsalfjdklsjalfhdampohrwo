using System;
using System.Text;
using System.Globalization;
using System.Collections.Generic;

#if (WPF)
using RectangleF = System.Windows.Rect;
using PointF = System.Windows.Point;
using real = System.Double;
using ColorBlend = System.Windows.Media.GradientStopCollection;

using System.Windows.Media;
#else
using real = System.Single;
using System.Drawing;
using System.Drawing.Drawing2D;
using FlexCel.Core;
using System.Diagnostics;
#endif


namespace FlexCel.Pdf
{
    #region Gradients
    internal enum TGradientType
    {
        Axial,
        Radial
    }

    /// <summary>
    /// A pdf gradient. It would be nice if we could use gradients of type 1 here because of 2 advantages:
    /// 1) They can be reused. A gradient of type 2 or 3 has the coordinates implicit, so 2 gradients for 2 different cells will need 2 different objects
    /// 2) They are stored in streams, so they can be compressed with object streams.
    /// Sadly, type 1 gradients don't look to be implemented by most alternate pdf viewers, and also using a linear interpolation function (type 0)
    /// doesn't work very well with gradient positions. A type 0 function assumes all values of the x parameters to be uniform in the range, and a gradient with
    /// might have positions anywhere. Note: An special case is SetSigmaBellShape, this will create 256 values with uniform positions. To avoid creating
    /// 256 type2 functions and then join them in a big stitching type3 function, we will use a type0 function for them.
    /// </summary>
    internal class TPdfGradient: TPdfPattern, IComparable
    {
        TGradientType GradientType;
        TPdfFunction BlendFunction;
        RectangleF Coords;
        RectangleF RotatedCoords;
        PointF CenterPoint;
        string DrawingMatrix;

        public TPdfGradient(int aPatternId, TGradientType aGradientType, ColorBlend aBlendColors, RectangleF aCoords, PointF aCenterPoint, RectangleF aRotatedCoords, string aDrawingMatrix, List<TPdfFunction> FunctionList): base(aPatternId, TPdfToken.GradientPrefix)
        {
            GradientType = aGradientType;
            BlendFunction = GetBlendFunction(aBlendColors, FunctionList);
            CenterPoint = aCenterPoint;
            Coords = aCoords;
            RotatedCoords = aRotatedCoords;
            DrawingMatrix = aDrawingMatrix;
        }

        #region Blend function
        private TPdfFunction GetBlendFunction(ColorBlend BlendColors, List<TPdfFunction> FunctionList)
        {
            TPdfFunction SearchFunction = CreateBlendFunction(BlendColors, FunctionList);
            int Index = FunctionList.BinarySearch(0, FunctionList.Count, SearchFunction, null);

            if (Index < 0)
            {
                FunctionList.Insert(~Index, SearchFunction);
            }
            else SearchFunction = FunctionList[Index];
            return SearchFunction;
        }

        private TPdfFunction CreateBlendFunction(ColorBlend BlendColors, List<TPdfFunction> FunctionList)
        {
            //Decide what kind of function we want:
            //For 2 values: type2:
            //For more values: If they are uniform, type 1, else type 3.

            if (BlendCount(BlendColors) == 1)
            {
                return new TPdfType2Function(new Double[] { 0, 1 }, null, BlendColorArray(BlendColors, 0), BlendColorArray(BlendColors, 0), 1);
            }

            if (BlendCount(BlendColors) == 2)
            {
                return new TPdfType2Function(new Double[] { 0, 1 }, null, BlendColorArray(BlendColors, 1), BlendColorArray(BlendColors, 0), 1);
            }

            if (IsUniform(BlendColors))
            {
                return new TPdfType0Function(new Double[] { 0, 1 }, new Double[] { 0, 1 , 0, 1, 0, 1}, GetType0Array(BlendColors), 8, new int[] { BlendCount(BlendColors) });
            }

            return new TPdfType3Function(new Double[] { 0, 1 }, null, GetFunctions(BlendColors, FunctionList), GetBounds(BlendColors), GetEncode(BlendColors));
        }

        private TPdfFunction[] GetFunctions(ColorBlend BlendColors, List<TPdfFunction> FunctionList)
        {
            int functionCount = BlendCount(BlendColors) - 1;

            TPdfFunction[] Result = new TPdfFunction[functionCount];
            for (int i = 1; i <= functionCount; i++)
            {
                TPdfFunction SearchFunction = new TPdfType2Function(new Double[] { 0, 1 }, null, BlendColorArray(BlendColors, i), BlendColorArray(BlendColors, i-1), 1);
                int Index = FunctionList.BinarySearch(0, FunctionList.Count, SearchFunction, null);

                if (Index < 0)
                {
                    FunctionList.Insert(~Index, SearchFunction);
                }
                else SearchFunction = FunctionList[Index];
                Result[functionCount - i] = SearchFunction;
            }
            return Result;
        }

        private static double[] GetBounds(ColorBlend BlendColors)
        {
            double[] Result = new double[BlendCount(BlendColors) - 2];
            for (int i = BlendCount(BlendColors) - 2; i > 0; i--) //First and last are not written.
            {
                Result[Result.Length - i] = 1 - BlendPosition(BlendColors, i);
            }
            return Result;
        }

        private double[] GetEncode(ColorBlend BlendColors)
        {
            int functionCount = BlendCount(BlendColors) - 1;
            double[] Result = new double[2 * functionCount];
            for (int i = 1; i < Result.Length; i+=2)
            {
                Result[i] = 1;
            }
            return Result;
        }

        private static bool IsUniform(ColorBlend BlendColors)
        {
            int bc = BlendCount(BlendColors);

            if (bc <= 2) return true;
            real diff = BlendPosition(BlendColors, 1) - BlendPosition(BlendColors, 0);

            for (int i = 2; i < bc; i++)
            {
                real newdiff = BlendPosition(BlendColors, i) - BlendPosition(BlendColors, i - 1);
                if (!AlmostEqual(diff, newdiff)) return false;
            }
            return true;
        }

        private static byte[] GetType0Array(ColorBlend BlendColors)
        {
            int bc = BlendCount(BlendColors);
            byte[] Result = new byte[bc * 3];

            int i = 0;
            for (int x = 0; x < bc; x++)
            {
                Color c = BlendColor(BlendColors, bc - x - 1);
                Result[i] = c.R; i++;
                Result[i] = c.G; i++;
                Result[i] = c.B; i++;
            }


            return Result;
        }
        #endregion

        #region Framework independent Blend functions
        private static Color BlendColor(ColorBlend BlendColors, int Position)
        {
#if (WPF)
            return BlendColors[Position].Color;
#else
            return BlendColors.Colors[Position];
#endif
        }

        private static double[] BlendColorArray(ColorBlend BlendColors, int Position)
        {
            Color c = BlendColor(BlendColors, Position);
            return new double[] { c.R / 255.0, c.G / 255.0, c.B / 255.0 };
        }

        private static real BlendPosition(ColorBlend BlendColors, int Position)
        {
#if (WPF)
            return BlendColors[Position].Offset;
#else
            return BlendColors.Positions[Position];
#endif
        }

        private static int BlendCount(ColorBlend BlendColors)
        {
#if (WPF)
                return BlendColors.Count;
#else
                return BlendColors.Colors.Length;
#endif
        }
        #endregion

        #region Coords
        private static bool AlmostEqual(real a1, real a2)
        {
            return Math.Abs(a1 - a2) < 0.0005; //We are outputting 4 digits to the pdf file.
        }

        private string GetCoords()
        {
            StringBuilder Result = new StringBuilder();

            Result.Append(TPdfTokens.GetString(TPdfToken.OpenArray));
            Result.Append(PdfConv.CoordsToString(RotatedCoords.Left)); Result.Append(" ");
            Result.Append(PdfConv.CoordsToString(RotatedCoords.Top)); Result.Append(" ");
            Result.Append(PdfConv.CoordsToString(RotatedCoords.Right)); Result.Append(" ");
            Result.Append(PdfConv.CoordsToString(RotatedCoords.Bottom));  
            Result.Append(TPdfTokens.GetString(TPdfToken.CloseArray));
            return Result.ToString();
        }

        private string GetRadialCoords()
        {
            StringBuilder Result = new StringBuilder();
            real x0 = CenterPoint.X;
            real y0 = CenterPoint.Y;

            PointF corner = Coords.Location;
            if (Math.Abs(CenterPoint.X - Coords.X) < Math.Abs(CenterPoint.X - Coords.Right)) corner.X = Coords.Right;
            if (Math.Abs(CenterPoint.Y - Coords.Y) < Math.Abs(CenterPoint.Y - Coords.Bottom)) corner.Y = Coords.Bottom;
            real r = (real)Math.Sqrt(Math.Pow(corner.X - CenterPoint.X, 2) + Math.Pow(corner.Y - CenterPoint.Y, 2));

            Result.Append(TPdfTokens.GetString(TPdfToken.OpenArray));
            Result.Append(PdfConv.CoordsToString(x0)); Result.Append(" ");
            Result.Append(PdfConv.CoordsToString(y0)); Result.Append(" ");
            Result.Append(PdfConv.CoordsToString(0)); Result.Append(" ");
            Result.Append(PdfConv.CoordsToString(x0)); Result.Append(" ");
            Result.Append(PdfConv.CoordsToString(y0)); Result.Append(" ");
            Result.Append(PdfConv.CoordsToString(r)); Result.Append(" ");
            Result.Append(TPdfTokens.GetString(TPdfToken.CloseArray));
            return Result.ToString();
        }
        #endregion

        private void WriteShadingDictionary(TPdfStream DataStream, TXRefSection XRef)
        {
            TDictionaryRecord.WriteLine(DataStream, TPdfTokens.GetString(TPdfToken.ShadingName));
            TDictionaryRecord.BeginDictionary(DataStream);
            if (GradientType == TGradientType.Axial) 
                TDictionaryRecord.SaveKey(DataStream, TPdfToken.ShadingTypeName, "2");
            else
                TDictionaryRecord.SaveKey(DataStream, TPdfToken.ShadingTypeName, "3");

            TDictionaryRecord.SaveKey(DataStream, TPdfToken.ColorSpaceName, TPdfTokens.GetString(TPdfToken.DeviceRGBName));
            TDictionaryRecord.SaveKey(DataStream, TPdfToken.ExtendName, TPdfTokens.GetString(TPdfToken.OpenArray)+ "true true" + TPdfTokens.GetString(TPdfToken.CloseArray));

            if (GradientType == TGradientType.Axial)
            {
                TDictionaryRecord.SaveKey(DataStream, TPdfToken.CoordsName, GetCoords());
            }
            else
            {
                TDictionaryRecord.SaveKey(DataStream, TPdfToken.CoordsName, GetRadialCoords());
            }

            TDictionaryRecord.SaveKey(DataStream, TPdfToken.FunctionName, TIndirectRecord.GetCallObj(BlendFunction.GetFunctionObjId(DataStream, XRef)));
            TDictionaryRecord.EndDictionary(DataStream);
        }

        public void WritePatternObject(TPdfStream DataStream, TXRefSection XRef)
        {
            XRef.SetObjectOffset(PatternObjId, DataStream);
            TIndirectRecord.SaveHeader(DataStream, PatternObjId);
            TDictionaryRecord.BeginDictionary(DataStream);
            TDictionaryRecord.SaveKey(DataStream, TPdfToken.TypeName, TPdfTokens.GetString(TPdfToken.PatternName));
            TDictionaryRecord.SaveKey(DataStream, TPdfToken.PatternTypeName, "2");

            TDictionaryRecord.SaveKey(DataStream, TPdfToken.MatrixName, DrawingMatrix);

            WriteShadingDictionary(DataStream, XRef);
            TDictionaryRecord.EndDictionary(DataStream);

            TIndirectRecord.SaveTrailer(DataStream);
        }

        public void Select(TPdfStream DataStream)
        {
            TPdfBaseRecord.WriteLine(DataStream,
                TPdfTokens.GetString(TPdfToken.PatternName) + " " + TPdfTokens.GetString(TPdfToken.Commandcs)+" " +
                TPdfTokens.GetString(TPdfToken.GradientPrefix) + PatternId.ToString(CultureInfo.InvariantCulture) + " " +
                TPdfTokens.GetString(TPdfToken.Commandscn));
        }

        internal string GetSMask()
        {
            StringBuilder Result = new StringBuilder();
            Result.Append(TPdfTokens.GetString(TPdfToken.PatternName)); Result.Append(" ");
            Result.Append(TPdfTokens.GetString(TPdfToken.Commandcs)); Result.Append(" ");
            Result.Append(TPdfTokens.GetString(TPdfToken.GradientPrefix));
            Result.Append(Convert.ToString(PatternId, CultureInfo.InvariantCulture) ); Result.Append(" ");
            Result.Append(TPdfTokens.GetString(TPdfToken.Commandscn)); Result.Append(" ");
            Result.Append(PdfConv.ToRectangleWH(Coords, false)); Result.Append(" ");
            Result.Append(TPdfTokens.GetString(TPdfToken.CommandRectangle)); Result.Append(" ");
            Result.Append(TPdfTokens.GetString(TPdfToken.CommandFillPath));
            return Result.ToString();
        }

        #region IComparable Members

        public int CompareTo(object obj)
        {
            TPdfGradient p2= obj as TPdfGradient;
            if (p2==null)
                return -1;
            
            
            int Result = GradientType.CompareTo(p2.GradientType);
            if (Result != 0) return Result;

            Result = CompareCoords(Coords, p2.Coords);
            if (Result != 0) return Result;

            Result = CompareCoords(RotatedCoords, p2.RotatedCoords);
            if (Result != 0) return Result;

            if (GradientType == TGradientType.Radial)
            {
                Result = CenterPoint.X.CompareTo(p2.CenterPoint.X);
                if (Result != 0) return Result;
                Result = CenterPoint.Y.CompareTo(p2.CenterPoint.Y);
                if (Result != 0) return Result;
            }

            Result = BlendFunction.CompareTo(p2.BlendFunction);
            if (Result != 0) return Result;

            Result = String.Compare(DrawingMatrix, p2.DrawingMatrix);
            if (Result != 0) return Result;

            return 0;
        }

        private static int CompareCoords(RectangleF Coords1, RectangleF Coords2)
        {
            int Result = Coords1.Left.CompareTo(Coords2.Left);
            if (Result != 0) return Result;
            Result = Coords1.Top.CompareTo(Coords2.Top);
            if (Result != 0) return Result;
            Result = Coords1.Right.CompareTo(Coords2.Right);
            if (Result != 0) return Result;
            Result = Coords1.Bottom.CompareTo(Coords2.Bottom);
            if (Result != 0) return Result;

            return 0;
        }

        #endregion
    }
    #endregion

    #region Functions
    internal abstract class TPdfFunction : IComparable
    {
        #region Variables
        protected readonly int FunctionType;
        private int FunctionObjId;

        protected double[] Domain;
        protected double[] Range;
        #endregion

        #region Constructor
        protected TPdfFunction(int aFunctionType, double[] aDomain, double[] aRange)
        {
            FunctionType = aFunctionType;
            Domain = aDomain;
            Range = aRange;
        }
        #endregion

        #region IComparable Members

        public virtual int CompareTo(object obj)
        {
            //Remember not to compare FunctionObjId
            TPdfFunction o2 = obj as TPdfFunction;
            if (o2 == null) return 1;
            int Result = FunctionType.CompareTo(o2.FunctionType);
            if (Result != 0) return Result;

            Result = FlxUtils.CompareArray(Domain, o2.Domain);
            if (Result != 0) return Result;

            Result = FlxUtils.CompareArray(Range, o2.Range);
            if (Result != 0) return Result;

            return 0;
        }

        public override bool Equals(object obj)
        {
            return CompareTo(obj) == 0;
        }

        protected static int GetHash(double[] d)
        {
            if (d == null) return 0;
            return d.GetHashCode();
        }

        public override int GetHashCode()
        {
            return HashCoder.GetHash(FunctionType.GetHashCode(), GetHash(Domain), GetHash(Range));
        }

        #endregion

        #region Save
        public void WriteFunctionObject(TPdfStream DataStream, TXRefSection XRef, bool Compress)
        {
            XRef.SetObjectOffset(GetFunctionObjId(DataStream, XRef), DataStream);
            TIndirectRecord.SaveHeader(DataStream, FunctionObjId);
            TDictionaryRecord.BeginDictionary(DataStream);

            TDictionaryRecord.SaveKey(DataStream, TPdfToken.FunctionTypeName, FunctionType);
            if (Domain != null)
            {
                TDictionaryRecord.SaveKey(DataStream, TPdfToken.DomainName, PdfConv.ToString(Domain, true));
            }

            if (Range != null)
            {
                TDictionaryRecord.SaveKey(DataStream, TPdfToken.RangeName, PdfConv.ToString(Range, true));
            }

            SaveExtraKeys(DataStream, XRef, Compress);

            TDictionaryRecord.EndDictionary(DataStream);
            SaveStream(DataStream, XRef, Compress);

            TIndirectRecord.SaveTrailer(DataStream);
        }

        protected abstract void SaveExtraKeys(TPdfStream DataStream, TXRefSection XRef, bool Compress);

        protected virtual void SaveStream(TPdfStream DataStream, TXRefSection XRef, bool Compress)
        {
            //By default don't save a stream.
        }

        #endregion

        internal int GetFunctionObjId(TPdfStream DataStream, TXRefSection XRef)
        {
            if (FunctionObjId == 0) FunctionObjId = XRef.GetNewObject(DataStream);
            return FunctionObjId;
        }
    }

    internal class TPdfType0Function : TPdfFunction
    {
        #region Variables
        private byte[] Data;
        private byte[] CompressedData;
        private int BitsPerSample;
        private int[] Size;
        #endregion

        #region Constructor
        internal TPdfType0Function(double[] aDomain, double[] aRange, byte[] aData, int aBitsPerSample, int[] aSize)
            : base(0, aDomain, aRange)
        {
            Data = aData;
            CompressedData = TPdfStream.CompressData(aData);
            BitsPerSample = aBitsPerSample;
            Size = aSize;
        }
        #endregion

        #region IComparable Members
        public override int CompareTo(object obj)
        {
            int Result = base.CompareTo(obj);
            if (Result != 0) return Result;

            TPdfType0Function o2 = obj as TPdfType0Function;
            Debug.Assert(o2 != null, "The object should be a TPdfType0Function, this has been checked in base Compare");

            Result = FlxUtils.CompareArray(Data, o2.Data);
            if (Result != 0) return Result;

            Result = FlxUtils.CompareArray(Size, o2.Size);
            if (Result != 0) return Result;

            Result = BitsPerSample.CompareTo(o2.BitsPerSample);
            if (Result != 0) return Result;

            return Result;
        }

        public override int GetHashCode()
        {
            return HashCoder.GetHash(FunctionType.GetHashCode(), GetHash(Domain), GetHash(Range), Data.GetHashCode(), BitsPerSample.GetHashCode());
        }

        #endregion

        #region Save
        private byte[] FunctionData(bool Compress)
        {
            if (NeedsCompression(Compress)) return CompressedData; else return Data;
        }

        private bool NeedsCompression(bool Compress)
        {
            return Compress && Data.Length > CompressedData.Length + 30; // There is some overhead in the flatedecode filter, so if both are the same, we prefer non compressed.
        }

        protected override void SaveExtraKeys(TPdfStream DataStream, TXRefSection XRef, bool Compress)
        {
            Debug.Assert(DataStream.Compress == false); //It shouldn't be turned on at this place.
            TDictionaryRecord.SaveKey(DataStream, TPdfToken.LengthName, FunctionData(Compress).Length.ToString(CultureInfo.InvariantCulture));
            if (NeedsCompression(Compress)) TStreamRecord.SetFlateDecode(DataStream);

            TDictionaryRecord.SaveKey(DataStream, TPdfToken.SizeName, PdfConv.ToString(Size, true));
            TDictionaryRecord.SaveKey(DataStream, TPdfToken.BitsPerSampleName, 8);            
        }

        protected override void SaveStream(TPdfStream DataStream, TXRefSection XRef, bool Compress)
        {
            TStreamRecord.BeginSave(DataStream);
            DataStream.Write(FunctionData(Compress));
            TStreamRecord.EndSave(DataStream);
        }
        #endregion

    }

    internal class TPdfType2Function : TPdfFunction
    {
        #region Variables
        private double[] C0;
        private double[] C1;
        private double N;
        #endregion

        #region Constructor
        internal TPdfType2Function(double[] aDomain, double[] aRange, double[] aC0, double[] aC1, double aN)
            : base(2, aDomain, aRange)
        {
            C0 = aC0;
            C1 = aC1;
            N = aN;
        }
        #endregion

        #region IComparable Members
        public override int CompareTo(object obj)
        {
            int Result = base.CompareTo(obj);
            if (Result != 0) return Result;

            TPdfType2Function o2 = obj as TPdfType2Function;
            Debug.Assert(o2 != null, "The object should be a TPdfType0Function, this has been checked in base Compare");

            Result = FlxUtils.CompareArray(C0, o2.C0);
            if (Result != 0) return Result;

            Result = FlxUtils.CompareArray(C1, o2.C1);
            if (Result != 0) return Result;

            Result = N.CompareTo(o2.N);
            if (Result != 0) return Result;

            return Result;
        }

        public override int GetHashCode()
        {
            return HashCoder.GetHash(FunctionType, GetHash(Domain), GetHash(Range), GetHash(C0), GetHash(C1), N.GetHashCode());
        }

        #endregion

        #region Save
        protected override void SaveExtraKeys(TPdfStream DataStream, TXRefSection XRef, bool Compress)
        {
            TDictionaryRecord.SaveKey(DataStream, TPdfToken.C0Name, PdfConv.ToString(C0, true));
            TDictionaryRecord.SaveKey(DataStream, TPdfToken.C1Name, PdfConv.ToString(C1, true));
            TDictionaryRecord.SaveKey(DataStream, TPdfToken.NName, PdfConv.DoubleToString(N));
        }
        #endregion

    }

    internal class TPdfType3Function : TPdfFunction
    {
        #region Variables
        private TPdfFunction[] Functions;
        private double[] Bounds;
        private double[] Encode;
        #endregion

        #region Constructor
        internal TPdfType3Function(double[] aDomain, double[] aRange, TPdfFunction[] aFunctions, double[] aBounds, double[] aEncode)
            : base(3, aDomain, aRange)
        {
            Functions = aFunctions;
            Bounds = aBounds;
            Encode = aEncode;
        }
        #endregion

        #region IComparable Members
        public override int CompareTo(object obj)
        {
            int Result = base.CompareTo(obj);
            if (Result != 0) return Result;

            TPdfType3Function o2 = obj as TPdfType3Function;
            Debug.Assert(o2 != null, "The object should be a TPdfType0Function, this has been checked in base Compare");

            Result = Functions.Length.CompareTo(o2.Functions.Length);
            if (Result != 0) return Result;

            for (int i = 0; i < Functions.Length; i++)
            {
                Result = Functions[i].CompareTo(o2.Functions[i]);
                if (Result != 0) return Result;
            }

            Result = FlxUtils.CompareArray(Bounds, o2.Bounds);
            if (Result != 0) return Result;

            Result = FlxUtils.CompareArray(Encode, o2.Encode);
            if (Result != 0) return Result;

            return Result;
        }

        public override int GetHashCode()
        {
            return HashCoder.GetHash(FunctionType, GetHash(Domain), GetHash(Range), Functions.GetHashCode(), GetHash(Bounds), GetHash(Encode));
        }

        #endregion

        #region Save
        protected override void SaveExtraKeys(TPdfStream DataStream, TXRefSection XRef, bool Compress)
        {
            TDictionaryRecord.SaveKey(DataStream, TPdfToken.FunctionsName, 
                TPdfTokens.GetString(TPdfToken.OpenArray) + FunctionCalls(DataStream, XRef) + TPdfTokens.GetString(TPdfToken.CloseArray));
            TDictionaryRecord.SaveKey(DataStream, TPdfToken.BoundsName, PdfConv.ToString(Bounds, true));
            TDictionaryRecord.SaveKey(DataStream, TPdfToken.EncodeName, PdfConv.ToString(Encode, true));
        }

        private string FunctionCalls(TPdfStream DataStream, TXRefSection XRef)
        {
            StringBuilder Result = new StringBuilder(Functions.Length * 10);
            foreach (TPdfFunction fn in Functions)
            {
                Result.Append(TIndirectRecord.GetCallObj(fn.GetFunctionObjId(DataStream, XRef)));
                Result.Append(" ");
            }

            return Result.ToString();
        }
        #endregion

    }
    
    #endregion

}
