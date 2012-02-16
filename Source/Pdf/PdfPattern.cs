using System;
using System.Globalization;
using System.Text;
using FlexCel.Core;

#if (WPF)
using System.Windows.Media;
#else
using System.Drawing;
using System.Drawing.Drawing2D;
#endif

namespace FlexCel.Pdf
{
    internal abstract class TPdfPattern
    {
        protected int PatternId;
        protected int PatternObjId;
        TPdfToken PatternPrefix;
        
        protected TPdfPattern(int aPatternId, TPdfToken aPatternPrefix)
        {
            PatternId = aPatternId;
            PatternPrefix = aPatternPrefix;
        }

        public void WritePattern(TPdfStream DataStream, TXRefSection XRef)
        {
            TPdfBaseRecord.Write(DataStream, TPdfTokens.GetString(PatternPrefix) + PatternId.ToString(CultureInfo.InvariantCulture) + " ");
            PatternObjId = XRef.GetNewObject(DataStream);
            TIndirectRecord.CallObj(DataStream, PatternObjId);
        }
    }
	
	internal abstract class TPdfTexture: TPdfPattern
	{
		protected TPdfTexture(int aPatternId, TPdfToken aPatternPrefix):base(aPatternId, aPatternPrefix){}
	}

	internal class TPdfHatch: TPdfTexture, IComparable
	{
		HatchStyle PatternStyle;


		public TPdfHatch(int aPatternId, HatchStyle aBrushStyle): base (aPatternId, TPdfToken.PatternPrefix)
		{
			PatternStyle = aBrushStyle;
		}

		private static string GetColorString (Color PatternColor)
		{
			return String.Format(CultureInfo.InvariantCulture, "/CsPAT cs {0} {1} {2}", PatternColor.R / 255F, PatternColor.G / 255F, PatternColor.B / 255F);
		}

		public static int WriteColorSpace(TPdfStream DataStream, TXRefSection XRef)
		{
			int Result = XRef.GetNewObject(DataStream);
			TPdfBaseRecord.Write(DataStream, TPdfTokens.GetString(TPdfToken.PatternColorSpacePrefix) + " ");
			TIndirectRecord.CallObj(DataStream, Result);

			return Result;
		}

		public static void WriteColorSpaceObject(TPdfStream DataStream, TXRefSection XRef, int ColorSpaceId)
		{
			XRef.SetObjectOffset(ColorSpaceId, DataStream);
			TIndirectRecord.SaveHeader(DataStream, ColorSpaceId);
			TPdfBaseRecord.Write(DataStream, TPdfTokens.GetString(TPdfToken.OpenArray));
			TDictionaryRecord.SaveKey(DataStream, TPdfToken.PatternName, TPdfTokens.GetString(TPdfToken.DeviceRGBName));
			TPdfBaseRecord.Write(DataStream, TPdfTokens.GetString(TPdfToken.CloseArray));
			TIndirectRecord.SaveTrailer(DataStream);
		}

		public void Select(TPdfStream DataStream, Color PatternColor)
		{
			TPdfBaseRecord.WriteLine(DataStream, GetColorString(PatternColor) + " " + 
				TPdfTokens.GetString(TPdfToken.PatternPrefix) + PatternId.ToString(CultureInfo.InvariantCulture) + " " +
				TPdfTokens.GetString(TPdfToken.Commandscn));
		}

		public void WritePatternObject(TPdfStream DataStream, TXRefSection XRef)
		{
			int MatrixSize;
			byte[] PatternDef = GetPattern(out MatrixSize);

			XRef.SetObjectOffset(PatternObjId, DataStream);
			TIndirectRecord.SaveHeader(DataStream, PatternObjId);
			TDictionaryRecord.BeginDictionary(DataStream);
			TDictionaryRecord.SaveKey(DataStream, TPdfToken.TypeName, TPdfTokens.GetString(TPdfToken.PatternName));
			TDictionaryRecord.SaveKey(DataStream, TPdfToken.PatternTypeName, "1");
			TDictionaryRecord.SaveKey(DataStream, TPdfToken.PaintTypeName, "2");
			TDictionaryRecord.SaveKey(DataStream, TPdfToken.TilingTypeName, "2");
			TDictionaryRecord.SaveKey(DataStream, TPdfToken.BBoxName, String.Format(CultureInfo.InvariantCulture, "[0 0 {0} {0}]", MatrixSize));
			TDictionaryRecord.SaveKey(DataStream, TPdfToken.XStepName, MatrixSize.ToString(CultureInfo.InvariantCulture));
			TDictionaryRecord.SaveKey(DataStream, TPdfToken.YStepName, MatrixSize.ToString(CultureInfo.InvariantCulture));
			TDictionaryRecord.SaveKey(DataStream, TPdfToken.LengthName, PatternDef.Length);
			TDictionaryRecord.WriteLine(DataStream, TPdfTokens.GetString(TPdfToken.ResourcesName));
			TDictionaryRecord.BeginDictionary(DataStream);
			TDictionaryRecord.EndDictionary(DataStream);

			TDictionaryRecord.EndDictionary(DataStream);

			TStreamRecord.BeginSave(DataStream);
			DataStream.Write(PatternDef);
			TStreamRecord.EndSave(DataStream);

			TIndirectRecord.SaveTrailer(DataStream);
		}

		#region IComparable Members
		public int CompareTo(object obj)
		{
			TPdfHatch p2= obj as TPdfHatch;
			if (p2==null)
				return obj.GetType().GUID.CompareTo(this.GetType().GUID);
			return PatternStyle.CompareTo(p2.PatternStyle);
		}
		#endregion

		#region GetPattern
		private byte[] GetPattern(out int MatrixSize)
		{
			return Encoding.UTF8.GetBytes(GetHatchPattern(out MatrixSize).ToString());
		}

		private StringBuilder GetHatchPattern(out int MatrixSize)
		{
			MatrixSize = 2;
			StringBuilder Result = new StringBuilder();

			switch (PatternStyle)
			{               
				case HatchStyle.Percent50:
					MatrixSize = 1;
					for (int r = 0; r<2; r++)
						Result.Append(String.Format(CultureInfo.InvariantCulture, "{0} {1} {2} {3} re f ", 0.5* (r%2), r*0.5, 0.5, 0.5));
					break;

				case HatchStyle.Percent75:
					MatrixSize = 2;
					for (int r = 0; r<2; r++)
						Result.Append(String.Format(CultureInfo.InvariantCulture, "{0} {1} {2} {3} re f ", 0.5, r, 1.5, 0.5));
					for (int r = 0; r<2; r++)
					{
						Result.Append(String.Format(CultureInfo.InvariantCulture, "{0} {1} {2} {3} re f ", 0, r+0.5, 1, 0.5));
						Result.Append(String.Format(CultureInfo.InvariantCulture, "{0} {1} {2} {3} re f ", 1.5, r+0.5, 0.6, 0.6));
					}
					break;

				case HatchStyle.Percent25:
					MatrixSize = 2;
					for (int r = 0; r<4; r++)
						Result.Append(String.Format(CultureInfo.InvariantCulture, "{0} {1} {2} {3} re f ", r%2, r * 0.5, 0.5, 0.5));
					break;

				case HatchStyle.DarkHorizontal:
					MatrixSize = 4;
					Result.Append(String.Format(CultureInfo.InvariantCulture, "{0} {1} {2} {3} re f ", 0, 0, MatrixSize+1, MatrixSize /2f));
					break;

				case HatchStyle.DarkVertical:
					MatrixSize = 4;
					Result.Append(String.Format(CultureInfo.InvariantCulture, "{0} {1} {2} {3} re f ", 0, 0, MatrixSize /2f, MatrixSize+1));
					break;

				case HatchStyle.DarkUpwardDiagonal:
					MatrixSize = 4;
					Result.Append("1.5 w -4 -4 m 8 8 l S ");
					Result.Append("1.5 w -8 -4 m 4 8 l S ");
					Result.Append("1.5 w 0 -4 m 12 8 l S ");
					break;

				case HatchStyle.DarkDownwardDiagonal:
					MatrixSize = 4;
					Result.Append("1.5 w -8 8 m 4 -4 l S ");
					Result.Append("1.5 w -4 8 m 8 -4 l S ");
					Result.Append("1.5 w 0 8 m 12 -4 l S ");
					break;

				case HatchStyle.SmallCheckerBoard:
					MatrixSize = 4;
					for (int r = 0; r<2; r++)
						Result.Append(String.Format(CultureInfo.InvariantCulture, "{0} {1} {2} {3} re f ", 2* (r%2), r*2, 2, 2));
					break;

				case HatchStyle.Percent70:
					MatrixSize = 2;
					for (int r = 0; r<2; r++)
						Result.Append(String.Format(CultureInfo.InvariantCulture, "{0} {1} {2} {3} re f ", 0.5, r, 1.5, 0.5));
					for (int r = 0; r<2; r++)
					{
						Result.Append(String.Format(CultureInfo.InvariantCulture, "{0} {1} {2} {3} re f ", 0, r+0.5, 1, 0.5));
						Result.Append(String.Format(CultureInfo.InvariantCulture, "{0} {1} {2} {3} re f ", 1.5, r+0.5, 0.6, 0.6));
					}
					break;

				case HatchStyle.LightHorizontal: //  thin horz lines
					MatrixSize = 4;
					Result.Append(String.Format(CultureInfo.InvariantCulture, "{0} {1} {2} {3} re f ", 0, 0, MatrixSize+1, MatrixSize /4f));
					break;
            
				case HatchStyle.LightVertical: //  thin vert lines
					MatrixSize = 4;
					Result.Append(String.Format(CultureInfo.InvariantCulture, "{0} {1} {2} {3} re f ", 0, 0, MatrixSize /4f, MatrixSize+1));
					break;

				case HatchStyle.LightUpwardDiagonal:
					MatrixSize = 4;
					Result.Append("0.5 w -4 -4 m 8 8 l S ");
					Result.Append("0.5 w -8 -4 m 4 8 l S ");
					Result.Append("0.5 w 0 -4 m 12 8 l S ");
					break;

				case HatchStyle.LightDownwardDiagonal:
					MatrixSize = 4;
					Result.Append("0.5 w -8 8 m 4 -4 l S ");
					Result.Append("0.5 w -4 8 m 8 -4 l S ");
					Result.Append("0.5 w 0 8 m 12 -4 l S ");
					break;

				case HatchStyle.SmallGrid:
					MatrixSize = 4;
					Result.Append(String.Format(CultureInfo.InvariantCulture, "{0} {1} {2} {3} re f ", 0, 0, MatrixSize+1, MatrixSize /4f));
					Result.Append(String.Format(CultureInfo.InvariantCulture, "{0} {1} {2} {3} re f ", 0, 0, MatrixSize /4f, MatrixSize+1));
					break;
            
				case HatchStyle.Percent60:
					MatrixSize = 4;
					Result.Append("0.5 w -4 -4 m 8 8 l S ");
					Result.Append("0.5 w -8 -4 m 4 8 l S ");
					Result.Append("0.5 w 0 -4 m 12 8 l S ");
					Result.Append("0.5 w -4 8 m 8 -4 l S ");
					break;
            
				case HatchStyle.Percent10:
					MatrixSize = 2;
					for (int r = 0; r<2; r++)
						Result.Append(String.Format(CultureInfo.InvariantCulture, "{0} {1} {2} {3} re f ", r%2, r * 1, 0.5, 0.5));
					break;
            
				case HatchStyle.Percent05:
					MatrixSize = 4;
					for (int r = 0; r<4; r++)
						Result.Append(String.Format(CultureInfo.InvariantCulture, "{0} {1} {2} {3} re f ", (r%2)*2, r, 0.5, 0.5));
					break;
			}//case

			return Result;

		}

		#endregion
	}

	internal class TPdfImageTexture: TPdfTexture, IComparable
	{
		TPdfImage ImageDef;
		int ImgWidth;
		int ImgHeight;
		float[] PatternMatrix;


		public TPdfImageTexture(int aPatternId, Image aImage, float[] aPatternMatrix): base (aPatternId, TPdfToken.PatternPrefix)
		{
			ImgWidth = aImage.Width;
			ImgHeight = aImage.Height;
			PatternMatrix = aPatternMatrix;

			ImageDef = new TPdfImage(aImage, 0, null, -1, false);
		}

		public void Select(TPdfStream DataStream)
		{
			TPdfBaseRecord.WriteLine(DataStream, TPdfTokens.GetString(TPdfToken.PatternName) + " " + TPdfTokens.GetString(TPdfToken.Commandcs) + " " +
				TPdfTokens.GetString(TPdfToken.PatternPrefix) + PatternId.ToString(CultureInfo.InvariantCulture) + " " +
				TPdfTokens.GetString(TPdfToken.Commandscn));
		}

		public void WritePatternObject(TPdfStream DataStream, TXRefSection XRef)
		{
			byte[] PatternDef = GetPatternDef();
			XRef.SetObjectOffset(PatternObjId, DataStream);
			TIndirectRecord.SaveHeader(DataStream, PatternObjId);
			TDictionaryRecord.BeginDictionary(DataStream);
			TDictionaryRecord.SaveKey(DataStream, TPdfToken.TypeName, TPdfTokens.GetString(TPdfToken.PatternName));
			TDictionaryRecord.SaveKey(DataStream, TPdfToken.PatternTypeName, "1");
			TDictionaryRecord.SaveKey(DataStream, TPdfToken.PaintTypeName, "1");
			TDictionaryRecord.SaveKey(DataStream, TPdfToken.TilingTypeName, "2");
			TDictionaryRecord.SaveKey(DataStream, TPdfToken.BBoxName, String.Format(CultureInfo.InvariantCulture, "[0 0 {0} {1}]", ImgWidth, ImgHeight));
			TDictionaryRecord.SaveKey(DataStream, TPdfToken.XStepName, PdfConv.CoordsToString(ImgWidth));
			TDictionaryRecord.SaveKey(DataStream, TPdfToken.YStepName, PdfConv.CoordsToString(ImgHeight));
			TDictionaryRecord.SaveKey(DataStream, TPdfToken.MatrixName, PdfConv.ToString(PatternMatrix,true));
			TDictionaryRecord.SaveKey(DataStream, TPdfToken.LengthName, PatternDef.Length);
			TDictionaryRecord.WriteLine(DataStream, TPdfTokens.GetString(TPdfToken.ResourcesName));
			
			TDictionaryRecord.BeginDictionary(DataStream);	    
			TPdfBaseRecord.WriteLine(DataStream, TPdfTokens.GetString(TPdfToken.XObjectName));
			TDictionaryRecord.BeginDictionary(DataStream);

			ImageDef.WriteImage(DataStream, XRef);
			TDictionaryRecord.EndDictionary(DataStream);
			TDictionaryRecord.EndDictionary(DataStream);
			TDictionaryRecord.EndDictionary(DataStream);

			TStreamRecord.BeginSave(DataStream);
			DataStream.Write(PatternDef);
			TStreamRecord.EndSave(DataStream);

			TIndirectRecord.SaveTrailer(DataStream);

			ImageDef.WriteImageObject(DataStream, XRef);

		}

		#region IComparable Members
		public int CompareTo(object obj)
		{
			TPdfImageTexture p2= obj as TPdfImageTexture;
			if (p2==null)
				return obj.GetType().GUID.CompareTo(this.GetType().GUID);

			for (int i=0; i < PatternMatrix.Length; i++)
			{
				int Result = PatternMatrix[i].CompareTo(p2.PatternMatrix[i]);
				if (Result != 0) return Result;
			}
			
			return ImageDef.CompareTo(p2.ImageDef);
		}

		#endregion

		#region GetPatternDef
		private byte[] GetPatternDef()
		{
			string s = 	String.Format(CultureInfo.InvariantCulture, "{0} {1} {2} {3} {4} {5} cm", 
				ImgWidth, 
				0, 0,
				ImgHeight,
				0, 
				0);

			s+=" " +TPdfTokens.GetString(TPdfToken.ImgPrefix) + "0 " + TPdfTokens.GetString(TPdfToken.CommandDo);

			return Encoding.UTF8.GetBytes(s);
		}

		#endregion
	}

}
