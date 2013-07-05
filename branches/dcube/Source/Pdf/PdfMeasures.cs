using System;
using System.Text;
using System.Collections.Generic;

namespace FlexCel.Pdf
{
	/// <summary>
	/// Holds a text string and the kerning with the latest block written.
	/// </summary>
	internal struct TKernedString
	{
		internal int Kern;
		internal string Text;

		internal TKernedString(int aKern, string aText)
		{
			Text = aText;
			Kern = aKern;
		}
	}

	internal sealed class FontMeasures
	{
		private FontMeasures(){}

		public static UInt32 MakeHash(long l, long r)
		{
			return (UInt32)((l<<16)+r);
		}

		public static float GlyphWidth(int gl, int[] GlyphWidths)
		{
			if (gl>= GlyphWidths.Length) gl = GlyphWidths.Length-1;
			if (gl< 0) gl = 0;
			return GlyphWidths[gl];
		}

		#region MeasureString
		public static float MeasureString(int[]b, int[] GlyphWidths, TKerningTable Kern, float UnitsPerEm, bool[] Ignore)
		{
			float Result =0;
			if (b.Length<=0)
				return Result;

			int si1 = b[0];
			if (!Ignore[0]) Result += GlyphWidth(si1, GlyphWidths)*1000/ UnitsPerEm;                

			for (int i=1;i<b.Length;i++)
			{
				int si = b[i];
				float k = 0;
				UInt32 key = MakeHash(si1,si);
				if (Kern!=null && Kern.ContainsKey(key))
					k = Kern[key]*1000/ UnitsPerEm;
				if (!Ignore[i]) Result += GlyphWidth(si, GlyphWidths)*1000/ UnitsPerEm + k;                
				si1 = si;
			}

			return Result;
		}

		#endregion

		#region KernString
		public static TKernedString[] KernString(string Text, int[]b, TKerningTable Kern, float UnitsPerEm)
		{
			if (b.Length<=0)
				return new TKernedString[0];

			List<TKernedString> Result = new List<TKernedString>();

			int LastPos =0;
			int pos =1;
			int LastK=0;
			while (pos<b.Length)
			{
				UInt32 key = MakeHash(b[pos],b[pos-1]);
				if (Kern!=null && Kern.ContainsKey(key))
				{
					Result.Add(new TKernedString(LastK, Text.Substring(LastPos, pos-LastPos)));
					LastK = (int)Math.Round(Kern[key]*1000/ UnitsPerEm);
					LastPos = pos;
				}
				pos++;		
			}

			if (LastPos<b.Length)
				Result.Add(new TKernedString(LastK, Text.Substring(LastPos, b.Length-LastPos)));

			return Result.ToArray();
		}
		#endregion
	}
}
