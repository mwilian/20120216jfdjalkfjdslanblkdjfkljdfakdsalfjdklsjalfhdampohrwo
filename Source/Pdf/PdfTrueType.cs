using System;
using System.IO;
using System.Text;
using System.Collections.Generic;
using System.Globalization;
using FlexCel.Core;
using System.Diagnostics;

namespace FlexCel.Pdf
{
	#region Utility Classes
	internal struct TTrueTypeInfo
	{
		public string FamilyName;
		public int FontFlags;

		public TTrueTypeInfo(string aFamilyName, int aFontFlags)
		{
			FamilyName = aFamilyName;
			FontFlags = aFontFlags;
		}
	}

	internal class TCharAndGlyph: IComparable
	{
		internal int Character;
		internal int Glyph;

		internal TCharAndGlyph(int aCharacter, int aGlyph)
		{
			Character = aCharacter;
			Glyph = aGlyph;
		}
		#region IComparable Members

		public int CompareTo(object obj)
		{
			TCharAndGlyph o2 = obj as TCharAndGlyph;
			if (o2 == null) return -1;
			return Character.CompareTo(o2.Character);
		}

		#endregion
	}

	internal struct TTableInfo
	{
		internal uint HeadOffset;
		internal uint DataOffset;

		internal TTableInfo(uint aHeadOffset, uint aDataOffset)
		{
			HeadOffset = aHeadOffset;
			DataOffset = aDataOffset;
		}
	}

	internal class THeadTable: IComparable
	{
		internal UInt32 OrigHeadPos;
		internal UInt32 CheckSum;
		internal UInt32 Offset;
		internal UInt32 DataLength;

		internal THeadTable(UInt32 aOrigHeadPos, UInt32 aCheckSum, UInt32 aOffset, UInt32 aDataLength)
		{
			OrigHeadPos = aOrigHeadPos;
			CheckSum = aCheckSum;
			Offset = aOffset;
			DataLength = aDataLength;
		}
		#region IComparable Members

		public int CompareTo(object obj)
		{
			THeadTable o2 = obj as THeadTable;
			if (o2 == null) return -1;
			return OrigHeadPos.CompareTo(o2.OrigHeadPos);
		}

		#endregion
	}


#if(FRAMEWORK20)
    internal sealed class TTableList : Dictionary<string, TTableInfo>
    {
        public TTableList(): base()
        {
        }
    }

    internal sealed class THeadTableList : List<THeadTable>
    {
        public THeadTableList(): base()
        {
        }
    }

	internal sealed class TCMap : Dictionary<int, int>
    {
    }

	internal sealed class TKerningTable : Dictionary<UInt32, int>
    {
    }

#else
	internal sealed class TTableList: Hashtable
	{
		internal bool TryGetValue(string key, out TTableInfo Result)
		{
			object obj = base[key];
			if (obj != null) Result = (TTableInfo) obj; else Result = new TTableInfo();
			return (obj != null);
		}

		public TTableInfo this[string key]
		{
			get
			{
				return (TTableInfo) base[key];
			}
			set
			{
				base[key] = value;
			}
		}


	}

	internal sealed class THeadTableList: ArrayList
	{
		public new THeadTable this[int index]
		{
			get
			{
				return (THeadTable) base[index];
			}
			set
			{
				base[index] = value;
			}
		}


	}

	internal sealed class TCMap: Hashtable
	{

		public int this[int key]
		{
			get
			{
				return (int)base[key];
			}
			set
			{
				base[key] = (int)value;
			}
		}

		public bool TryGetValue(int key, out int Glyph)
		{
			object o = base[key];
			if (o == null) {Glyph = 0; return false;}
			Glyph = (int)o;
			return true;
		}
	}

	internal sealed class TKerningTable : Hashtable
	{
		public int this[UInt32 key]
		{
			get
			{
				return (int) base[key];
			}
			set
			{
				base[key] = value;
			}
		}
	}

#endif

	internal struct PdfRectangle
	{
		internal int x1;
		internal int y1;
		internal int x2;
		internal int y2;

		internal PdfRectangle(int ax1, int ay1, int ax2, int ay2)
		{
			x1 = ax1;
			x2 = ax2;
			y1 = ay1;
			y2 = ay2;
		}

		internal string GetBBox(int UnitsPerEm)
		{
			return  
			    ((int) (x1*1000 / UnitsPerEm)).ToString(CultureInfo.InvariantCulture)+ " "+
				((int) (y1*1000 / UnitsPerEm)).ToString(CultureInfo.InvariantCulture)+ " "+
				((int) (x2*1000 / UnitsPerEm)).ToString(CultureInfo.InvariantCulture)+ " "+
				((int) (y2*1000 / UnitsPerEm)).ToString(CultureInfo.InvariantCulture);
		}
	}
	#endregion

	/// <summary>
	/// Encapsulates a True-Type font.
	/// </summary>
	internal class TPdfTrueType
	{
		#region Privates
		private byte[] FFontData; 
		private uint FontStart;
		private TTableList Tables;
		private string FPostcriptName;
		private string FFamilyName;
		private string FSubFamilyName;
		private string FUniqueFontName;
		private PdfRectangle FBoundingBox;
		private int FUnitsPerEm;
		private int FLocFormat;
		private float FItalicAngle;
		private int FUnderlinePosition;
		private int FAscent;
		private int FDescent;
		private int FLineGap;
		private int FCapHeight;
		private int FFontFlags;

		private int FNumberOfHMetrics;
		private int[] GlyphWidths;

		private TKerningTable Kern;

		private TCMap CMap10;
		private TCMap CMap30;
		private TCMap CMap31;

		private TGlyphMap MissingChars;

		private bool FKerning;
		#endregion

		#region Constructor
		private TPdfTrueType(byte[] aFontData)
		{
			MissingChars = new TGlyphMap();
			FFontData=aFontData;
		}

		public TPdfTrueType(byte[] aFontData, string FontName, bool UseKerning) : this(aFontData)
		{
			FKerning = UseKerning;
			GetFontFromCollection(FontName);
		}
		#endregion

		#region Properties
		public byte[] FontData {get{return FFontData;} }
		public PdfRectangle BoundingBox {get{return FBoundingBox;}}

		public string PostcriptName {get{return FPostcriptName;} }
		public string FamilyName {get{return FFamilyName;}}
		public string SubFamilyName {get{return FSubFamilyName;}}
		public string UniqueFontName {get{return FUniqueFontName;}}
		public int UnitsPerEm {get{return FUnitsPerEm;}}

		public float ItalicAngle {get{return FItalicAngle;}}
		public int UnderlinePosition {get{return FUnderlinePosition;}}
		public int Descent {get{return FDescent;}}
		public int Ascent {get{return FAscent;}}
		public int LineGap {get{return FLineGap;}}
		public int CapHeight {get{return FCapHeight;}}
		public int FontFlags {get{return FFontFlags;}}

		#endregion

		#region GetBytes
		private UInt16 GetMotorolaUInt16(long Start)
		{
			return (UInt16)((FFontData[Start]<<8)+FFontData[Start+1]);
		}
		
		private Int16 GetMotorolaInt16(long Start)
		{
			unchecked
			{
				return (Int16)((FFontData[Start]<<8)+FFontData[Start+1]);
			}
		}

		private UInt32 GetMotorolaUInt32(long Start)
		{
			unchecked
			{
				return (UInt32)(FFontData[Start]<<24)+(UInt32)(FFontData[Start+1]<<16)+
					(UInt32)(FFontData[Start+2]<<8)+FFontData[Start+3];
			}
		}

        private static UInt32 GetMotorolaUInt32(byte[] Data, long Start)
		{
			unchecked
			{
				return (UInt32)(Data[Start]<<24)+(UInt32)(Data[Start+1]<<16)+
					(UInt32)(Data[Start+2]<<8)+Data[Start+3];
			}
		}

		private Int32 GetMotorolaInt32(long Start)
		{
			unchecked
			{
				return (Int32)(FFontData[Start]<<24)+(Int32)(FFontData[Start+1]<<16)+
					(Int32)(FFontData[Start+2]<<8)+(Int32)FFontData[Start+3];
			}
		}

		private float GetMotorolaFixed(long Start)
		{
			unchecked
			{
				float x= GetMotorolaInt32(Start);
				return x/65536;
			}
		}

		private string GetUnicodeString(long Start, int Size)
		{
			return Encoding.BigEndianUnicode.GetString(FFontData, (int)Start, Size);
		}

		private string GetRomanString(long Start, int Size)
		{
			StringBuilder sb = new StringBuilder(Size);
			for (int i=(int)Start; i<Start+Size;i++)
				sb.Append((char)FFontData[i]);
			return sb.ToString();
		}

		private static void SetMotorolaUInt32(byte[] Data, int Start, long value)
		{
			unchecked
			{
				Data[Start]   = (byte)(value >> 24);
				Data[Start+1] = (byte)(value >> 16);
				Data[Start+2] = (byte)(value >> 8);
				Data[Start+3] = (byte)(value >> 0);
			}
		}

		private static void WriteMotorolaUInt32(MemoryStream Data, long value)
		{
			unchecked
			{
				Data.WriteByte((byte)(value >> 24));
				Data.WriteByte((byte)(value >> 16));
				Data.WriteByte((byte)(value >> 8));
				Data.WriteByte((byte)(value >> 0));
			}
		}

		private static void IncMotorolaUInt32(MemoryStream Data, UInt32 ofs)
		{
			UInt32 OrigValue;
			unchecked
			{
				OrigValue =	(UInt32)(Data.ReadByte()<<24);
				OrigValue += (UInt32)(Data.ReadByte()<<16);
				OrigValue += (UInt32)(Data.ReadByte()<<8);
				OrigValue += (UInt32)Data.ReadByte();
			}

			Data.Position = Data.Position - 4;
			WriteMotorolaUInt32(Data, OrigValue + ofs);
		}

		private static void WriteMotorolaUInt16(MemoryStream Data, long value)
		{
			unchecked
			{
				Data.WriteByte((byte)(value >> 8));
				Data.WriteByte((byte)(value >> 0));
			}
		}

		private static UInt16 ReadMotorolaUInt16(MemoryStream Data)
		{
			int b1 = Data.ReadByte();
			Debug.Assert(b1>=0, "At end of stream");
			int b2 = Data.ReadByte();
			Debug.Assert(b2>=0, "At end of stream");
			return (UInt16)((b1<<8)+b2);
		}

		#endregion

		#region Collection
		private readonly static byte[] Signature = new byte[]{(byte)'t',(byte)'t',(byte)'c',(byte)'f'};

		public static TTrueTypeInfo[] GetColection(byte[] aFontData)
		{
			TPdfTrueType Tmp = new TPdfTrueType(aFontData);
			return Tmp.GetCollection();
		}

		private TTrueTypeInfo[] GetCollection()
		{
			if (FFontData.Length < Signature.Length || !FlxUtils.CompareMem(Signature, FFontData, 0)) //font is not ttc
			{
				ReadNames();
                return new TTrueTypeInfo[] { new TTrueTypeInfo(FamilyName, FontFlags) };
			}

			UInt32 NumFonts = GetMotorolaUInt32(8);
			int p = 12;
			TTrueTypeInfo[] Result = new TTrueTypeInfo[NumFonts];
			for (UInt32 i = 0; i < NumFonts; i++)
			{
				FontStart = GetMotorolaUInt32(p);
				ReadNames();
				Result[i] = new TTrueTypeInfo(FamilyName, FontFlags);
				p+=4;
			}

			return Result;
		}


		private void GetFontFromCollection(string FontName)
		{
			if (FFontData.Length < Signature.Length || !FlxUtils.CompareMem(Signature, FFontData, 0)) //font is not ttc
			{
				LoadTables();
				ParseNames();
				ParseRest();
				return;
			}

			UInt32 NumFonts = GetMotorolaUInt32(8);
			int p = 12;
			for (UInt32 i = 0; i < NumFonts; i++)
			{
				FontStart = GetMotorolaUInt32(p);
				LoadTables();
				ParseNames();
				ParseHead();
				if (FFamilyName == FontName) 
				{
					ParseRest();
					return;
				}
				p+=4;
			}
		}
		#endregion

		#region Parse

		private void LoadTables()
		{
			FFontFlags=0;
			Tables = new TTableList();

			uint p = FontStart+12;
			for (int i=0;i<NumTables(FontStart);i++)
			{
				Tables[TableName(p)]=new TTableInfo(p, FontOffset(p));
				p+=16;
			}
		}

		private void ParseRest()
		{
			Kern = new TKerningTable();
			ParseHead();          
			ParseHHead();          
			ParseCMaps();
			ParsePost();
			ParseOS2();
			ParseHmtx();
			if (FKerning) ParseKerning();
		}

		private void ReadNames()
		{
			LoadTables();
			ParseNames();
			ParseHead();
		}

		private UInt16 NumTables(uint Start) {return GetMotorolaUInt16(Start+4);}

		private UInt32 Tag(uint Start) {return GetMotorolaUInt32(Start);}
		private UInt32 FontOffset(uint Start) {return GetMotorolaUInt32(Start+8);}
		private UInt32 FontLength(uint Start) {return GetMotorolaUInt32(Start+12);}
		
		private string TableName(uint Start) 
		{
			UInt32 t=Tag(Start); 
			return 
				((char)((t>>24)&0xFF)).ToString(CultureInfo.InvariantCulture)+
				((char)((t>>16)&0xFF)).ToString(CultureInfo.InvariantCulture)+
				((char)((t>>8)&0xFF)).ToString(CultureInfo.InvariantCulture)+
				((char)(t&0xFF)).ToString(CultureInfo.InvariantCulture);
		}

		private void ParseNames()
		{
			const int PostscriptFontName=6;
			const int FamilyFontName=1;
			const int SubFamilyFontName=2;
			const int UniqueFontNameId=3;
			const int EnglishLanguage = 0x409;

			long Start = Tables["name"].DataOffset;
			int Count = GetMotorolaUInt16(Start+2);
			int MainOffset = GetMotorolaUInt16(Start+4);

			FPostcriptName = String.Empty;
			FFamilyName = String.Empty;
			FSubFamilyName = String.Empty;
			FUniqueFontName = String.Empty;
			bool BestPs=false; bool BestFn=false; bool BestSfn=false;

			long Pos = Start+6;
			for (int i=0;i<Count;i++)
			{
				int Id = GetMotorolaUInt16(Pos+6);
				if (Id==PostscriptFontName || Id==FamilyFontName || Id==SubFamilyFontName || Id == UniqueFontNameId)
				{
					int PlatformID=GetMotorolaUInt16(Pos+0);
					int EncodingID=GetMotorolaUInt16(Pos+2);
					int LanguageID=GetMotorolaUInt16(Pos+4);

					int Length=GetMotorolaUInt16(Pos+8);
					int TableOffset=GetMotorolaUInt16(Pos+10);

					if (PlatformID == 3 && EncodingID == 1 && LanguageID == EnglishLanguage) //Windows
					{
						if (Id==PostscriptFontName) 
						{
							FPostcriptName = GetUnicodeString(Start + MainOffset + TableOffset, Length);
							BestPs = true;
						}
						else if (Id==FamilyFontName) 
						{
							FFamilyName = GetUnicodeString(Start + MainOffset + TableOffset, Length);
							BestFn = true;
						}
						else if (Id==SubFamilyFontName) 
						{
							FSubFamilyName = GetUnicodeString(Start + MainOffset + TableOffset, Length);
							BestSfn = true;
						}
						else if (Id==UniqueFontNameId) 
						{
							FUniqueFontName = GetUnicodeString(Start + MainOffset + TableOffset, Length);
							BestFn = true;
						}
					}

					if (PlatformID == 1 && EncodingID == 0 && LanguageID == 0) //Mac
					{
						if (!BestPs && Id==PostscriptFontName) FPostcriptName = GetRomanString(Start + MainOffset + TableOffset, Length);
						else if (!BestFn && Id==FamilyFontName) FFamilyName = GetRomanString(Start + MainOffset + TableOffset, Length);
						else if (!BestSfn && Id==SubFamilyFontName) FSubFamilyName = GetRomanString(Start + MainOffset + TableOffset, Length);
						else if (!BestSfn && Id==UniqueFontNameId) FUniqueFontName = GetRomanString(Start + MainOffset + TableOffset, Length);
					}


				}

				Pos+=12;
			}
		}

		private void ParseHead()
		{
			long Start = Tables["head"].DataOffset;

			FUnitsPerEm = GetMotorolaUInt16(Start+18);

			FBoundingBox.x1 = GetMotorolaInt16(Start+36);
			FBoundingBox.y1 = GetMotorolaInt16(Start+38);
			FBoundingBox.x2 = GetMotorolaInt16(Start+40);
			FBoundingBox.y2 = GetMotorolaInt16(Start+42);

			int MacStyle = GetMotorolaUInt16(Start+44);
			if ((MacStyle & 1) !=0) FFontFlags |= (1<<18); //bold
			if ((MacStyle & 2) !=0) FFontFlags |= (1<<6); //italic

			FLocFormat = GetMotorolaUInt16(Start + 50);
		}

		private void ParseHHead()
		{
			long Start = Tables["hhea"].DataOffset;

			FNumberOfHMetrics = GetMotorolaUInt16(Start+34);
			FAscent = GetMotorolaInt16(Start+4);
			FDescent = GetMotorolaInt16(Start+6);
			FLineGap = GetMotorolaInt16(Start+8);
		}

		private void ParseCMaps()
		{
			long Start = Tables["cmap"].DataOffset;
			int Count = GetMotorolaUInt16(Start+2);
			long CMap31Offset = -1; //windows
			long CMap10Offset = -1; //mac
			long CMap30Offset = -1; //symbol

			long Pos = Start+4;
			for (int i=0;i<Count;i++)
			{
				int PlatformID=GetMotorolaUInt16(Pos+0);
				int EncodingID=GetMotorolaUInt16(Pos+2);

				if (PlatformID == 3 && EncodingID == 1) //Windows
				{
					CMap31Offset = GetMotorolaUInt32(Pos+4);
				}

				if (PlatformID == 3 && EncodingID == 0) //Symbol
				{
					CMap30Offset = GetMotorolaUInt32(Pos+4);
				}

				if (PlatformID == 1 && EncodingID == 0) //Mac
				{
					CMap10Offset = GetMotorolaUInt32(Pos+4);
				}

				Pos+=8;
			}

			if (CMap31Offset>0)
				ReadFormatTable(ref CMap31, Start+CMap31Offset, 0xFFFF);

			if (CMap30Offset>0)
				ReadFormatTable(ref CMap30, Start+CMap30Offset, 0xFF);

			if (CMap10 == null && CMap10Offset>0)
				ReadFormatTable(ref CMap10, Start+CMap10Offset, 0xFF);

			if (CMap30Offset>0) FFontFlags |= (1<<2); //symbolic
			else
				FFontFlags |= 1<<5;  //NonSymbolic
		}

		private void ReadFormatTable(ref TCMap CMap, long CMapOffset, int Mask)
		{
			int Format = GetMotorolaInt16(CMapOffset);
			switch (Format)
			{
				case 0: CMap = new TCMap(); ReadFormat0(CMap, CMapOffset);break;
				case 4: CMap = new TCMap(); ReadFormat4(CMap, CMapOffset, Mask);break;
				case 6: CMap = new TCMap(); ReadFormat6(CMap, CMapOffset);break;
			}
		}

		private void ReadFormat0(TCMap CMap, long CMapOffset)
		{
			for (int i=0;i <256; i++)
				CMap[i] = FFontData[CMapOffset+6+i];
		}

		private void ReadFormat4(TCMap CMap, long CMapOffset, int Mask)
		{
			int SegCountx2 = GetMotorolaUInt16(CMapOffset + 6);

			for (int i=0;i <SegCountx2; i+=2)
			{
				int EndChar = GetMotorolaUInt16(CMapOffset + 14 + i);
				int StartChar = GetMotorolaUInt16(CMapOffset + SegCountx2 + 16 + i);
				int RangeOffsetPos = (int)CMapOffset + SegCountx2+ SegCountx2+ SegCountx2 + 16 + i;
				int RangeOffset = GetMotorolaUInt16(RangeOffsetPos);
				int Delta = GetMotorolaInt16(CMapOffset + SegCountx2+ SegCountx2 + 16 + i);
				if (RangeOffset != 0)
				{
					for (int k=StartChar; k<= EndChar; k++)
					{
						int gindexPos = RangeOffset +(k-StartChar)*2+ RangeOffsetPos;
						int gindex = gindexPos>=FFontData.Length? 0: (int)GetMotorolaUInt16(gindexPos);
						if (gindex>0)
						{
							int g = (gindex+Delta) & 0xFFFF;
							if (g>0) CMap[k & Mask] = g;
						}
					}
				}
				else
				{
					for (int k=StartChar; k<= EndChar; k++)
					{
						int g = (k+Delta) & 0xFFFF;
						if (g>0) CMap[k & Mask] = g;
					}
				}			
			}
		}

		private void ReadFormat6(TCMap CMap, long CMapOffset)
		{
			int FirstChar = GetMotorolaUInt16(6);
			int CharCount = GetMotorolaUInt16(8);
			for (int i=FirstChar;i <CharCount; i++)
				CMap[i] = FFontData[CMapOffset+6+i];
		}

		private void ParsePost()
		{
			long Start = Tables["post"].DataOffset;

			FItalicAngle = GetMotorolaFixed(Start+4);
			FUnderlinePosition = GetMotorolaInt16(Start+8);
			UInt32 IsFixed = GetMotorolaUInt32(Start+12);
			FFontFlags |= (byte)(IsFixed & 1);
		}

		private void ParseOS2()
		{
			long Start = Tables["OS/2"].DataOffset;
			int Version = GetMotorolaUInt16(Start);

			int FamilyClass = GetMotorolaInt16(Start+30);
			byte FamClass = (byte)(FamilyClass >>8);

			if (FamClass <8 && FamClass>0 ) FFontFlags |=(1<<1); //serif
			if (FamClass ==12) FFontFlags |=(1<<2); //symbolic
			if (FamClass ==10) FFontFlags |=(1<<3); //script

			int asc = GetMotorolaInt16(Start+68);
			if (asc>0) FAscent = asc;
			int desc = GetMotorolaInt16(Start+70);
			if (FDescent>0) FDescent = desc;

			int lg = GetMotorolaInt16(Start+72);
			if (lg>0) FLineGap = lg;
			if (Version>1)
				FCapHeight = GetMotorolaInt16(Start+88);
		}

		private void ParseHmtx()
		{
			long Start = Tables["hmtx"].DataOffset;

			GlyphWidths = new int[FNumberOfHMetrics];
			long Pos = Start;
			for (int i=0;i<FNumberOfHMetrics;i++)
			{
				GlyphWidths[i] = GetMotorolaInt16(Pos);
				Pos+=4;
			}
		}

		private void ParseKerning()
		{
			TTableInfo TblInfo;
			if (!Tables.TryGetValue("kern", out TblInfo)) return;
			long Start = TblInfo.DataOffset;

			int Count = GetMotorolaUInt16(Start+2);
			long Pos = Start+4;
			for (int i=0;i<Count;i++)
			{
				int Format = FFontData[Pos+4];
				int Coverage = FFontData[Pos+5];
				if (Format!=0 || (Coverage & 3)!=1)
				{
					int Len = GetMotorolaUInt16(Pos+2);
					Pos+=Len;
					continue;
				}

				ReadKern0(Pos+6);
			}
		}

		private void ReadKern0(long Start)
		{
			int Count = GetMotorolaUInt16(Start);
			long Pos = Start+8;
			for (int i=0;i<Count;i++)
			{
				Kern[FontMeasures.MakeHash(GetMotorolaUInt16(Pos), GetMotorolaUInt16(Pos+2))] = GetMotorolaInt16(Pos+4);
				Pos+=6;
			}
		}

		#endregion

		#region Public Methods
		public float GlyphWidth(int gl)
		{
			return FontMeasures.GlyphWidth(gl, GlyphWidths)*1000/ UnitsPerEm;
		}

		public bool HasGlyph(int c)
		{
			if (CMap31!=null)
			{
				int gl;
				if (!CMap31.TryGetValue(c, out gl)) return false;
				return true;
			}

			if (CMap30!=null)  //symbolic
			{
				int gl;
				if (!CMap30.TryGetValue(c, out gl)) return false;
				return true;
			}

			return false;				 
		}

		public int Glyph(int c, bool LogError)
		{
			if (CMap31!=null)
			{
				int gl;
				if (!CMap31.TryGetValue(c, out gl)) return GlypNotFound(c, LogError);
				return gl;
			}

			if (CMap30!=null)  //symbolic
			{
				int gl;
				if (!CMap30.TryGetValue(c, out gl)) return GlypNotFound(c, LogError);
				return gl;
			}

			return GlypNotFound(c, LogError);				 
		}

		private int GlypNotFound(int c, bool LogError)
		{
			if (LogError && FlexCelTrace.HasListeners && c > 31 && !MissingChars.ContainsKey(c)) 
			{
				MissingChars.Add(c,c);
				FlexCelTrace.Write(new TPdfGlyphNotInFontError(FlxMessages.GetString(FlxErr.ErrGlyphNotFound, FamilyName, c, (char)c), FamilyName, c));
			}
			return 0;
		}

		public float MeasureString(int[] b, bool[] ignore)
		{
			return FontMeasures.MeasureString(b, GlyphWidths, Kern, UnitsPerEm, ignore);
		}

		public TKernedString[] KernString(string s, int[] b)
		{
			return FontMeasures.KernString(s, b, Kern, UnitsPerEm);
		}

		public byte[] SubsetFontData(TGlyphMap UsedGlypMap)
		{
			if (UsedGlypMap == null) return FontData;

			AddCompositeGlyphs(UsedGlypMap);
			int[] GlyphList = UsedGlypMap.ToList();
			return GetSubset(GlyphList, UsedGlypMap);
		}

		public bool NeedsEmbed(TFontEmbed aEmbed)
		{
			return aEmbed == TFontEmbed.OnlySymbolFonts && CMap30 != null;
		}

		#endregion

		#region Font Subset

		// Optimal order for table bodies is (from font validator help): head, hhea, maxp, OS/2, hmtx, LTSH, VDMX, hdmx, cmap, fpgm, prep, cvt, loca, glyf, kern, name, post, gasp, PCLT, DSIG
		// Table headers should be ordered in alphabetical order.
		// Tables required by Acrobat are (from pdf reference, sec 5.8):
		//The following TrueType tables are always required: "head," "hhea," "loca," "maxp," "cvt ", "prep," "glyf," "hmtx," and "fpgm."

		private static readonly string[] TablesToCopy = 
			{"head", "hhea", "maxp", "hmtx", "cmap", "fpgm", 
			 "prep", "cvt ", "loca", "glyf"};  

		private byte[] GetSubset(int[] UsedGlyps, TGlyphMap UsedGlypMap)
		{
			using (MemoryStream ResultFont = new MemoryStream())
			{
				THeadTableList HeadList = new THeadTableList();
				int HeadPos = -1;
				for (int i=0;i<TablesToCopy.Length;i++)
				{
					TTableInfo TblInfo;

					if (!Tables.TryGetValue(TablesToCopy[i], out TblInfo)) continue;

					if (i == 0) HeadPos = (int)ResultFont.Position + 8; //checkSumAdjustment
					ProcessTable(UsedGlyps, UsedGlypMap, TablesToCopy[i], TblInfo.DataOffset, FontLength(TblInfo.HeadOffset), TblInfo.HeadOffset, ResultFont, HeadList);
				}

				using (MemoryStream HeaderData = GetHeadStream(HeadList))
				{
					FixHeadCheckSum(HeaderData, ResultFont, HeadPos);

					byte[] Result = new byte[HeaderData.Length + ResultFont.Length];
					HeaderData.Position = 0;
					if (HeaderData.Read(Result, 0, (int)HeaderData.Length) != HeaderData.Length) FlxMessages.ThrowException(FlxErr.ErrInternal);

					ResultFont.Position = 0;
					if (ResultFont.Read(Result, (int)HeaderData.Length, (int)ResultFont.Length) != ResultFont.Length) FlxMessages.ThrowException(FlxErr.ErrInternal);

					return Result;
				}
			}	
		}

		private MemoryStream GetHeadStream(THeadTableList HeadList)
		{
			MemoryStream Result = new MemoryStream();

			Result.Write(FontData, (int)FontStart, 4); //sfnt
			WriteMotorolaUInt16(Result, (UInt16) HeadList.Count);
			int Log2 = (int)Math.Floor(Math.Log(HeadList.Count, 2));
			int MaxPow2 = (int)Math.Pow(2, Log2);
			int SearchRange = MaxPow2 << 4;
			WriteMotorolaUInt16(Result, (UInt16) SearchRange);
			WriteMotorolaUInt16(Result, (UInt16) Log2);
			WriteMotorolaUInt16(Result, (UInt16) ((HeadList.Count << 4) - SearchRange));

			//write table.
			HeadList.Sort(); //Sort the header items as they were in the original font. We assume it was the correct order.
			UInt32 HeadOfs = 12 + (UInt32)HeadList.Count * 16;

			for (int i = 0; i < HeadList.Count; i++)
			{
				Result.Write(FontData, (int)(HeadList[i].OrigHeadPos), 4);
				WriteMotorolaUInt32(Result, HeadList[i].CheckSum); //checksum
				WriteMotorolaUInt32(Result, HeadList[i].Offset + HeadOfs); //offset. doesn't include headerdata full length because we don't know it yet.
				WriteMotorolaUInt32(Result, HeadList[i].DataLength); //length not padded.
			}

			return Result;

		}

		private void FixHeadCheckSum(MemoryStream HeaderData, MemoryStream ResultFont, int HeadPos)
		{
			if (HeadPos < 0) PdfMessages.ThrowException(PdfErr.ErrInvalidFont, FFamilyName); //table should always have head.
			UInt32 ChkSum = GetCheckSum(HeaderData, 0, 0);
			unchecked
			{
				ChkSum = 0xB1B0AFBA - GetCheckSum(ResultFont, ChkSum, 0);
			}

			ResultFont.Position = HeadPos;
			WriteMotorolaUInt32(ResultFont, ChkSum);
		}

		private void ProcessTable(int[] UsedGlyps, TGlyphMap UsedGlypMap, string TableName, long ofs, long len, uint p,  MemoryStream ResultFont, THeadTableList HeadList)
		{
			long StartTable = ResultFont.Position;
			if (!TrimTable(UsedGlyps, UsedGlypMap, TableName, ofs, len, ResultFont)) return;
			UpdateHeader(HeadList, StartTable, p, ResultFont);
		}

		private void UpdateHeader(THeadTableList HeadList, long StartTable, uint p, MemoryStream ResultFont)
		{
			//Pad the table to 4-byte multiple. Note that Len is the non-padded lenght.
			int Len = (int)(ResultFont.Position - StartTable);
			for (int i = 0; i < 4 - (Len % 4); i++)
			{
				ResultFont.WriteByte(0);
			}

			HeadList.Add(new THeadTable(p, GetCheckSum(ResultFont, 0, StartTable), (UInt32)StartTable, (UInt32)Len));
		}

		private UInt32 GetCheckSum(MemoryStream ResultFont, UInt32 Sum, long StartTable)
		{
			unchecked
			{
				byte[] OneLong = new byte[4];
				ResultFont.Position = StartTable;

				while (ResultFont.Read(OneLong, 0, 4) == 4)
				{
					Sum += GetMotorolaUInt32(OneLong, 0);
				}
				return Sum;
			}
		}

		private bool TrimTable(int[] UsedGlyps, TGlyphMap UsedGlypMap, string TableName, long ofs, long len, MemoryStream ResultFont)
		{
			switch (TableName)
			{
				case "cmap": if (CMap30 == null && CMap31 == null) return false; TrimCMap(UsedGlyps, UsedGlypMap, ofs, len, ResultFont);return true;
				case "head": TrimHead(UsedGlyps, ofs, len, ResultFont);return true;
				case "hhea": TrimHhea(UsedGlyps, ofs, len, ResultFont);return true;
				case "hmtx": TrimHmtx(UsedGlyps, ofs, len, ResultFont);return true;

				case "loca": TrimLoca(UsedGlyps, ofs, len, ResultFont);return true;
				case "glyf": TrimGlyf(UsedGlyps, UsedGlypMap, ofs, len, ResultFont);return true;
	
				case "maxp": TrimMaxp(UsedGlyps, ofs, len, ResultFont);return true;

				case "fpgm":
				case "cvt ":
				case "prep":

				default:
					ResultFont.Write(FontData, (int)ofs, (int)len); return true;
				
			}
		}

		private void TrimHead(int[] UsedGlyps, long ofs, long len, MemoryStream ResultFont)
		{
			ResultFont.Write(FontData, (int)ofs, 8);
			ResultFont.Write(new byte[4], 0, 4); //clean checksum
			int flen = 38;
			ofs += 8+4;
			if (flen + 2 > len - 8 - 4)PdfMessages.ThrowException(PdfErr.ErrInvalidFont, FFamilyName);
			ResultFont.Write(FontData, (int)ofs, flen);

			ofs += flen + 2;
			WriteMotorolaUInt16(ResultFont, 1);//change  indexToLocFormat to 1.
			ResultFont.Write(FontData, (int)ofs, (int)(len - 8 - 4 - flen - 2));
		}

		private void TrimCMap(int[] UsedGlyps, TGlyphMap UsedGlypMap, long ofs, long len, MemoryStream ResultFont)
		{
			//we only need the cmap for 8bit characters. Unicode fonts use directly the glyph id.
			WriteMotorolaUInt16(ResultFont, 0); //version
			int MapCount = 1;
			WriteMotorolaUInt16(ResultFont, MapCount);
			
			WriteMotorolaUInt16(ResultFont, 3);
			if (CMap30 != null) WriteMotorolaUInt16(ResultFont, 0); else WriteMotorolaUInt16(ResultFont, 1);
			WriteMotorolaUInt32(ResultFont, 12); //offset

			long StartTable = ResultFont.Position;
			WriteMotorolaUInt16(ResultFont, 4); //table format for microsoft encodings
			long LenPos = ResultFont.Position;
			WriteMotorolaUInt16(ResultFont, 0); //length, will be updated later

			WriteMotorolaUInt16(ResultFont, 0); //language

			//Segments
			WriteMotorolaUInt16(ResultFont, 4); //segments * 2
			WriteMotorolaUInt16(ResultFont, 4); //2* the largest power of 2 that is less than or equal to segCount
			WriteMotorolaUInt16(ResultFont, 1); //log2(2)
			WriteMotorolaUInt16(ResultFont, 0); 

#if (FRAMEWORK20)
            List<TCharAndGlyph> AsciiChars = new List<TCharAndGlyph>();
#else
			ArrayList AsciiChars = new ArrayList();
#endif
			TCMap CMap = CMap30!= null? CMap30: CMap31;
			int CMapOfs = CMap30 != null? 0xF000 : 0;  //see http://partners.adobe.com/public/developer/opentype/index_recs.html
			foreach (int chuni in CMap.Keys)
			{
				if (CharUtils.IsWin1252(chuni))
				{
					int GlId = CMap[chuni];
					int newgl;
					if (UsedGlypMap.TryGetValue(GlId, out newgl))
					{
						AsciiChars.Add(new TCharAndGlyph(chuni + CMapOfs, newgl));
					}
				}
			}

			AsciiChars.Sort();
            int FirstChar = 0;
			int LastChar = 0;
			if (AsciiChars.Count > 0)
			{
				FirstChar = ((TCharAndGlyph)AsciiChars[0]).Character;
				LastChar = ((TCharAndGlyph)AsciiChars[AsciiChars.Count - 1]).Character;
			}

			//End chars
			WriteMotorolaUInt16(ResultFont, LastChar); 
			WriteMotorolaUInt16(ResultFont, 0xFFFF); 

			//pad
			WriteMotorolaUInt16(ResultFont, 0); 

			//Start chars
			WriteMotorolaUInt16(ResultFont, FirstChar); 
			WriteMotorolaUInt16(ResultFont, 0xFFFF); 

			WriteMotorolaUInt16(ResultFont, 0);  //delta
			WriteMotorolaUInt16(ResultFont, 1);  //delta, will map to char 0.


			//Idrangeoffset
			WriteMotorolaUInt16(ResultFont, 4); //This is the offset from this position where the map is.  As we only have one uint16 between this and the map, and other for the rangeoffset itself, this is 4.
			WriteMotorolaUInt16(ResultFont, 0); 

			//map
			if (AsciiChars.Count == 0) WriteMotorolaUInt16(ResultFont, 0);
			for (int i = 0; i < AsciiChars.Count; i++)
			{
				if (i > 0)
				{
					for (int k = ((TCharAndGlyph)AsciiChars[i - 1]).Character + 1; k < ((TCharAndGlyph)AsciiChars[i]).Character; k++)
					{
						WriteMotorolaUInt16(ResultFont, 0); 
					}
				}
				WriteMotorolaUInt16(ResultFont, ((TCharAndGlyph)AsciiChars[i]).Glyph); 
			}

			//write the length
			ResultFont.Position = LenPos;
			WriteMotorolaUInt16(ResultFont, ResultFont.Length - StartTable);
			ResultFont.Position = ResultFont.Length;
		}
		
		private void TrimLoca(int[] UsedGlyps, long ofs, long len, MemoryStream ResultFont)
		{
			int EntrySize = FLocFormat == 1? 4: 2;

			UInt32 CurrentOffset = 0;
			WriteMotorolaUInt32(ResultFont, CurrentOffset);  //first glyph is always at offset 0.
			for (int i = 0; i < UsedGlyps.Length; i++)
			{
				int oldPos = UsedGlyps[i];
				long OldOffset = ofs + oldPos * EntrySize;
				UInt32 glylen;
				if (FLocFormat == 1)
				{
					glylen = GetMotorolaUInt32(OldOffset + EntrySize) - GetMotorolaUInt32(OldOffset) ;
				}
				else
				{
					glylen = (UInt32) (GetMotorolaUInt16(OldOffset + EntrySize) - GetMotorolaUInt16(OldOffset)) * 2;
				}
				
				CurrentOffset += glylen;
				WriteMotorolaUInt32(ResultFont, CurrentOffset);  
			}
		}

		private void TrimGlyf(int[] UsedGlyps, TGlyphMap UsedGlypMap, long ofs, long len, MemoryStream ResultFont)
		{
			TTableInfo loca;
			if (!Tables.TryGetValue("loca", out loca)) PdfMessages.ThrowException(PdfErr.ErrInvalidFont, FFamilyName);

			int EntrySize = FLocFormat == 1? 4: 2;

			for (int i = 0; i < UsedGlyps.Length; i++)
			{
				int oldPos = UsedGlyps[i];
				long OldOffset = loca.DataOffset + oldPos * EntrySize;
				int glylen;
				int glystart;
				if (FLocFormat == 1)
				{
					glystart = (int) GetMotorolaUInt32(OldOffset);
					glylen = (int) GetMotorolaUInt32(OldOffset + EntrySize) - glystart ;
				}
				else
				{
					glystart = GetMotorolaUInt16(OldOffset) * 2;
					glylen = GetMotorolaUInt16(OldOffset + EntrySize) * 2 - glystart;
				}
				
				long BeforePosition = ResultFont.Position;
				ResultFont.Write(FFontData, (int)(ofs + glystart), glylen);
				long AfterPosition = ResultFont.Position;

				if (IsComposite(ofs + glystart, glylen))
				{
					ResultFont.Position = BeforePosition;
					RemapComposite(ResultFont, UsedGlypMap);
					ResultFont.Position = AfterPosition;
				}
			}
		}

		private void TrimMaxp(int[] UsedGlyps, long ofs, long len, MemoryStream ResultFont)
		{
			ResultFont.Write(FFontData, (int)ofs, 4);
			WriteMotorolaUInt16(ResultFont, UsedGlyps.Length);
			if (len - 6 > 0) ResultFont.Write(FFontData, (int)ofs + 6, (int)(len - 6));
		}


		private int NewNumberOfHMetrics(int[] UsedGlyps)
		{
			if (UsedGlyps.Length == 0) return 0;

			long ofs = Tables["hmtx"].DataOffset;
			int LastWidth = -1;
			for (int i = UsedGlyps.Length - 1; i >= 0; i--)
			{
				int metric = UsedGlyps[i];
				if (metric > FNumberOfHMetrics - 1) metric = FNumberOfHMetrics - 1;

				UInt16 w = GetMotorolaUInt16(ofs + metric * 4);
				if (LastWidth < 0) LastWidth = w;
				if (w != LastWidth) 
				{
					return i + 2;
				}
			}
            return 1;
		}

		private void TrimHhea(int[] UsedGlyps, long ofs, long len, MemoryStream ResultFont)
		{
			ResultFont.Write(FFontData, (int)ofs, 34);
		    WriteMotorolaUInt16(ResultFont, NewNumberOfHMetrics(UsedGlyps));
			ResultFont.Write(FFontData, (int)(ofs + 34 + 2), (int)(len - 34 - 2));
		}

		private void TrimHmtx(int[] UsedGlyps, long ofs, long len, MemoryStream ResultFont)
		{
			int n = NewNumberOfHMetrics(UsedGlyps);
			for (int i = 0; i < n; i++)
			{
				int metric = UsedGlyps[i];
				if (metric > FNumberOfHMetrics - 1) metric = FNumberOfHMetrics - 1;
				ResultFont.Write(FFontData, (int)(ofs + metric * 4), 4);
			}

			//leftsidebearing

			for (int i = n; i< UsedGlyps.Length; i++)
			{
				int metric = UsedGlyps[i];
				Int16 lsb = 0;
				if (metric < FNumberOfHMetrics) 
				{
					lsb = GetMotorolaInt16(ofs + metric * 4 + 2);
				}
				else
				{
					int lsbofs = FNumberOfHMetrics *4 + (metric - FNumberOfHMetrics) * 2;
					if (lsbofs < len)
					{
						lsb = GetMotorolaInt16(ofs + lsbofs);
					}

				}

				unchecked
				{
					WriteMotorolaUInt16(ResultFont, (UInt16)lsb);
				}
			}
		}

		#endregion

		#region Composite Glyphs

		private void AddCompositeGlyphs(TGlyphMap UsedGlypMap)
		{
			long StartLoca = Tables["loca"].DataOffset;
			long StartGlyf = Tables["glyf"].DataOffset;
			int EntrySize = FLocFormat == 1? 4: 2;

			UInt32List GlypList = new UInt32List(UsedGlypMap.Count + 20); //We can't store the new items in GlypMap and then iterate in GlypMap, since it is a hashtable. So we will use a parallel list.

			foreach (int oldGlyph in UsedGlypMap.Keys)
			{
				GlypList.Add((UInt32) oldGlyph);
			}

			int i = 0;
			while (i < GlypList.Count)  //note that GlypList.Count might grow inside the loop.
			{
				long OldOffset = StartLoca + GlypList[i] *EntrySize;
				long glystart;
				if (FLocFormat == 1)
				{
					glystart = (int) GetMotorolaUInt32(OldOffset);
				}
				else
				{
					glystart = GetMotorolaUInt16(OldOffset) * 2;
				}
				ProcessGlyph(StartGlyf + glystart, GlypList, UsedGlypMap);
				i++;
			}
		}

		private bool IsComposite(long Start, int GlyLen)
		{
			if (GlyLen <= 0) return false;
			int numberOfCountours = GetMotorolaInt16(Start);
			return numberOfCountours == -1;
		}

		private void ProcessGlyph(long Start, UInt32List GlypList, TGlyphMap UsedGlypMap)
		{
			const int ARG_1_AND_2_ARE_WORDS = 1;
			const int WE_HAVE_A_SCALE = 8;
			const int MORE_COMPONENTS = 32;
			const int WE_HAVE_AN_X_AND_Y_SCALE = 64;
			const int WE_HAVE_A_TWO_BY_TWO = 128;


			int numberOfCountours = GetMotorolaInt16(Start);
			if (numberOfCountours >= 0) return;

			long StartComposite = Start + 10;
			UInt16 flags = 0;
			do 
			{
				flags = GetMotorolaUInt16(StartComposite);
				UInt16 glyphIndex = GetMotorolaUInt16(StartComposite + 2);

				if (!UsedGlypMap.ContainsKey((int)glyphIndex)) 
				{
					UsedGlypMap.Add((int)glyphIndex, UsedGlypMap.Count);
					GlypList.Add(glyphIndex);
				}

				StartComposite += 4;

				if ((flags & ARG_1_AND_2_ARE_WORDS) != 0) StartComposite += 4; else StartComposite += 2;

				if ((flags & WE_HAVE_A_SCALE) != 0) 
				{
					StartComposite += 2;
				} 
				else if ((flags & WE_HAVE_AN_X_AND_Y_SCALE) != 0) 
				{
					StartComposite += 4;
				} 
				else if ((flags & WE_HAVE_A_TWO_BY_TWO) != 0) 
				{
					StartComposite += 8;
				}
			} 
			while ((flags & MORE_COMPONENTS) != 0);
		}


		private static void RemapComposite(MemoryStream ResultFont, TGlyphMap UsedGlyphMap)
		{
			const int ARG_1_AND_2_ARE_WORDS = 1;
			const int WE_HAVE_A_SCALE = 8;
			const int MORE_COMPONENTS = 32;
			const int WE_HAVE_AN_X_AND_Y_SCALE = 64;
			const int WE_HAVE_A_TWO_BY_TWO = 128;

			// this method must have been called when IsComposite = true. 

			ResultFont.Position += 10;
			UInt16 flags = 0;
			do 
			{
				flags = ReadMotorolaUInt16(ResultFont);
				UInt16 glyphIndex = ReadMotorolaUInt16(ResultFont);

				int newGlyph;
				if (!UsedGlyphMap.TryGetValue(glyphIndex, out newGlyph)) FlxMessages.ThrowException(FlxErr.ErrInternal);
				ResultFont.Position -= 2;
				WriteMotorolaUInt16(ResultFont, newGlyph);
                
				if ((flags & ARG_1_AND_2_ARE_WORDS) != 0) ResultFont.Position += 4; else ResultFont.Position += 2;

				if ((flags & WE_HAVE_A_SCALE) != 0) 
				{
					ResultFont.Position += 2;
				} 
				else if ((flags & WE_HAVE_AN_X_AND_Y_SCALE) != 0) 
				{
					ResultFont.Position += 4;
				} 
				else if ((flags & WE_HAVE_A_TWO_BY_TWO) != 0) 
				{
					ResultFont.Position += 8;
				}
			} 
			while ((flags & MORE_COMPONENTS) != 0);
		}
		#endregion

	}

}
