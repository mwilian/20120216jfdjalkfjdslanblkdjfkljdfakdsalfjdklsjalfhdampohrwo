using System;
using System.Text;
using System.Globalization;
using FlexCel.Core;

namespace FlexCel.XlsAdapter
{
	/// <summary>
	/// An Excel unicode string.
	/// </summary>
	/// 

	internal enum TStrLenLength {is8bits=1, is16bits=2}
	internal class TExcelString
	{
		private TStrLenLength StrLenLength;
		internal byte OptionFlags;
		internal string Data;
		internal byte[] RichTextFormats=null;
		internal byte[] FarEastData=null;
		internal int Hash;

		private string GetHashString()
		{
			/*string s="";
			if (Data==null) s="0"; else s=Data.Length.ToString(CultureInfo.InvariantCulture)+Data;
			if (RichTextFormats==null) s+="0"; else s+=RichTextFormats.Length.ToString(CultureInfo.InvariantCulture); //+RichTextFormats.ToString();  not really needed. ToString don't work here, and anyway, this is a hash
			if (FarEastData==null) s+="0"; else s+=FarEastData.Length.ToString(CultureInfo.InvariantCulture);//+FarEastData.ToString();
			return s;
            */
            return Data;

		}

		internal TExcelString(TStrLenLength aStrLenLength, ref TxBaseRecord aRecord, ref int Ofs)
		{
			StrLenLength=aStrLenLength;
			byte[]tmpLen= new byte[(byte)StrLenLength];
			int StrLen=0;
			BitOps.ReadMem(ref aRecord, ref Ofs, tmpLen);
			if (StrLenLength==TStrLenLength.is8bits) StrLen=tmpLen[0]; else StrLen=BitConverter.ToUInt16(tmpLen,0);

			byte[] of1=new byte[1];
			BitOps.ReadMem(ref aRecord, ref Ofs, of1);
			OptionFlags=of1[0];

			if (HasRichText) 
			{
				byte [] NumberRichTextFormatsArray=new byte[2];
				BitOps.ReadMem(ref aRecord, ref Ofs, NumberRichTextFormatsArray);
				RichTextFormats= new byte[4* BitConverter.ToUInt16(NumberRichTextFormatsArray,0)];
			}
			else RichTextFormats=null;

			if (HasFarInfo) 
			{
				byte [] FarEastDataSizeArray=new byte[4];
				BitOps.ReadMem(ref aRecord, ref Ofs, FarEastDataSizeArray);
				FarEastData=new byte[BitConverter.ToUInt32(FarEastDataSizeArray,0)];
			}
			else FarEastData=null;

			StringBuilder s=new StringBuilder(StrLen);
			StrOps.ReadStr(ref aRecord, ref Ofs, s, OptionFlags, ref OptionFlags, StrLen);
			Data=s.ToString();

			if (RichTextFormats!=null)
			{
				BitOps.ReadMem(ref aRecord, ref Ofs, RichTextFormats);
			}

			if (FarEastData!=null)
			{
				BitOps.ReadMem(ref aRecord, ref Ofs, FarEastData);
			}

			//We have to include all data on the hash.
			Hash=GetHashString().GetHashCode();


		}

		internal TExcelString(TStrLenLength aStrLenLength, string s, TRTFRun[] RTFRuns, bool ForceWide)
		{
			StrLenLength=aStrLenLength;
			if (StrLenLength== TStrLenLength.is8bits)
			{
				if (s.Length> 0xFF) XlsMessages.ThrowException(XlsErr.ErrInvalidStringRecord);
			}

			OptionFlags=0;
			if (ForceWide || StrOps.IsWide(s)) OptionFlags=1;

			if ((RTFRuns!=null) &&(RTFRuns.Length>0))
			{
				OptionFlags= (byte)(OptionFlags | 8);
				RichTextFormats= TRTFRun.ToByteArray(RTFRuns);
			}
			else
				RichTextFormats=null;
		
			FarEastData=null;

			Data=s;

			//We have to include all data on the hash.
			Hash=GetHashString().GetHashCode();

		}


		public override int GetHashCode() 
		{
			return Hash;
		}

		public override bool Equals(object o)
		{
            TExcelString x= o as TExcelString;
            if (x==null) return false;
			return (x.OptionFlags==OptionFlags) && (x.Data==Data)&&(BitOps.CompareMem(x.RichTextFormats, RichTextFormats)) && (BitOps.CompareMem(x.FarEastData, x.FarEastData));
		}

		internal bool HasFarInfo
		{
			get
			{
				return (OptionFlags & 0x4) == 0x4;
			}
		}

		internal bool HasRichText
		{
			get
			{
				return (OptionFlags & 0x8) == 0x8;
			}
		}

		internal int CharSize
		{
			get
			{
				if ((OptionFlags & 0x1) == 0x0) return 1; else return 2;
			}
		}

		internal int TotalSize()
		{
			int Result=
				(int)StrLenLength+
				1 + //SizeOf(OptionFlags)+
				Data.Length* CharSize;

			//Rich text
			if (HasRichText)
				Result+= 2+ RichTextFormats.Length; // SizeOf(NumberRichTextFormats)+ 4* NumberRichTextFormats;

			//FarEast
			if (HasFarInfo) 
				Result+=4 + FarEastData.Length; // SizeOf(FarEastDataSize) + FarEastDataSize;

			return Result;
		}

        internal void CopyToPtr(byte [] pData, int ofs)
        {
            CopyToPtr(pData, ofs, true);
        }
  
        internal void CopyToPtr(byte [] pData, int ofs, bool IncludeLen)
        {
            if (IncludeLen)
            {
                switch (StrLenLength)
                {
                    case TStrLenLength.is8bits:
                        pData[ofs]=(byte)Data.Length;
                        ofs++;
                        break;
                    case TStrLenLength.is16bits: 
                        BitConverter.GetBytes((UInt16)Data.Length).CopyTo(pData, ofs);
                        ofs+=2;
                        break;
                }
            }
			pData[ofs]=OptionFlags;
			ofs++;

			if (HasRichText)
			{
				BitConverter.GetBytes((UInt16)(RichTextFormats.Length/4)).CopyTo(pData, ofs);
				ofs+= 2;
			}

			if (HasFarInfo)
			{
				BitConverter.GetBytes((UInt32)FarEastData.Length).CopyTo(pData, ofs);
				ofs+= 4;
			}

			if (Data.Length>0)
			{
				if (CharSize== 1)
				{
					if (!StrOps.CompressUnicode(Data, pData, ofs)) XlsMessages.ThrowException(XlsErr.ErrInvalidStringRecord);
					ofs+=Data.Length;
				}
				else
				{
					Encoding.Unicode.GetBytes(Data, 0, Data.Length, pData, ofs);
					ofs+=Data.Length*CharSize;

				}
			}

			if (HasRichText)
			{
				RichTextFormats.CopyTo(pData, ofs);
				ofs+= RichTextFormats.Length;
			}

			if (HasFarInfo)
			{
				FarEastData.CopyTo(pData, ofs);
				ofs+= FarEastData.Length;
			}
		}


	}
}
