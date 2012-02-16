using System;
using System.IO;
using System.Runtime.CompilerServices;
using System.Reflection;
using FlexCel.Core;


namespace FlexCel.XlsAdapter
{
	/// <summary>
	/// Converts Image formats from standard WMF to internal XLS WMF representation.
	/// For this class to work, you need to install Visual J# runtime from http://msdn.microsoft.com/vjsharp/downloads/howtoget.asp
	/// </summary>
    internal sealed class XlsMetafiles
    {
        private XlsMetafiles(){}
        //Current implementation of TCompressor is not thread safe.
        //private static TCompressor Cmp=null; //Wont be created until needed. //STATIC*
            
        //[MethodImpl(MethodImplOptions.NoInlining)]   //We don't want to check for ZLib until we are here.
        internal static void ToXls(Byte[] ImgData, Stream OutStream, bool IsEMF)
        {
#if (FRAMEWORK30 || !COMPACTFRAMEWORK)
            int HeadOfs=0;
            if (!IsEMF && ImgData.Length>4 && BitConverter.ToUInt32(ImgData,0)==0x9AC6CDD7) HeadOfs+=22; //A wmf file might have a placeable header or not. If it does, it begins with the magic number 9AC6CDD7 and it must be stripped.  
            
			using (TCompressor Cmp= new TCompressor())
			{
				Cmp.Deflate(ImgData, HeadOfs, OutStream);
			}
#else
            throw new FlexCelException("Operation not supported in CF 2.0");
#endif
        }


        internal static UInt16 ComputeAldusCheckSum(byte[] Data)
        {
            UInt16 Result=0;
 
            for (int i=0;i<Data.Length-2;i+=2)  //-2 is to skip the own checksum from the checksum.
                Result ^= BitConverter.ToUInt16(Data,i);
            return Result;
        }

        //[MethodImpl(MethodImplOptions.NoInlining)] //We don't want to check for ZLib until we are here.
        internal static void ToWMF(byte[] XlsData, int Offset, Stream OutStream, bool IsEMF)
        {
            if (!IsEMF)
            {
                WmfHeader WmfHead= new WmfHeader();
                WmfHead.Key=0x9AC6CDD7;

                //On Xls format coords are 32 bits, on wmf file format they are 16. That's why we have to convert.
                //catch overflows. We could use unchecked here, but this way we setup the final result.
                try{WmfHead.Left=(Int16)BitConverter.ToInt32(XlsData, Offset+4+0);}
                catch(OverflowException){WmfHead.Left=0;}
                try{WmfHead.Top=(Int16)BitConverter.ToInt32(XlsData, Offset+4+4);}
                catch(OverflowException){WmfHead.Top=0;}
                try{WmfHead.Right=(Int16)BitConverter.ToInt32(XlsData, Offset+4+8);}
                catch(OverflowException){WmfHead.Right=0xFFF;}
                try{WmfHead.Bottom=(Int16)BitConverter.ToInt32(XlsData, Offset+4+12);}
                catch(OverflowException){WmfHead.Bottom=0xFFF;}

                WmfHead.Inch=96;
                WmfHead.CheckSum=ComputeAldusCheckSum(WmfHead.Data);
                OutStream.Write(WmfHead.Data,0,WmfHead.Data.Length);
            }
            //Common part on EMF and WMF
            int IsCompressed=XlsData[Offset+32];

#if (FRAMEWORK30 || !COMPACTFRAMEWORK)
            if (IsCompressed==0)  //Data is compressed.
            {
				using (TCompressor Cmp= new TCompressor())
				{
					Cmp.Inflate(XlsData, Offset+34, OutStream);
				}
            }
#else
            throw new FlexCelException("Operation not supported in CF 2.0");
#endif
        }
    }

    /// <summary>
    /// A wmf file header
    /// </summary>
    internal class WmfHeader
    {
        private byte[] FData;
        internal WmfHeader()
        {
            FData=new byte[4+2+8+2+4+2];
        }

        internal UInt32 Key {get {return BitConverter.ToUInt32(FData, 0);} set{BitConverter.GetBytes((UInt32)value).CopyTo(FData,0);}}
        internal Int16 Handle {get {return BitConverter.ToInt16(FData, 4);} set{BitConverter.GetBytes((Int16)value).CopyTo(FData,4);}}
        internal Int16 Left {get {return BitConverter.ToInt16(FData, 6);} set{BitConverter.GetBytes((Int16)value).CopyTo(FData,6);}}
        internal Int16 Top  {get {return BitConverter.ToInt16(FData, 8);} set{BitConverter.GetBytes((Int16)value).CopyTo(FData,8);}}
        internal Int16 Right {get {return BitConverter.ToInt16(FData, 10);} set{BitConverter.GetBytes((Int16)value).CopyTo(FData,10);}}
        internal Int16 Bottom {get {return BitConverter.ToInt16(FData, 12);} set{BitConverter.GetBytes((Int16)value).CopyTo(FData,12);}}

        internal UInt16 Inch {get {return BitConverter.ToUInt16(FData, 14);} set{BitConverter.GetBytes((UInt16)value).CopyTo(FData,14);}}
        internal int Reserved {get {return BitConverter.ToInt32(FData, 16);} set{BitConverter.GetBytes((Int32)value).CopyTo(FData,16);}}
        internal UInt16 CheckSum {get {return BitConverter.ToUInt16(FData, 20);} set{BitConverter.GetBytes((UInt16)value).CopyTo(FData,20);}}

        internal byte[] Data {get {return FData;}}
    }
}
