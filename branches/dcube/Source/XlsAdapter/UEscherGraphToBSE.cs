using System;
using System.IO;
using FlexCel.Core;

namespace FlexCel.XlsAdapter
{
	/// <summary>
	/// Converts a bitmap to BSE representation.
	/// </summary>
	internal sealed class EscherGraphToBSE
	{
		private EscherGraphToBSE(){}

		/*
		 *   TBSEHeader= packed record
				btWin32:        byte;                  // Required type on Win32
				btMacOS:        byte;                  // Required type on Mac
				rgbUid:         TMd5Digest;            // Identifier of blip
				tag:            word;                  // currently unused
				size:           Cardinal;              // Blip size in stream
				cRef:           Cardinal;              // Reference count on the blip
				foDelay:        Cardinal;              // File offset in the delay stream
				usage:          byte;                  // How this blip is used (MSOBLIPUSAGE)
				cbName:         byte;                  // length of the blip name
				unused2:        byte;                  // for the future
				unused3:        byte;                  // for the future
			  end;
		*/

		enum TBSEHeader 
		{
			btWin32        = 0,                  // Required type on Win32
			btMacOS        = btWin32+1,          // Required type on Mac
			rgbUid         = btMacOS+1,          // Identifier of blip
			tag            = rgbUid+16,          // currently unused
			size           = tag+2,              // Blip size in stream
			cRef           = size+4,             // Reference count on the blip
			foDelay        = cRef+4,             // File offset in the delay stream
			usage          = foDelay+4,          // How this blip is used (MSOBLIPUSAGE)
			cbName         = usage+1,            // length of the blip name
			unused2        = cbName+1,           // for the future
			unused3        = unused2+1,          // for the future
		    length         = unused3+1
		}


        internal static TEscherBSERecord Convert(byte[] Data, TXlsImgType DataType,
            TEscherDwgGroupCache DwgGroupCache, TEscherDwgCache DwgCache)
        {
            byte[] BSEHeader=TCompactFramework.GetBSEHeader(Data, (int)TBSEHeader.length, (int)TBSEHeader.rgbUid);
         
            using (MemoryStream BlipData=new MemoryStream())
            {
                //Common header
                BlipData.Write(BSEHeader, (int)TBSEHeader.rgbUid, (int)TBSEHeader.tag - (int)TBSEHeader.rgbUid);

                // Specific info
                if ((DataType == TXlsImgType.Jpeg) || (DataType == TXlsImgType.Png))
                    LoadDataBitmap(Data, DataType, BlipData); else
				if (DataType == TXlsImgType.Bmp)
					LoadDataBmp(Data, DataType, BlipData); else
					LoadDataWMF(Data, DataType, BlipData);

                BSEHeader[(int)TBSEHeader.btWin32]= (byte) XlsEscherConsts.XlsImgConv(DataType);
                BSEHeader[(int)TBSEHeader.btMacOS]= (byte) msoblip.PICT;

                BitOps.SetWord(BSEHeader, (int)TBSEHeader.tag, 0xFF);
                BitOps.SetCardinal(BSEHeader, (int)TBSEHeader.size, BlipData.Length+ XlsEscherConsts.SizeOfTEscherRecordHeader);
                BitOps.SetCardinal(BSEHeader,(int)TBSEHeader.cRef,0);
                BitOps.SetCardinal(BSEHeader,(int)TBSEHeader.foDelay,0);

                TEscherRecordHeader Eh= new TEscherRecordHeader();
                Eh.Id= (int)Msofbt.BSE;
                Eh.Pre=2 + ( (int)XlsEscherConsts.XlsImgConv(DataType) << 4);
                Eh.Size=BitOps.GetCardinal(BSEHeader,(int)TBSEHeader.size) + BSEHeader.Length;
                TEscherBSERecord Result= new TEscherBSERecord(Eh, DwgGroupCache, DwgCache, DwgGroupCache.BStore);

                TEscherRecordHeader BlipHeader=new TEscherRecordHeader();
                BlipHeader.Id= (int)XlsEscherConsts.XlsBlipHeaderConv(DataType);
                BlipHeader.Pre= (int) XlsEscherConsts.XlsBlipSignConv(DataType) << 4;
                BlipHeader.Size= BlipData.Length;

                BlipData.Position=0;
                Result.CopyFromData(BSEHeader, BlipHeader, BlipData);
            
                return Result;
            }
		}

		private static void LoadDataBitmap(byte[] Data, TXlsImgType DataType, Stream BlipData)
		{
			BlipData.WriteByte(0xFF); //Tag
			BlipData.Write(Data, 0, Data.Length);
		}

		private static void LoadDataBmp(byte[] Data, TXlsImgType DataType, Stream BlipData) //A BMP is a DIB with a 14 byte header. We need to split the header.
		{
			int Ofs = 0;
			if (BitOps.GetWord(Data, 0) == 0x4D42) Ofs = 14; //bitmap type ("BM"). If this is not this way, we will understand this is a DIB.

			BlipData.WriteByte(0xFF); //Tag
			BlipData.Write(Data, Ofs, Data.Length - Ofs);
		}


/*
		TWMFBlipHeader = packed record
			m_rgbUid: TMd5Digest;  { The secondary, or data, UID - should always be set. }

		Metafile Blip overhead = 34 bytes. m_cb gives the number of
		bytes required to store an uncompressed version of the file, m_cbSave
		is the compressed size.  m_mfBounds gives the boundary of all the
		drawing calls within the metafile (this may just be the bounding box
		or it may allow some whitespace, for a WMF this comes from the
		SetWindowOrg and SetWindowExt records of the metafile). 
		
			m_cb: integer;           // Cache of the metafile size
			m_rcBounds: Array[0..3] of integer;     // Boundary of metafile drawing commands
			m_ptSize: Array[0..1] of integer;       // Size of metafile in EMUs
			m_cbSave: integer;       // Cache of saved size (size of m_pvBits)
			m_fCompression: byte; // MSOBLIPCOMPRESSION
			m_fFilter: byte;      // always msofilterNone
       end;
*/

        //1 point = 12700 emu.   
        private static void LoadDataWMF(byte[] Data, TXlsImgType DataType, Stream BlipData)
        {

            byte[] cb = BitConverter.GetBytes((UInt32) Data.Length);
            BlipData.Write(cb, 0, cb.Length);

            //This one is used only on metafiles.
            if (DataType== TXlsImgType.Wmf)
            {
                //On Xls format coords are 32 bits, on wmf file format they are 16. That's why we have to convert.
                BlipData.Write(BitConverter.GetBytes((Int32)BitConverter.ToInt16(Data, 6)),0,4);
                BlipData.Write(BitConverter.GetBytes((Int32)BitConverter.ToInt16(Data, 8)),0,4);
                BlipData.Write(BitConverter.GetBytes((Int32)BitConverter.ToInt16(Data, 10)),0,4);
                BlipData.Write(BitConverter.GetBytes((Int32)BitConverter.ToInt16(Data, 12)),0,4);

                //byte[] ptSize = {0x18,0xF0,0x01,0x00,  0x18,0xF0,0x01,0x00};  //100 points x 100 points. This one is usd on EMF
                byte[] ptSize = {0x18,0xF0,0xFF,0x00,  0x18,0xF0,0xFF,0x00};  //something bigger. This will be the default size. This one is usd on EMF

                BlipData.Write(ptSize, 0, ptSize.Length);
            }
            else
            {
                byte[] rcBounds = new byte[4*4];
                Array.Copy(Data, 8, rcBounds, 0, rcBounds.Length);
                BlipData.Write(rcBounds, 0, rcBounds.Length);

                //ptSize
                int WidthMM = BitConverter.ToInt32(Data, 24 + 8) - BitConverter.ToInt32(Data, 24);
                int HeightMM = BitConverter.ToInt32(Data, 24 + 12) - BitConverter.ToInt32(Data, 24 + 4);
                byte[] WidthEmu = BitConverter.GetBytes(WidthMM * 360);
                BlipData.Write(WidthEmu, 0, WidthEmu.Length);
                byte[] HeightEmu = BitConverter.GetBytes(HeightMM * 360);
                BlipData.Write(HeightEmu, 0, HeightEmu.Length);
            }

            byte[]OtherDat = {0,0,0,0,  0, 254};
            long StreamPos= BlipData.Position;
            BlipData.Write(OtherDat, 0, 6);
            if (DataType== TXlsImgType.Emf)
                XlsMetafiles.ToXls(Data, BlipData, true);
            else
                XlsMetafiles.ToXls(Data, BlipData, false);

            //GoBack and set m_cbSave
            BlipData.Position=StreamPos;
            BlipData.Write(BitConverter.GetBytes((UInt32)( BlipData.Length-StreamPos-6)),0,4);
            BlipData.Position=BlipData.Length;
		}


	}
}
