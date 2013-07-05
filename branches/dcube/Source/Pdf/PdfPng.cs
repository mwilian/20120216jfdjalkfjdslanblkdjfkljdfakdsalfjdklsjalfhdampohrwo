#region Using directives

using System;
using System.Text;
using System.IO;
using FlexCel.Core;
using System.Diagnostics;

#endregion

namespace FlexCel.Pdf
{

	enum TChunkType
	{
		IHDR = (((byte)'I')<<24)+(((byte)'H')<<16)+(((byte)'D')<<8)+(((byte)'R')<<0),
		IDAT = (((byte)'I')<<24)+(((byte)'D')<<16)+(((byte)'A')<<8)+(((byte)'T')<<0),
		PLTE = (((byte)'P')<<24)+(((byte)'L')<<16)+(((byte)'T')<<8)+(((byte)'E')<<0),
		tRNS = (((byte)'t')<<24)+(((byte)'R')<<16)+(((byte)'N')<<8)+(((byte)'S')<<0)
	}

	internal sealed class TPdfPngData
	{
		internal long Height;
		internal long Width;
		internal byte BitDepth;
		internal byte ColorType;
		internal byte CompressionMethod;
		internal byte FilterMethod;
		internal byte InterlaceMethod;

		internal MemoryStream Data;
		internal byte[] PLTE;
		internal byte[] tRNS;
		internal byte[] SMask;
		internal byte[] OneBitMask;

		internal TPdfPngData(MemoryStream aData)
		{
			Data = aData;
		}
	}

	/// <summary>
	/// Basic information about a PNG file.
	/// </summary>
	public sealed class TPngInformation
	{
		internal long FHeight;
		internal long FWidth;
		internal byte FBitDepth;
		internal byte FColorType;
		internal byte FCompressionMethod;
		internal byte FFilterMethod;
		internal byte FInterlaceMethod;

		/// <summary>
		/// Creates a new TPngInformation class.
		/// </summary>
		public TPngInformation()
		{
		}

		internal TPngInformation(TPdfPngData Data)
		{
			FHeight    = Data.Height;
			FWidth     = Data.Width;
			FBitDepth  = Data.BitDepth;
			FColorType = Data.ColorType;
			FCompressionMethod = Data.CompressionMethod;
			FFilterMethod = Data.FilterMethod;
			FInterlaceMethod = Data.InterlaceMethod;
		}

		#region Properties
		/// <summary>
		/// Height of the image on pixels.
		/// </summary>
		public long Height {get {return FHeight;} set {FHeight = value;}}
		/// <summary>
		/// Width of the image on pixels.
		/// </summary>
		public long Width {get {return FWidth;} set {FWidth = value;}}
		/// <summary>
		/// Bith depth of the image.
		/// </summary>
		public byte BitDepth {get {return FBitDepth;} set {FBitDepth = value;}}
		/// <summary>
		/// Png ColorType (See png reference for more information)
		/// </summary>
		public byte ColorType {get {return FColorType;} set {FColorType = value;}}
		/// <summary>
		/// Png Compression method (See png reference for more information)
		/// </summary>
		public byte CompressionMethod {get {return FCompressionMethod;} set {FCompressionMethod = value;}}
		#endregion
	}

	/// <summary>
	/// A class for reading a PNG image. Mostly for internal use, but it can
	/// return some very basic information on a PNG file too.
	/// </summary>
	public sealed class TPdfPng
	{
		private TPdfPng()
		{
		}

		private static readonly byte[] Signature ={ 137, 80, 78, 71, 13, 10, 26, 10 };

		internal static UInt32 GetUInt32(Stream Data)
		{
			unchecked
			{
				return (UInt32)((Data.ReadByte() << 24) + (Data.ReadByte() << 16) + (Data.ReadByte() << 8) + Data.ReadByte());
			}
		}

		internal static bool CheckHeaders(Stream PngImageData)
		{
			if (PngImageData.Length <10) return false;
			PngImageData.Position = 0;

			for (int i = 0; i < Signature.Length; i++)
			{
				if (Signature[i] != PngImageData.ReadByte())
					return false;
			}
			return true;
		}

		internal static bool IsOkPng(Stream PngImageData)
		{
			if (!CheckHeaders(PngImageData)) return false;

			TPdfPngData OutData = new TPdfPngData(null);
			ReadChunk(PngImageData, OutData);     
			if (OutData.InterlaceMethod!=0 || OutData.BitDepth>8) return false;

			return true;
		}

		internal static void ProcessPng(Stream PngImageData, TPdfPngData OutData)
		{
			if (!CheckHeaders(PngImageData))
				PdfMessages.ThrowException(PdfErr.ErrInvalidPngImage);

			while (PngImageData.Position< PngImageData.Length)
			{
				ReadChunk(PngImageData, OutData);     
			}

			if (OutData.ColorType == 4)
				SuppressAlpha(OutData, 1);

			if (OutData.ColorType == 6)
				SuppressAlpha(OutData, 3);
		}

		/// <summary>
		/// Returns the basic information on a png file. Null if the file is not PNG.
		/// </summary>
		/// <param name="PngImageData">Stream with the image data.</param>
		/// <returns>Null if the image is invalid, or the image properties otherwise.</returns>
		public static TPngInformation GetPngInfo(Stream PngImageData)
		{
			if (!CheckHeaders(PngImageData)) return null;

			TPdfPngData OutData = new TPdfPngData(null);
			ReadChunk(PngImageData, OutData);     
			return new TPngInformation(OutData);
		}


		internal static void ReadChunk(Stream PngImageData, TPdfPngData OutData)
		{
            if (PngImageData.Length - PngImageData.Position < 3 * 4)
            {
                PngImageData.Position = PngImageData.Length;
                return;
            }
			UInt32 Len = GetUInt32(PngImageData);
			UInt32 ChunkType = GetUInt32(PngImageData);
            
			switch ((TChunkType)ChunkType)
			{
				case TChunkType.IHDR: ReadIHDR(Len, PngImageData, OutData); break;
				case TChunkType.IDAT: ReadIDAT(Len, PngImageData, OutData); break;
				case TChunkType.PLTE: ReadPLTE(Len, PngImageData, OutData); break;
				case TChunkType.tRNS: ReadtRNS(Len, PngImageData, OutData); break;
				default: PngImageData.Position+=Len;break; //ignore chunk
			}

			/*UInt32 CRC = */GetUInt32(PngImageData);
		}

		internal static void ReadIHDR(UInt32 Len, Stream PngImageData, TPdfPngData OutData)
		{
			OutData.Width  = GetUInt32(PngImageData);
			OutData.Height = GetUInt32(PngImageData);

			OutData.BitDepth = (byte)PngImageData.ReadByte();
			OutData.ColorType = (byte)PngImageData.ReadByte();
			OutData.CompressionMethod = (byte)PngImageData.ReadByte();
			OutData.FilterMethod = (byte)PngImageData.ReadByte();
			OutData.InterlaceMethod = (byte)PngImageData.ReadByte();
		}

		internal static void ReadIDAT(UInt32 Len, Stream PngImageData, TPdfPngData OutData)
		{
			const int size = 4096;
			byte[] bytes = new byte[size];
			int Remaining=(int)Len;
			int numBytes;

			while(Remaining>0)
			{
				numBytes = PngImageData.Read(bytes, 0, Math.Min(size, Remaining));
				OutData.Data.Write(bytes, 0, numBytes);
				Remaining-=numBytes;
			}
		}

		internal static void ReadPLTE(UInt32 Len, Stream PngImageData, TPdfPngData OutData)
		{
			OutData.PLTE = new byte[Len];
			Sh.Read(PngImageData, OutData.PLTE, 0, (int)Len);
		}

		internal static void ReadtRNS(UInt32 Len, Stream PngImageData, TPdfPngData OutData)
		{
			OutData.tRNS = new byte[Len];
			Sh.Read(PngImageData, OutData.tRNS, 0, (int)Len);
		}

		//We can't crop the image here, without changing the filters.
		private static void SuppressAlpha(TPdfPngData OutData, byte IntBytes)
		{
			using (TCompressor Cmp = new TCompressor())
			{
				using (TCompressor Cmp2 = new TCompressor())
				{
					using (TCompressor Cmp3 = new TCompressor())
					{

						const int BuffLen=4096;

						bool MaskEmpty = true;
						bool IsOneBitMask = true;

						int h= (int)OutData.Height;
						int w= (int)OutData.Width;

#if (FRAMEWORK20 || ICSHARP) && !COMPACTFRAMEWORK
            byte[] NewData = new byte[BuffLen + IntBytes];
            byte[] SMask = new byte[BuffLen + 1];
			byte[] OneBitMask = new byte[BuffLen + 1];
#else
						sbyte[] NewData = new sbyte[BuffLen+IntBytes];
						sbyte[] SMask = new sbyte[BuffLen + 1];
						sbyte[] OneBitMask = new sbyte[BuffLen + 1];
#endif

						using (MemoryStream InflatedStream = new MemoryStream())
						{
							unchecked
							{
								int NewDataPos=0;
								int SMaskPos=0;
								int OneBitMaskPos=0;

								Cmp.Inflate(OutData.Data.ToArray(), 0, InflatedStream);
								InflatedStream.Position=0;

								Cmp.BeginDeflate();
								Cmp2.BeginDeflate();
								Cmp3.BeginDeflate();
                    
								OutData.Data.SetLength(0); Stream OutStream=OutData.Data;
								using (MemoryStream SMaskStream = new MemoryStream())
								{
									using (MemoryStream OneBitMaskStream= new MemoryStream())
									{
										int OneBitMaskInnerPos = 128;
										for (int r=0; r<h; r++)
										{
											byte LastSMask = 0;

#if (FRAMEWORK20 || ICSHARP) && !COMPACTFRAMEWORK
                                byte RowFilter = (byte)InflatedStream.ReadByte();
#else
											sbyte RowFilter = (sbyte)InflatedStream.ReadByte();
#endif

											#region inlined //for speed on debug mode
											if (NewDataPos>=BuffLen)
											{
												Cmp.IncDeflate(NewData, 0, NewDataPos, OutStream);
												NewDataPos=0;
											}
											if (SMaskPos>=BuffLen)
											{
												Cmp2.IncDeflate(SMask, 0, SMaskPos, SMaskStream);
												SMaskPos=0;
											}
											if (OneBitMaskPos>=BuffLen)
											{
												Cmp3.IncDeflate(OneBitMask, 0, OneBitMaskPos, OneBitMaskStream);
												OneBitMaskPos=0;
												OneBitMask[OneBitMaskPos] = 0;
												OneBitMaskInnerPos=128;
											}
											#endregion

											NewData[NewDataPos++] = RowFilter;
											SMask[SMaskPos++] = RowFilter;
											if (IsOneBitMask) 
											{
												if (OneBitMaskInnerPos < 128) 
												{
													OneBitMaskPos++; //finish row.
													OneBitMask[OneBitMaskPos] = 0;
													OneBitMaskInnerPos = 128;
												}
											}

											for (int c=0; c< w; c++)
											{
												#region inlined //for speed on debug mode
												if (NewDataPos>=BuffLen)
												{
													Cmp.IncDeflate(NewData, 0, NewDataPos, OutStream);
													NewDataPos=0;
												}
												if (SMaskPos>=BuffLen)
												{
													Cmp2.IncDeflate(SMask, 0, SMaskPos, SMaskStream);
													SMaskPos=0;
												}
												if (OneBitMaskPos>=BuffLen)
												{
													Cmp3.IncDeflate(OneBitMask, 0, OneBitMaskPos, OneBitMaskStream);
													OneBitMaskPos=0;
													OneBitMask[OneBitMaskPos] = 0;
													OneBitMaskInnerPos=128;
												}
												#endregion

												for (int b=0;b<IntBytes;b++)
#if (FRAMEWORK20 || ICSHARP) && !COMPACTFRAMEWORK
									{    
										NewData[NewDataPos++] = (byte)InflatedStream.ReadByte();
									}
									
									byte SMaskData = (byte)InflatedStream.ReadByte(); 
									SMask[SMaskPos++] = SMaskData;
#else
												{
													NewData[NewDataPos++] = (sbyte)InflatedStream.ReadByte();
												}

												byte SMaskData = (byte)InflatedStream.ReadByte(); 
												SMask[SMaskPos++] = (sbyte)SMaskData;
#endif
												if (MaskEmpty && SMaskData != 0xFF)
												{
													MaskEmpty = false;
												}

												if (IsOneBitMask)
												{
													//SMaskData might have been flushed, and so contain invalid data.
													//So we need a separate LastSMask for sub filter, and the whole last scanline for the others.
													//As this is only an optimization (it will work the same with an SMask), we will only contemplate filters 0 and 1.
													if (RowFilter == 1) //sub
													{
                                                        unchecked { SMaskData += LastSMask; }  
														LastSMask = SMaskData;
													}

													if (RowFilter > 1 || (SMaskData != 0xFF && SMaskData != 0)) 
													{
														IsOneBitMask = false;
													}
													else
													{

#if (FRAMEWORK20 || ICSHARP) && !COMPACTFRAMEWORK
											OneBitMask[OneBitMaskPos] |= (byte)(~SMaskData &  OneBitMaskInnerPos);
#else
														OneBitMask[OneBitMaskPos] = (sbyte)((byte)OneBitMask[OneBitMaskPos] | (~(byte)SMaskData &  OneBitMaskInnerPos));
#endif
														if (OneBitMaskInnerPos > 1)
														{
															OneBitMaskInnerPos >>= 1;
														}
														else
														{
															OneBitMaskPos++;
															OneBitMask[OneBitMaskPos] = 0;
															OneBitMaskInnerPos = 128;
														}
													}
												}
											}
										}
                    

										#region inlined //for speed on debug mode
										if (NewDataPos>0)
										{
											Cmp.IncDeflate(NewData, 0, NewDataPos, OutStream);
											NewDataPos=0;
										}
										if (SMaskPos>0)
										{
											Cmp2.IncDeflate(SMask, 0, SMaskPos, SMaskStream);
											SMaskPos=0;
										}
										if (OneBitMaskPos>0)
										{
											if (OneBitMaskInnerPos < 128) OneBitMaskPos++; //finish row.
											Cmp3.IncDeflate(OneBitMask, 0, OneBitMaskPos, OneBitMaskStream);
											OneBitMaskPos=0;
											OneBitMask[OneBitMaskPos] = 0;
										}
										#endregion

										Cmp.EndDeflate(OutStream);
										Cmp2.EndDeflate(SMaskStream);
										Cmp3.EndDeflate(OneBitMaskStream);

										if (!MaskEmpty) 
										{
											if (IsOneBitMask) OutData.OneBitMask = OneBitMaskStream.ToArray();
											OutData.SMask = SMaskStream.ToArray();
										}

									}

								}
							}
						}
					}
				}
			}
		}

		private static byte[] UnFilter(byte[] Img, int ScanLine, int h)
		{
			//See http://www.w3.org/TR/PNG/

			int ScanPos =0;
			for (int r = 0; r<h; r++)
			{
				byte RowFilter = Img[ScanPos];

				switch (RowFilter)
				{
					case 0:  //identity.
						break;

					case 1:  //Sub.
						Img[ScanPos]=0; //Reset row filter.
						for (int c=2; c<ScanLine; c++)
						{
							unchecked
							{
								Img[ScanPos+c]+= Img[ScanPos+c-1];
							}
						}
						break;

					case 2:  //Up.
						Img[ScanPos]=0; //Reset row filter.
						if (r<=0) break;

						for (int c=1; c<ScanLine; c++)
						{
							unchecked
							{
								Img[ScanPos+c]+= Img[ScanPos+c-ScanLine];
							}
						}
						break;

					case 3:  //Average.
						Img[ScanPos]=0; //Reset row filter.

						for (int c=1; c<ScanLine; c++)
						{
							//Do not overflow this part. We use ints.
							int up = r > 0? (int)Img[ScanPos+c-ScanLine]: 0;
							int left = Img[ScanPos+c-1];
							int delta = (int)((uint)(up + left) >> 1);

							unchecked
							{
								Img[ScanPos+c]= (byte)(Img[ScanPos+c] + delta);
							}
						}
						break;

					
					case 4: //Paeth
						Img[ScanPos]=0; //Reset row filter.

						for (int c=1; c<ScanLine; c++)
						{
							//Do not overflow this part. We use ints.
							int b = r > 0? (int)Img[ScanPos+c-ScanLine]: 0;
							int a = Img[ScanPos+c-1];
							int cp = r > 0? (int)Img[ScanPos+c-ScanLine - 1]: 0;

							unchecked
							{
								Img[ScanPos+c]= (byte)(Img[ScanPos+c] + Paeth(a,b,cp));
							}
						}
						break;
				}

				ScanPos += ScanLine;
			}

			return Img;
		}

		private static int Paeth(int a, int b, int c)
		{
			int p = a + b - c;
			int pa = Math.Abs(p - a);
			int pb = Math.Abs(p - b);
			int pc = Math.Abs(p - c);

			int Pr;
			if (pa <= pb && pa <= pc) Pr = a;
			else if (pb <= pc) Pr = b;
			else Pr = c;
			return Pr;
		}

		private static byte[] DecodeImage(byte[] Img, int w, int h, int ScanLine)
		{
			using (TCompressor Cmp = new TCompressor())
			{

				if (w<=0) return new byte[2];
				byte[] Result = new byte[h * ScanLine];
				using (MemoryStream InflatedStream = new MemoryStream())
				{
					unchecked
					{
						Cmp.Inflate(Img, 0, InflatedStream);
						InflatedStream.Position=0;
                    
						Sh.Read(InflatedStream, Result, 0, h*ScanLine);
						return UnFilter(Result, ScanLine, h);

					}
				}
			}
		}

		internal static byte[] GetIndexedSMask(byte[] Img, int w, int h, byte[] tRNS, int BitsPerPixel)
		{
			int ScanLine = 1+ (w-1)*BitsPerPixel/8 +1;  //RowFilter+Bytes needed for a row.

			byte[] RawImage = DecodeImage(Img, w, h, ScanLine);
			byte[] SMask = BitsPerPixel ==8? RawImage: new byte[h * (w+1)];  //no need for a new array if the image depth is 8.
			
			int ScanPos = 0; 
			int SMaskPos = 1; //first byte is the row filter, already in 0.

			byte MaxBit = (byte) (8/BitsPerPixel);

			byte[] Sh=new byte[8]; //Cached for speed.
			byte[] Sr=new byte[8]; //Cached for speed.
			for (int i=0;i<BitsPerPixel;i++)
				Sh[0] |= (byte)(1<<i);

			for (int bit=1; bit<MaxBit; bit++)
			{
				Sh[bit]=(byte)(Sh[bit-1]<<BitsPerPixel);
				Sr[bit]=(byte)(Sr[bit-1]+BitsPerPixel); //Sr[0]=0
			}

			for (int r=0; r<h; r++)
			{
				int MaxSMaskPos = SMaskPos+w;  //Some images might use only a part of the last byte.
				for (int c=1; c<ScanLine; c++)
				{
					byte b = RawImage[ScanPos + c];

					for (int bit=MaxBit-1; bit>=0; bit--)
					{
						if (SMaskPos>= MaxSMaskPos) break;
						byte b0 = (byte)((b & Sh[bit])>>Sr[bit]); byte x = b0 <tRNS.Length? (byte)tRNS[b0] : (byte)255;
						SMask[SMaskPos] = x; 
						SMaskPos++;
					}						

				}
				ScanPos += ScanLine;
				SMaskPos ++; //Rowfilter
			}

			using (TCompressor Cmp = new TCompressor())
			{
				using (MemoryStream ms = new MemoryStream())
				{
					Cmp.Deflate(SMask, 0, ms);
					return ms.ToArray();
				}
			}

		}

	}

}
