using System;
using System.Runtime.InteropServices;
using System.Security.Permissions;

using System.Drawing;
using FlexCel.Core;
using System.Security;
using System.Drawing.Imaging;

namespace FlexCel.Render
{
	/// <summary>
	/// Converts a true color image to black and white using  Floyd-Steinberg dithering.
    /// Needs UNMANAGED permissions in order to process the bits of the image.
	/// </summary>
	public sealed class FloydSteinbergDither
	{
		private FloydSteinbergDither()	{}

        /// <summary>
        /// Converts a true color image to black and white, using Floyd-Steinberg error diffusion.
        /// </summary>
        /// <param name="Source">Image to convert. Must be on Format32bppPArgb.</param>
        /// <param name="Result">Here we will draw the converted image. Must be Format1bppIndexed and have the same size as Source.</param>
#if (!FULLYMANAGED)
#if (FRAMEWORK40)
        [SecurityCritical]
#else
        [SecurityPermission(SecurityAction.LinkDemand, Flags = SecurityPermissionFlag.UnmanagedCode)]
#endif
        public static void ConvertToBlackAndWhite(Bitmap Source, Bitmap Result)
		{
			if (Source.PixelFormat != PixelFormat.Format32bppPArgb)
				throw new ArgumentException("Please convert the image to Format32bppPArgb before calling this method");
			if (Result.PixelFormat != PixelFormat.Format1bppIndexed)
				throw new ArgumentException("Result image should be Format1bppIndexed");
			int[][] RowBrightness = new int[2][];
			RowBrightness[0] = new int[Result.Width + 1];
			RowBrightness[1] = new int[Result.Width + 1];

			BitmapData DestBits = Result.LockBits(new Rectangle(0, 0, Result.Width, Result.Height), ImageLockMode.ReadWrite, PixelFormat.Format1bppIndexed);      
			try
			{
				int SourceWidth = Source.Width;
				int SourceHeight = Source.Height;
				//lock the bits of the original bitmap
				BitmapData SourceBits = Source.LockBits(new Rectangle(0, 0, SourceWidth, SourceHeight), ImageLockMode.ReadOnly, Source.PixelFormat);
				try
				{
					int SourceBitsStride = SourceBits.Stride;
					int DestBitsStride = DestBits.Stride;
					long SourceBitsScan0 = (long)SourceBits.Scan0;
                    long DestBitsScan0 = (long)DestBits.Scan0;

					//scan through the pixels Y by X
                    byte[] SourceRow = new byte[SourceWidth * 4];  //Read entire lines into memory.
                    byte[] DestRow = new byte[DestBitsStride];  //write entire lines into memory.
                    for (int y = 0; y < SourceHeight; y++)
					{
						int RightPixel = 0;

						UInt32 DestData = 0;
                        int DestIndex = 0;
						byte MaskByte = 0;  //This is little endian. Bytes grow to the right.
						UInt32 MaskBit  = 0x80; //Bits grow to the left.

                        Marshal.Copy((IntPtr)(SourceBitsScan0), SourceRow, 0, SourceRow.Length);
                        SourceBitsScan0 += SourceBitsStride;
						for(int x = 0; x < SourceWidth; x++)
						{
							//check brightness
							int Brightness =
								(
								76 * SourceRow[x*4+2] + 
								150 * SourceRow[x*4+1] + 
								29 * SourceRow[x*4]
								) / 255 +
								RightPixel + RowBrightness[1][x];
						
							Brightness = Brightness < 0 ? 0: (Brightness > 255 ? 255: Brightness);

							int Error;
							if (Brightness > 127)
							{
								Error = Brightness - 255;
								//set dest pixel if its bright.
								DestData |= (MaskBit << MaskByte);
							}
							else
							{
								Error = Brightness;
								//No need to unset the pixel.
							}

							unchecked
							{
								if ((MaskByte == 8*3 && MaskBit == 1) || x == SourceWidth - 1)
								{
									//Marshal.WriteInt32(DestBitsScan0, DestIndex, (Int32)DestData);  This is too slow on .NET 20
                                    DestRow[DestIndex] = (byte)(DestData & 0xFF);
                                    DestRow[DestIndex + 1] = (byte)((DestData >> 8) & 0xFF);
                                    DestRow[DestIndex + 2] = (byte)((DestData >> 16) & 0xFF);
                                    DestRow[DestIndex + 3] = (byte)((DestData >> 24) & 0xFF);
                                    DestIndex += 4;
                                    MaskByte = 0;
									MaskBit = 0x80;
									DestData = 0;
								}
								else
								{
									MaskBit >>= 1;
									if (MaskBit == 0)
									{
										MaskByte += 8;
										MaskBit = 0x80;
									}
								}
							}

							//Diffuse to the right
							RightPixel = (Error * 7) / 16;

							//Diffuse to the top
							if (x > 0)
							{
								RowBrightness[0][x-1] += (Error * 3) / 16;
							}
							RowBrightness[0][x] += (Error * 5) / 16;
							RowBrightness[0][x+1] += (Error * 1) / 16;
						}

                        Marshal.Copy(DestRow, 0, (IntPtr)(DestBitsScan0), DestRow.Length);
                        DestBitsScan0 += DestBitsStride;

						int[] Tmp = RowBrightness[1]; //Avoid creating other array.
						RowBrightness[1] = RowBrightness[0];
						RowBrightness[0] = Tmp;
						Array.Clear(RowBrightness[0], 0, RowBrightness[0].Length);
					}

				}
				finally
				{
					Source.UnlockBits(SourceBits);
				}
			}
			finally
			{
				Result.UnlockBits(DestBits);
			}
		}
#else
        public static void ConvertToBlackAndWhite(Bitmap Source, Bitmap Result)
		{
            FlxMessages.ThrowException(FlxErr.ErrNeedsUnmanaged);
        }
#endif
        /// <summary>
        /// Converts a true color image to black and white, using Floyd-Steinberg error diffusion.
        /// </summary>
        /// <param name="Source">The image we want to convert. Must be on Format32bppPArgb</param>
        /// <returns>The converted black and white image.</returns>
#if (!FULLYMANAGED)
#if (FRAMEWORK40)
        [SecurityCritical]
#else
        [SecurityPermission(SecurityAction.LinkDemand, Flags = SecurityPermissionFlag.UnmanagedCode)]
#endif
        public static Bitmap ConvertToBlackAndWhite(Bitmap Source)
		{
			Bitmap Result = BitmapConstructor.CreateBitmap(Source.Width, Source.Height, PixelFormat.Format1bppIndexed);
			Result.SetResolution(Source.HorizontalResolution, Source.VerticalResolution);
			ConvertToBlackAndWhite(Source, Result);
			return Result;
		}
#else
        public static Bitmap ConvertToBlackAndWhite(Bitmap Source)
		{
            FlxMessages.ThrowException(FlxErr.ErrNeedsUnmanaged);
            return null;
        }
#endif
	}
}
