#if (!WPF)
//This code is from http://msdn.microsoft.com/library/default.asp?url=/library/en-us/dnaspp/html/colorquant.asp

// 
//  THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF 
//  ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO 
//  THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A 
//  PARTICULAR PURPOSE. 
//  
//    This is sample code and is freely distributable. 
//

using System;
using System.Runtime.InteropServices;
using System.Security.Permissions;
using System.Security;

using System.Drawing;
using System.Drawing.Imaging;
using FlexCel.Core;

namespace FlexCel.Render
{
	/// <summary>
	/// A generic implementation for any Color Quantizer.
	/// </summary>
	public abstract class Quantizer
	{
		/// <summary>
		/// Constructs the quantizer.
		/// </summary>
		/// <param name="singlePass">If true, the quantization only needs to loop through the Source pixels once</param>
		/// <remarks>
		/// If you construct this class with a true value for singlePass, then the code will, when quantizing your image,
		/// only call the 'QuantizeImage' function. If two passes are required, the code will call 'InitialQuantizeImage'
		/// and then 'QuantizeImage'.
		/// </remarks>
		protected Quantizer(bool singlePass)
		{
			_singlePass = singlePass;
		}

		/// <summary>
		/// Quantize an image and return the resulting Result bitmap
		/// </summary>
		/// <param name="Source">The image to quantize</param>
        /// <param name="Result">A quantized version of the image</param>
#if (!FULLYMANAGED)
#if (FRAMEWORK40)
        [SecurityCritical]
#else
        [SecurityPermission(SecurityAction.LinkDemand, Flags = SecurityPermissionFlag.UnmanagedCode)]
#endif
        public void Quantize(Image Source, Bitmap Result)
		{
			if (Source.PixelFormat != PixelFormat.Format32bppPArgb && Source.PixelFormat != PixelFormat.Format32bppArgb)
				throw new ArgumentException("Please convert the image to Format32bppPArgb before calling this method");
			if (Result.PixelFormat != PixelFormat.Format8bppIndexed)
				throw new ArgumentException("Result image should be Format8bppIndexed");

			// Get the size of the Source image
			int height = Source.Height;
			int width = Source.Width;

			// And construct a rectangle from these dimensions
			Rectangle bounds = new Rectangle(0, 0, width, height);

			// Define a pointer to the bitmap data
			BitmapData sourceData = null;
            Bitmap BmpSource = (Bitmap)Source;
            try
			{
				// Get the Source image bits and lock into memory
				sourceData = BmpSource.LockBits(bounds, ImageLockMode.ReadOnly, Source.PixelFormat);

				// Call the FirstPass function if not a single pass algorithm.
				// For something like an octree quantizer, this will run through
				// all image pixels, build a data structure, and create a palette.
				if (!_singlePass)
					FirstPass(sourceData, width, height);

				// Then set the color palette on the Result bitmap. I'm passing in the current palette 
				// as there's no way to construct a new, empty palette.
				Result.Palette = this.GetPalette(Result.Palette);

				// Then call the second pass which actually does the conversion
				SecondPass(sourceData, Result, width, height, bounds);
			}
			finally
			{
				// Ensure that the bits are unlocked
				BmpSource.UnlockBits(sourceData);
			}
		}
#else
        public void Quantize(Image Source, Bitmap Result)
		{
            FlxMessages.ThrowException(FlxErr.ErrNeedsUnmanaged);
        }
#endif

#if (!FULLYMANAGED)
		/// <summary>
		/// Execute the first pass through the pixels in the image
		/// </summary>
		/// <param name="sourceData">The Source data</param>
		/// <param name="width">The width in pixels of the image</param>
		/// <param name="height">The height in pixels of the image</param>
#if (FRAMEWORK40)
        [SecurityCritical]
#else
        [SecurityPermission(SecurityAction.LinkDemand, Flags = SecurityPermissionFlag.UnmanagedCode)]
#endif
        protected internal virtual void FirstPass(BitmapData sourceData, int width, int height)
		{
			// Define the Source data pointers. The Source row is a byte to
			// keep addition of the stride value easier (as this is in bytes)
			long SourceScan0 = (long)sourceData.Scan0;
			int pSourceRow = 0;
			int SourceStride = sourceData.Stride;

            Int32[] SourceRow = new Int32[width];
			// Loop through each row
			for (int row = 0; row < height; row++)
			{
				// Set the Source pixel to the first pixel in this row
				int pSourcePixel = 0;

                Marshal.Copy((IntPtr)(SourceScan0 + pSourceRow), SourceRow, 0, SourceRow.Length);

				// And loop through each column
				for (int col = 0; col < width; col++, pSourcePixel++)
					// Now I have the pixel, call the FirstPassQuantize function...
                    InitialQuantizePixel(SourceRow[pSourcePixel]);

				// Add the stride to the Source row
				pSourceRow += SourceStride;
			}
		}

		/// <summary>
		/// Execute a second pass through the bitmap
		/// </summary>
		/// <param name="sourceData">The Source bitmap, locked into memory</param>
		/// <param name="Result">The Result bitmap</param>
		/// <param name="width">The width in pixels of the image</param>
		/// <param name="height">The height in pixels of the image</param>
		/// <param name="bounds">The bounding rectangle</param>
#if (FRAMEWORK40)
        [SecurityCritical]
#else
        [SecurityPermission(SecurityAction.LinkDemand, Flags = SecurityPermissionFlag.UnmanagedCode)]
#endif
        protected internal virtual void SecondPass(BitmapData sourceData, Bitmap Result, int width, int height, Rectangle bounds)
		{
			BitmapData outputData = null;

			try
			{
				// Lock the Result bitmap into memory
				outputData = Result.LockBits(bounds, ImageLockMode.WriteOnly, PixelFormat.Format8bppIndexed);

				// Define the Source data pointers. The Source row is a byte to
				// keep addition of the stride value easier (as this is in bytes)
				long SourceScan0 = (long)sourceData.Scan0;
				int pSourceRow = 0;
				int SourceStride = sourceData.Stride;

				// Now define the destination data pointers
				long DestScan0 = (long)outputData.Scan0;
				int pDestRow = 0;
				int DestStride = outputData.Stride;

				// And convert the first pixel, so that I have values going into the loop
				int LastPixel = Marshal.ReadInt32((IntPtr)SourceScan0, pSourceRow);
				byte pixelValue = QuantizePixel(LastPixel);

				// Assign the value of the first pixel
				//Marshal.WriteByte(DestScan0, pDestRow, pixelValue);  //not needed, it iwll be set on the loop.

                Int32[] SourceRow = new Int32[width];
                byte[] DestRow = new byte[DestStride];
				// Loop through each row
				for (int row = 0; row < height; row++)
				{
					int pSourcePixel = 0;
                    Marshal.Copy((IntPtr)(SourceScan0 + pSourceRow), SourceRow, 0, SourceRow.Length);

					// Loop through each pixel on this scan line
					for (int col = 0; col < width; col++, pSourcePixel ++)
					{
						// Check if this is the same as the last pixel. If so use that value
						// rather than calculating it again. This is an inexpensive optimisation.
						int NextPixel = SourceRow[pSourcePixel];
						if (LastPixel != NextPixel)
						{
							// Quantize the pixel
							pixelValue = QuantizePixel(NextPixel);

							// And setup the previous pointer
							LastPixel = NextPixel;
						}

						// And set the pixel in the Result
						DestRow[col] = pixelValue;
					}

                    Marshal.Copy(DestRow, 0, (IntPtr)(DestScan0 + pDestRow), DestRow.Length);

					// Add the stride to the Source row
					pSourceRow += SourceStride;

					// And to the destination row
					pDestRow += DestStride;
				}
			}
			finally
			{
				// Ensure that I unlock the Result bits
				Result.UnlockBits(outputData);
			}
		}
#endif
		/// <summary>
		/// Override this to process the pixel in the first pass of the algorithm
		/// </summary>
		/// <param name="pixel">The pixel to quantize</param>
		/// <remarks>
		/// This function need only be overridden if your quantize algorithm needs two passes,
		/// such as an Octree quantizer.
		/// </remarks>
		protected virtual void InitialQuantizePixel(Int32 pixel)
		{
		}

		/// <summary>
		/// Override this to process the pixel in the second pass of the algorithm
		/// </summary>
		/// <param name="pixel">The pixel to quantize</param>
		/// <returns>The quantized value</returns>
		protected abstract byte QuantizePixel(Int32 pixel);

		/// <summary>
		/// Retrieve the palette for the quantized image
		/// </summary>
		/// <param name="original">Any old palette, this is overrwritten</param>
		/// <returns>The new color palette</returns>
		protected abstract ColorPalette GetPalette(ColorPalette original);

		/// <summary>
		/// Flag used to indicate whether a single pass or two passes are needed for quantization.
		/// </summary>
		private bool _singlePass;
	}
}
#endif
