using System;
using System.IO;
using System.Globalization;
using System.Text;

#if (WPF)
using TImage = System.Windows.Media.ImageSource;
#else
using TImage = System.Drawing.Image;
using System.Drawing;
using System.Drawing.Imaging;
#endif


namespace FlexCel.Pdf
{
	internal class TPdfImage : IComparable
	{
		private byte[] FImage;  //We won't use an image here because it is disposable and it could be disposed outside this class.
		private int Id;
		private int FImageWidth;
		private int FImageHeight;
		private int FBitsPerComponent;
		private string FColorSpace;
		private string FFilterName;
		private string FDecodeParmsName;
		private string FDecodeParmsSMask;
		private string FMask;
		private byte[] FSMask;
		private byte[] OneBitMask;

		private long FTransparentColor;

		private int ImgObjId;
		private int SMaskId;

		public TPdfImage(TImage aImage, int aId, Stream aImageData, long transparentColor, bool defaultToJpg)
		{
			Id = aId;
			FMask=null;
			FSMask=null;
			FTransparentColor = transparentColor;

			if (aImageData !=null && TPdfPng.IsOkPng(aImageData))
			{
				aImageData.Position=0;
				ReadPng(aImageData);
			}
			else
			{
                TImage NewImage = aImage;
                try
                {
                    if (aImage == null && aImageData != null)
                    {
                        NewImage = TImage.FromStream(aImageData);
                    }


                    //Convert to PNG / JPEG
                    if (defaultToJpg || NewImage.RawFormat.Equals(System.Drawing.Imaging.ImageFormat.Jpeg))
                        ReadJpeg(NewImage, aImageData);
                    else
                        ReadPng(NewImage);
                }
                finally
                {
                    if (NewImage != null && NewImage != aImage) NewImage.Dispose();
                }
			}          
		}

		private void ReadPng(TImage aImage)
		{
			using (MemoryStream Ms = new MemoryStream())
			{
				aImage.Save(Ms, System.Drawing.Imaging.ImageFormat.Png);
				Ms.Position=0;
				ReadPng(Ms);
			}
		}

		private void ReadJpeg(TImage aImage, Stream aImageData)
		{
			MemoryStream Msi= (aImageData as MemoryStream);
			if (Msi!=null) FImage = Msi.ToArray(); else FImage=null;
			FBitsPerComponent=8;
			FDecodeParmsName=null;
			FDecodeParmsSMask=null;
			if (aImage.PixelFormat == PixelFormat.Format8bppIndexed)
			{
				FColorSpace= TPdfTokens.GetString(TPdfToken.DeviceGrayName);  
			}
			else
			{
				FColorSpace= TPdfTokens.GetString(TPdfToken.DeviceRGBName);  //We don't care what the JPEG suggested colorspace is (in APPE/X'FFEE), we will show this on RGB. (Internally, JPEG Colorspace is always YCC)
			}
			FFilterName = TPdfTokens.GetString(TPdfToken.DCTDecodeName);
			FImageWidth = aImage.Width;
			FImageHeight = aImage.Height;
			if (FImage==null)
			{
				using (MemoryStream Ms = new MemoryStream())
				{
					aImage.Save(Ms, System.Drawing.Imaging.ImageFormat.Jpeg);
					FImage = Ms.ToArray();
				}
			}

			if (FTransparentColor!=~0L)
			{
				int r=(int)((FTransparentColor>>0) & 0xFF);
				int g=(int)((FTransparentColor>>8) &0xFF);
				int b=(int)((FTransparentColor>>16) &0xFF);
				FMask = 
					TPdfTokens.GetString(TPdfToken.OpenArray)+
					String.Format(CultureInfo.InvariantCulture, "{0} {1} {2} {3} {4} {5}", Math.Max(r-15,0), Math.Min(r+15,255), Math.Max(g-15,0), Math.Min(g+15,255), Math.Max(b-15,0), Math.Min(b+15,255) )+
					TPdfTokens.GetString(TPdfToken.CloseArray);
			}
		}


		private void ReadPng(Stream Ms)
		{
			using (MemoryStream OutMs = new MemoryStream())
			{
				TPdfPngData ImgParsedData = new TPdfPngData(OutMs);
				TPdfPng.ProcessPng(Ms, ImgParsedData);

				FImage = OutMs.ToArray();
				FImageWidth = (int)ImgParsedData.Width;
				FImageHeight = (int)ImgParsedData.Height;
				FBitsPerComponent = ImgParsedData.BitDepth;

				if (FBitsPerComponent==16) PdfMessages.ThrowException(PdfErr.ErrInvalidPngImage);

				int Colors = 1;
				switch (ImgParsedData.ColorType)
				{
					case 0: //GrayScale
						FColorSpace = TPdfTokens.GetString(TPdfToken.DeviceGrayName);
						if (ImgParsedData.tRNS!=null && ImgParsedData.tRNS.Length>=2)
						{
							FMask = 
								TPdfTokens.GetString(TPdfToken.OpenArray)+
								String.Format(CultureInfo.InvariantCulture, "{0} {0}", (ImgParsedData.tRNS[0]<<8)+ImgParsedData.tRNS[1])+
								TPdfTokens.GetString(TPdfToken.CloseArray);
						}
                        
						break;
					case 2:  //TrueColor
						FColorSpace = TPdfTokens.GetString(TPdfToken.DeviceRGBName);
						Colors=3; 
						if (ImgParsedData.tRNS!=null && ImgParsedData.tRNS.Length>=6)
						{
							FMask = 
								TPdfTokens.GetString(TPdfToken.OpenArray)+
								String.Format(CultureInfo.InvariantCulture, 
								"{0} {0} {1} {1} {2} {2}", (ImgParsedData.tRNS[0]<<8)+ImgParsedData.tRNS[1],
								(ImgParsedData.tRNS[2]<<8)+ImgParsedData.tRNS[3],
								(ImgParsedData.tRNS[4]<<8)+ImgParsedData.tRNS[5]
								)+
								TPdfTokens.GetString(TPdfToken.CloseArray);
						}
						break;
					case 3: //Indexed Color
						FColorSpace = GetPalette(ImgParsedData.PLTE);
						if (ImgParsedData.tRNS!=null && ImgParsedData.tRNS.Length>0)
						{
							ImgParsedData.SMask = TPdfPng.GetIndexedSMask(FImage, FImageWidth, FImageHeight, ImgParsedData.tRNS, FBitsPerComponent);
						}
						break;

					case 4:  //GrayScale + Alpha
						FColorSpace = TPdfTokens.GetString(TPdfToken.DeviceGrayName);
						break;

					case 6: //TrueColor + Alpha
						FColorSpace = TPdfTokens.GetString(TPdfToken.DeviceRGBName);
						Colors = 3;
						break;
					default: PdfMessages.ThrowException(PdfErr.ErrInvalidPngImage);break;
				}

				FSMask = ImgParsedData.SMask;
				OneBitMask = ImgParsedData.OneBitMask;

				if (ImgParsedData.InterlaceMethod!=0) 
					PdfMessages.ThrowException(PdfErr.ErrInvalidPngImage);

				FFilterName =  TPdfTokens.GetString(TPdfToken.FlateDecodeName);

				FDecodeParmsName = 
					TPdfTokens.GetString(TPdfToken.StartDictionary)+
					String.Format(CultureInfo.InvariantCulture, TPdfTokens.GetString(TPdfToken.PredictorName) + " {0} ", 15)+ 
					String.Format(CultureInfo.InvariantCulture, TPdfTokens.GetString(TPdfToken.ColorsName) + " {0} ", Colors)+ 
					String.Format(CultureInfo.InvariantCulture, TPdfTokens.GetString(TPdfToken.BitsPerComponentName) + " {0} ", FBitsPerComponent)+ 
					String.Format(CultureInfo.InvariantCulture, TPdfTokens.GetString(TPdfToken.ColumnsName) + " {0} ", FImageWidth)+ 
					TPdfTokens.GetString(TPdfToken.EndDictionary);

				FDecodeParmsSMask = 
					TPdfTokens.GetString(TPdfToken.StartDictionary)+
					String.Format(CultureInfo.InvariantCulture, TPdfTokens.GetString(TPdfToken.PredictorName) + " {0} ", 15)+ 
					String.Format(CultureInfo.InvariantCulture, TPdfTokens.GetString(TPdfToken.ColorsName) + " {0} ", 1)+ 
					String.Format(CultureInfo.InvariantCulture, TPdfTokens.GetString(TPdfToken.BitsPerComponentName) + " {0} ", 8)+ 
					String.Format(CultureInfo.InvariantCulture, TPdfTokens.GetString(TPdfToken.ColumnsName) + " {0} ", FImageWidth)+ 
					TPdfTokens.GetString(TPdfToken.EndDictionary);
			}
		}

		private static string GetPalette(byte[] PLTE)
		{
			string Result = 
				TPdfTokens.GetString(TPdfToken.OpenArray) +
				TPdfTokens.GetString(TPdfToken.IndexedName)+" "+
				TPdfTokens.GetString(TPdfToken.DeviceRGBName)+" "+
				String.Format(CultureInfo.InvariantCulture, "{0}", PLTE.Length /3 -1) +" ";


			StringBuilder Pal = new StringBuilder();
			for (int i=0;i<PLTE.Length;i++)
			{
				Pal.Append(String.Format("{0:X2}", PLTE[i]));
			}

			return Result+ "<" + Pal.ToString()+ ">"+
				TPdfTokens.GetString(TPdfToken.CloseArray);
		}

		public void WriteImage(TPdfStream DataStream, TXRefSection XRef)
		{
			TPdfBaseRecord.Write(DataStream, TPdfTokens.GetString(TPdfToken.ImgPrefix) + Id.ToString(CultureInfo.InvariantCulture) + " ");
			ImgObjId = XRef.GetNewObject(DataStream);
			TIndirectRecord.CallObj(DataStream, ImgObjId);
			if (FSMask!=null) 
			{
				SMaskId = XRef.GetNewObject(DataStream);
			}
		}

		public void Select(TPdfStream DataStream)
		{
			TPdfBaseRecord.WriteLine(DataStream, TPdfTokens.GetString(TPdfToken.ImgPrefix) + Id.ToString(CultureInfo.InvariantCulture) + " " +
				TPdfTokens.GetString(TPdfToken.CommandDo));
		}

		public void WriteImageObject(TPdfStream DataStream, TXRefSection XRef)
		{
			WriteImageOrMaskObject(DataStream, XRef, false, ImgObjId);
			if (FSMask!=null)
				WriteImageOrMaskObject(DataStream, XRef, true, SMaskId);
		}

		private void WriteImageOrMaskObject(TPdfStream DataStream, TXRefSection XRef, bool IsSMask, int ObjId)
		{
			XRef.SetObjectOffset(ObjId, DataStream);
			TIndirectRecord.SaveHeader(DataStream, ObjId);
			TDictionaryRecord.BeginDictionary(DataStream);
			TDictionaryRecord.SaveKey(DataStream, TPdfToken.TypeName, TPdfTokens.GetString(TPdfToken.XObjectName));
			TDictionaryRecord.SaveKey(DataStream, TPdfToken.SubtypeName, TPdfTokens.GetString(TPdfToken.ImageName));
			TDictionaryRecord.SaveKey(DataStream, TPdfToken.WidthName, FImageWidth);
			TDictionaryRecord.SaveKey(DataStream, TPdfToken.HeightName, FImageHeight);

			if (!IsSMask)
			{
				TDictionaryRecord.SaveKey(DataStream, TPdfToken.BitsPerComponentName, FBitsPerComponent);
				if (FSMask!=null)TDictionaryRecord.SaveKey(DataStream, MaskOrSMask, TIndirectRecord.GetCallObj(SMaskId));
				if (FMask!=null)TDictionaryRecord.SaveKey(DataStream, TPdfToken.MaskName, FMask);
				TDictionaryRecord.SaveKey(DataStream, TPdfToken.LengthName, (int)FImage.Length);
				TDictionaryRecord.SaveKey(DataStream, TPdfToken.ColorSpaceName, FColorSpace);  //This value is stored on the PropertyTagExifColorSpace property of the image for JPEG.
				if (FDecodeParmsName!=null)TDictionaryRecord.SaveKey(DataStream, TPdfToken.DecodeParmsName, FDecodeParmsName);
			}
			else
			{
				if (OneBitMask != null && FMask == null) SaveSMaskAsMask(DataStream);
				else
				{
					TDictionaryRecord.SaveKey(DataStream, TPdfToken.BitsPerComponentName, 8);
					TDictionaryRecord.SaveKey(DataStream, TPdfToken.LengthName, (int)FSMask.Length);
					TDictionaryRecord.SaveKey(DataStream, TPdfToken.ColorSpaceName, TPdfTokens.GetString(TPdfToken.DeviceGrayName));  //This value is stored on the PropertyTagExifColorSpace property of the image for JPEG.
					if (FDecodeParmsSMask!=null)TDictionaryRecord.SaveKey(DataStream, TPdfToken.DecodeParmsName, FDecodeParmsSMask);
				}
			}

			TDictionaryRecord.SaveKey(DataStream, TPdfToken.FilterName, FFilterName);


			TDictionaryRecord.EndDictionary(DataStream);

			TStreamRecord.BeginSave(DataStream);
			if (!IsSMask)
			{
				DataStream.Write(FImage);
			}
			else
			{
				if (OneBitMask != null) DataStream.Write(OneBitMask); else	DataStream.Write(FSMask);
			}
			TStreamRecord.EndSave(DataStream);

			TIndirectRecord.SaveTrailer(DataStream);
		}

		private TPdfToken MaskOrSMask
		{
			get
			{
				if (OneBitMask != null) return TPdfToken.MaskName; else return TPdfToken.SMaskName;
			}
		}

		//This is not only needed because it will be smaller, but also because when using tansparency (SMasks), Acrobat switches the color space, 
		//And things like antialiasing text might not look well in the screen. A Single object with SMask will turn all the document into a "Transparent"
		//Document. See "7.6.1 Color Spaces for Transparency Groups" in the pdf reference.
		//So we will try as much as possible to use Masks instead of SMasks.
		private void SaveSMaskAsMask(TPdfStream DataStream)
		{
			TDictionaryRecord.SaveKey(DataStream, TPdfToken.BitsPerComponentName, 1);
			TDictionaryRecord.SaveKey(DataStream, TPdfToken.LengthName, (int)OneBitMask.Length);
			TDictionaryRecord.SaveKey(DataStream, TPdfToken.ImageMaskName, TPdfTokens.GetString(TPdfToken.TrueText));  

		}

		#region IComparable Members
		public int CompareTo(object obj)
		{
			TPdfImage i2= (obj as TPdfImage);
			if (i2==null) return 1;
			int l = FImage.Length.CompareTo(i2.FImage.Length);
			if (l!=0)return l;
			for (int i = 0; i < FImage.Length; i++)
			{
				int z= FImage[i].CompareTo(i2.FImage[i]);
				if (z != 0) return z;
			}

			int tp = FTransparentColor.CompareTo(i2.FTransparentColor);
			if (tp!=0) return tp;
            
			if (FSMask==null)
				if (i2.FSMask !=null) return -1;
				else
					return 0;

			if (i2.FSMask ==null) return 1;

			l = FSMask.Length.CompareTo(i2.FSMask.Length);
			if (l!=0)return l;
			for (int i = 0; i < FSMask.Length; i++)
			{
				int z= FSMask[i].CompareTo(i2.FSMask[i]);
				if (z != 0) return z;
			}


			return 0;
		}

		#endregion
	}
}
