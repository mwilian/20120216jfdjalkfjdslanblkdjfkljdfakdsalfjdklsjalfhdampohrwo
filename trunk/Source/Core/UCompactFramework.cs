using System;
using System.IO;
using System.Runtime.CompilerServices;
using System.Reflection;
using System.Globalization;
#if (!SILVERLIGHT)
using System.Data;
#endif

#if (MONOTOUCH)
    using Color = MonoTouch.UIKit.UIColor;
    using Image = MonoTouch.UIKit.UIImage;
    using System.Drawing;
#else
  #if (WPF)
  using System.Windows.Controls;
  using System.Windows.Media;
  using System.Windows.Media.Imaging;
  #else
  using System.Drawing;
  using System.Drawing.Imaging;
  #endif
#endif


namespace FlexCel.Core
{
	/// <summary>
	/// Alternate implementations to Methods that do not exist on CF.
	/// </summary>
	internal sealed class TCompactFramework
	{
		private TCompactFramework(){}

		#region ImgConvert
#if (WPF)
        public static byte[] ImgConvert(byte[] data, ref TXlsImgType imgType)
        {  
            return data;
        }
#else
#if(!COMPACTFRAMEWORK)
        [MethodImpl(MethodImplOptions.NoInlining)]
        private static byte[] DoImgConvert(byte[] data, ref TXlsImgType imgType)
        {
            using (MemoryStream DataStream= new MemoryStream(data))
            {
                using (MemoryStream Result= new MemoryStream())
                {
                    using (Image Img = ImageExtender.FromStream(DataStream))
                    {
                        Img.Save(Result, ImageFormat.Png);  
                        imgType=TXlsImgType.Png;
                        return Result.ToArray();
                    }
                }
            }
        }

        /// <summary>
        /// No need for threadstatic.
        /// </summary>
        private static bool MissingFrameworkImageSave;

        public static byte[] ImgConvert(byte[] data, ref TXlsImgType imgType)
        {
            if (MissingFrameworkImageSave) return data;
            try
            {
                byte[] Result= DoImgConvert(data, ref imgType);
                return Result;
            }
            catch (MissingMethodException)
            {
                MissingFrameworkImageSave=true;
                //Nothing. It could be a bmp, and it is valid. (even when it would be better to convert it to a png)
            }
            catch (ArgumentException)
            {
                FlxMessages.ThrowException(FlxErr.ErrInvalidImage);
            }
            return data;
        }
#else
		public static byte[] ImgConvert(byte[] data, ref TXlsImgType imgType)
		{
			return data;
		}
#endif
#endif

		#endregion

		#region ImageFromStream
#if(!COMPACTFRAMEWORK)
        [MethodImpl(MethodImplOptions.NoInlining)]
        private static Image DoGetImage(Stream aStream)
        {
            return ImageExtender.FromStream(aStream);
        }
        
        /// <summary>
        /// No need for threadstatic.
        /// </summary>
        private static bool MissingFrameworkImageFromStream;

        public static Image GetImage(Stream aStream)
        {
            if (MissingFrameworkImageFromStream) return null;
            try
            {
                Image Result= DoGetImage(aStream);
                return Result;
            }
            catch (MissingMethodException)
            {
                MissingFrameworkImageFromStream=true;
                //Nothing. 
            }
            catch (ArgumentException)
            {
                FlxMessages.ThrowException(FlxErr.ErrInvalidImage);
            }
            return null;
        }
#else
		public static Image GetImage(Stream aStream)
		{
			return null;
		}
#endif

		#endregion

		#region EnumValues
		private static Array SlowEnumGetValues(Type enumType)
		{
			//Using Reflection on the type we get an array of FieldInfo objects which describe the contained fields. We have specified that we only want the Public and Static fields contain by the type.

			//get the public static fields (members of the enum)
			System.Reflection.FieldInfo[] fi = enumType.GetFields(

				BindingFlags.Static | BindingFlags.Public);

			// Now we have information on the Fields and their count we can create a new array to contain the values.

   
			//create a new enum array
			System.Enum[] values = new System.Enum[fi.Length];

			//We now loop through the FieldInfo collection. With each item we return the value and place it into the output array. Note the Value is the value as an enumeration member, not just the underlying numerical value.

			//populate with the values
			for(int iEnum = 0; iEnum < fi.Length; iEnum++)
			{
				values[iEnum] = (System.Enum)fi[iEnum].GetValue(null);
			}

			//Finally the array of enumeration values is returned.

			//return the array
			return values;
		}

#if(!COMPACTFRAMEWORK && !SILVERLIGHT) 
		[MethodImpl(MethodImplOptions.NoInlining)]
		private static Array TryEnumGetValues(Type enumType)
		{
			return Enum.GetValues(enumType);
		}

		/// <summary>
		/// No need for threadstatic.
		/// </summary>
		private static bool MissingFrameworkEnumValues;

		/// <summary>
		/// A <see cref="Enum.GetValues"/> that works on CF
		/// </summary>
		public static Array EnumGetValues(Type enumType)
		{
			if (MissingFrameworkEnumValues)
				return SlowEnumGetValues(enumType);

			try
			{
				Array Result = TryEnumGetValues(enumType);
				return Result;
			}
			catch (MissingMethodException)
			{
				MissingFrameworkEnumValues=true;
				return SlowEnumGetValues(enumType);
			}
		}
#else
		public static Array EnumGetValues(Type enumType)
		{
			return SlowEnumGetValues(enumType);
		}
#endif
		#endregion

		#region EnumNames
		private static string SlowEnumGetName(Type enumType, object value)
		{
			//Using Reflection on the type we get an array of FieldInfo objects which describe the contained fields. We have specified that we only want the Public and Static fields contain by the type.

			//get the public static fields (members of the enum)
			System.Reflection.FieldInfo[] fi = enumType.GetFields(BindingFlags.Static | BindingFlags.Public);

			for (int iEnum = 0; iEnum < fi.Length; iEnum++)
			{
				if ((Int32)fi[iEnum].GetValue(null) == (Int32)value)
					return fi[iEnum].Name;
			}

			return String.Empty;
		}

#if(!COMPACTFRAMEWORK) 
		/// <summary>
		/// A <see cref="Enum.GetName"/> that works on CF
		/// </summary>
		public static string EnumGetName(Type enumType, object value)
		{
          return Enum.GetName(enumType, value);
		}
#else
		public static string EnumGetName(Type enumType, object value)
		{
			return SlowEnumGetName(enumType, value);
		}
#endif
		#endregion

		#region NewLine
#if(!COMPACTFRAMEWORK)
        [MethodImpl(MethodImplOptions.NoInlining)]
        private static string TryEnvironmentNewLine()
        {
            return Environment.NewLine;
        }

        [ThreadStatic]
        private static string FNewLine; //STATIC*  It could be different on different threads. Remember, do not initialize threadstatic members
        /// <summary>
        /// A NewLine that works on CF.
        /// </summary>
        public static string NewLine
        {
            get
            {
                if (FNewLine!=null) return FNewLine;
                try
                {
                    FNewLine= TryEnvironmentNewLine();
                }
                catch (MissingMethodException)
                {
                    FNewLine="\r\n";
                }
                return FNewLine;
            }
        }
#else
		public static string NewLine
		{
			get
			{
				return "\r\n";
			}
		}
#endif
		#endregion

		#region MD5
		private static byte[] GetInternalMD5Hash(byte[] Data)
		{
			byte[] hash=null;
			using(MD5 MD5Alg = MD5CryptoServiceProvider.Create("MD5")) 
			{
				hash = MD5Alg.ComputeHash(Data);
			}
			return hash;
		}

#if(!COMPACTFRAMEWORK && !SILVERLIGHT)
        [MethodImpl(MethodImplOptions.NoInlining)]
        private static byte[] GetNativeMD5Hash(byte[] Data)
        {
            byte[] hash=null;
            using(System.Security.Cryptography.MD5 MD5Alg = System.Security.Cryptography.MD5CryptoServiceProvider.Create("MD5")) 
            {
                hash = MD5Alg.ComputeHash(Data);
            }
            return hash;
        }

        /// <summary>
        /// No need for threadstatic.
        /// </summary>
        private static bool MissingFrameworkMD5;

        public static byte[] GetBSEHeader(byte[] Data, int TBSEHeaderLength, int TBSEHeaderRgbUid)
        {
            byte[] hash= null;
            if (MissingFrameworkMD5) 
                hash = GetInternalMD5Hash(Data);
            else
            {
                try
                {
                    hash=GetNativeMD5Hash(Data); //Not supported on CF!
                }
                catch (TypeLoadException)
                {
                    hash=GetInternalMD5Hash(Data);
                    MissingFrameworkMD5=true;
                }
                catch (MissingMethodException)
                {
                    hash=GetInternalMD5Hash(Data);
                    MissingFrameworkMD5=true;
                }
            }

            byte[] BSEHeader=new byte[TBSEHeaderLength];
            hash.CopyTo(BSEHeader, TBSEHeaderRgbUid);

            return BSEHeader;
        }

        public static byte[] GetMD5Hash(byte[] Data)
        {
            byte[] hash= null;
            if (MissingFrameworkMD5) 
                hash = GetInternalMD5Hash(Data);
            else
            {
                try
                {
                    hash=GetNativeMD5Hash(Data); //Not supported on CF!
                }
                catch (TypeLoadException)
                {
                    hash=GetInternalMD5Hash(Data);
                    MissingFrameworkMD5=true;
                }
                catch (MissingMethodException)
                {
                    hash=GetInternalMD5Hash(Data);
                    MissingFrameworkMD5=true;
                }
            }

            return hash;
        }

#else
		public static byte[] GetBSEHeader(byte[] Data, int TBSEHeaderLength, int TBSEHeaderRgbUid)
		{
			byte[] hash = GetInternalMD5Hash(Data);
			byte[] BSEHeader=new byte[TBSEHeaderLength];
			hash.CopyTo(BSEHeader, TBSEHeaderRgbUid);

			return BSEHeader;
		}

		public static byte[] GetMD5Hash(byte[] Data)
		{
			byte[] hash = GetInternalMD5Hash(Data);
			return hash;
		}
#endif
		#endregion

		#region TryParse


		/// <summary>
		/// Slower implementation to run on CF.
		/// </summary>
        private static bool CFConvertToNumber(string value, out double Result, CultureInfo Culture)
        {
            Result = 0;
            try
            {
                Result = Double.Parse(value, NumberStyles.Any, Culture); // Note that NumberStyles.Any does not include NumberStyles.Hexadecimal, and that is ok.
                return true;
            }
            catch (FormatException)
            {
            }
            catch (OverflowException)
            {
            }
            catch (ArgumentNullException)
            {
            }
            return false;
        }

        public static bool ConvertToNumber(string aValue, CultureInfo Culture, out double Result)
        {
            TNumberFormat tmp;
            return ConvertToNumber(aValue, Culture, out Result, out tmp);
        }

#if(!COMPACTFRAMEWORK)
        [MethodImpl(MethodImplOptions.NoInlining)] //we need this method not to be inlined so cf can throw a missingexception method.
        private static bool NormalConvertToNumber(string value, out double Result, CultureInfo Culture)
        {
            return (Double.TryParse(value, NumberStyles.Any, Culture, out Result)); // Note that NumberStyles.Any does not include NumberStyles.Hexadecimal, and that is ok.
        }

        /// <summary>
        /// No need for threadstatic.
        /// </summary>
        private static bool MissingFrameworkDoubleTryParse;

        public static bool ConvertToNumber(string aValue, CultureInfo Culture, out double Result, out TNumberFormat NumberFormat)
        {
            string sValue;
            NumberFormat = new TNumberFormat(aValue, Culture, out sValue);

            if (MissingFrameworkDoubleTryParse)
            {
                bool Converted = CFConvertToNumber(sValue, out Result, Culture);
                if (Converted && NumberFormat.HasPercent) Result /= 100;
                return Converted;
            }

            bool ConversionOk = false;
            Result = 0;

            try
            {
                ConversionOk = NormalConvertToNumber(sValue, out Result, Culture);
            }
            catch (MissingMethodException)
            {
                MissingFrameworkDoubleTryParse = true;
            }

            if (MissingFrameworkDoubleTryParse) ConversionOk = CFConvertToNumber(sValue, out Result, Culture);

            if (ConversionOk && NumberFormat.HasPercent) Result /= 100;

            return ConversionOk;
        }
#else
        public static bool ConvertToNumber(string aValue, CultureInfo Culture, out double Result, out TNumberFormat NumberFormat)
		{
            string sValue;
            NumberFormat = new TNumberFormat(aValue, Culture, out sValue);
            bool Converted = CFConvertToNumber(sValue, out Result, Culture);
            if (Converted && NumberFormat.HasPercent) Result /= 100;
            return Converted;
        }
#endif
		#endregion

		#region ConvertToDate
        public static bool ConvertDateToNumber(string sValue, out DateTime DateResult)
        {
            return ConvertDateToNumber(sValue, null, out DateResult);
        }

        public static bool ConvertDateToNumber(string sValue, string[] DateFormats, out DateTime DateResult)
		{
			//Until we get a DateTime.TryParse... THIS IS *REALLY* SLOW because of exceptions. On 2.0 we should use Tryparse
#if(!FRAMEWORK20 || COMPACTFRAMEWORK)
			bool IsDateTime=false;
			DateResult= new DateTime(1,1,1);
			if ((sValue!=null)&&(sValue.Length>2) && 
				sValue[0]>='0' && sValue[0]<='9'
				//&& 
				//(sValue.IndexOf(CultureInfo.CurrentCulture.DateTimeFormat.DateSeparator) > 0 || 
				//sValue.IndexOf(CultureInfo.CurrentCulture.DateTimeFormat.TimeSeparator) > 0)
				)  //this is a hack, but allows us to not check many strings, and speed up a lot. As numbers are already checked, we will only check for a little portions of the strings.
				//Bad for strings like "july 15, 2004", but it is a price we have to pay. On framework20 there is no problem.
			{
				try
				{
					DateResult=DateTime.Parse(sValue, CultureInfo.CurrentCulture, DateTimeStyles.NoCurrentDateDefault);
					IsDateTime=true;
				}
				catch (FormatException)
				{
					IsDateTime=false;
				}
			}
#else
            DateResult=DateTime.MinValue;
            bool IsDateTime = false;
            if (DateFormats == null)
            {
                IsDateTime = (sValue != null) && (sValue.Length > 0) && DateTime.TryParse(sValue, CultureInfo.CurrentCulture, DateTimeStyles.NoCurrentDateDefault, out DateResult);
            }
            else
            {
                IsDateTime = (sValue != null) && (sValue.Length > 0) && DateTime.TryParseExact(sValue, DateFormats, CultureInfo.CurrentCulture, DateTimeStyles.NoCurrentDateDefault, out DateResult);               
            }

            double dt;
            if (IsDateTime && (!FlxDateTime.TryToOADate(DateResult, false, out dt) || !FlxDateTime.TryToOADate(DateResult, true, out dt))) IsDateTime = false;
#endif

			return IsDateTime;
		}
		#endregion

#if (!SILVERLIGHT)
		#region Dispose tables
#if(!COMPACTFRAMEWORK   || FRAMEWORK20)
		/// <summary>
		/// No need for threadstatic.
		/// </summary>
		private static bool MissingFrameworkDisposeTables;

		[MethodImpl(MethodImplOptions.NoInlining)] //we need this method not to be inlined so cf can throw a missingexception method.
		private static void NormalDisposeTable(DataTable table)
		{
			table.Dispose();
		}

		public static void DisposeDataTable(DataTable table)
		{
			try
			{
				if (!MissingFrameworkDisposeTables) NormalDisposeTable(table);
			}
			catch (MissingMethodException)
			{
				MissingFrameworkDisposeTables=true;
				//Nothing. 
			}
		}

		[MethodImpl(MethodImplOptions.NoInlining)] //we need this method not to be inlined so cf can throw a missingexception method.
		private static void NormalDisposeDataView(DataView dv)
		{
			dv.Dispose();
		}

		public static void DisposeDataView(DataView dv)
		{
			try
			{
				if (!MissingFrameworkDisposeTables) NormalDisposeDataView(dv);
			}
			catch (MissingMethodException)
			{
				MissingFrameworkDisposeTables=true;
				//Nothing. 
			}
		}

#else
		public static void DisposeDataTable(DataTable table)
		{
		}

		public static void DisposeDataView(DataView table)
		{
		}
#endif
		#endregion
#endif

		#region Decimal OACurrency
		internal static object DecimalFromOACurrency(long p)
		{
#if (COMPACTFRAMEWORK || SILVERLIGHT)
			return null;
#else
            try
            {
                return Decimal.FromOACurrency(p);
            }
            catch (MissingMethodException)
            {
                return null;
            }
#endif            
		}
		#endregion


        internal static bool TryParse(string p, out int i)
        {
#if (FRAMEWORK20 && !COMPACTFRAMEWORK)
            return Int32.TryParse(p, NumberStyles.Any, CultureInfo.InvariantCulture, out i);
#else
			try
			{
				i =Int32.Parse(p);
			}
			catch
			{
				i = 0;
				return false;
			}
			return true;
#endif
        }

        internal static bool TryParse(string p, out long i)
        {
#if (FRAMEWORK20 && !COMPACTFRAMEWORK)
            return long.TryParse(p, NumberStyles.Any, CultureInfo.InvariantCulture, out i);
#else
			try
			{
				i =long.Parse(p);
			}
			catch
			{
				i = 0;
				return false;
			}
			return true;
#endif
        }

        internal static int HexToNumber(string number)
        {
            return int.Parse(number, NumberStyles.HexNumber, CultureInfo.InvariantCulture);
        }
    }

	#region SilverLight
#if (SILVERLIGHT)
    /// <summary>
    /// This is a "Fake" SerializableAttribute just to compile in SilverLight. Serializing objects is not supported in SilverLight.
    /// </summary>
	public class SerializableAttribute: Attribute
	{
	}

    /// <summary>
    /// ICloneable was dropped, see http://silverlight.net/forums/p/11495/36821.aspx#36821
    /// </summary>
    internal interface ICloneable 
    {
        object Clone();
    }
#endif
	#endregion

    internal struct TNumberFormat
    {
        internal bool HasPercent;
        internal bool HasCurrency;
        internal bool HasExp;

        internal TNumberFormat(string aValue, CultureInfo Culture, out string sValue)
        {
            HasPercent = false;
            HasCurrency = false;
            HasExp = false;

#if (!COMPACTFRAMEWORK && FRAMEWORK20)
            HasExp = aValue.ToUpperInvariant().IndexOf("E") > 0;
#else
            HasExp = aValue.ToUpper(CultureInfo.InvariantCulture).IndexOf("E") > 0;
#endif
            CheckPercent(aValue, out sValue);
            if (HasPercent) return;

            CheckCurrency(sValue, Culture.NumberFormat);
        }


        private void CheckPercent(string aValue, out string sValue)
        {
            string TrimValue = aValue.Trim();
            if (TrimValue.StartsWith("%"))
            {
                HasPercent = true;
                sValue = TrimValue.Substring(1);
                return;
            }

            if (TrimValue.EndsWith("%"))
            {
                HasPercent = true;
                sValue = TrimValue.Substring(0, TrimValue.Length - 1);
                return;
            }

            sValue = aValue;
            HasPercent = false;
        }

        private void CheckCurrency(string aValue, NumberFormatInfo NumberFormat)
        {
            HasCurrency = false;
            if (NumberFormat.CurrencySymbol == null) return;

            string TrimValue = aValue.Trim();
            int CurrIndex = TrimValue.IndexOf(NumberFormat.CurrencySymbol);

            if (CurrIndex >= 0)
            {
                //look for the "nearest side"
                if (CurrIndex < 3)
                {
                    bool Ok = true;
                    for (int i = 0; i < CurrIndex; i++)
                    {
                        if (TrimValue[i] != '-' && TrimValue[i] != '(')
                        {
                            Ok = false;
                            break;
                        }
                    }

                    if (Ok)
                    {
                        HasCurrency = true;
                        return;
                    }
                }

                if (CurrIndex > aValue.Length - 3 - NumberFormat.CurrencySymbol.Length)
                {
                    bool Ok = true;
                    for (int i = CurrIndex; i < aValue.Length - NumberFormat.CurrencySymbol.Length; i++)
                    {
                        if (TrimValue[i] != '-' && TrimValue[i] != ')')
                        {
                            Ok = false;
                            break;
                        }
                    }

                    if (Ok)
                    {
                        HasCurrency = true;
                        return;
                    }
                }

            }

        }

    }

#if(!FRAMEWORK30 || COMPACTFRAMEWORK)
internal delegate TResult Func<T, TResult>(T arg);
#endif

}
