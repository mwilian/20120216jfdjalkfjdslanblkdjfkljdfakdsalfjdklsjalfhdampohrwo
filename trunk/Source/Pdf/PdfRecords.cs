#region Using directives

using System;
using System.IO;
using System.Text;
using System.Globalization;
using FlexCel.Core;

#if (WPF)
using real = System.Double;
#else
using real = System.Single;
#endif

#endregion

namespace FlexCel.Pdf
{
    /// <summary>
    /// Base for all PDF objects.
    /// </summary>
    internal abstract class TPdfBaseRecord
    {
        internal static readonly Encoding Coder = Encoding.GetEncoding(1252); //WinAnsiEncoding
        protected TPdfBaseRecord() { }

        public static void Write(TPdfStream DataStream, string Text)
        {
            byte[] Buffer = Coder.GetBytes(Text);
            DataStream.Write(Buffer);
        }

        public static void WriteLine(TPdfStream DataStream, string Text)
        {
            Write(DataStream, Text+ TPdfTokens.NewLine);
        }

        public static void UnicodeWrite(TPdfStream DataStream, string Text)
        {
            bool NeedsUnicode = false;
            for (int i = 0; i < Text.Length; i++)
            {
                if ((int)Text[i] > 128) NeedsUnicode = true;
            }
            UnicodeWrite(DataStream, Text, NeedsUnicode, "\uFEFF");

        }

		public static void UnicodeWrite(TPdfStream DataStream, string Text, bool NeedsUnicode, string UnicodePrefix)
		{
			byte[] Buffer = null;
			if (NeedsUnicode)
				Buffer = Encoding.BigEndianUnicode.GetBytes(UnicodePrefix + Text);
			else
				Buffer = Coder.GetBytes(Text);
			Byte[] Buff2 = TPdfStringRecord.EscapeString(Buffer);
			DataStream.Write(Buff2);
		}
	
		public static void UnicodeWrite(TPdfStream DataStream, string Text, TPdfFont aFont)
		{
			byte[] Buffer = null;
			Buffer = aFont.EncodeString(Text);
			Byte[] Buff2 = TPdfStringRecord.EscapeString(Buffer);
			DataStream.Write(Buff2);
		}
	}

    /// <summary>
    /// The PDF Header.
    /// It must be followed with a comment with non alphanumeric characters.
    /// </summary>
    internal sealed class TPdfHeaderRecord: TPdfBaseRecord
    {
		private TPdfHeaderRecord(){}

        public static void SaveToStream(TPdfStream DataStream, string aText)
        {
            WriteLine(DataStream, aText); 
        }


    }

    /// <summary>
    /// An ASCII string.
    /// Syntax:
    ///          (this is a \(string\))  
    /// "(", ")" and "\" should be escaped. Strings cannot be longer than 65535
    /// </summary>
    internal sealed class TPdfStringRecord : TPdfBaseRecord
    {
        private TPdfStringRecord(){}

        /// <summary>
        /// Text should be escaped AFTER it has been converted to bytes.
        /// For example, unicode "\" would be written as "\00" and it should be escaped as "\\00" and not as "\00\00"
        /// </summary>
        /// <param name="aText"></param>
        /// <returns></returns>
        public static byte[] EscapeString(byte[] aText)
        {
            using (MemoryStream Ms = new MemoryStream((int)(aText.Length * 1.2)))
            {
                byte os = (byte)TPdfTokens.GetString(TPdfToken.OpenString)[0];
                byte cs = (byte)TPdfTokens.GetString(TPdfToken.CloseString)[0];
                byte es = (byte)TPdfTokens.GetString(TPdfToken.EscapeString)[0];
                for (int i = 0; i < aText.Length; i++)
                {
                    byte c = aText[i];

					switch (c)
					{
						case 0x0A: //line feed. 
							Ms.WriteByte(es);
							Ms.WriteByte((byte)'n');
							break;

						case 0x0C: //form feed
							Ms.WriteByte(es);
							Ms.WriteByte((byte)'f');
							break;

						case 0x0D: //Carriage Return.   reference says it is not needed, but that is false.
							Ms.WriteByte(es);
							Ms.WriteByte((byte)'r');
							break;

						case 0x09: //tab
							Ms.WriteByte(es);
							Ms.WriteByte((byte)'t');
							break;

						default:
							if (c == os || c == cs || c == es)
								Ms.WriteByte(es);
							Ms.WriteByte(c);
							break;
					}
                }
                return Ms.ToArray();
            }
        }

		private static void WriteSimpleString(TPdfStream DataStream, string Text, TPdfFont aFont, string EndNewText)
		{
			if (Text == null || Text.Length == 0) return;
			Write(DataStream, TPdfTokens.GetString(TPdfToken.OpenString));
			UnicodeWrite(DataStream, Text, aFont);
			Write(DataStream, TPdfTokens.GetString(TPdfToken.CloseString));
			TPdfBaseRecord.WriteLine(DataStream, EndNewText);
		}

		public static void WriteStringInStream(TPdfStream DataStream, string Text, real FontSize, TPdfFont aFont, ref string LastFont, string EndNewText, string StartNewText2, string EndNewText2, TTracedFonts TracedFonts)
		{
			TPdfFont LastFallbackFont = aFont;
			int StartText = 0;

			int TLen = Text.Length;
			for (int i = 0; i <= TLen; i++)
			{
				TPdfFont FallbackFont = null;
				
				if (i < TLen) 
				{
					FallbackFont = aFont.Fallback(Text[i], 0);
					if (FallbackFont == null) FallbackFont = aFont;
				}

				if (FallbackFont != LastFallbackFont) 
				{
					WriteSimpleString(DataStream, Text.Substring(StartText, i - StartText), LastFallbackFont, EndNewText);

					StartText = i;

					if (FallbackFont != null)
					{
						TPdfBaseRecord.Write(DataStream, EndNewText2);
						FallbackFont.Select(DataStream, FontSize, ref LastFont);
						
						if (FlexCelTrace.HasListeners && FallbackFont != aFont && !TracedFonts.ContainsKey(aFont.FontName+ ";" + FallbackFont.FontName))
						{
							TracedFonts.Add(aFont.FontName+ ";" + FallbackFont.FontName, String.Empty);
							FlexCelTrace.Write(new TPdfUsedFallbackFontError(FlxMessages.GetString(FlxErr.ErrUsedFallbackFont, aFont.FontName, FallbackFont.FontName), aFont.FontName, FallbackFont.FontName ));
						}

						TPdfBaseRecord.Write(DataStream, StartNewText2);
						LastFallbackFont = FallbackFont;
					}
				}

			}
		}
    }

    /// <summary>
    /// A dictionary List.
    /// </summary>
    internal class TDictionaryRecord : TPdfBaseRecord
    {
        public static void BeginDictionary(TPdfStream DataStream)
        {
            WriteLine(DataStream, TPdfTokens.GetString(TPdfToken.StartDictionary));
        }

        public static void EndDictionary(TPdfStream DataStream)
        {
            WriteLine(DataStream, TPdfTokens.GetString(TPdfToken.EndDictionary));
        }

        public static void SaveKey(TPdfStream DataStream, TPdfToken Tk, string Value)
        {
            WriteLine(DataStream, String.Format("{0} {1}", TPdfTokens.GetString(Tk), Value));
        }
        
        public static void SaveUnicodeKey(TPdfStream DataStream, TPdfToken Tk, string Value)
        {
            Write(DataStream, TPdfTokens.GetString(Tk)+" "+TPdfTokens.GetString(TPdfToken.OpenString));
            UnicodeWrite(DataStream, Value);
            WriteLine(DataStream, TPdfTokens.GetString(TPdfToken.CloseString));
        }

        public static void SaveKey(TPdfStream DataStream, TPdfToken Tk, int Value)
        {
            SaveKey(DataStream, Tk, PdfConv.LongToString(Value));
        }

    }


    /// <summary>
    /// A stream with text or images.
    /// Syntax:
    ///         &lt;&lt; /Length 534
    ///         Filter [/ASCII85Decode /LZWDecode]
    ///         &gt;&gt;
    ///         stream
    ///         ....
    ///         endstream
    /// </summary>
    internal sealed class TStreamRecord : TPdfBaseRecord
    {
		private TStreamRecord(){}

		public static void BeginSave(TPdfStream DataStream, int LengthId, bool Compress)
		{
			BeginSave(DataStream, LengthId, Compress, -1);
		}

        public static void BeginSave(TPdfStream DataStream, int LengthId, bool Compress, int Length1)
        {
            TDictionaryRecord.BeginDictionary(DataStream);
            Write(DataStream, TPdfTokens.GetString(TPdfToken.LengthName) + " ");
            TIndirectRecord.CallObj(DataStream, LengthId);

			if (Length1>=0)
				WriteLine(DataStream, TPdfTokens.GetString(TPdfToken.Length1Name) + " "+ Length1.ToString(CultureInfo.InvariantCulture));

            if (Compress) SetFlateDecode(DataStream);
            TDictionaryRecord.EndDictionary(DataStream);
            BeginSave(DataStream);
        }

        public static void SetFlateDecode(TPdfStream DataStream)
        {
            TDictionaryRecord.SaveKey(DataStream, TPdfToken.FilterName,
            TPdfTokens.GetString(TPdfToken.OpenArray) +
            TPdfTokens.GetString(TPdfToken.FlateDecodeName) +
            TPdfTokens.GetString(TPdfToken.CloseArray));
        }

        public static void BeginSave(TPdfStream DataStream)
        {
            WriteLine(DataStream, TPdfTokens.GetString(TPdfToken.Stream));
        }

        public static void EndSave(TPdfStream DataStream)
        {
            WriteLine(DataStream, TPdfTokens.GetString(TPdfToken.EndStream));
        }

    }

    /// <summary>
    /// A reference to other record.
    /// Syntax: 
    ///         id id2 obj
    ///         ....
    ///         endobj
    /// 
    /// Referenced with id id2 R
    /// </summary>
    internal sealed class TIndirectRecord : TPdfBaseRecord
    {
		private TIndirectRecord(){}
        public static void SaveHeader(TPdfStream DataStream, int Id)
        {
            WriteLine(DataStream, String.Format(TPdfTokens.GetString(TPdfToken.Obj), Id));
        }

        public static void SaveTrailer(TPdfStream DataStream)
        {
            WriteLine(DataStream, TPdfTokens.GetString(TPdfToken.EndObj));
        }

        public static void CallObj(TPdfStream DataStream, int Id)
        {
            WriteLine(DataStream, String.Format(TPdfTokens.GetString(TPdfToken.CallObj), Id));
        }

        public static string GetCallObj(int Id)
        {
            return String.Format(CultureInfo.InvariantCulture, TPdfTokens.GetString(TPdfToken.CallObj), Id);
        }

    }

    /// <summary>
    /// A date
    /// </summary>
    internal sealed class TDateRecord
    {
		private TDateRecord(){}

        internal static string GetDate(DateTime d)
        {
            return d.ToString("(D:yyyyMMddHHmmsszz\\'00\\')", CultureInfo.InvariantCulture);
        }
    }
}
