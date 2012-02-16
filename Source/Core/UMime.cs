using System;
using System.Text;
using System.Globalization;
using System.IO;

using System.Collections.Generic;

namespace FlexCel.Core
{
	/// <summary>
	/// Different ways to define a multipart arcvhive.
	/// </summary>
	public enum TMultipartType
	{
		/// <summary>
		/// Parts inside this container are related, for example this message could contain an HTML file and its external images.
		/// </summary>
		Related,

		/// <summary>
		/// Parts inside this container ara an alternative one form the other. For example, a message could be sent in HTML and plain Text, once inside a
		/// different MIME part, and the mail reader should chose the best alternative of the parts to display.
		/// </summary>
		Alternative,

		/// <summary>
		/// Parts inside this MIME container are not related, e.g. different attachments.
		/// </summary>
		Mixed,

		/// <summary>
		/// A compilation of messages, used in forwarded emails.
		/// </summary>
		Digest,

		/// <summary>
		/// A part containing files and one signature.
		/// </summary>
		Signed
	}

	/// <summary>
	/// Defines how a part of a MIME message will be coded.
	/// </summary>
	public enum TContentTransferEncoding
	{
		/// <summary>
		/// Use the Quoted Printable algorithm (RFC 2045 section 6.7).
		/// You will normally use this encoding for text.
        /// When using this option, you need to write the part content using <see cref="TMimeWriter.WriteQuotedPrintable"/> 
		/// or a similar method.
		/// </summary>
		QuotedPrintable,

		/// <summary>
		/// Use base64 algorithm (RFC 2045 section 6.7).
		/// You would normally use this encoding for binary files.
		/// When using this option, you need to write the part content using <see cref="TMimeWriter.WriteBase64"/> 
		/// or a similar method.
		/// </summary>
		Base64
	}

    /// <summary>
    /// Defines how a string returned by Q-Encode will be handled.
    /// </summary>
    public enum TQEncodeMetaInfo
    {
        /// <summary>
        /// This will return the raw Q-Encoded string, without "=?charset?encoding?" at the beginning or "?=" at the end.
        /// </summary>
        None,

        /// <summary>
        /// This will always return the Q-Encoded string with "=?charset?encoding?" at the beginning and "?=" at the end.
        /// </summary>
        Always,

        /// <summary>
        /// This will return the string with "=?charset?encoding?" only if the original string has special characters that need encoding,  if not it will return the original string without encoding.
        /// </summary>
        OnlyIfNeeded
    }

	/// <summary>
	/// A simple class used to create MIME formatted messages. While it does not provide much functionality, it gives enough to create simple multipart archives.
	/// </summary>
	public class TMimeWriter
	{
		#region Privates
		private string Boundary;
		private int QuotedPrintablePosition;
		#endregion

		#region Mime Headers
		/// <summary>
		/// Creates the headers for a multipart MIME file. This must be the first method to call in order to create a MIME file.
		/// After this, you need to call <see cref="AddPartHeader"/> and start adding the parts of the message, and you <b>always</b>
		/// need to end the message by calling <see cref="EndMultiPartMessage"/>.
		/// </summary>
		/// <param name="message"></param>
		/// <param name="multipartType">Type of multipart for this file.</param>
		/// <param name="contentType">Type of the header as defined in the MIME standard, e.g. "text/plain", "text/html", etc. This is the type of the main part on a related message. Set it to null if there is no main part.</param>
		/// <param name="contentLocation">The location for the whole mime file. null if you do not want to set a location. for this to work in ie/opera, etc, this should be something like "file:///filename.ext"</param>
		public void CreateMultiPartMessage(TextWriter message, TMultipartType multipartType, string contentType, Uri contentLocation)
		{
            //IsMailNewsSave does not seem to be working in mono 1.2.5.1
			if (!FlxUtils.IsMonoRunning() && !message.Encoding.IsMailNewsSave) FlxMessages.ThrowException(FlxErr.ErrInvalidEncodingForMIME, message.Encoding.EncodingName);

			WriteLine(message, "MIME-Version: 1.0");

			//Encoding on the contentLocation must be done ONLY in the header.
			//QEncode has problems with OPERA 9, it will not recognize it.
			//if (contentLocation != null) WriteLine(message, "Content-Location: " + QEncode(contentLocation, TQEncodeMetaInfo.OnlyIfNeeded));
			
			if (contentLocation != null) WriteLine(message, "Content-Location: " + contentLocation.AbsoluteUri); //Use AbsoluteUri instead of ToString() so unicode characters are escaped.

			string MType = "related";
			switch (multipartType)
			{
				case TMultipartType.Alternative: MType = "alternative"; break;
				case TMultipartType.Mixed: MType = "mixed"; break;
				case TMultipartType.Digest: MType = "digest"; break;
				case TMultipartType.Signed: MType = "signed"; break;
			}

			WriteLine(message, "Content-Type: multipart/" + MType + ";");
			if (contentType != null) WriteLine(message, "\ttype = \"" + contentType + "\";");
			Boundary = GetBoundary();
			WriteLine(message, "\tboundary = \"" + Boundary + "\"");

			WriteLine(message);
			WriteLine(message, "This is a multi-part message in MIME format.");
			WriteLine(message);
			Write(message, "--" + Boundary);

		}

		/// <summary>
		/// Adds the header for a part in a multipart Mime message. After calling this method, you need to write your data content into
		/// the TextWriter using <see cref="WriteQuotedPrintable"/> or <see cref="WriteBase64"/> and after that always call <see cref="EndPart"/>.
		/// </summary>
		/// <param name="message">TextWriter where you are saving the message.</param>
		/// <param name="contentType">Type of the header as defined in the MIME standard, e.g. "text/plain", "text/html", etc.</param>
		/// <param name="contentTransferEncoding">How the part will be codified. Use base64 for binarty TextWriter and quotedprintable for text.</param>
		/// <param name="contentLocation">The location for this resource. null if you do not want to set a location.</param>
		/// <param name="contentId">Content id of the resource, if you want to use it in cid: links. Null if you do not want to specify a content-id. Note that this string will not be escaped by the framework, so make sure it 
		/// contains valid characters. As it needs to be globally unique, normally a GUID can be a good option here.</param>
		/// <param name="encodingName">Name for the encoding on this part. Null means do not write an encoding. (for example in binary parts)</param>
		public void AddPartHeader(TextWriter message, string contentType, TContentTransferEncoding contentTransferEncoding, Uri contentLocation, string contentId, string encodingName)
		{
			WriteLine(message); //to move down from the last boundary.
			string encoding = encodingName == null? String.Empty: "; charset=" + encodingName;
			WriteLine(message, "Content-Type: " + contentType + encoding);
			string cte = "quoted-printable";
			if (contentTransferEncoding == TContentTransferEncoding.Base64) cte = "base64";

			WriteLine(message, "Content-Transfer-Encoding: " + cte);

			//Encoding on the contentLocation must be done ONLY in the header.
			//QEncode has problems with OPERA 9, it will not recognize it.
			//if (contentLocation != null) WriteLine(message, "Content-Location: " + QEncode(contentLocation, TQEncodeMetaInfo.OnlyIfNeeded));
			
			if (contentLocation != null) WriteLine(message, "Content-Location: " + contentLocation.AbsoluteUri); //Use AbsoluteUri instead of ToString() so unicode characters are escaped.
			if (contentId != null) WriteLine(message, "Content-Id: <" + contentId + ">");
			
			WriteLine(message);

			QuotedPrintablePosition = 0;
		}

        /// <summary>
        /// Ends a MIME part started with <see cref="AddPartHeader"/>.
        /// </summary>
        /// <param name="message">TextWriter where you are saving the message.</param>
        public void EndPart(TextWriter message)
		{
            WriteLine(message);
			Write(message, "--" + Boundary);
		}

	    /// <summary>
	    /// Call this method after the last call to EndPart, to finish the MIME message.
	    /// </summary>
	    /// <param name="message"></param>
		public void EndMultiPartMessage(TextWriter message)
		{
			WriteLine(message, "--");
		}

		#endregion

		#region QEncode

		/// <summary>
		/// Returns the Q-encode of a string, used in the MIME Headers.  //RFC 2047
		/// </summary>
		/// <param name="s">String to encode</param>
		/// <param name="addMetaInfo">Defines if to add the string  "=?charset?encoding?" will be appended at the begining, and "?=" at the end.</param>
		/// <returns></returns>
		public static string QEncode(string s, TQEncodeMetaInfo addMetaInfo)	
		{
			StringBuilder Result = new StringBuilder();
			
			int i = -1;
			while (i < s.Length - 1)
			{
				i++;

				byte b = (byte)s[i];

				if (b >= 33 && b <= 60) 
				{
					Result.Append((char)b);
					continue;
				}

				if (b != 63 && b!= 95 && b >= 62 && b <= 126)  //? and _ are not valid in Q-encoding.
				{
					Result.Append((char)b);
					continue;
				}
				
				if (b == 0x20)
				{
					Result.Append('_');
					continue;
				}

				byte d1 = (byte)(b/16); byte d1Base = d1<10? (byte)'0': (byte)((byte)'A' - 10);
				byte d2 = (byte)(b%16); byte d2Base = d2<10? (byte)'0': (byte)((byte)'A' - 10);
				Result.Append('=');
				Result.Append((char)(d1Base + d1));
				Result.Append((char)(d2Base + d2));
			}

			if (addMetaInfo == TQEncodeMetaInfo.Always || (addMetaInfo == TQEncodeMetaInfo.OnlyIfNeeded && Result.ToString() != s))
			{
				Result.Insert(0, "=?utf-8?Q?");
				Result.Append("?=");
			}

			return Result.ToString();
		}
		#endregion

		#region Quoted Printable encoding

		/// <summary>
		/// Writes the Quoted Printable encoding of a string, as defined in RFC 2045 section 6.7
		/// This method keeps state and breaks the line  every time it is longer than 76 characters. 
		/// The state is reset each time <see cref="AddPartHeader"/> is called.
		/// </summary>
		/// <param name="Data">TextWriter where we will write the data.</param>
		/// <param name="s">String to Encode.</param>
		public void WriteQuotedPrintable(TextWriter Data, string s)	
		{
			byte[] encoded = Data.Encoding.GetBytes(s);
			int i = -1;
			while (i < encoded.Length - 1)
			{
				i++;

				byte b = encoded[i];

				if (b >= 33 && b <= 60) 
				{
					QPWriteByte(Data, b);
					continue;
				}

				if (b >= 62 && b <= 126) 
				{
					QPWriteByte(Data, b);
					continue;
				}
				
				if (b == 0x09 || b == 0x20) 
				{
					if (NextIsEnter(encoded, i+1))  //Last character cannot be an space.
					{
						QPWriteByte(Data, '=', (char)((byte)'0' + b/16), (char)((byte)'0' + b%16));
					}
					else
					{
						QPWriteByte(Data, b);
					}
					continue;
				}

				if (NextIsEnter(encoded, i))
				{
					WriteLine(Data);
					QuotedPrintablePosition = 0;
					i++;
					continue;
				}

				byte d1 = (byte)(b/16); byte d1Base = d1<10? (byte)'0': (byte)((byte)'A' - 10);
				byte d2 = (byte)(b%16); byte d2Base = d2<10? (byte)'0': (byte)((byte)'A' - 10);
				QPWriteByte(Data, '=', (char)(d1Base + d1), (char)(d2Base + d2));
			}
		}

		private static bool NextIsEnter(byte[] encoded, int bytePos)
		{
			if (bytePos > encoded.Length - 1) return false;
			byte b = encoded[bytePos];
			if (b != 13) return false;
			bytePos++;
			b = encoded[bytePos];
			if (b != 10) return false;
			return true;
		}

		private void QPWriteByte(TextWriter Data, byte b)
		{
			if (QuotedPrintablePosition >= 75)  
			{
				Data.Write('=');
				WriteLine(Data);
				QuotedPrintablePosition = 0;
			}

			Data.Write((char)b);
			QuotedPrintablePosition++;
		}
		
		private void QPWriteByte(TextWriter Data, char a, char b, char c)
		{
			if (QuotedPrintablePosition >= 73)  
			{
				Data.Write('=');
				WriteLine(Data);
				QuotedPrintablePosition = 0;
			}

			Data.Write(a);
			Data.Write(b);
			Data.Write(c);
			QuotedPrintablePosition += 3;
		}

		#endregion

		#region Base64 encoding

		/// <summary>
		/// Writes the base64 encoding of a byte array, as defined in RFC 2045 section 6.8
		/// This method does not keep the state, so all binary data must be supplied at once.
		/// As an alternative you could use <see cref="System.Convert.ToBase64String(byte[], int, int)"/>, but this method
		/// avoids creating a temporary string, and then doubling the memory needed for the encoding.
		/// This method will also correctly split the string at 76 characters, while Convert.ToBase64 in .NET 1.1 will not. (2.0 added this support)
		/// </summary>
		/// <param name="Data">TextWriter where we will write the data.</param>
		/// <param name="s">String to Encode.</param>
		public void WriteBase64(TextWriter Data, byte[] s)	
		{

			char[] ToBase64 =
					  {
						  'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 
						  'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', 'a', 'b', 'c', 'd', 'e', 'f', 'g', 'h', 
						  'i', 'j', 'k', 'l', 'm', 'n', 'o', 'p', 'q', 'r', 's', 't', 'u', 'v', 'w', 'x', 'y', 
						  'z', '0', '1', '2', '3', '4', '5', '6', '7', '8', '9', '+', '/' 
					  };

			int LinePos = 0;
			for (int i = 0; i < s.Length; i+=3)
			{
				unchecked
				{
					Data.Write(ToBase64[s[i] >> 2]);
					
					if (i + 1 < s.Length)
					{
						Data.Write(ToBase64[((s[i] & 0x03)<< 4) | (s[i+1] >> 4)]);
					}
					else
					{
						Data.Write(ToBase64[((s[i] & 0x03)<< 4)]);
						Data.Write('=');
						Data.Write('=');
						return;
					}

					if (i + 2 < s.Length)
					{
						Data.Write(ToBase64[((s[i+1] & 0x0F)<< 2) | (s[i+2] >> 6)]);
						Data.Write(ToBase64[(s[i+2] & 0x3F)]);
					}
					else
					{
						Data.Write(ToBase64[((s[i+1] & 0x0F)<< 2)]);
						Data.Write('=');
						return;
					}
					
					LinePos++;
					if (LinePos >= 19)  //19*4 = 76
					{
						WriteLine(Data);
						LinePos = 0;
					}
				}
			}
		}
		#endregion

		#region General Utilites

		private static string GetBoundary()
		{
			Guid g1 = Guid.NewGuid();
			Guid g2 = Guid.NewGuid();
			return "----=_f.l.e.x.c.e.l." + g1.ToString("D") + "." + g2.ToString("D"); 
		}

		private static void WriteLine(TextWriter message)
		{
			//No need to use encoding.ASCII here, this is just 7 bits printable chars, and can be made faster just by writing the bytes.
			message.Write("\r\n");
		}

		private static void Write(TextWriter message, string s)
		{
			//No need to use encoding.ASCII here, this is just 7 bits printable chars, and can be made faster just by writing the bytes.

			message.Write(s);
		}

		private static void WriteLine(TextWriter message, string s)
		{
			//No need to use encoding.ASCII here, this is just 7 bits printable chars, and can be made faster just by writing the bytes.

			message.Write(s);
			WriteLine(message);
		}
		#endregion

	}
}
