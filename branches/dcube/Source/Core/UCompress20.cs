//Framework 20 has built in zupport for DEFLATE (RFC 1951) and GZIP (RFC 1952) but not ZLIB (RFC 1950)
//To convert a RFC1951 to RFC1950 we need to add a 2 byte Header: 0x78 0x9C and an Adler32 CheckSum at the end.

#if (FRAMEWORK20 && (FRAMEWORK30 || !COMPACTFRAMEWORK) && !SILVERLIGHT)  
using System;
using System.IO;
using System.Runtime.CompilerServices;
using System.Reflection;
using System.IO.Compression;


namespace FlexCel.Core
{
    /// <summary>
    /// Compress a stream of bytes using GZIP compatible compression.
    /// This class needs .NET Framework 2.0. A slower implementation using Vj# is available if 
    /// you don't define "FRAMEWORK20" (it also has a little better compression because compression level on .NET 20 is not configurable)
    /// </summary>
    internal sealed class TCompressor: IDisposable
    {
        DeflateStream Deflater;
        UInt32 FAdler;

        internal TCompressor()
        {
        }

        static readonly byte[] Header = { 0x78, 0x9C };

        #region Normal compress
        internal void Deflate(byte[] InputData, int Offset, Stream OutStream)
        {
            OutStream.Write(Header, 0, Header.Length);
            using (DeflateStream Def = new DeflateStream(OutStream, CompressionMode.Compress, true))
            {
                Def.Write(InputData, Offset, InputData.Length - Offset);
            }

            WriteAdler32(InputData, Offset, OutStream);
        }

        internal void Inflate(byte[] InputData, int Offset, Stream OutStream)
        {
            using (MemoryStream ms = new MemoryStream(InputData, Offset + 2, InputData.Length - Offset - 2 - 4))  //extract addler sum and header.
            {
                using (DeflateStream Inf = new DeflateStream(ms, CompressionMode.Decompress, true))
                {
                    Sh.CopyStream(Inf, OutStream);
                }
            }
        }
        #endregion

        #region Progressive compress
        internal void BeginDeflate()
        {
            //nothing here   
        }

        internal void IncDeflate(byte[] Data, int Offset, Stream OutStream)
        {
            if (Deflater == null)
            {
                OutStream.Write(Header, 0, Header.Length);
                Deflater = new DeflateStream(OutStream, CompressionMode.Compress, true);
                FAdler = 1;
            }


            Deflater.Write(Data, Offset, Data.Length - Offset);
            FAdler = Adler32(FAdler, Data, Offset, Data.Length - Offset);

        }

        internal void IncDeflate(byte[] InData, int Offset, int Length, Stream OutStream)
        {
            if (Deflater == null)
            {
                OutStream.Write(Header, 0, Header.Length);
                Deflater = new DeflateStream(OutStream, CompressionMode.Compress, true);
            }

            Deflater.Write(InData, Offset, Length);
            FAdler = Adler32(FAdler, InData, Offset, Length);
        }

        internal void EndDeflate(Stream OutStream)
        {
            if (Deflater != null) Deflater.Close();
            Deflater = null;
            WriteAdler32(OutStream, FAdler);
        }
        #endregion

        #region Adler 32
        //This method is adapted from Zlib sources.
        #region Copyright Notice
        /* zlib.h -- interface of the 'zlib' general purpose compression library
        version 1.2.2, October 3rd, 2004

        Copyright (C) 1995-2004 Jean-loup Gailly and Mark Adler

        This software is provided 'as-is', without any express or implied
        warranty.  In no event will the authors be held liable for any damages
        arising from the use of this software.

        Permission is granted to anyone to use this software for any purpose,
        including commercial applications, and to alter it and redistribute it
        freely, subject to the following restrictions:

        1. The origin of this software must not be misrepresented; you must not
            claim that you wrote the original software. If you use this software
            in a product, an acknowledgment in the product documentation would be
            appreciated but is not required.
        2. Altered source versions must be plainly marked as such, and must not be
            misrepresented as being the original software.
        3. This notice may not be removed or altered from any source distribution.

        Jean-loup Gailly        Mark Adler
        jloup@gzip.org          madler@alumni.caltech.edu


        The data format used by the zlib library is described by RFCs (Request for
        Comments) 1950 to 1952 in the files http://www.ietf.org/rfc/rfc1950.txt
        (zlib format), rfc1951.txt (deflate format) and rfc1952.txt (gzip format).
        */
#endregion

        const int BASE = 65521; /* largest prime smaller than 65536 */
        const int NMAX = 5552; /* NMAX is the largest n such that 255n(n+1)/2 + (n+1)(BASE-1) <= 2^32-1 */

        static UInt32 Adler32(UInt32 adler, byte[] buf, int first, int len)
        {
            unchecked
            {
                UInt32 s1 = adler & 0xffff;
                UInt32 s2 = (adler >> 16) & 0xffff;
                int k;

                int Pos = first;
                if (buf == null) return 1;

                while (len > 0)
                {
                    k = len < NMAX ? (int)len : NMAX;
                    len -= k;
                    while (k > 0)
                    {
                        s1 += buf[Pos];
                        s2 += s1;
                        Pos++;
                        k--;
                    }
                    s1 %= BASE;
                    s2 %= BASE;
                }
                return (s2 << 16) | s1;
            }
        }

        private static void WriteAdler32(byte[] InputData, int Offset, Stream OutStream)
        {
            UInt32 adl = Adler32(1, InputData, Offset, InputData.Length - Offset);
            WriteAdler32(OutStream, adl);
        }

        private static void WriteAdler32(Stream OutStream, UInt32 adl)
        {
            OutStream.WriteByte((byte)((adl >> 24) & 0xFF));
            OutStream.WriteByte((byte)((adl >> 16) & 0xFF));
            OutStream.WriteByte((byte)((adl >> 8) & 0xFF));
            OutStream.WriteByte((byte)((adl >> 0) & 0xFF));
        }

        #endregion

        #region IDisposable Members

        public void Dispose()
        {
            if (Deflater != null) Deflater.Dispose();
            GC.SuppressFinalize(this);
        }

        #endregion
    }
}
#endif
