using System;
using System.IO;

namespace FlexCel.Core
{
	/// <summary>
	/// Stream Helper. To read/write full streams.
	/// </summary>
	internal sealed class Sh
	{
		private Sh(){}

        internal static void Read(Stream aStream, byte[] data, int iniOfs, int count)
        {
            Read(aStream, data, iniOfs, count, true);
        }

		internal static void Read(Stream aStream, byte[] data, int iniOfs, int count, bool ThrowOnEOF)
		{
			int offset=iniOfs;
			int remaining = count;
			while (remaining > 0)
            {
				int read = aStream.Read(data, offset, remaining);
			
				if (read <= 0)
				{
					//we Are at EOF.
					//Sector might not end if it is the end of the stream :-(  .We read it anyway and fill it with zeros
						
					if (remaining==count || ThrowOnEOF)
						throw new EndOfStreamException (FlxMessages.GetString(FlxErr.ErrEofReached, remaining));
					for (int i=count-remaining; i<count;i++)
						data[iniOfs + i]=0;
					read=remaining;

				}
				offset+=read;
				remaining -= read;
			}
		}

		/// <summary>
		/// For consistency. It should be inlined anyway....
		/// </summary>
		/// <param name="aStream"></param>
		/// <param name="data"></param>
		/// <param name="offset"></param>
		/// <param name="count"></param>
		internal static void Write(Stream aStream, byte[] data, int offset, int count)
		{
			aStream.Write(data, offset, count);
		}
		
		internal static void CopyStream(Stream Inf, Stream OutStream)
		{
			byte[] buff = new byte[4096];
        	int Read = 0;
        	while ((Read = Inf.Read(buff, 0, 4096)) > 0)
        	{
				OutStream.Write(buff, 0, Read);
        	}
		}

	}
}
