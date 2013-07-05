#region Using directives

using System;
using System.Text;
using FlexCel.Core;
using System.IO;
using System.Security.Cryptography;
using System.Diagnostics;

#endregion

namespace FlexCel.Pdf
{
    /// <summary>
    /// A wrapper for a normal stream that could compress or encrypt the data.
    /// </summary>
    internal class TPdfStream : IDisposable
    {
        private Stream FDataStream;
		private TSignedStream SignedStream;// Is the same as FDataStream or null if not signed.
        private bool FCompress;
        private TCompressor CompressEngine;
		internal bool PendingEndText; //optimization.
		private byte[] EndTextData;
		private TPdfSignature Signature;
		private TPdfSigner Signer;

        internal TPdfStream(Stream aDataStream, TPdfSignature aSignature)
        {
			EndTextData = TPdfBaseRecord.Coder.GetBytes(TPdfTokens.GetString(TPdfToken.CommandEndText) + TPdfTokens.NewLine);

			Signature = aSignature;
			if (Signature != null)
			{
                if (Signature.SignerFactory != null) 
                    Signer = Signature.SignerFactory.CreateSigner(); 
                else PdfMessages.ThrowException(PdfErr.ErrUnassingedSignerFactory);
			}

            if (Signer == null) FDataStream = aDataStream; else FDataStream = new TSignedStream(aDataStream, Signer);  //When signing we use a special stream that will compute the hash.
			SignedStream = FDataStream as TSignedStream;
        }

		public bool HasSignature
		{
			get
			{
				return Signature != null;
			}
		}

        public bool Compress 
        {
            get 
            {
                return FCompress;
            } 
            set 
            {
                if (value && !FCompress)
                {
                    if (CompressEngine==null)
                        CompressEngine= new TCompressor();
                    CompressEngine.BeginDeflate();
                }
                else
                    if (!value && FCompress && CompressEngine != null)
                {
                    CompressEngine.EndDeflate(FDataStream);
                }
                
                FCompress=value;
            }
        }

        public static byte[] CompressData(byte[] Data)
        {
            using (TCompressor Cmp = new TCompressor())
            {
                using (MemoryStream ms = new MemoryStream())
                {
                    Cmp.Deflate(Data, 0, ms);
                    return ms.ToArray();
                }
            }

        }

		public void FlushEndText()
		{
			if (FCompress)
			{
				if (PendingEndText) 
				{
					CompressEngine.IncDeflate(EndTextData, 0, FDataStream);
					PendingEndText = false;
				}
			}
			else
			{
				if (PendingEndText) 
				{
					FDataStream.Write(EndTextData, 0, EndTextData.Length);
					PendingEndText = false;
				}
			}
		}

        public void Write(byte[] Data)
        {
			if (FCompress)
			{
				if (PendingEndText) 
				{
					CompressEngine.IncDeflate(EndTextData, 0, FDataStream);
					PendingEndText = false;
				}
				CompressEngine.IncDeflate(Data, 0, FDataStream); 
			}
			else
			{
				if (PendingEndText) 
				{
					FDataStream.Write(EndTextData, 0, EndTextData.Length);
					PendingEndText = false;
				}
				FDataStream.Write(Data, 0, Data.Length);
			}
        }
        
        public long Length
        {
            get
            {
                return FDataStream.Length;
            }
        }

        public long Position
        {
            get
            {
                return FDataStream.Position;
            }
		}

        public long SignPosition
        {
            get
            {
                Debug.Assert(SignedStream != null);
                return SignedStream.SignPosition;
            }
        }

		public void GoVirtual(int aSignSize)
		{
			Debug.Assert(SignedStream != null);
			SignedStream.GoVirtual(aSignSize);
		}

		public void GoReal(byte[] AfterSign, int PKCSSize)
		{
			Debug.Assert(SignedStream != null);
			SignedStream.GoReal(AfterSign, PKCSSize);
		}

        internal int GetEstimatedLength()
        {
            Debug.Assert(Signer != null);
            return Signer.EstimateLength();
        }

		#region IDisposable Members

		public void Dispose()
		{
            if (Signer != null) Signer.Dispose();
			if (CompressEngine != null) CompressEngine.Dispose();
			if (SignedStream != null) SignedStream.Close();
			GC.SuppressFinalize(this);
		}

		#endregion

    }

	#region Sign
	internal class TSignedStream : Stream
	{
		private Stream FDataStream;
		private TPdfSigner Signer;
		private THelperStream HelperStream;
		private int SignOffset;
        
		internal TSignedStream(Stream aDataStream, TPdfSigner aSigner)
		{
			FDataStream = aDataStream;
			Signer = aSigner;
			HelperStream = null;
			SignOffset = 0;
		}

		#region Stream Methods
		public override bool CanRead
		{
			get
			{
				return false;
			}
		}

		public override bool CanSeek
		{
			get
			{
				return false;
			}
		}

		public override bool CanWrite
		{
			get
			{
				return VirtualStream.CanWrite;
			}
		}

		public override void Flush()
		{
			if (HelperStream != null) FlxMessages.ThrowException(FlxErr.ErrInternal);		
			VirtualStream.Flush();
		}

		public override long Length
		{
			get
			{
                if (HelperStream == null) return FDataStream.Length;
                return FDataStream.Length + SignOffset + HelperStream.Length;
			}
		}
        
		public override long Position
		{
			get
			{
                if (HelperStream == null) return FDataStream.Position;
				return FDataStream.Position + SignOffset + HelperStream.Length;
			}
			set
			{
				FlxMessages.ThrowException(FlxErr.ErrInternal);
			}
		}

		public override void Write(byte[] buffer, int offset, int count)
		{
			if (HelperStream == null) Signer.Write(buffer, offset, count);
			VirtualStream.Write(buffer, offset, count);
		}

		public override int Read(byte[] buffer, int offset, int count)
		{
			FlxMessages.ThrowException(FlxErr.ErrInternal);
			return 0;
		}

		public override long Seek(long offset, SeekOrigin origin)
		{
			FlxMessages.ThrowException(FlxErr.ErrInternal);
			return 0;
		}

		public override void SetLength(long value)
		{
			FlxMessages.ThrowException(FlxErr.ErrInternal);
		}

		#endregion

		private Stream VirtualStream
		{
			get
			{
				if (HelperStream != null) return HelperStream.CurrentStream;
				return FDataStream;
			}
		}

        internal long SignPosition
        {
            get
            {
                return FDataStream.Length;
            }
        }
		
		internal void GoVirtual(int aSignOffset)
		{
			Debug.Assert(HelperStream == null);
			SignOffset = aSignOffset;
			HelperStream = new THelperStream(FDataStream);
		}


        private static byte[] Pad(byte Character, int Count)
        {
            byte[] Result = new byte[Count];
            for (int i = 0; i < Count; i++)
            {
                Result[i] = Character;
            }

            return Result;
        }

        internal void GoReal(byte[] AfterSign, int PKCSSize)
		{
			Debug.Assert(HelperStream != null);
            Signer.Write(AfterSign, 0, AfterSign.Length); 
            
            if (PKCSSize + AfterSign.Length > SignOffset)
            {
                PdfMessages.ThrowException(PdfErr.ErrSigningLengthToSmall);
            }
            byte[] PaddedByteCount = Pad(0x20, SignOffset - (PKCSSize + AfterSign.Length));
            Signer.Write(PaddedByteCount, 0, PaddedByteCount.Length);

			byte[] bt = HelperStream.CurrentStream.ToArray();
			HelperStream.Dispose();
			HelperStream = null;

            Signer.Write(bt, 0, bt.Length);
			byte[] sg = Signer.GetSignature();
            byte[] hexsg = TPdfBaseRecord.Coder.GetBytes("<" + PdfConv.ToHexString(sg, false));

            FDataStream.Write(hexsg, 0, hexsg.Length);

            if (hexsg.Length > PKCSSize - 1)
            {
                PdfMessages.ThrowException(PdfErr.ErrSigningLengthToSmall);
            }

            for (int i = hexsg.Length; i < PKCSSize - 1; i++) //pad the certificate.
            {
                FDataStream.WriteByte(0);
            }

            FDataStream.Write(TPdfBaseRecord.Coder.GetBytes(">"),0,1);

			FDataStream.Write(AfterSign, 0, AfterSign.Length);
            FDataStream.Write(PaddedByteCount, 0, PaddedByteCount.Length);  //Pad the whole thing.
            
            FDataStream.Write(bt, 0, bt.Length);
		}

#if (FRAMEWORK20)
    protected override void Dispose(bool disposing)
	{
            if (disposing)
            {
                if (HelperStream != null) HelperStream.Dispose();
                HelperStream = null;

            }

            base.Dispose(disposing);
        }
    }
#else
		protected virtual void Dispose(bool disposing)
		{
			if (disposing)
			{
				if (HelperStream != null) HelperStream.Dispose();
				HelperStream = null;

			}

			base.Close();
		}

		/// <summary>
		/// Closes the stream.
		/// </summary>
		public void Dispose()
		{
			Dispose(true);
			GC.SuppressFinalize(this);
		}
	}

#endif

	
	internal class THelperStream: IDisposable
	{
		private MemoryStream AuxStream;
		private Stream MainStream;

		internal THelperStream(Stream aMainStream)
		{
			MainStream = aMainStream;
			AuxStream = new MemoryStream();
		}

		internal MemoryStream CurrentStream
		{
			get
			{
				return AuxStream;
			}
		}

		internal long Length
		{
			get
			{
				return AuxStream.Length;
			}
		}


		#region IDisposable Members

		public void Dispose()
		{
			AuxStream.Close();
            GC.SuppressFinalize(this);
		}

		#endregion
	}
	#endregion
}
