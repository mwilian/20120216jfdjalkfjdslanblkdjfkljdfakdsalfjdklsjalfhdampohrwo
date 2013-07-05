using System;
using System.Security.Cryptography;
#if (FRAMEWORK20)
using System.Security.Cryptography.Pkcs;
#endif
using System.Drawing;
using System.IO;

using FlexCel.Core;

namespace FlexCel.Pdf
{
    /// <summary>
    /// Changes allowed in a signed PDF document.
    /// </summary>
    public enum TPdfAllowedChanges
    {
        /// <summary>
        /// No changes to the document are permitted; any change to the document invalidates the signature.
        /// </summary>
        None,

        /// <summary>
        /// Permitted changes are filling in forms, instantiating page templates, and signing; other changes invalidate the signature.
        /// </summary>
        FillingForms_PageTemplates_Signing,

        /// <summary>
        /// Permitted are filling in forms, instantiating page templates, and signing, as well as annotation creation, deletion, and modification; other changes invalidate the signature
        /// </summary>
        FillingForms_PageTemplates_Signing_Annotations
    }

	/// <summary>
	/// Describes a non visible signature for a PDF file. For a visible signature, use <see cref="TPdfVisibleSignature"/>
	/// </summary>
	public class TPdfSignature
	{
		#region Privates
		private string FName;
		private string FReason;
		private string FLocation;
		private string FContactInfo;
        private DateTime FSignDate;
        private TPdfAllowedChanges FAllowedChanges;

        private TPdfSignerFactory FSignerFactory;
		#endregion

		/// <summary>
		/// Creates an invisible signature. For a visible signature, create a <see cref="TPdfVisibleSignature"/> class.
		/// </summary>
        /// <param name="aSignerFactory">See <see cref="SignerFactory"/></param>
        /// <param name="aName">See <see cref="Name"/></param>
		/// <param name="aReason">See <see cref="Reason"/></param>
		/// <param name="aLocation">See <see cref="Location"/></param>
		/// <param name="aContactInfo">See <see cref="ContactInfo"/></param>
		public TPdfSignature(TPdfSignerFactory aSignerFactory, string aName, string aReason, string aLocation, string aContactInfo)
		{
            FSignerFactory = aSignerFactory;
			Name = aName; //without F, so it is checked.
			FReason = aReason;
			FLocation = aLocation;
			FContactInfo = aContactInfo;
            FSignDate = DateTime.MinValue;
            FAllowedChanges = TPdfAllowedChanges.FillingForms_PageTemplates_Signing;
		}

        /// <summary>
        /// Object that implements the actual signing.
        /// </summary>
        public TPdfSignerFactory SignerFactory { get { return FSignerFactory; } set { FSignerFactory = value; } }

		/// <summary>
		/// Name to be given to the signature. This will be displayed in the "signatures" tab, and acrobat normally names it "Signature".
        /// It cannot be null.
		/// <b>Note:</b> Signature names cannot contain dots. An exception will be thrown if you try to enter a name with a dot here.
		/// </summary>
		public string Name {get {return FName;} 
			set
			{
                if (value == null || value.Length == 0) PdfMessages.ThrowException(PdfErr.ErrNoDotsInSigName, String.Empty);
                if (value.IndexOf(".") >= 0) PdfMessages.ThrowException(PdfErr.ErrNoDotsInSigName, value);
				FName = value;

			}
		}

		/// <summary>
		/// The reason for the signing, such as "I agree...". Leave it null if you do not want to specify a reason.
		/// </summary>
		public string Reason {get {return FReason;} set{FReason = value;}}

		/// <summary>
		/// The CPU host name or physical location of the signing. Leave it null for not specifying a location.
		/// </summary>
		public string Location {get {return FLocation;} set{FLocation = value;}}

		/// <summary>
		/// Information provided by the signer to enable a recipient to contact the signer to verify the signature; for example, a phone number.
		/// </summary>
		public string ContactInfo {get {return FContactInfo;} set{FContactInfo = value;}}

        /// <summary>
        /// Sign Date. Use DateTime.MinValue to use the current date.
        /// </summary>
        public DateTime SignDate { get { return FSignDate; } set { FSignDate = value; } }

        /// <summary>
        /// Specifies which changes are allowed in the signed pdf.
        /// </summary>
        public TPdfAllowedChanges AllowedChanges { get { return FAllowedChanges; } 
            set 
            {
                if (!Enum.IsDefined(typeof(TPdfAllowedChanges), value))
                    PdfMessages.ThrowException(PdfErr.ErrInvalidAllowedChanges, (int)value);
                FAllowedChanges = value; 
            } 
        }

        internal int AllowedChangesValue
        {
            get
            {
                switch (FAllowedChanges)
                {
                    case TPdfAllowedChanges.None:
                        return 1;
                    case TPdfAllowedChanges.FillingForms_PageTemplates_Signing:
                        return 2;
                    case TPdfAllowedChanges.FillingForms_PageTemplates_Signing_Annotations:
                        return 3;
                    default:
                        return 1;
                }
            }
        }
 
    }

    /// <summary>
    /// Describes a visible signature in a PDF file. For an invisible signature, see <see cref="TPdfSignature"/>.
    /// </summary>
	public class TPdfVisibleSignature: TPdfSignature
	{
		#region Privates
		private int FPage;
		private RectangleF FRect;
		private byte[] FImageData;
		#endregion

		/// <summary>
		/// Creates a new visible signature for a PDF file.
		/// </summary>
        /// <param name="aSignerFactory">See <see cref="TPdfSignature.SignerFactory"/></param>
        /// <param name="aName">See <see cref="TPdfSignature.Name"/></param>
        /// <param name="aReason">See <see cref="TPdfSignature.Reason"/></param>
        /// <param name="aLocation">See <see cref="TPdfSignature.Location"/></param>
        /// <param name="aContactInfo">See <see cref="TPdfSignature.ContactInfo"/></param>
        /// <param name="aPage">See <see cref="Page"/></param>
		/// <param name="aRect">See <see cref="Rect"/></param>
		/// <param name="aImageData">See <see cref="ImageData"/></param>
        public TPdfVisibleSignature(TPdfSignerFactory aSignerFactory, string aName, string aReason, string aLocation, string aContactInfo, 
			int aPage, RectangleF aRect, byte[] aImageData): base(aSignerFactory, aName, aReason, aLocation, aContactInfo)
		{
			FPage = aPage;
			FRect = aRect;
			FImageData = aImageData;
		}

		/// <summary>
		/// Page where the signature will go. (1 based). Use 0 to place the signature at the last page.
		/// </summary>
		public int Page {get {return FPage;} set{FPage = value;}}

		/// <summary>
		/// Rectangle where the signature will go in the page. It is measured in points (1/72 of an inch) from the left lower corner of the page.
		/// </summary>
		public RectangleF Rect {get {return FRect;} set{FRect = value;}}

		/// <summary>
		/// The image that will be shown in the signature as an array of bytes.
		/// </summary>
		public byte[] ImageData {get {return FImageData;} set{FImageData = value;}}


	}

    /// <summary>
    /// Represents an abstract class to create a pdf PKCS7 DER encoded signature.
    /// Descend from this class to create your own SignerFactory implementations.
    /// </summary>
    public abstract class TPdfSigner : IDisposable
    {
        /// <summary>
        /// This method is called each time new data is added to the pdf. When overwriting this method, use it to incrementally calculate the hash of the data.
        /// </summary>
        /// <param name="buffer">Data written to the pdf.</param>
        /// <param name="offset">Offset in buffer where to start writing data.</param>
        /// <param name="count">Number of bytes from buffer that will be written.</param>
        public abstract void Write(byte[] buffer, int offset, int count);

        /// <summary>
        /// This method is called only once at the end of the pdf creation. It should release all handles and temporary memory used to calculate 
        /// the data hash, and return a PKCS7 DER-encoded signature.
        /// </summary>
        /// <returns>a PKCS7 DER-encoded signature.</returns>
        public abstract byte[] GetSignature();

        /// <summary>
        /// Returns the estimated length for the data that will be returned in <see cref="GetSignature"/>. 
        /// Note that this method will be called <b>before</b> finishing the pdf, so you still don't know what the final signature will be.
        /// You can return a number larger than what <see cref="GetSignature"/> will return at the end of the pdf creation, but <b>never smaller</b>.
        /// </summary>
        /// <returns></returns>
        public abstract int EstimateLength();

        #region IDisposable Members

        /// <summary>
        /// Disposes this instace.
        /// </summary>
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        /// <summary>
        /// Override this method to dispose your own objects in your descending classes.
        /// </summary>
        /// <param name="disposing">If true, this is a direct call to Dispose(), and you need to dispose both managed and unmanaged resources.
        /// If false, this is called by the finalizer, and you only need to release unmanaged resources.</param>
        protected virtual void Dispose(bool disposing)
        {
            //nothing here.
        }

        /// <summary>
        /// Finalizer for the class.
        /// </summary>

        ~TPdfSigner()
        {
            Dispose(false);
        }

        #endregion
    }

    /// <summary>
    /// Override this factory when creating your own <see cref="TPdfSigner"/> class, so it is returned here.
    /// </summary>
    public abstract class TPdfSignerFactory
    {
        /// <summary>
        /// This method should return an instance of your customized <see cref="TPdfSigner"/> class.
        /// </summary>
        /// <returns>A new TPdfSigner instance. Do not reuse instances, always return a new instance here since it will be disposed.</returns>
        public abstract TPdfSigner CreateSigner();
    }

#if (FRAMEWORK20)
    /// <summary>
    /// This class will create instances of the Built-in singer.
    /// </summary>
    public class TBuiltInSignerFactory: TPdfSignerFactory
    {
        private CmsSigner FSigner;

        /// <summary>
        /// Creates a new instance of this class.
        /// </summary>
        /// <param name="aSigner">CmsSigner used to sign the pdf files.</param>
        public TBuiltInSignerFactory(CmsSigner aSigner)
        {
            FSigner = aSigner;
        }

        /// <summary>
        /// Creates a new Builtin Signer.
        /// </summary>
        /// <returns></returns>
        public override TPdfSigner CreateSigner()
        {
            return new TBuiltInSigner(FSigner);
        }
    }

    internal class TBuiltInSigner : TPdfSigner
    {
        private MemoryStream DataBuffer;
        private CmsSigner FSigner;

        internal TBuiltInSigner(CmsSigner aSigner)
        {
            DataBuffer = new MemoryStream();
            FSigner = aSigner;
        }

        public CmsSigner Signer { get { return FSigner; }}


        public override void Write(byte[] buffer, int offset, int count)
        {
            DataBuffer.Write(buffer, offset, count);
        }

        public override byte[] GetSignature()
        {
            //This is incredibly wrong, but there is no method in the framework to create an incremental signature.
            //So we need to buffer it all, convert it to a byte array and sign that. :(
            SignedCms Message = new SignedCms(new ContentInfo(DataBuffer.ToArray()), true);

            DataBuffer.Close();  //Release the memory as soon as possible, it can be quite big since the whole doc is here.
            DataBuffer = null;

            Message.ComputeSignature(FSigner, true);
            return Message.Encode();
        }

        public override int EstimateLength()
        {
            byte[] dummyData = { 1, 2, 3 };  //Length shouldn't matter since signature is detached.
            SignedCms Message = new SignedCms(new ContentInfo(dummyData), true);
            Message.ComputeSignature(FSigner, true);
            return Message.Encode().Length;
        }

        protected override void Dispose(bool disposing)
        {
            base.Dispose(disposing);
            if (disposing)
            {
                if (DataBuffer != null)
                {
                    DataBuffer.Close();
                    DataBuffer = null;
                }
            }
        }
    }

#endif
}
