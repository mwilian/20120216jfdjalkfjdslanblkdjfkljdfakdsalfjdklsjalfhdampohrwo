using System;
using System.Resources;
using System.Reflection;
using System.Diagnostics;
using System.Runtime.Serialization;
using System.Security.Permissions;

using FlexCel.Core;
using System.Security;

namespace FlexCel.Pdf
{
    /// <summary>
    /// Error Codes. We use this and not actual strings to make sure all are correctly spelled.
    /// </summary>
    public enum PdfErr
    {
        /// <summary>
        /// Internal error.
        /// </summary>
        ErrInternal,

        /// <summary>The png image is corrupt, invalid, or not in a format FlexCel can understand.</summary>
        ErrInvalidPngImage,

		/// <summary>This PaperKind is not supported. Please use a custom Paper size.</summary>
		ErrInvalidPageSize,

		/// <summary>
		/// The font file for this font is invalid.
		/// </summary>
		ErrInvalidFont,

        /// <summary>
        /// The font was not found.
        /// </summary>
        ErrFontNotFound,

		/// <summary>
		/// Invalid page number.
		/// </summary>
		ErrInvalidPageNumber,

		/// <summary>
		/// A pdf file must be signed before calling BeginDoc.
		/// </summary>
		ErrTryingToSignStartedDocument,

		/// <summary>
		/// Signature names cannot contain dots.
		/// </summary>
		ErrNoDotsInSigName,

        /// <summary>
        /// The estimated size for the signature was smaller than the final size.
        /// </summary>
        ErrSigningLengthToSmall,

        /// <summary>
        /// There is no signer associated with the signature.
        /// </summary>
        ErrUnassingedSignerFactory,

        /// <summary>
        /// AllowedChanges value is invalid.
        /// </summary>
        ErrInvalidAllowedChanges

    }

    /// <summary>
    /// FlexCel Native PDF Constants. It reads the resources from the active locale, and
    /// returns the correct string.
    /// If your language is not supported and you feel like translating the messages,
    /// please send us a copy. We will include it on the next FlexCel version. 
    /// <para>To add a new language:
    /// <list type="number">
    /// <item>
    ///    Copy the file pdfmsg.resx to your language (for example, pdfmsg.es.resx to translate to spanish)
    /// </item><item>
    ///    Edit the new file and change the messages(you can do this visually with visual studio)
    /// </item><item>
    ///    Add the .resx file to the FlexCel project
    /// </item>
    /// </list>
    /// </para>
    /// </summary>
    public sealed class PdfMessages
    {
        private PdfMessages() { }
        internal static readonly ResourceManager rm = new ResourceManager("FlexCel.Pdf.pdfmsg", Assembly.GetExecutingAssembly()); //STATIC*

        /// <summary>
        /// Reruns a string based on the PdfErr enumerator, formatted with args.
        /// This method is used to get an Exception error message. 
        /// </summary>
        /// <param name="ResName">Error Code.</param>
        /// <param name="args">Params for this error.</param>
        /// <returns></returns>
        public static string GetString(PdfErr ResName, params object[] args)
        {
            if (args.Length == 0) return rm.GetString(ResName.ToString()); //To test without args
            return (String.Format(rm.GetString(ResName.ToString()), args));
        }

        /// <summary>
        /// Throws a standard FlexCelPdfException.
        /// </summary>
        /// <param name="ResName">Error Code.</param>
        /// <param name="args">Params for this error.</param>
        public static void ThrowException(PdfErr ResName, params object[] args)
        {
            throw new FlexCelPdfException(GetString(ResName, args), ResName);
        }

        /// <summary>
        /// Throws a standard FlexCelPdfException with innerException.
        /// </summary>
        /// <param name="e">Inner exception.</param>
        /// <param name="ResName">Error Code.</param>
        /// <param name="args">Params for this error.</param>
        public static void ThrowException(Exception e, PdfErr ResName, params object[] args)
        {
            throw new FlexCelPdfException(GetString(ResName, args), ResName, e);
        }
    }

    /// <summary>
    /// Exception thrown when an exception on the PDF engine happens. 
    /// </summary>
    [Serializable]
    public class FlexCelPdfException : FlexCelException
    {
#if (FRAMEWORK40)
        [NonSerialized]
#endif
        private FlexCelPdfExceptionState FState;

        private void InitState(PdfErr aErrorCode)
        {
            FState.ErrorCode = aErrorCode;
#if(FRAMEWORK40)
            SerializeObjectState += delegate(object exception,
                SafeSerializationEventArgs eventArgs)
            {
                eventArgs.AddSerializedState(FState);
            };
#endif
        }

        /// <summary>
        /// Creates a new FlexCelPdfException
        /// </summary>
        public FlexCelPdfException(): base() { InitState(PdfErr.ErrInternal); }

        /// <summary>
        /// Creates a new FlexCelPdfException with an error message.
        /// </summary>
        /// <param name="message">Error Message</param>
        public FlexCelPdfException(string message): base(message) { InitState(PdfErr.ErrInternal); }

        /// <summary>
        /// Creates a new FlexCelPdfException with an error message and an exception code.
        /// </summary>
        /// <param name="message">Error Message</param>
        /// <param name="aErrorCode">Error code of the exception.</param>
        public FlexCelPdfException(string message, PdfErr aErrorCode): base(message) { InitState(aErrorCode); }

        /// <summary>
        /// Creates a nested Exception.
        /// </summary>
        /// <param name="message">Error Message.</param>
        /// <param name="inner">Inner Exception.</param>
        public FlexCelPdfException(string message, Exception inner):base(message, inner) { InitState(PdfErr.ErrInternal); }

        /// <summary>
        /// Creates a nested Exception.
        /// </summary>
        /// <param name="message">Error Message.</param>
        /// <param name="inner">Inner Exception.</param>
        /// <param name="aErrorCode">Error code of the exception.</param>
        public FlexCelPdfException(string message, PdfErr aErrorCode, Exception inner):base(message, inner) { InitState(aErrorCode); }

#if(!COMPACTFRAMEWORK && !SILVERLIGHT && !FRAMEWORK40)
        /// <summary>
        /// Creates an exception from a serialization context.
        /// </summary>
        /// <param name="info">Serialization information.</param>
        /// <param name="context">Streaming Context.</param>
		[SecurityPermissionAttribute(SecurityAction.Demand,SerializationFormatter=true)]
        protected FlexCelPdfException(SerializationInfo info, StreamingContext context) : base(info, context) 
        { 
            if (info != null) FState.ErrorCode = (PdfErr)info.GetInt32("FErrorCode"); 
        }

        /// <summary>
        /// Implements standard GetObjectData.
        /// </summary>
        /// <param name="info"></param>
        /// <param name="context"></param>
		[SecurityPermissionAttribute(SecurityAction.Demand,SerializationFormatter=true)]
        public override void GetObjectData(SerializationInfo info, StreamingContext context)
        {
            if (info != null) info.AddValue("FErrorCode", (int)FState.ErrorCode);
            base.GetObjectData (info, context);
        }
#endif

        /// <summary>
        /// Error code on the Exception.
        /// </summary>
        public PdfErr ErrorCode
        {
            get
            {
                return FState.ErrorCode;
            }
        }

        [Serializable]
        private struct FlexCelPdfExceptionState : ISafeSerializationData
        {
            private PdfErr FErrorCode;

            public PdfErr ErrorCode
            {
                get { return FErrorCode; }
                set { FErrorCode = value; }
            }

            void ISafeSerializationData.CompleteDeserialization(object obj)
            {
                FlexCelPdfException ex = obj as FlexCelPdfException;
                ex.FState = this;
            }
        }
    }

}
