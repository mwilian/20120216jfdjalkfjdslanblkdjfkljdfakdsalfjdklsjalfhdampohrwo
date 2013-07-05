using System;
using System.Resources;
using System.Reflection;

namespace TVCDesigner
{
    /// <summary>
    /// Message Codes. We use this and not actual strings to make sure all are correctly spelled.
    /// </summary>
    public enum FldMsg
    {
        /// <summary>
        /// Datasets
        /// </summary>
        strDatasets,
        
        /// <summary>
        /// Tags
        /// </summary>
        strExtras,

        /// <summary>
        /// A full config sheet.
        /// </summary>
        strConfigSheet,

        /// <summary>
        /// Datasets defined on the excel sheet.
        /// </summary>
        strUserDefined,

        /// <summary>
        /// Report variables
        /// </summary>
        strReportVars,

        /// <summary>
        /// Report Expressions
        /// </summary>
        strReportExpressions,

        /// <summary>
        /// User defined formats.
        /// </summary>
        strFormats,

		/// <summary>
		/// Commands for the config sheet.
		/// </summary>
		strConfig
    }	

		
    /// <summary>
    /// FlexCelDesigner Constants. It reads the resources from the active locale, and
    /// returns the correct string.
    /// If your language is not supported and you feel like transating the messages,
    /// please send me a copy. I will include it on the next FlexCel version. 
    /// <para>To add a new language:
    /// <list type="number">
    /// <item>
    ///    Copy the file fldmsg.resx to your language (for example, fldmsg.es.resx to translate to spanish)
    /// </item><item>
    ///    Edit the new file and change the messages(you can do this visually with visual studio)
    /// </item><item>
    ///    Add the .resx file to the FlexCelDesigner project
    /// </item>
    /// </list>
    /// </para>
    /// </summary>
    public sealed class FldMessages
    {
        private FldMessages(){}
        internal static readonly ResourceManager rm = new ResourceManager("TVCDesigner.FldMsg", Assembly.GetExecutingAssembly()); //STATIC*

        /// <summary>
        /// Retruns a string based on the FldMsg enumerator, formated with args.
        /// This method is used to get an Exception error message. 
        /// </summary>
        /// <param name="ResName">Message Code.</param>
        /// <param name="args">Params for this message.</param>
        /// <returns></returns>
        public static string GetString( FldMsg ResName, params object[] args)
        {
            if (args.Length==0) return rm.GetString(ResName.ToString()); //To test without args
            return (String.Format(rm.GetString(ResName.ToString()), args));
        }
    }
}
