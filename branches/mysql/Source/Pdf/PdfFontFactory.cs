#if (!WPF) //WPF doesn't need this.

using System;
using System.IO;
using System.Globalization;
using System.Collections.Generic;

using FlexCel.Core;

#if (WPF)
#else
using System.Drawing;
#endif

namespace FlexCel.Pdf
{
	/// <summary>
	/// Holds a list of system installed fonts and their postscript names.
	/// </summary>
	internal sealed class PdfFontFactory
	{
		private static TPsFontList FontList = new TPsFontList(); //STATIC*
		private static object FontListAccess = new object();
		private PdfFontFactory(){}

		public static byte[] GetFontData (Font aFont, string aFontPath)
		{
			string s = null;
			lock (FontListAccess)  //we could use a multireader lock here, but they are much more expensive. Most of the time, there will be no much contention anyway. Since we do not use readwrite locks, we need to lock() on READ and WRITE, as only way to be safe.
			{
				s = FontList.GetFont(aFontPath, aFont.Name, aFont.Style);
			}

			if (s!=null) return OpenFile(s);
				else return null;
		}

		private static byte[] OpenFile(string Path)
		{
			using (FileStream fs = new FileStream(Path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)) //FileShare.ReadWrite is the way we have to open a file even if it is being used by excel.))
			{
				byte[] b = new byte[fs.Length];
				Sh.Read(fs, b, 0, (int)fs.Length);
				return b;
			}
		}

	}
#if(FRAMEWORK20)
	internal class StringHash: Dictionary<string, string>
	{
        internal new string this[string key]
        {
            get
            {
                string Result = null;
                if (TryGetValue(key, out Result))
                    return Result;
                return null;
            }
			set
			{
				base[key] = value;
			}
        }

	}
#else
	internal class StringHash: Hashtable
	{
	}
#endif

	internal class TPsFontList
	{
		private StringHash FFontList;
		private StringHash FFontFiles;
		private static DateTime LastRefreshDate = DateTime.MinValue;
        private static string LastRereshPath = null;
		
		internal TPsFontList ()
		{
			FFontList = new StringHash();
			FFontFiles = new StringHash();
		}

		private static string MakeHash(string name, FontStyle style)
		{
			string s = name;
			if ((style & FontStyle.Bold)!=0) s+="\u0001"; else s+="\u0000";
			if ((style & FontStyle.Italic)!=0) s+="\u0001"; else s+="\u0000";
			return s;
		}

		/*private FontStyle GetStyle(string SubFamily)
		{
			FontStyle Result = FontStyle.Regular;
			SubFamily = SubFamily.ToUpper(CultureInfo.InvariantCulture);
			if (SubFamily.IndexOf(TPdfTokens.GetString(TPdfToken.FamilyItalic))>=0) Result |= FontStyle.Italic;
			if (SubFamily.IndexOf(TPdfTokens.GetString(TPdfToken.FamilyOblique))>=0) Result |= FontStyle.Italic;
			if (SubFamily.IndexOf(TPdfTokens.GetString(TPdfToken.FamilyBold))>=0) Result |= FontStyle.Bold;
					
			return Result;
		}*/

		internal static FontStyle GetStyle(int FontFlags)
		{
			FontStyle Result = FontStyle.Regular;
			if ((FontFlags & (1<<18)) != 0) Result |= FontStyle.Bold;
			if ((FontFlags & (1<<6)) != 0) Result |= FontStyle.Italic;
					
			return Result;
		}
		
		internal string GetFont(string FontPath, string FontName, FontStyle Style)
		{
            string s = (string)FFontList[MakeHash(FontName, Style)];

			if (s==null) 
			{
				RefreshList(FontPath, FontName);
                s = (string)FFontList[MakeHash(FontName, Style)];
			}

            if (s == null && Style != FontStyle.Regular)
            {
                if (((Style & FontStyle.Bold) != 0) && ((Style & FontStyle.Italic) != 0))
                {
                    s = (string)FFontList[MakeHash(FontName, FontStyle.Bold)]; //try only bold.
                    if (s != null)
                    {
                        return s;
                    }

                    s = (string)FFontList[MakeHash(FontName, FontStyle.Italic)]; //try only italic.
                    if (s != null)
                    {
                        return s;
                    }
                }

                if (s == null)
                {
                    s = (string)FFontList[MakeHash(FontName, FontStyle.Regular)]; //try without style.
                    if (s != null)
                    {
                        return s;
                    }
                }
            }

			return s;
		}

		private void RefreshList(string FontPath, string FontName)
		{
			const int WaitTime = 1500; // 25 minutes.

			if (DateTime.Now - LastRefreshDate < TimeSpan.FromSeconds(WaitTime) && string.Equals(FontPath, LastRereshPath, StringComparison.InvariantCultureIgnoreCase)) return; //Avoid too much failed refreshs.
			// Create a reference to the current directory.
			DirectoryInfo di = new DirectoryInfo(FontPath);
            LoadAllFonts(FontPath, di, TPdfTokens.GetString(TPdfToken.TTFExtension), true);
            LoadAllFonts(FontPath, di, TPdfTokens.GetString(TPdfToken.TTCExtension), false);
			LastRefreshDate = DateTime.Now;
            LastRereshPath = FontPath;
		}

        private void LoadAllFonts(string FontPath, DirectoryInfo di, string FontExtension, Boolean ErrorIfEmpty)
        {
            // Create an array representing the files in the current directory.
            FileInfo[] fi = di.GetFiles(FontExtension);
            if (ErrorIfEmpty && fi.Length == 0) FlxMessages.ThrowException(FlxErr.ErrEmptyFolder, FontPath, FontExtension);


            foreach (FileInfo fontFile in fi)
            {
                if (FFontFiles.ContainsKey(fontFile.Name))
                    continue;

                FFontFiles.Add(fontFile.Name, null);
                try
                {
                    LoadFont(fontFile.FullName);
                }
                catch (Exception ex)
                {
                    //Invalid font. nothing to do, just continue reading the other fonts.
                    // What we will do is just ignore it here, and throw an exception when (and if) the user actually tries to use this font.

                    if (FlexCelTrace.HasListeners) FlexCelTrace.Write(new TPdfCorruptFontInFontFolderError(ex.Message, fontFile.FullName));
                }
            }
        }
		

		private void LoadFont(string path)
		{
			byte[] b = null;
			using (FileStream fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)) //FileShare.ReadWrite is the way we have to open a file even if it is being used by excel.))
			{
				b = new byte[fs.Length];
				Sh.Read(fs, b, 0, (int)fs.Length);
			}

            LoadFont(path, b);
		}

        private void LoadFont(string path, byte[] b)
        {
            TTrueTypeInfo[] ttfs = TPdfTrueType.GetColection(b);

            foreach (TTrueTypeInfo ttf in ttfs)
            {
                //FFontList[MakeHash(ttf.FamilyName, GetStyle(ttf.SubFamilyName))]=path;
                FFontList[MakeHash(ttf.FamilyName, GetStyle(ttf.FontFlags))] = path;
            }
        }
	}
}
#endif
