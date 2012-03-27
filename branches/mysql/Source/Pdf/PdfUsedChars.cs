using System;
using System.Collections.Generic;
using System.Text;
using FlexCel.Core;
using System.Globalization;


namespace FlexCel.Pdf
{
#if(FRAMEWORK20)
    internal class TWKeyList : List<char>
    {
        internal TWKeyList(ICollection<char> c): base(c)
        {
        }
    }
#else
    internal class TWKeyList : ArrayList
    {
		internal TWKeyList(ICollection c): base(c){}
    }
#endif

    /// <summary>
    /// A list holding the characters used on an Unicode True Type font.
	/// </summary>
	internal class TUsedCharList
	{
#if(FRAMEWORK20)
		private Dictionary<char, float> FList;
#else
		private Hashtable FList;
#endif
		private int FDW=1000;

		public TUsedCharList()
		{
#if(FRAMEWORK20)
			FList = new Dictionary<char, float>();
#else
			FList = new Hashtable();
#endif
		}

		public void Add(char newc, float GlyphWidth)
		{
			if (FList.ContainsKey(newc)) return;
			FList[newc] = GlyphWidth;
		}

		public int DW()
		{
			return FDW;
		}

		public string W()
		{
			StringBuilder sb = new StringBuilder();
			sb.Append(TPdfTokens.GetString(TPdfToken.OpenArray));
			TWKeyList a = new TWKeyList(FList.Keys);
			a.Sort();
			if (a.Count<=0) 
			{
				sb.Append(TPdfTokens.GetString(TPdfToken.CloseArray));
				return sb.ToString();
			}


			int LastW = Convert.ToInt32(FList[a[0]], CultureInfo.InvariantCulture);
			int LastI = 0;
			int LastKey = Convert.ToInt32(a[0]);
			bool ContinueKey = false;
			bool ArrayOpened = false;

			for (int i =1; i<= a.Count; i++)
			{
				char key = i==a.Count? (char)a[i-1]: (char)a[i];
				int w= Convert.ToInt32(FList[key]);

				if (w!=LastW || i==a.Count) 
				{
					if (LastI+1<i)
					{
						if (ArrayOpened) sb.Append(TPdfTokens.GetString(TPdfToken.CloseArray)+TPdfTokens.NewLine);
						ArrayOpened = false;
						if (LastW!=FDW)
						{
							int z=1; if (i==a.Count) z=0;
							sb.Append(FlxConvert.ToString(Convert.ToInt32(LastKey)) +" "+FlxConvert.ToString(Convert.ToInt32(key-z)+" "));
							sb.Append(LastW+TPdfTokens.NewLine);
						}
					}
					else
					{
						if (!ContinueKey)
						{
							if (ArrayOpened) sb.Append(TPdfTokens.GetString(TPdfToken.CloseArray)+TPdfTokens.NewLine);
							ArrayOpened = false;
							sb.Append(Convert.ToInt32(LastKey)+" "+TPdfTokens.GetString(TPdfToken.OpenArray));
							ArrayOpened = true;
						}
						else
							sb.Append(" ");
						sb.Append(LastW);
					}

					ContinueKey = LastKey+1 == key;
					LastKey = key;
					LastI = i;
					LastW = w;
				}
			}
			if (ArrayOpened) sb.Append(TPdfTokens.GetString(TPdfToken.CloseArray)+TPdfTokens.NewLine);
			sb.Append(TPdfTokens.GetString(TPdfToken.CloseArray));
			return sb.ToString();
		}
	}

	/// <summary>
	/// Maps a new glyph to an old one.
	/// </summary>
#if (FRAMEWORK20)
	internal class TGlyphMap :Dictionary<int, int>{
#else
	internal class TGlyphMap : Hashtable
	{
		internal bool TryGetValue(int OldGlyph, out int NewGlyph)
		{
			object obj = this[OldGlyph];
			if (obj == null) {NewGlyph = 0; return false;}

			NewGlyph = (int)obj;
			return true;
		}

#endif

		internal int[] ToList()
		{
			int[] Result = new int[Count];
			foreach (int k in Keys)
			{
				Result[(int)this[k]] = k;
			}

			return Result;
		}
	}

}
