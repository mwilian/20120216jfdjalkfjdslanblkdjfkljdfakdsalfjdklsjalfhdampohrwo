using System;
using System.Collections.Generic;

using System.Text;
using System.Globalization;

namespace FlexCel.Pdf
{
#if(FRAMEWORK20)
    internal class TIntKeyList : List<int>
    {
        internal TIntKeyList(ICollection<int> c): base(c)
        {
        }
    }
#else
    internal class TIntKeyList : ArrayList
    {
		internal TIntKeyList(ICollection c): base(c){}
    }
#endif

    /// <summary>
	/// Handles the information needed to write a ToUnicode entry on a pdf file.
	/// </summary>
	internal class TToUnicode
	{
#if(FRAMEWORK20)
		private Dictionary<int, int> FList;
#else
		private Hashtable FList;
#endif
		public TToUnicode()
		{
#if(FRAMEWORK20)
			FList = new Dictionary<int, int>();
#else
			FList = new Hashtable();
#endif
		}

		public void Add(int ccode, int cUnicode)
		{
			if (ccode==0) {FList[0] = 0; return;}  //a non-existing unicode char will be mapped to glyph 0, so cUnicode might not be 0. Without this line, all other non existing chars that appear later would be mapped to cunicode.
			if (FList.ContainsKey(ccode)) return;
			FList[ccode] = cUnicode;
		}

		private static string StrToHex(int n)
		{
			return String.Format(CultureInfo.InvariantCulture, "<{0:X4}>", n);
		}

		public string GetData()
		{
			StringBuilder CharResult = new StringBuilder();
			StringBuilder RangeResult = new StringBuilder();

			TIntKeyList a = new TIntKeyList(FList.Keys);
			a.Sort();
			if (a.Count<=0) 
			{
				return String.Empty;
			}

			int RangeEntries=0;
			int CharEntries=0;
			int FirstKey = Convert.ToInt32(a[0], CultureInfo.InvariantCulture);
			int LastKey = FirstKey;
			int LastUnicode = (int)FList[FirstKey];

			bool OnARoll = true;
			for (int i =1; i<= a.Count; i++)
			{
				int key = i==a.Count? LastKey: (int)a[i];
				int Unicode= i==a.Count?0: (int)FList[key];
                
				bool CanKeepRolling = (Unicode == LastUnicode +1) && ((Unicode &0xFF) !=0) && ((key &0xFF) !=0);
				if (LastKey+1== key &&
					(!OnARoll || key-FirstKey<3 || CanKeepRolling)) 
				{
					LastKey = key;
					if (OnARoll) OnARoll = CanKeepRolling;
					LastUnicode = Unicode;
					continue;
				}

				if (FirstKey<LastKey)
				{
					RangeResult.Append(StrToHex(FirstKey));
					RangeResult.Append(StrToHex(LastKey));
					if (OnARoll)
						RangeResult.Append(StrToHex((int)FList[FirstKey]));
					else
					{
						RangeResult.Append(TPdfTokens.GetString(TPdfToken.OpenArray));
						for (int k=FirstKey; k<= LastKey; k++)
							RangeResult.Append(StrToHex((int)FList[k]));
						RangeResult.Append(TPdfTokens.GetString(TPdfToken.CloseArray));
					}
					RangeResult.Append("\n");
					RangeEntries++;
				}
				else
				{
					CharResult.Append(StrToHex(FirstKey));
					CharResult.Append(StrToHex((int)FList[FirstKey]));
					CharResult.Append("\n");
					CharEntries++;
				}

				LastKey = key;
				FirstKey = key;
				OnARoll = true;
				LastUnicode = Unicode;
			}

			string Result = String.Empty;
			if (CharEntries>0)
			{
				Result+= CharEntries.ToString(CultureInfo.InvariantCulture)+" "+ 
					TPdfTokens.GetString(TPdfToken.beginbfchar)+"\n"+
					CharResult.ToString()+
					TPdfTokens.GetString(TPdfToken.endbfchar)+"\n";
			}

			if (RangeEntries>0)
			{
				Result+= RangeEntries.ToString(CultureInfo.InvariantCulture)+" "+ 
					TPdfTokens.GetString(TPdfToken.beginbfrange)+"\n"+
					RangeResult.ToString()+
					TPdfTokens.GetString(TPdfToken.endbfrange)+"\n";
			}
			return Result;
		}
	}
}
