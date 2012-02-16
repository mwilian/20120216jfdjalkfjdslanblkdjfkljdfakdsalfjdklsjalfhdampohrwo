using System;
using System.Collections.Generic;

namespace FlexCel.Core
{
	/// <summary>
	/// A small class to do pattern matching using widcards.
	/// </summary>
	internal sealed class TWildcardMatch
	{
		private TWildcardMatch() {}

		/// <summary>
		/// Returns all matches of pattern inside s, considering the special "?" character.
		/// </summary>
		/// <param name="pattern"></param>
		/// <param name="s"></param>
		/// <param name="start_s"></param>
		/// <param name="ProtectedChars"></param>
		/// <returns></returns>
		private static TIndexPosList LiteralFindIndexes(string pattern, string s, int start_s, bool[] ProtectedChars)
		{
			TIndexPosList Result = new TIndexPosList();

			for (int k=start_s; k<s.Length-pattern.Length+1; k++)
			{
				bool Found = true;
				for (int i=0; i<pattern.Length; i++)
				{
					if ((ProtectedChars[i] || pattern[i]!='?') && pattern[i]!=s[k+i])
					{
						Found = false;
						break;
					}
				}
				if (Found)
				{
					Result.Add(new TIndexPos(k, k+pattern.Length-1));
				}
			}
			return Result;
		}

		/// <summary>
		/// This will find pattern on s considering wildcards.
		/// </summary>
		/// <param name="pattern"></param>
		/// <param name="s"></param>
		/// <param name="start_s"></param>
		/// <returns></returns>
		private static TIndexPosList FindIndexes(string pattern, string s, int start_s)
		{
			bool[] ProtectedChars = new bool[pattern.Length];
			int i=0; 
			while (i<pattern.Length)
			{
				if (pattern[i]=='~')
				{
					ProtectedChars[i]=true;
					pattern = pattern.Remove(i,1);
					i++;
					continue;
				}

				if (pattern[i]=='*')
				{
					TIndexPosList Pos1 = LiteralFindIndexes(pattern.Substring(0, i), s, start_s, ProtectedChars);  //there are no * or ~ here.
					TIndexPosList Pos2 = FindIndexes(pattern.Substring(i+1), s, start_s);

					TIndexPosList Result = new TIndexPosList();
					for (int p1=0;p1<Pos1.Count;p1++)
					{
						for (int p2=0;p2<Pos2.Count;p2++)
						{
                            if (Pos1[p1].Last<Pos2[p2].First)
								Result.Add(new TIndexPos(Pos1[p1].First, Pos2[p2].Last));
						}
					}
					return Result;
				}

				i++;
			}

			return LiteralFindIndexes(pattern, s, start_s, ProtectedChars);
		}

		internal static int IndexOf(string pattern, string s, int start)
		{
			TIndexPosList Pos = FindIndexes(pattern, s, start);
			if (Pos.Count<=0) return -1;
			return Pos[0].First;
		}

		internal static bool Matches(string pattern, string s)
		{
			TIndexPosList Pos = FindIndexes(pattern, s, 0);
			for (int i=0; i<Pos.Count; i++)
			{
				if(Pos[i].First>0) return false; //list is ordered.
				if (Pos[i].Last==s.Length-1) return true; //full match
			}

			return false;
		}
	}

	internal class TIndexPos
	{
		internal int First;
		internal int Last;

		internal TIndexPos(int aFirst, int aLast)
		{
			First = aFirst;
			Last = aLast;
		}

	}

#if (FRAMEWORK20)
	internal class TIndexPosList: List<TIndexPos>
	{
	}
#else
	internal class TIndexPosList: ArrayList
	{
		public new TIndexPos this[int index]
		{
			get 
			{
				return (TIndexPos)base[index];
			}
		}
	}
#endif
}
