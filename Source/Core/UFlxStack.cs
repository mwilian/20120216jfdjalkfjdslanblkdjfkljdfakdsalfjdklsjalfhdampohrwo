using System;
using System.Collections.Generic;

namespace FlexCel.Core
{
    /// <summary>
    /// Specialized string stack used on formula parsing.
    /// </summary>
    internal class TStringStack
    {
        private Stack<String> FList;
         
        public TStringStack()
        {
            FList = new Stack<String>();
        }

        public int Count
        {
            get 
            {
                return FList.Count;
            }
        }

        public bool AtLeast(int ACount)
        {
            return FList.Count>=ACount;
        }
        public virtual void Push(string s)
        {
            FList.Push(s);
        }

        public string Pop()
        {
            return (string) FList.Pop();
        }

        public string Peek()
        {
            return (string) FList.Peek();
        }        
    }

    /// <summary>
    /// A string stack supporting spaces, so we can reconstruct the original formula from the RPN.
    /// </summary>
    internal class TFormulaStack: TStringStack
    {
        public string FmSpaces;
        public string FmPreSpaces;
        public string FmPostSpaces;

        public TFormulaStack()
        {
            FmSpaces="";
            FmPreSpaces="";
            FmPostSpaces="";
        }

        public override void Push(string s)
        {
            base.Push (s);
            FmSpaces="";
            FmPreSpaces="";
            FmPostSpaces="";
        }

    }

    /// <summary>
    /// It holds one whitespace keyword.
    /// </summary>
    internal class TWhiteSpace
    {
        public byte EnterCount;
        public FormulaAttr EnterKind;
        public byte SpaceCount;
        public FormulaAttr SpaceKind;
    }


    /// <summary>
    /// Specialized stack for keeping the whitespace.
    /// </summary>
    internal class TWhiteSpaceStack
    {
#if (FRAMEWORK20)
        private Stack<TWhiteSpace> FList;
#else
        private  Stack FList;
#endif
         
        public TWhiteSpaceStack()
        {
#if (FRAMEWORK20)
            FList = new Stack<TWhiteSpace>();
#else
            FList=new Stack();
#endif
        }

        public int Count
        {
            get 
            {
                return FList.Count;
            }
        }

        public bool AtLeast(int ACount)
        {
            return FList.Count >= ACount;
        }

        public virtual void Push(TWhiteSpace s)
        {
            FList.Push(s);
        }

        public TWhiteSpace Pop()
        {
            return (TWhiteSpace) FList.Pop();
        }

        internal void NormalizeLastWhiteSpace()
        {
            TWhiteSpace Ws = (TWhiteSpace)FList.Peek();
            Ws.EnterKind = FormulaAttr.bitFEnter;
            Ws.SpaceKind = FormulaAttr.bitFSpace;
        }
    }

}
