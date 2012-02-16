using System;
using System.Collections.Generic;

namespace FlexCel.Core
{
    internal struct TFunctionKey
    {
        public int Index;
        public ptg PtgId;
        public int ArgCount;

        public TFunctionKey(int aIndex, ptg aPtgId, int aArgCount)
        {
            Index = aIndex;
            PtgId = aPtgId;
            ArgCount = aArgCount;
        }
    }

    /// <summary>
    /// A cache of used Function tokens so we don't create them each time.
    /// </summary>
    internal sealed class FunctionCache
    {
        [ThreadStatic] //So we don't need to lock() access to its members. We can't initialize threadstatic members.
#if (FRAMEWORK20)
        private static Dictionary<TFunctionKey, TBaseFunctionToken> FList;  //STATIC*
#else
		private static Hashtable FList;  //STATIC*
#endif

        private FunctionCache()
        {
        }

        private static void EnsureFList()
        {
#if (FRAMEWORK20)
            if (FList == null) FList = new Dictionary<TFunctionKey, TBaseFunctionToken>();
#else
			if (FList == null) FList = new Hashtable();
#endif
        }

        internal static bool TryGetValue(int Index, ptg PtgId, int ArgCount, out TBaseFunctionToken ResultValue)
        {
            EnsureFList();
#if (FRAMEWORK20)
            return FList.TryGetValue(new TFunctionKey(Index, PtgId, ArgCount), out ResultValue);
#else
            ResultValue = FList[GetHash(Index, PtgId, ArgCount)] as TBaseFunctionToken;
			return ResultValue != null;
#endif
        }

        internal static void Add(int Index, ptg PtgId, int ArgCount, TBaseFunctionToken ResultValue)
        {
            EnsureFList();
            if (FList.Count > 1024) FList.Clear(); //Shouldn't happen, but just in case avoid this growing ad infinitum.
            FList.Add(new TFunctionKey(Index, PtgId, ArgCount), ResultValue);
        }

    }
}
