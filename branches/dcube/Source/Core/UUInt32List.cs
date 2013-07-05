using System;
using System.Collections.Generic;

namespace FlexCel.Core 
{
    internal class UInt32List : List<UInt32>
    {
        internal UInt32List() { }
        internal UInt32List(int capacity) : base(capacity) { }
    }

    internal class Int32List : List<Int32>
    {
        internal Int32List() { }
        internal Int32List(int capacity) : base(capacity) { }
    }
}
