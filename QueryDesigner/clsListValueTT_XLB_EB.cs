using System;
using System.Collections.Generic;

using System.Text;
using FlexCel.Core;

namespace dCube
{
    public static class  clsListValueTT_XLB_EB
    {        
        static Dictionary<TPoint, object> _values = new Dictionary<TPoint, object>();

        public static Dictionary<TPoint, object> Values
        {
            get { return _values; }
            set { _values = value; }
        }        
    }
}
