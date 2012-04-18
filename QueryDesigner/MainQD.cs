using System;
using System.Collections.Generic;
using System.Text;
using System.Collections;

namespace dCube
{
    class MainQD
    {
        public static string MessageCaption = "Universal Query Designer";
        public static string GetColumnFriendlyName(Janus.Windows.GridEX.GridEXColumn column)
        {
            if (column.Caption.Length == 0)
            {
                if (column.Tag != null)
                {
                    return System.Convert.ToString(column.Tag);
                }
                else
                {
                    return column.Key;
                }
            }
            else
            {
                return column.Caption;
            }
        }
        public static string GetCustomGroupName(ICollection groupRows, string objName, int index)
        {

            string propName = objName + index;
            foreach (Janus.Windows.GridEX.GridEXCustomGroup customGroup in groupRows)
            {
                if (string.Compare(customGroup.Key, propName, true) == 0)
                {
                    return GetCustomGroupName(groupRows, objName, index + 1);
                }
            }

            return propName;

        }
        public static string GetCustomGroupRowName(ICollection groupRows, string objName, int index)
        {

            string propName = objName + index;
            foreach (Janus.Windows.GridEX.GridEXCustomGroupRow groupRow in groupRows)
            {
                if (string.Compare(groupRow.GroupCaption, propName, true) == 0)
                {
                    return GetCustomGroupRowName(groupRows, objName, index + 1);
                }
            }

            return propName;

        }
    }
}
