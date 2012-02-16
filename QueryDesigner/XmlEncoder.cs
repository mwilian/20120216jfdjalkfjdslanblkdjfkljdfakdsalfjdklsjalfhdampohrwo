using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;

namespace QueryDesigner
{
    public class XmlEncoder
    {
        static List<string[]> partern = new List<string[]>();

        public static void Init()
        {
            if (partern.Count == 0)
            {
                partern.Add(new string[] { ">", "%3E" });
                partern.Add(new string[] { "/", "%2F" });
                partern.Add(new string[] { " ", "%20" });
                partern.Add(new string[] { "=", "%3D" });
                partern.Add(new string[] { ":", "%3A" });
                partern.Add(new string[] { ";", "%3B" });
                partern.Add(new string[] { "(", "%28" });
                partern.Add(new string[] { ")", "%29" });
                partern.Add(new string[] { "\"", "%22" });
                partern.Add(new string[] { "\\", "%5C" });
                partern.Add(new string[] { "{", "%7B" });
                partern.Add(new string[] { "&", "%26" });
                partern.Add(new string[] { "}", "%7D" });
                partern.Add(new string[] { "'", "%27" });
            }
        }
        public static string Encode(string input)
        {
            Init();
            foreach (string[] x in partern)
            {
                input = input.Replace(x[0], x[1]);
            }
            return input;
            //return System.Xml.XmlConvert.EncodeName(input);
        }
        public static string Decode(string input)
        {
            Init();
            foreach (string[] x in partern)
            {
                input = input.Replace(x[1], x[0]);
            }
            return input;
        }
    }
}
