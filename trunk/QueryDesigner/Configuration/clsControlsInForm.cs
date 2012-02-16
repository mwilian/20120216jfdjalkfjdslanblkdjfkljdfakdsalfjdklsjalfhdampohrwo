using System;
using System.Data;
using System.Configuration;




namespace QueryDesigner.Configuration
{
    public class clsControlsInForm
    {
        string _id = "";

        public string Id
        {
            get { return _id; }
            set { _id = value; }
        }
        string _type = "";

        public string Type
        {
            get { return _type; }
            set { _type = value; }
        }
        string _index = "";

        public string Index
        {
            get { return _index; }
            set { _index = value; }
        }
        string _languageCode = "";

        public string LanguageCode
        {
            get { return _languageCode; }
            set { _languageCode = value; }
        }
        public clsControlsInForm(string id, string type,string index,string languageCode)
        {
            _id = id;
            _type = type;
            _index = index;
            _languageCode = languageCode;
        }
    }
}
