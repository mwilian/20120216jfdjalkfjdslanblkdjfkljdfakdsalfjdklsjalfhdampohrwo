using System;
using System.Collections.Generic;
using System.Text;

namespace DTO
{
    public class License
    {
        private string _companyName;
        private string _serialNumber;
        private int _numUsers;
        private string _modules;
        private int _expiryDate;
        private string _key;
        private string _serialCPU;

        public string SerialCPU
        {
            get { return _serialCPU; }
            set { _serialCPU = value; }
        }
        public string CompanyName
        {
            get { return _companyName; }
            set { _companyName = value; }
        }
        public string SerialNumber
        {
            get { return _serialNumber; }
            set { _serialNumber = value; }
        }
        public int NumUsers
        {
            get { return _numUsers; }
            set { _numUsers = value; }
        }
        public string Modules
        {
            get { return _modules; }
            set { _modules = value; }
        }
        public int ExpiryDate
        {
            get { return _expiryDate; }
            set { _expiryDate = value; }
        }
        public string Key
        {
            get { return _key; }
            set { _key = value; }
        }
    }
}
