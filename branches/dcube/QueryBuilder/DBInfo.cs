//using pbs.Helper;
using System;
using System.Data;
using System.Data.SqlClient;

using System.Collections;
using System.Collections.Generic;

using System.Diagnostics;

namespace QueryBuilder
{
    //  Class DBInfo
    [Serializable()]
    public class DBInfo
    {

        #region '" Business Properties and Methods "'

        //  declare members
        private string _dbId = string.Empty;
        private string _description = string.Empty;
        private string _dataAccessGroup = string.Empty;

        private string _dateFormat = string.Empty;
        private string _referenceFileDrive = string.Empty;
        private string _ledgerFileDrive = string.Empty;
        private string _salesOrderFileDrive = string.Empty;
        private string _purchaseOrderFileDrive = string.Empty;
        private string _inventoryFileDrive = string.Empty;
        private string _backupFileDrive = string.Empty;
        private string _printFileDrive = string.Empty;
        private string _referenceCreated = string.Empty;
        private string _ledgerCreated = string.Empty;
        private string _decimalPlaces = "2";
        private string _decimalSeparator = ".";
        private string _thousandSeparator = ",";
        private string _primaryBudget = string.Empty;
        private string _commitmentLedger = string.Empty;

        private string _sunbusinessDecimalPlaces = string.Empty;

        //  Property DbId
        [System.ComponentModel.DataObjectField(true, false)]
        public string DbId
        {
            get
            {
                return _dbId;
            }
        }
        //  Property Code
        public string Code
        {
            get
            {
                return _dbId;
            }
        }
        //  Property LookUp
        public string LookUp
        {
            get
            {
                return string.Empty;
            }
        }

        //  Property Description
        public string Description
        {
            get
            {
                return _description;
            }
        }

        //  Property DataAccessGroup
        public string DataAccessGroup
        {
            get
            {
                return _dataAccessGroup;
            }
        }

        //  Property DateFormat
        public string DateFormat
        {
            get
            {
                switch (_dateFormat.Trim().ToUpper())
                {
                    case "B":
                        return "dd/MM/yyyy";
                    case "A":
                        return "MM/dd/yyyy";
                    default:
                        return "yyyy/MM/dd";
                }

            }
        }

        //  Property ReferenceFileDrive
        public string ReferenceFileDrive
        {
            get
            {
                return _referenceFileDrive;
            }
        }

        //  Property LedgerFileDrive
        public string LedgerFileDrive
        {
            get
            {
                return _ledgerFileDrive;
            }
        }

        //  Property BackupFileDrive
        public string BackupFileDrive
        {
            get
            {
                return _backupFileDrive;
            }

        }

        //  Property PrintFileDrive
        public string PrintFileDrive
        {
            get
            {
                return _printFileDrive;
            }
        }

        //  Property ReferenceCreated
        public bool ReferenceCreated
        {
            get
            {
                return _referenceCreated.Trim() == "Y";
            }

        }

        //  Property LedgerCreated
        public string LedgerCreated
        {
            get
            {
                return _ledgerCreated;
            }
        }

        //  Property DecimalPlaces
        public int DecimalPlaces
        {
            get
            {

                return (Int32.Parse((Security.NZ(_decimalPlaces, "2"))));
            }
        } //  default is 2
        //  Property BDecimalPlaces
        public int BDecimalPlaces
        {
            get
            {

                return (Int32.Parse((Security.NZ(this._sunbusinessDecimalPlaces, "5"))));
            }
        } //  default is 2

        //  Property DecimalSeparator
        public string DecimalSeparator
        {
            get
            {
                return Security.NZ(_decimalSeparator, ".");
            }
        }

        //  Property BaseAmountFormatString
        public string BaseAmountFormatString
        {
            get
            {
                string pattern = "d";
                switch (DecimalPlaces)
                {
                    case 0:
                        pattern = "#,###;#,###;-";
                        break;
                    default:
                        pattern = "#,###." + "".PadLeft(DecimalPlaces, '0') + ";#,###." + "".PadLeft(DecimalPlaces, '0') + ";-";
                        break;
                }

                return pattern;
            }
        }

        //  Property ThousandSeparator
        public string ThousandSeparator
        {
            get
            {
                return Security.NZ(_thousandSeparator, ",");
            }
        }

        //  Property PrimaryBudget
        public string PrimaryBudget
        {
            get
            {
                return _primaryBudget;
            }

        }

        //  Property CommitmentLedger
        public string CommitmentLedger
        {
            get
            {
                return _commitmentLedger;
            }
        }


        #endregion //  Business Properties and Methods

        #region '" Factory Methods "'

        //  Method GetDBInfo
        public static DBInfo GetDBInfo(SqlDataReader dr)
        {
            return new DBInfo(dr);
        }


        //  Method EmptyDBInfo
        public static DBInfo EmptyDBInfo()
        {
            return new DBInfo();
        }


        private DBInfo()
        {
            //_dbId = pbs.Helper.SystemDTB; 
        }

        private DBInfo(SqlDataReader dr)
        {
            Fetch(dr);
        }

        #endregion //  Factory Methods

        #region '" Data Access "'

        #region '" Data Access - Fetch "'

        //  Method Fetch
        private void Fetch(SqlDataReader dr)
        {
            FetchObject(dr);

        }


        //  Method FetchObject
        private void FetchObject(SqlDataReader dr)
        {
            _dbId = dr["DB"].ToString().Trim();
            //_dbId = dr["DB_ID"].ToString().Trim();

            //_description = dr["Description"].ToString().Trim();
            _description = dr["DESCRIPTION"].ToString().Trim(); 
            //_dataAccessGroup = dr["Data_Access_Group"].ToString().Trim();


            //_dateFormat = dr["Date_Format"].ToString().Trim();

            //_referenceFileDrive = dr["Reference_File_Drive"].ToString().Trim();

            //_ledgerFileDrive = dr["Ledger_File_Drive"].ToString().Trim();

            //_salesOrderFileDrive = dr["Sales_Order_File_Drive"].ToString().Trim();

            //_purchaseOrderFileDrive = dr["Purchase_Order_File_Drive"].ToString().Trim();

            //_inventoryFileDrive = dr["Inventory_File_Drive"].ToString().Trim();

            //_backupFileDrive = dr["Backup_File_Drive"].ToString().Trim();

            //_printFileDrive = dr["Print_File_Drive"].ToString().Trim();

            //_referenceCreated = dr["Reference_Created"].ToString().Trim();

            //_ledgerCreated = dr["Ledger_Created"].ToString().Trim();


            //_decimalPlaces = dr["Decimal_Places"].ToString().Trim();

            //_decimalSeparator = dr["Decimal_Separator"].ToString().Trim();

            //_thousandSeparator = dr["Thousand_Separator"].ToString().Trim();

            //_primaryBudget = dr["Primary_Budget"].ToString().Trim();

            //_commitmentLedger = dr["Commitment_Ledger"].ToString().Trim();

            //_sunbusinessDecimalPlaces = dr["SunBusiness_Decimal_Places"].ToString().Trim();
        }



        #endregion //  Data Access - Fetch

        #endregion

    }



}
