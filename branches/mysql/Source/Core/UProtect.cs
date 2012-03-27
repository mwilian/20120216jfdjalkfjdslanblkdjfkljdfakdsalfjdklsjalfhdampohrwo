using System;
using System.Reflection;
using System.Globalization;
using System.Text;
using System.Collections.Generic;

namespace FlexCel.Core
{
    /// <summary>
    /// How the file is encrypted. This applies only to xls files. Xlsx files are encrypted using the Agile xlsx encryption.
    /// </summary>
    public enum TEncryptionType
    {
        /// <summary>
        /// Excel 95 XOR encryption.
        /// </summary>
        Xor,
        /// <summary>
        /// Excel 97/2000 encryption.
        /// </summary>
        Standard,
        /// <summary>
        /// Excel XP/2003 encryption.
        /// </summary>
        Strong
    }

    /// <summary>
    /// Encryption algorithms supported in xlsx encrypted files.
    /// </summary>
    public enum TEncryptionAlgorithm
    {
        /// <summary>
        /// AES 128. This is the default in Excel 2007/2010
        /// </summary>
        AES_128,

        /// <summary>
        /// AES 192
        /// </summary>
        AES_192,

        /// <summary>
        /// AES 256
        /// </summary>
        AES_256
    }


    /// <summary>
    /// Use this class to supply a password to open an encrypted file.
    /// </summary>
    public class OnPasswordEventArgs: EventArgs
    {
        private string FPassword;
        private ExcelFile FXls;

        /// <summary>
        /// Creates a new Argument.
        /// </summary>
        public OnPasswordEventArgs(ExcelFile aXls)
        {
            FXls=aXls;
        }

        /// <summary>
        /// The password needed to open the file.
        /// </summary>
        public string Password {get {return FPassword;} set {FPassword=value;}}

		/// <summary>
		/// Excel file we are trying to open.
		/// </summary>
		public ExcelFile Xls {get {return FXls;}}
    }

    /// <summary>
    /// This event fires when opening an encrypted file and no password has been supplied.
    /// </summary>
    public delegate void OnPasswordEventHandler(OnPasswordEventArgs e);



    /// <summary>
    /// Options for protecting the workbook.
    /// </summary>
    public sealed class TWorkbookProtectionOptions
    {
        private bool FWindow;
        private bool FStructure;

        /// <summary>
        /// Initializes empty protection.
        /// </summary>
        public TWorkbookProtectionOptions(){}
        /// <summary>
        /// Initializes the class.
        /// </summary>
        /// <param name="aStructure">True if the structure will be protected.</param>
        /// <param name="aWindow">True if the window will be protected.</param>
        public TWorkbookProtectionOptions(bool aWindow, bool aStructure)
        {
            Window=aWindow;
            Structure=aStructure;
        }

        /// <summary>
        /// Window is protected.
        /// </summary>
        public bool Window {get {return FWindow;} set {FWindow=value;}}

        /// <summary>
        /// Structure is protected.
        /// </summary>
        public bool Structure {get {return FStructure;} set {FStructure=value;}}

    }

    /// <summary>
    /// Options for protecting the change list in a shared workbook. In Excel you can change this settings in Protection-&gt;Protect Shared Workbook.
    /// </summary>
    public sealed class TSharedWorkbookProtectionOptions
    {
        private bool FSharingWithTrackChanges;

        /// <summary>
        /// Initializes empty protection.
        /// </summary>
        public TSharedWorkbookProtectionOptions() { }
        
        /// <summary>
        /// Initializes the class.
        /// </summary>
        /// <param name="aSharingWithTrackChanges">True to protect the change history from being removed.</param>
        public TSharedWorkbookProtectionOptions(bool aSharingWithTrackChanges)
        {
            SharingWithTrackChanges = aSharingWithTrackChanges;
        }

        /// <summary>
        /// True to protect the change history from being removed.
        /// </summary>
        public bool SharingWithTrackChanges { get { return FSharingWithTrackChanges; } set { FSharingWithTrackChanges = value; } }

    }

	/// <summary>
	/// Indicates how a sheet will be protected.
	/// </summary>
	public enum TProtectionType
	{
		/// <summary>
		/// All things will be protected.
		/// </summary>
		All,

		/// <summary>
		/// Nothing will be protected.
		/// </summary>
		None
	}
    /// <summary>
    /// Options for protecting a sheet.
    /// </summary>
    public sealed class TSheetProtectionOptions
    {
        private bool FContents;
        private bool FObjects;
        private bool FScenarios;

        private bool FCellFormatting;
        private bool FColumnFormatting;
        private bool FRowFormatting;
        private bool FInsertColumns;
        private bool FInsertRows;
        private bool FInsertHyperlinks;
        private bool FDeleteColumns;
        private bool FDeleteRows;
        private bool FSelectLockedCells;
        private bool FSortCellRange;
        private bool FEditAutoFilters;
        private bool FEditPivotTables;
        private bool FSelectUnlockedCells;

        /// <summary>
        /// Initializes a not protected block.
        /// </summary>
        public TSheetProtectionOptions()
        {
        }

        /// <summary>
        /// Creates a protected or unprotected sheet.
        /// </summary>
        /// <param name="allTrue">If true, all properties on this class will be set to true.
        /// This means cells on the sheet will be protected, and things like Cell formatting will be not. 
        /// If false, all properties will be set to false.</param>
        public  TSheetProtectionOptions(bool allTrue)
        {

            FContents = allTrue;
            FObjects = allTrue;
            FScenarios = allTrue;

            FCellFormatting = allTrue;
            FColumnFormatting = allTrue;
            FRowFormatting = allTrue;
            FInsertColumns = allTrue;
            FInsertRows = allTrue;
            FInsertHyperlinks = allTrue;
            FDeleteColumns = allTrue;
            FDeleteRows = allTrue;
            FSelectLockedCells = allTrue;
            FSortCellRange = allTrue;
            FEditAutoFilters = allTrue;
            FEditPivotTables = allTrue;
            FSelectUnlockedCells = allTrue;   
        }

		/// <summary>
		/// Creates a protected or unprotected sheet.
		/// </summary>
		/// <param name="protectAll">If true, all things in the sheet will be protected. this means a property like <see cref="Contents"/> 
		/// will be true, and others like <see cref="CellFormatting"/>  will be false.
		/// If false, nothing will.</param>
		public TSheetProtectionOptions(TProtectionType protectAll)
		{
			bool p = protectAll == TProtectionType.All;

			FContents = p;
			FObjects = p;
			FScenarios = p;

			FCellFormatting = !p;
			FColumnFormatting = !p;
			FRowFormatting = !p;
			FInsertColumns = !p;
			FInsertRows = !p;
			FInsertHyperlinks = !p;
			FDeleteColumns = !p;
			FDeleteRows = !p;
			FSelectLockedCells = !p;
			FSortCellRange = !p;
			FEditAutoFilters = !p;
			FEditPivotTables = !p;
			FSelectUnlockedCells = !p;
		}

        /// <summary>
        /// Sheet contents are protected
        /// </summary>
        public bool Contents {get {return FContents;} set {FContents=value;}}

        /// <summary>
        /// Objects on the sheet are protected.
        /// </summary>
        public bool Objects {get {return FObjects;} set {FObjects=value;}}

        /// <summary>
        /// Scenarios on the sheet are protected.
        /// </summary>
        public bool Scenarios {get {return FScenarios;} set {FScenarios=value;}}

        /// <summary>If TRUE, users are allowed to change this setting. Set it to FALSE to disable this property. Only on Excel &gt;= XP.</summary>
        public bool CellFormatting {get {return FCellFormatting;} set {FCellFormatting=value;}}
		/// <summary>If TRUE, users are allowed to change this setting. Set it to FALSE to disable this property. Only on Excel &gt;= XP.</summary>
		public bool ColumnFormatting {get {return FColumnFormatting;} set {FColumnFormatting=value;}}
		/// <summary>If TRUE, users are allowed to change this setting. Set it to FALSE to disable this property. Only on Excel &gt;= XP.</summary>
		public bool RowFormatting {get {return FRowFormatting;} set {FRowFormatting=value;}}
		/// <summary>If TRUE, users are allowed to change this setting. Set it to FALSE to disable this property. Only on Excel &gt;= XP.</summary>
		public bool InsertColumns {get {return FInsertColumns;} set {FInsertColumns=value;}}
		/// <summary>If TRUE, users are allowed to change this setting. Set it to FALSE to disable this property. Only on Excel &gt;= XP.</summary>
		public bool InsertRows {get {return FInsertRows;} set {FInsertRows=value;}}
		/// <summary>If TRUE, users are allowed to change this setting. Set it to FALSE to disable this property. Only on Excel &gt;= XP.</summary>
		public bool InsertHyperlinks {get {return FInsertHyperlinks;} set {FInsertHyperlinks=value;}}
		/// <summary>If TRUE, users are allowed to change this setting. Set it to FALSE to disable this property. Only on Excel &gt;= XP.</summary>
		public bool DeleteColumns {get {return FDeleteColumns;} set {FDeleteColumns=value;}}
		/// <summary>If TRUE, users are allowed to change this setting. Set it to FALSE to disable this property. Only on Excel &gt;= XP.</summary>
		public bool DeleteRows {get {return FDeleteRows;} set {FDeleteRows=value;}}
		/// <summary>If TRUE, users are allowed to change this setting. Set it to FALSE to disable this property. Only on Excel &gt;= XP.</summary>
		public bool SelectLockedCells {get {return FSelectLockedCells;} set {FSelectLockedCells=value;}}
		/// <summary>If TRUE, users are allowed to change this setting. Set it to FALSE to disable this property. Only on Excel &gt;= XP.</summary>
		public bool SortCellRange {get {return FSortCellRange;} set {FSortCellRange=value;}}
		/// <summary>If TRUE, users are allowed to change this setting. Set it to FALSE to disable this property. Only on Excel &gt;= XP.</summary>
		public bool EditAutoFilters {get {return FEditAutoFilters;} set {FEditAutoFilters=value;}}
		/// <summary>If TRUE, users are allowed to change this setting. Set it to FALSE to disable this property. Only on Excel &gt;= XP.</summary>
		public bool EditPivotTables {get {return FEditPivotTables;} set {FEditPivotTables=value;}}
		/// <summary>If TRUE, users are allowed to change this setting. Set it to FALSE to disable this property. Only on Excel &gt;= XP.</summary>
		public bool SelectUnlockedCells {get {return FSelectUnlockedCells;} set {FSelectUnlockedCells=value;}}

    }


    /// <summary>
    /// Encryption data for an Excel sheet.
    /// </summary>
    public sealed class TProtection
    {
        private TEncryptionType FEncryptionType;
        private TEncryptionAlgorithm FEncryptionAlgorithmXlsx;
        private string FOpenPassword;
        private ExcelFile FXls;
        //private TProtectedRangeList FUserProtectedRanges;
        private OnPasswordEventHandler FOnPassword;


        /// <summary>
        /// Creates a new TProtection instance and initializes it.
        /// </summary>
        internal TProtection(ExcelFile xls)
        {
            EncryptionType = TEncryptionType.Standard;
            EncryptionAlgorithmXlsx = TEncryptionAlgorithm.AES_128;
            FXls = xls;
        }

        /// <summary>
        /// Encryption mode for xls files .
        /// </summary>
        public TEncryptionType EncryptionType {get {return FEncryptionType;} set {FEncryptionType=value;}}

        /// <summary>
        /// Encryption algorithm for xlsx files.
        /// </summary>
        public TEncryptionAlgorithm EncryptionAlgorithmXlsx { get { return FEncryptionAlgorithmXlsx; } set { FEncryptionAlgorithmXlsx = value; } }


        /// <summary>
        /// Sets the password to open the file. When set, the file will be encrypted. On Excel go to Options->Security to check it.
        /// Set this to null to clear it.
        /// </summary>
        public string OpenPassword
        {
            get
            {
                if (FOpenPassword ==null) return String.Empty; else return FOpenPassword;
            }
            set
            {
                FOpenPassword=value;
            }
        }

        /// <summary>
        /// Sets the password for modifying the file. It won't encrypt the file, it just won't let Excel save the file. On Excel goto Options->Security to check it.
        /// Note that you can only set it, there is no way to retrieve an existing password.
        /// </summary>
        /// <param name="modifyPassword">The new password. Set it to null to clear it.</param>
        /// <param name="recommendReadOnly">When true, Excel will recommend read only when opening a file.</param>
        /// <param name="reservingUser">The user that reserves the file. It will appear on the password dialog on Excel.</param>
        public void SetModifyPassword(string modifyPassword, bool recommendReadOnly, string reservingUser)
        {
            FXls.SetModifyPassword(modifyPassword, recommendReadOnly, reservingUser);
        }

        /// <summary>
        /// Returns true if the file has a password to modify.
        /// </summary>
        public bool HasModifyPassword
        {
            get 
            {
                return FXls.HasModifyPassword;
            }
        }

        /// <summary>
        /// Returns true if the file is recommended to open read-only.
        /// </summary>
        public bool RecommendReadOnly
        {
            get 
            {
                return FXls.RecommendReadOnly;
            }
            set
            {
                FXls.RecommendReadOnly = value;
            }
        }

        /// <summary>
        /// Protects the workbook. On Excel goto Protect->Workbook to check it.
        /// </summary>
        /// <param name="workbookPassword">Password to protect the file. You can set it to null to clear it.</param>
        /// <param name="workbookProtectionOptions">The options to protect.</param>
        public void SetWorkbookProtection (string workbookPassword, TWorkbookProtectionOptions workbookProtectionOptions)
        {
            FXls.SetWorkbookProtection(workbookPassword, workbookProtectionOptions);
        }


        /// <summary>
        /// Returns true if the workbook is protected with a password.
        /// </summary>
        public bool HasWorkbookPassword
        {
            get 
            {
                return FXls.HasWorkbookPassword;
            }
        }

        /// <summary>
        /// Reads the Workbook protection options for a file.
        /// </summary>
        public TWorkbookProtectionOptions GetWorkbookProtectionOptions()
        {
            return FXls.WorkbookProtectionOptions;
        }

        /// <summary>
        /// Sets the workbook protection options for a file.
        /// </summary>
        /// <param name="value">Workbook protection options.</param>
        public void SetWorkbookProtectionOptions(TWorkbookProtectionOptions value)
        {
            FXls.WorkbookProtectionOptions = value;
        }

        /// <summary>
        /// Protects the change history from being removed. On Excel goto Protect->Protect Shared Workbook to check it.
        /// </summary>
        /// <param name="sharedWorkbookPassword">Password to protect this setting. You can set it to null to clear it.</param>
        /// <param name="sharedWorkbookProtectionOptions">The options to protect.</param>
        public void SetSharedWorkbookProtection(string sharedWorkbookPassword, TSharedWorkbookProtectionOptions sharedWorkbookProtectionOptions)
        {
            FXls.SetSharedWorkbookProtection(sharedWorkbookPassword, sharedWorkbookProtectionOptions);
        }


        /// <summary>
        /// Returns true if the change history is protected with a password.
        /// </summary>
        public bool HasSharedWorkbookPassword
        {
            get
            {
                return FXls.HasSharedWorkbookPassword;
            }
        }

        /// <summary>
        /// Reads the protection options for the change history.
        /// </summary>
        public TSharedWorkbookProtectionOptions GetSharedWorkbookProtectionOptions()
        {
            return FXls.SharedWorkbookProtectionOptions;
        }

        /// <summary>
        /// Sets the change history protection options for a file.
        /// </summary>
        /// <param name="value">Protection options.</param>
        public void SetSharedWorkbookProtectionOptions(TSharedWorkbookProtectionOptions value)
        {
            FXls.SharedWorkbookProtectionOptions = value;
        }

        /// <summary>
        /// Protects a sheet. On Excel goto Protect->Sheet to check it.
        /// </summary>
        /// <param name="sheetPassword">Password to protect the active sheet. You can set it to null to clear it.</param>
        /// <param name="sheetProtectionOptions">The options to protect.</param>
        public void SetSheetProtection (string sheetPassword, TSheetProtectionOptions sheetProtectionOptions)
        {    
            FXls.SetSheetProtection(sheetPassword, sheetProtectionOptions);
        }

        /// <summary>
        /// Returns true if the active sheet is protected with a password.
        /// </summary>
        public bool HasSheetPassword
        {
            get 
            {
                return FXls.HasSheetPassword;
            }
        }

        /// <summary>
        /// Return the sheet protection options for the file.
        /// </summary>
        public TSheetProtectionOptions GetSheetProtectionOptions()
        {
            return FXls.SheetProtectionOptions;
        }

        /// <summary>
        /// Sets the sheet protection options for the file.
        /// </summary>
        /// <param name="value">Protection options.</param>
        public void SetSheetProtectionOptions(TSheetProtectionOptions value)
        {
            FXls.SheetProtectionOptions = value;
        }


        /// <summary>
        /// It is called when opening a password protected file, so you can supply the correct password.
        /// If you know beforehand that the file is protected you do not need this event, just use the <see cref="OpenPassword"/> method on this object.
        /// </summary>
        public OnPasswordEventHandler OnPassword {get {return FOnPassword;} set {FOnPassword=value;}}

        internal ExcelFile Xls  {get{return FXls;}}


        /// <summary>
        /// Reads or sets the user writing the file. Useful to know which user opened the file in Excel when
        /// you want to save and the file is in use.
        /// </summary>
        public string WriteAccess
        {
            get 
            {
                return FXls.WriteAccess;
            }
            set
            {
                FXls.WriteAccess = value;
            }
        }

        //public TProtectedRangeList UserProtectedRanges { get { return FUserProtectedRanges; } }
    }

    /// <summary>
    /// Specifies a protected range in a sheet. You can define those ranges in Excel 2007 by going to "Review" tab and selecting "Allow Users to Edit Ranges"
    /// In Excel 2003, they are available under "Menu->Tools->Protection".
    /// </summary>
    public sealed class TProtectedRange
    {
        private string FPassword;
        private TXlsCellRange[] FRanges;
        private string FName;
        private string FPasswordHash;

        /// <summary>
        /// Creates an empty protected range.
        /// </summary>
        public TProtectedRange()
        {
        }

        /// <summary>
        /// Creates and initializes a protected range.
        /// </summary>
        /// <param name="aName">Name for the protected range.</param>
        /// <param name="aPassword">Password to modify the cells. Keep it null if you don't want to set a password.</param>
        /// <param name="aRanges">Ranges of cells where this protected range applies. You might pass an array of TXlsCellRanges here or just a single TXlsCellRange object.</param>
        public TProtectedRange(string aName, string aPassword, params TXlsCellRange[] aRanges)
        {
            Password = aPassword;
            FName = aName;
            Ranges = aRanges; //Not FRanges, we want to go through the setter.
        }

        /// <summary>
        /// Password used to protect the range. Use empty or null to have no password. <b>Note:</b> As this password is not saved in the file,
        /// when you open a file this property will be empty. You can know if a file has a password by looking at <see cref="PasswordHash"/>. 
        /// Setting this property will clear the <see cref="PasswordHash"/> property.
        /// </summary>
        public string Password
        {
            get { return FPassword; }
            set
            {
                FPassword = value;
                FPasswordHash = null;
            }
        }        
        
        /// <summary>
        /// This is the hash for the password that is stored in the file.You shouldn't set this property directly unless you are copying the hash from other place.
        /// When you set this property, <see cref="Password"/> will be reset.
        /// </summary>
        public string PasswordHash
        {
            get { return FPasswordHash; }
            set
            {
                FPasswordHash = value;
                FPassword = null;
            }
        }

        /// <summary>
        /// Ranges of cells this protection applies to. You can specify more than one range of cells for the same ProtectedRange.
        /// </summary>
        public TXlsCellRange[] Ranges { get { return FRanges; } 
            set 
            {
                if (value == null) FlxMessages.ThrowException(FlxErr.ErrNullParameter, "aRanges");
                FRanges = new TXlsCellRange[value.Length];
                for (int i = 0; i < FRanges.Length; i++)
                {
                    if (value[i] == null) FlxMessages.ThrowException(FlxErr.ErrNullParameter, "aRanges[" + i.ToString() + "]");
                    FRanges[i] = (TXlsCellRange)value[i].Clone();
                }
            } 
        }

        /// <summary>
        /// Name of the protected range.
        /// </summary>
        public string Name { get { return FName; } set { FName = value; } }

        /// <summary>
        /// Returns a deep copy of the object.
        /// </summary>
        /// <returns></returns>
        public TProtectedRange Clone()
        {
            TProtectedRange Result = (TProtectedRange)MemberwiseClone();
            if (Ranges != null) Result.Ranges = Ranges;//will clone
            return Result;
        }
    }

    /// <summary>
    /// A list of protected ranges where the user can edit a protected sheet.
    /// </summary>
    public sealed class TProtectedRangeList
    {
        private Dictionary<string, TProtectedRange> FDict;
        private List<TProtectedRange> FList;

        /// <summary>
        /// Creates a new TProtectedRangeList object.
        /// </summary>
        public TProtectedRangeList()
        {
            FDict = new Dictionary<string, TProtectedRange>(StringComparer.OrdinalIgnoreCase);
            FList = new List<TProtectedRange>();
        }

        /// <summary>
        /// Clears the list.
        /// </summary>
        public void Clear()
        {
            FList.Clear();
            FDict.Clear();
        }

        /// <summary>
        /// Adds a new Protected range to the list.
        /// </summary>
        /// <param name="range">Range to add.</param>
        public void Add(TProtectedRange range)
        {
            TProtectedRange range2 = Validate(range, -1);
            FList.Add(range2);
            FDict.Add(range2.Name, range2);
        }

        private TProtectedRange Validate(TProtectedRange range, int index)
        {
            if (range == null) FlxMessages.ThrowException(FlxErr.ErrNullParameter, "range");
            if (string.IsNullOrEmpty(range.Name)) FlxMessages.ThrowException(FlxErr.ErrNullParameter, "range.Name");
            if (!ValidateName(range.Name)) FlxMessages.ThrowException(FlxErr.ErrInvalidName, range.Name);
            if (index < 0)
            {
                if (FDict.ContainsKey(range.Name)) FlxMessages.ThrowException(FlxErr.ErrRangeNameAlreadyExists, range.Name);
            }
            else
            {
                if (FDict[range.Name] != FList[index] && FDict.ContainsKey(range.Name)) FlxMessages.ThrowException(FlxErr.ErrRangeNameAlreadyExists, range.Name);
            }
         
            return range.Clone();
        }

        private bool ValidateName(string name)
        {
            if (!char.IsLetter(name[0])) return false;

            for (int i = 0; i < name.Length; i++)
            {
                char c = name[i];
                if (!Char.IsLetterOrDigit(c) && c != '_' && c != '.') return false;
            }
            return true;
        }

        /// <summary>
        /// Number of protected ranges in the list.
        /// </summary>
        public int Count
        {
            get
            {
                return FList.Count;
            }
        }

        /// <summary>
        /// The range at position "index" (0 based). This method returns a copy of the range.
        /// </summary>
        /// <param name="index"></param>
        /// <returns></returns>
        public TProtectedRange this[int index]
        {
            get
            {
                return FList[index].Clone();
            }
            set
            {
                TProtectedRange range2 = Validate(value, index);
                FDict.Remove(FList[index].Name);
                FList[index] = range2;
                FDict.Add(range2.Name, range2);

            }
        }

        internal TProtectedRangeList Clone()
        {
            TProtectedRangeList Result = new TProtectedRangeList();
            foreach (TProtectedRange rng in FList)
            {
                Result.Add(rng);
            }

            return Result;
        }
    }
}

