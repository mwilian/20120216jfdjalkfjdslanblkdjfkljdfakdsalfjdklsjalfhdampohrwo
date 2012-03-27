using System; 
using System.ComponentModel;
using System.Reflection;
using System.IO; 
using System.Globalization;
using System.Data;
using System.Data.Common;
using System.Text;

using FlexCel.Core;
using System.Collections.Generic;

#if (FRAMEWORK30 && !COMPACTFRAMEWORK)
using System.Linq;
#endif
#if (MONOTOUCH)
using System.Drawing;
using Color = MonoTouch.UIKit.UIColor;
using Image = MonoTouch.UIKit.UIImage;
#else
    #if (WPF)
    using System.Windows.Media;
    #else
    using System.Drawing;
    using Colors = System.Drawing.Color;
    #endif
#endif

namespace FlexCel.Report
{
    /// <summary>
    /// Component for creating reports on Excel based on a template. It will read an xls file, replace tags with data read from
    /// a database or memory, and save a new file with the data.
    /// </summary>
    public class FlexCelReport : Component, IDataTableFinder
    {
        #region Consts
        private readonly int FirstConfigRow = 10;
        private readonly int ConfigColTableName = 1;
        private readonly int ConfigColSourceName = 2;
        private readonly int ConfigColFilter = 3;
        private readonly int ConfigColSort = 4;

        private readonly int ConfigColFormatName = 8;
        private readonly int ConfigColFormatDef = 9;
        private readonly int ConfigColExpName = 13;
        private readonly int ConfigColExpDef = 14;

        private readonly int MaxNestedIncludes = 16;

        #endregion

        #region Private variables
        private volatile bool FCanceled;
        private volatile FlexCelReportProgress FProgress;

        private bool FDeleteEmptyRanges;
        private bool FResetCellSelections;
        private TErrorActions FErrorActions;
        private bool FAllowOverwritingFiles;
        private bool FSemiAbsoluteReferences;
        private TRecalcMode FRecalcMode;
        private bool FRecalcForced;
        private bool FHtmlMode;
        private bool FEnterFormulas;
        private bool FTryToConvertStrings;
        private string FSqlParameterReplace;
        private TSqlParametersType FSqlParametersType;

        private ExcelFile Workbook;

        internal TFormatList FormatList;
        internal TExpressionList ExpressionList;
        internal TExpressionList StaticExpressionList;
        internal TValueList ValueList;
        internal TUserFunctionList UserFunctionList;
        internal TRelationshipList ExtraRelations;
        internal TRelationshipList StaticRelations;

        internal bool FErrorsInResultFile;
        internal bool FDebugExpressions;
        internal bool IntErrorsInResultFile;
        internal bool IntDebugExpressions;

        private TDataSourceInfoList DataSourceList;
        private List<Object> ObjectsToDispose;

        /// <summary>
        /// Holds the data for the config sheet.
        /// </summary>
        private TDataSourceInfoList ConfigDataSourceList;

        private TDataAdapterList AdapterList;
        private TSqlParameterList SqlParameterList;

        /// <summary>
        /// The level of nesting of this report. So we can detect recursive includes.
        /// </summary>
        private int FNestedIncludeLevel;
        #endregion

        #region Internal
        internal int NestedIncludeLevel { get { return FNestedIncludeLevel; } }
        #endregion

        #region Constructors
        /// <summary>
        /// Creates a new FlexCelReport component.
        /// </summary>
        public FlexCelReport()
            : this(false, false)
        {
        }

        private FlexCelReport(bool IsInclude, bool dummy)
        {
            FCanceled = false;
            FProgress = new FlexCelReportProgress();
            DataSourceList = new TDataSourceInfoList(!IsInclude, this);
            ConfigDataSourceList = null;
            ObjectsToDispose = new List<object>();
            AdapterList = new TDataAdapterList();
            SqlParameterList = new TSqlParameterList();
            if (!IsInclude)
            {
                ValueList = new TValueList();
                StaticExpressionList = new TExpressionList();
                UserFunctionList = new TUserFunctionList();
                ExtraRelations = new TRelationshipList();
                StaticRelations = new TRelationshipList();
                FDeleteEmptyRanges = true;
                FSqlParametersType = TSqlParametersType.Automatic;
            }

            FRecalcMode = TRecalcMode.Smart;
            FRecalcForced = true;
            FHtmlMode = false;
        }

        /// <summary>
        /// Creates a new FlexCelReport component and sets the desired Overwrite mode for files. <seealso cref="AllowOverwritingFiles"/>
        /// </summary>
        /// <param name="aAllowOverwritingFiles">If false, FlexCelReport will never overwrite an existing file, and you have to delete it before creating it again.</param>
        public FlexCelReport(bool aAllowOverwritingFiles)
            : this()
        {
            AllowOverwritingFiles = aAllowOverwritingFiles;
        }

        /// <summary>
        /// Creates a new FlexCelReport component to be used on #include tags.
        /// </summary>
        internal FlexCelReport(int aNestedIncludeLevel, string tagText,
            TDataSourceInfoList dsInfoList,
            FlexCelReport parentReport)
            : this(true, true)
        {
            ValueList = parentReport.ValueList;
            UserFunctionList = parentReport.UserFunctionList;
            FNestedIncludeLevel = aNestedIncludeLevel;
            FDeleteEmptyRanges = parentReport.DeleteEmptyRanges;
            ExtraRelations = parentReport.ExtraRelations;
            StaticRelations = parentReport.StaticRelations;

            FSqlParameterReplace = parentReport.SqlParameterReplace;
            FSqlParametersType = parentReport.SqlParametersType;

            AllowOverwritingFiles = false;
            SemiAbsoluteReferences = parentReport.SemiAbsoluteReferences;
            RecalcMode = parentReport.RecalcMode;
            RecalcForced = parentReport.RecalcForced;
            HtmlMode = parentReport.HtmlMode;
            EnterFormulas = parentReport.EnterFormulas;
            TryToConvertStrings = parentReport.TryToConvertStrings;

            if (NestedIncludeLevel > MaxNestedIncludes) //Avoid infinite recursion
                FlxMessages.ThrowException(FlxErr.ErrTooManyNestedIncludes, tagText);
            foreach (TDataSourceInfo di in dsInfoList.Values)
                DataSourceList.Add(di.Name, di);

            foreach (string daName in parentReport.AdapterList)
                AdapterList.Add(daName, parentReport.AdapterList[daName]);

            foreach (string paName in parentReport.SqlParameterList)
                SqlParameterList.Add(paName, parentReport.SqlParameterList[paName]);

            ExpressionList = new TExpressionList();
            foreach (string aKey in parentReport.ExpressionList.Keys)
                ExpressionList.Add(aKey, parentReport.ExpressionList[aKey]);

            ErrorActions = parentReport.ErrorActions;
            DeleteEmptyRanges = parentReport.DeleteEmptyRanges;
            GetImageData = parentReport.GetImageData;
            GetInclude = parentReport.GetInclude;
        }

        #endregion

        #region Public Methods

        /// <summary>
        /// Executes the report, reading from a file and writing to a file.
        /// </summary>
        /// <param name="templateFileName">File with the template to use.</param>
        /// <param name="outFileName">File to create. Note the it must not exist.</param>
        public void Run(string templateFileName, string outFileName)
        {
            ExcelFile aWorkbook = new XlsAdapter.XlsFile(AllowOverwritingFiles);
            aWorkbook.ErrorActions = ConvertErrorActions(ErrorActions);
            aWorkbook.SemiAbsoluteReferences = SemiAbsoluteReferences;
            aWorkbook.RecalcMode = RecalcMode;
            aWorkbook.RecalcForced = RecalcForced;
            OnBeforeReadTemplate(new GenerateEventArgs(aWorkbook));
            aWorkbook.Open(templateFileName);
            Run(aWorkbook);
            if (Canceled) return;
            aWorkbook.Save(outFileName);
        }

        /// <summary>
        /// Executes the report, reading from a stream and writing to a stream.
        /// </summary>
        /// <param name="templateStream">Stream with the template.</param>
        /// <param name="outStream">Stream where the result will be written.</param>
        public void Run(Stream templateStream, Stream outStream)
        {
            FCanceled = false;
            ExcelFile aWorkbook = new XlsAdapter.XlsFile();
            aWorkbook.ErrorActions = ConvertErrorActions(ErrorActions);
            aWorkbook.SemiAbsoluteReferences = SemiAbsoluteReferences;
            aWorkbook.RecalcMode = RecalcMode;
            aWorkbook.RecalcForced = RecalcForced;
            OnBeforeReadTemplate(new GenerateEventArgs(aWorkbook));
            aWorkbook.Open(templateStream);
            Run(aWorkbook);
            if (Canceled) return;
            aWorkbook.Save(outStream);
        }

        /// <summary>
        /// Executes the report, reading from an ExcelFile and writing the results to it again.
        /// </summary>
        /// <remarks>
        /// Note that <see cref="RecalcMode"/>, <see cref="SemiAbsoluteReferences"/> <see cref="RecalcForced"/> and <see cref="AllowOverwritingFiles"/> values used will be the ones on 
        /// aWorkbook, not the ones on this report.
        /// </remarks>
        /// <param name="aWorkbook">ExcelFile that contains the template file, and that will contain the generated file once this method runs.</param>
        public void Run(ExcelFile aWorkbook)
        {
            ExtraRelations.Clear();
            IntDebugExpressions = DebugExpressions;
            IntErrorsInResultFile = ErrorsInResultFile;

            Workbook = aWorkbook;
            try
            {
                FProgress.Clear();
                FProgress.SetPhase(FlexCelReportProgressPhase.ReadTemplate);
                try
                {
                    FormatList = new TFormatList();
                    try
                    {
                        if (ExpressionList == null) ExpressionList = new TExpressionList();  //If this is a nested subreport, ExpressionList has the values of the parent.
                        try
                        {
                            ConfigDataSourceList = new TDataSourceInfoList(true, this);
                            try
                            {
                                ObjectsToDispose = new List<object>();
                                try
                                {
                                    try
                                    {
                                        OnBeforeGenerateWorkbook(new GenerateEventArgs(Workbook));

                                        TBandSheetList DataSheetList = new TBandSheetList();
                                        try
                                        {
                                            TBoolArray UsePreviousTemplate = new TBoolArray();
                                            int ConfigSheet = -1;
                                            InsertSheets(DataSheetList, UsePreviousTemplate, ref ConfigSheet);

                                            TBand MainBand = null;
                                            try
                                            {

                                                List<TKeepTogether> KeepRowsTogether = new List<TKeepTogether>();
                                                List<TKeepTogether> KeepColsTogether = new List<TKeepTogether>();

                                                for (int i = 1; i <= Workbook.SheetCount; i++)
                                                {
                                                    Workbook.ActiveSheet = i;
                                                    if (Workbook.SheetName.Length > 1 && Workbook.SheetName.StartsWith(ReportTag.StrExcludeSheet))
                                                    {
                                                        Workbook.SheetName = Workbook.SheetName.Substring(ReportTag.StrExcludeSheet.Length);
                                                        continue;
                                                    }

                                                    if (!UsePreviousTemplate[i])
                                                    {
                                                        if (MainBand != null) MainBand.Dispose(); MainBand = null;
                                                        KeepRowsTogether = new List<TKeepTogether>();
                                                        KeepColsTogether = new List<TKeepTogether>();
                                                    }

                                                    if (i != ConfigSheet) ProcessSheet(i, DataSheetList, ref MainBand, KeepRowsTogether, KeepColsTogether);
                                                    if (FResetCellSelections) Workbook.SelectCell(1, 1, true);
                                                    if (Canceled) return;
                                                }   //For each page

                                                if (Canceled) return;

                                                //Delete or hide (if we can't delete it) the config sheet.
                                                RemoveSheet(ConfigSheet);

                                                GotoFirstVisibleSheet();
                                                OnAfterGenerateWorkbook(new GenerateEventArgs(Workbook));
                                            }
                                            finally
                                            {
                                                if (MainBand != null) MainBand.Dispose();
                                            }
                                        }
                                        finally
                                        {
                                            DataSheetList.Dispose();
                                        }
                                    }
                                    finally
                                    {
                                        DataSourceList.DeleteTempTables();
                                    }
                                }
                                finally
                                {
                                    foreach (object o in ObjectsToDispose)
                                    {
                                        IDisposable disp = o as IDisposable; //Remember CF might not have those objects disposable.
                                        if (disp != null) disp.Dispose();
                                    }
                                    ObjectsToDispose.Clear();
                                }
                            }
                            finally
                            {
                                ConfigDataSourceList.Dispose();
                                ConfigDataSourceList = null;
                            }
                        }
                        finally
                        {
                            ExpressionList = null;
                        }
                    }
                    finally
                    {
                        FormatList = null;
                    }
                }
                finally
                {
                    FProgress.SetPhase(FlexCelReportProgressPhase.Done);
                } //finally
            }
            finally
            {
                Workbook = null;
            }

            aWorkbook.Recalc(false);
        }

        /// <summary>
        /// Cancels a running report. This method is equivalent to setting <see cref="Canceled"/> = true.
        /// </summary>
        public void Cancel()
        {
            Canceled = true;
        }


        #region AddTable overloads
        /// <inheritdoc cref="AddTable(string, DataTable, TDisposeMode)" />
        public void AddTable(string tableName, DataTable table)
        {
            DataSourceList.Add(tableName, table, false);
        }

        /// <summary>
        /// Use this method to tell FlexCel which DataTables or DataViews are available for the report. <b>Note: </b> If you don't know the tables before 
        /// running the report (and you are not using User Tables or Direct SQL) you can use the <see cref="LoadTable"/> event to load them in demand instead of using AddTable.
        /// This way you will only load the tables you need.
        /// </summary>
        /// <param name="tableName">Name that the table will have on the report.</param>
        /// <param name="table">Table that will be available to the report.</param>
        /// <param name="disposeMode">When disposeMode is TDisposeMode.DisposeTable, FlexCel will take care of
        /// disposing this table after running the report. Use it when adding tables created on the fly, so you do not have to dispose them yourself.
        /// </param>
        /// <example>
        /// To allow flexCelReport1 to use Customers and Orders tables, you can use the code:
        /// <code>
        ///    public Form1()
        ///    {
        ///        InitializeComponent();
        ///        InitializeReports();
        ///    }
        ///    
        ///    private InitializeReports()
        ///    {
        ///        flexCelReport1.AddTable("Customers", DataSet1.Customers); //Add datatable Customers to the report.
        ///        flexCelReport1.AddTable("Orders", OrdersDataView); //Add dataview OrdersDataView to the report.
        ///        flexCelReport1.AddTable(MyDataSet); //Add all the tables on MyDataSet to the report.
        ///        flexCelReport1.AddTable(Form1, true);    //Add all tables and dataviews on Form1, and on all the components inside Form1.
        ///    }
        /// </code>
        /// </example>
        public void AddTable(string tableName, DataTable table, TDisposeMode disposeMode)
        {
            DataSourceList.Add(tableName, table, disposeMode == TDisposeMode.DisposeAfterRun);
        }

        /// <inheritdoc cref="AddTable(string, DataTable, TDisposeMode)" />
        public void AddTable(string tableName, DataView table)
        {
            DataSourceList.Add(tableName, table, false);
        }

        /// <inheritdoc cref="AddTable(string, DataTable, TDisposeMode)" />
        public void AddTable(string tableName, DataView table, TDisposeMode disposeMode)
        {
            DataSourceList.Add(tableName, table, disposeMode == TDisposeMode.DisposeAfterRun);
        }

        /// <summary>
        /// Use this method to load all tables on a dataset at once. This is equivalent to calling <see cref="AddTable(System.String,System.Data.DataTable)"/>
        /// for each of the tables on the dataset.
        /// </summary>
        /// <param name="tables">Dataset containing the tables to add.</param>
        /// <example>See <see cref="FlexCelReport.AddTable(System.String, System.Data.DataTable)"/> for an example.</example>
        public void AddTable(DataSet tables)
        {
            AddTable(tables, TDisposeMode.DoNotDispose);
        }

        /// <summary>
        /// Use this method to load all tables on a dataset at once. When disposeMode is DoNotDispose, this is equivalent to calling <see cref="AddTable(System.String, System.Data.DataTable)"/>
        /// for each of the tables on the dataset. If disposeMode is DisposeTable, this will make the same as calling AddTable() with disposeTable equal to false for
        /// each table on the dataset. and when fisnished all the dataset will be disposed.
        /// </summary>
        /// <param name="tables">Dataset containing the tables to add.</param>
        ///<param name="disposeMode">When disposeMode is TDisposeMode.DisposeTable, FlexCel will take care of
        ///disposing this dataset after running the report. Use it when adding datasets created on the fly, so you do not have to dispose them yourself.</param>
        /// <example>See <see cref="FlexCelReport.AddTable(System.String, System.Data.DataTable)"/> for an example.</example>
        public void AddTable(DataSet tables, TDisposeMode disposeMode)
        {
            if (disposeMode == TDisposeMode.DisposeAfterRun) ObjectsToDispose.Add(tables);
            foreach (DataTable dt in tables.Tables)
            {
                AddTable(dt.TableName, dt, TDisposeMode.DoNotDispose);  //the whole dataset will be disposed on ObjectsToDispose.Dispose().
            }
        }

        /// <summary>
        /// Use this method to add any custom object as a datasource for FlexCel. Make sure to read "Virtual Datasets" on UsingFlexCelReports.pdf.
        /// Use <see cref="FlexCelReport.AddTable(System.String, System.Data.DataTable)"/> to add datasets.
        /// </summary>
        /// <param name="tableName">Name that the table will have on the report.</param>
        /// <param name="table">Table that will be available in the report.</param>
        public void AddTable(string tableName, VirtualDataTable table)
        {
            DataSourceList.Add(tableName, table, false);
        }

        /// <summary>
        /// Use this method to add any custom object as a datasource for FlexCel. Make sure to read "Virtual Datasets" on UsingFlexCelReports.pdf.
        /// Use <see cref="FlexCelReport.AddTable(System.String, System.Data.DataTable)"/> to add datasets.
        /// </summary>
        /// <param name="tableName">Name that the table will have on the report.</param>
        ///<param name="disposeMode">When disposeMode is TDisposeMode.DisposeTable, FlexCel will take care of
        ///disposing this table after running the report. Use it when adding tables created on the fly, so you do not have to dispose them yourself.</param>
        /// <param name="table">Table that will be available in the report.</param>
        public void AddTable(string tableName, VirtualDataTable table, TDisposeMode disposeMode)
        {
            DataSourceList.Add(tableName, table, disposeMode == TDisposeMode.DisposeAfterRun);
        }

        #if (FRAMEWORK30 && !COMPACTFRAMEWORK)
        /// <summary>
        /// Use this method to add any IEnumerable as a datasource for FlexCel. You can use this method
        /// to add LINQ tables.
        /// Use <see cref="FlexCelReport.AddTable(System.String, System.Data.DataTable)"/> to add datasets.
        /// </summary>
        /// <param name="tableName">Name that the table will have on the report.</param>
        /// <param name="table">Table that will be available in the report.</param>
        public void AddTable<T>(string tableName, IEnumerable<T> table)
        {
            if (table == null) FlxMessages.ThrowException(FlxErr.ErrDataSetNull);
            DataSourceList.Add(tableName, new TEFDataTable<T>(tableName, null, table.AsQueryable()), false);
        }

        /// <summary>
        /// Use this method to add any IEnumerable as a datasource for FlexCel. You can use this method
        /// to add LINQ tables.
        /// Use <see cref="FlexCelReport.AddTable(System.String, System.Data.DataTable)"/> to add datasets.
        /// </summary>
        /// <param name="tableName">Name that the table will have on the report.</param>
        ///<param name="disposeMode">When disposeMode is TDisposeMode.DisposeTable, FlexCel will take care of
        ///disposing this table after running the report. Use it when adding tables created on the fly, so you do not have to dispose them yourself.</param>
        /// <param name="table">Table that will be available in the report.</param>
        public void AddTable<T>(string tableName, IEnumerable<T> table, TDisposeMode disposeMode)
        {
            if (table == null) FlxMessages.ThrowException(FlxErr.ErrDataSetNull);
            DataSourceList.Add(tableName, new TEFDataTable<T>(tableName, null, table.AsQueryable()), disposeMode == TDisposeMode.DisposeAfterRun);
        }


        #endif

        #endregion

        #region Other table methods
        /// <summary>
        /// Clear the collection of tables available to the report. Use <see cref="AddTable(DataSet)"/> to add new tables to it.
        /// </summary>
        public void ClearTables()
        {
            DataSourceList.Clear();
        }

        /// <summary>
        /// Returns the VirtualDataTable with the specified name that was added to the report.
        /// </summary>
        /// <param name="tableName">Name of the table to retrieve.</param>
        /// <returns></returns>
        public VirtualDataTable GetTable(string tableName)
        {
            return DataSourceList[tableName].Table;
        }
        #endregion

        #region Add Connections
        /// <summary>
        /// Adds an adapter to use from the template on the DIRECT SQL commands.
        /// <b>For security reasons, make sure this adapter ONLY GRANTS READONLY ACCESS TO THE DATA</b>
        /// </summary>
        /// <param name="connectionName">This is the name you will use on the template to refer to this adapter.</param>
        /// <param name="adapter">The adapter that will be used.</param>
        /// <param name="locale">The locale for the created tables.</param>
        public void AddConnection(string connectionName, IDbDataAdapter adapter, CultureInfo locale)
        {
            AddConnection(connectionName, adapter, locale, false);
        }

        /// <summary>
        /// Adds an adapter to use from the template on the DIRECT SQL commands.
        /// <b>For security reasons, make sure this adapter ONLY GRANTS READONLY ACCESS TO THE DATA</b>
        /// </summary>
        /// <param name="connectionName">This is the name you will use on the template to refer to this adapter.</param>
        /// <param name="adapter">The adapter that will be used.</param>
        /// <param name="locale">The locale for the created tables.</param>
        /// <param name="caseSensitive">When true strings will be case sensitive, and "a" will be different from "A"</param>
        public void AddConnection(string connectionName, IDbDataAdapter adapter, CultureInfo locale, bool caseSensitive)
        {
            if (adapter == null)
                FlxMessages.ThrowException(FlxErr.ErrAdapterNull);

            AdapterList.Add(connectionName, new TAdapterData(adapter, locale, caseSensitive));
        }


        /// <summary>
        /// Clear the collection of connections available to the report. Use <see cref="AddConnection(string, IDbDataAdapter, CultureInfo, bool)"/> to add new connections to it.
        /// </summary>
        public void ClearConnections()
        {
            AdapterList.Clear();
        }
        #endregion

        #region Add SqlParams
        /// <summary>
        /// Adds an SQL parameter to use from the template on the DIRECT SQL commands.
        /// Note that the parameter must have a name even if you are using
        /// positional parameters ("?") because on the template you should always write 
        /// named parameters.
        /// </summary>
        /// <param name="parameterName">The name of the parameter without special symbols, as it is written
        /// on the DIRECT SQL string. This might be different from the real parameter name.
        /// For example, on SQL Server, parameterName might be "MyParameter" while
        /// parameter.ParameterName would be "@MyParameter"</param>
        /// <param name="parameter">The parameter to add.</param>
        public void AddSqlParameter(string parameterName, IDbDataParameter parameter)
        {
            if (parameter == null)
                FlxMessages.ThrowException(FlxErr.ErrInvalidSqlParams);

            SqlParameterList.Add(parameterName, parameter);
        }

        /// <summary>
        /// Clear the collection of SQL parameters available to the report. Use <see cref="AddSqlParameter"/> to add new parameters to it.
        /// </summary>
        public void ClearParameters()
        {
            SqlParameterList.Clear();
        }
        #endregion

        /// <summary>
        /// Sets a variable for the report. 
        /// </summary>
        /// <param name="name">Name of the variable to set. Case insensitive.</param>
        /// <param name="value">Value of the variable.</param>
        /// <example>
        /// You can define a variable "CurrentDate" on the following way:
        /// <code>
        ///   flexCelReport1.SetValue("CurrentDate", DateTime.Now);
        /// </code>
        /// Then, if you write &lt;#CurrentDate&gt; on a cell, the date will be shown.
        /// Note that the name is case insensitive, so both "CURRENTDATE" and "currentdate" refer to the same variable.
        /// </example>
        public void SetValue(string name, object value)
        {
            ValueList[name] = value;
        }


        /// <summary>
        /// Destroys all variables on the report. To add new variables, use <see cref="SetValue"/>
        /// </summary>
        public void ClearValues()
        {
            ValueList.Clear();
        }

        /// <summary>
        /// Sets an user-defined expression to be used in the report. Different from <see cref="SetValue"/> this method will evaluate the
        /// &lt;#tags&gt; in "value". This allows you to provide formula functionality to end users, and to reuse the same report for 
        /// different formulas.
        /// </summary>
        /// <param name="name">Name of the expression to set. Case insensitive.</param>
        /// <param name="value">Value of the expression.</param>
        /// <example>
        /// You could ask the user for an expression in an edit box, and then before running the report do:
        /// <code>
        ///   flexCelReport1.SetExpression("MyExpression", EditBox.Text);
        /// </code>
        /// Then, if you write &lt;#MyExpresion&gt; on a cell, the expression will be evaluated and entered into the cell.
        /// If the user enters "&lt;#evaluate(&lt;#Order.Amount&gt; * &lt;#Order.Vat&gt;)&gt;" in the edit box, this formula will
        /// be evaluated into the cell. The user can write any &lt;#tag&gt; that could be used in an expression defined directly in the template.
        /// </example>
        public void SetExpression(string name, object value)
        {
            StaticExpressionList.Add(name, new TExpression(null, value));
        }

        /// <summary>
        /// Destroys all user-defined expressions on the report. To add new expressions, use <see cref="SetExpression"/>
        /// </summary>
        public void ClearExpressions()
        {
            StaticExpressionList.Clear();
        }


        /// <summary>
        /// Adds a new user defined function to be used with the report. 
        /// For information on how to create the user function, see <see cref="TFlexCelUserFunction"/> 
        /// </summary>
        /// <param name="name">Name that the function will have on the report. Case insensitive.</param>
        /// <param name="functionImplementation">An implementation of the user function.</param>
        /// <example>
        /// You can define a function "MyFunc" on the following way:
        /// <code>
        ///   flexCelReport1.SetValue("MyFunc", MyFuncImpl);
        /// </code>
        /// Then, if you write &lt;#MyFunc(param1, param2... paramn)&gt; on a cell, the function will be called and the result shown.
        /// Note that the name is case insensitive, so both "MYFUNC" and "myfunc" refer to the same function.
        /// </example>
        public void SetUserFunction(string name, TFlexCelUserFunction functionImplementation)
        {
            UserFunctionList[name] = functionImplementation;
        }

        /// <summary>
        /// Destroys all user defined functions on the report. To add new functions, use <see cref="SetUserFunction"/>
        /// </summary>
        public void ClearUserFunctions()
        {
            UserFunctionList.Clear();
        }

        #region Relationships

        /// <inheritdoc cref="AddRelationship(string, string, string[], string[])" />
        public void AddRelationship(string masterTable, string detailTable, string masterKeyFields, string detailKeyFields)
        {
            AddRelationship(masterTable, detailTable, new string[] { masterKeyFields }, new string[] { detailKeyFields });
        }


        /// <summary>
        /// Adds a relationship between two tables. Note that if your tables are datasets, or come from Entity Framework,
        /// or your master detail tables are nested, you probably don't need to call this method. FlexCel will detect the relationships automatically
        /// and use them. This method is only for when you have custom data with no DataRelationships (as in a dataset) and where detail
        /// tables are not nested to the master (as in Entity Framework).
        /// </summary>
        /// <remarks>You might use this method before adding the tables with AddTable. Tables must exist when running the report, not when adding the relationship.</remarks>
        /// <param name="masterTable">Master table for the relationship.</param>
        /// <param name="detailTable">Detail table for the relationship.</param>
        /// <param name="masterKeyFields">Key fields in the master table that relate with the key fields in the detail table.
        /// Must not be null, and normally will be an array with a single field. It must be the same length as detailKeyFields/></param>
        /// <param name="detailKeyFields">Key fields in the detail table that relate with the key fields in the master table.
        /// Must not be null, and normally will be an array with a single field.</param> 
        public void AddRelationship(string masterTable, string detailTable, string[] masterKeyFields, string[] detailKeyFields)
        {
            if (masterTable == null) FlxMessages.ThrowException(FlxErr.ErrNullParameter, "masterTable");
            if (detailTable == null) FlxMessages.ThrowException(FlxErr.ErrNullParameter, "detailTable");
            if (masterKeyFields == null || masterKeyFields.Length == 0)
            {
                FlxMessages.ThrowException(FlxErr.ErrNullOrEmptyParameter, "masterKeyFields");
            }
            
            if (detailKeyFields == null || detailKeyFields.Length == 0)
            {
                FlxMessages.ThrowException(FlxErr.ErrNullOrEmptyParameter, "detailKeyFields");
            }

            if (detailKeyFields.Length != masterKeyFields.Length)
            {
                FlxMessages.ThrowException(FlxErr.ErrInvalidManualRelationshipFieldCount);
            }

            StaticRelations.Add(new TRelationship(masterTable, detailTable, masterKeyFields, detailKeyFields));
        }

        /// <summary>
        /// Clears all relationships added by <see cref="AddRelationship(string, string, string, string)"/>.
        /// </summary>
        public void ClearRelationships()
        {
            StaticRelations.Clear();
        }
        #endregion

        #endregion

        #region Properties

        /// <summary>
        /// If true the report has been canceled with <see cref="Cancel"/> method.
        /// You can't set this variable to false, and setting it true is the same as calling <see cref="Cancel"/>.
        /// </summary>
        [Browsable(false),
        DefaultValue(false)]
        public bool Canceled
        {
            get { return FCanceled; }
            set
            {
                if (value == true) FCanceled = true; //Don't allow to uncancel.
            }
        }

        /// <summary>
        /// Progress of the report. This variable must be accessed from other thread.
        /// </summary>
        [Browsable(false)]
        public FlexCelReportProgress Progress
        {
            get { return FProgress; }
        }

        /// <summary>
        /// Determines if FlexCel will delete or just clear ranges with empty datasets (0 records).
        /// </summary>
        [Category("Behavior"),
        Description("Determines if FlexCel will delete or just clear ranges with empty datasets (0 records)."),
        DefaultValue(true)]
        public bool DeleteEmptyRanges
        {
            get { return FDeleteEmptyRanges; }
            set { FDeleteEmptyRanges = value; }
        }

        /// <summary>
        /// Determines if FlexCel will automatically delete existing files or not.
        /// </summary>
        [Category("Behavior"),
        Description("Determines if FlexCel will automatically delete existing files or not."),
        DefaultValue(false)]
        public bool AllowOverwritingFiles { get { return FAllowOverwritingFiles; } set { FAllowOverwritingFiles = value; } }

        /// <summary>
        /// When true, all sheets will selections will be reset to A1. This way, you do not need to care about setting the correct selection when editing the template.
        /// </summary>
        [Category("Behavior"),
        Description("When true, all sheets cell selections will be reset to A1. This way, you do not need to care about setting the correct selection when editing the template."),
        DefaultValue(false)]
        public bool ResetCellSelections
        {
            get { return FResetCellSelections; }
            set { FResetCellSelections = value; }
        }

        /// <summary>
        /// When true, FlexCel will interpret the text as HTML, and honor the tags that it can understand.
        /// Note that when in HtmlMode, many consecutive spaces will be interpreted as one, and carriage returns
        /// will be interpreted as spaces. To enter real carriage returns you need to enter a &lt;br&gt; tag (unless the text is inside &lt;pre&gt; tags). 
        /// Also &amp; symbols need to be escaped. For more info on HTML syntax supported, see <see cref="ExcelFile.SetCellFromHtml(int, int, string, int)"/>
        /// <br/>Note that the &lt;#HTML&gt; tag can overwrite the Html behavior on a cell by cell basis.
        /// </summary>
        [Category("Behavior"),
        Description("When true, FlexCel will interpret the text as HTML, and honor the tags that it can understand."),
        DefaultValue(false)]
        public bool HtmlMode
        {
            get { return FHtmlMode; }
            set { FHtmlMode = value; }
        }

        /// <summary>
        /// When true, FlexCel will try to enter any string starting with "=" as a formula instead of text. 
        /// If this property is true, any string you enter that starts with "=" must be a valid formula, or an error will be raised.
        /// When you know a priori which cells will have formulas, you might want to use the &lt;#formula&gt; tag instead.
        /// </summary>
        [Category("Behavior"),
        Description("When true, FlexCel will try to enter any string starting with \"=\" as a formula instead of text."),
        DefaultValue(false)]
        public bool EnterFormulas
        {
            get { return FEnterFormulas; }
            set { FEnterFormulas = value; }
        }

        /// <summary>
        /// When true, FlexCel will try to convert strings to numbers or dates before entering them into the cells. 
        /// <b>USE THIS PROPERTY WITH CARE!</b>  You shouldn't normally need to use this property, since FlexCel automatically
        /// enters numbers or dates in the DataSets as number or dates in the Excel file. If you need to use this property,
        /// it means that data in your database is stored as strings when they should not be. So the correct fix is to fix the columns
        /// you know should have numbers to have numbers, NOT to use this property. This is just a workaround when you can't do anything else about it.<br/>
        /// Note also that this method is not efficient since it has to "guess" what a string might be if anything, and it might be wrong.
        /// Also, it might have issues with locales: Does the string "1/2/2008" means January 2 or February 1? Depends on the locale.
        /// </summary>
        [Category("Behavior"),
        Description("When true, FlexCel will try to convert strings to numbers or dates before entering them into the cells. READ HELP BEFORE USING!"),
        DefaultValue(false)]
        public bool TryToConvertStrings
        {
            get { return FTryToConvertStrings; }
            set { FTryToConvertStrings = value; }
        }

        /// <summary>
        /// Determines if the report will be recalculated before saving.
        /// See <see cref="FlexCel.Core.ExcelFile.RecalcMode"/> for more info.
        /// </summary>
        [Category("Recalculation"),
        Description("Determines if the report will be recalculated before saving."),
        DefaultValue(TRecalcMode.Smart)]
        public TRecalcMode RecalcMode { get { return FRecalcMode; } set { FRecalcMode = value; } }

        /// <summary>
        /// <b>Before changing this property, look at <see cref="FlexCel.Core.ExcelFile.RecalcForced"/></b> Determines if the formulas will be recalculated when Excel opens them.
        /// </summary>
        [Category("Recalculation"),
        Description("Do not change without reading the documentation"),
        DefaultValue(true)]
        public bool RecalcForced { get { return FRecalcForced; } set { FRecalcForced = value; } }

        /// <summary>
        /// Determines if FlexCel will throw Exceptions or just ignore errors on specific situations. When the errors are ignored, they will
        /// be logged into the <see cref="FlexCel.Core.FlexCelTrace"/> class.
        /// </summary>
        [Category("Behavior"),
        Description("Determines if FlexCel will throw Exceptions or just ignore errors on specific situations."),
        DefaultValue(TErrorActions.None)]
        public TErrorActions ErrorActions { get { return FErrorActions; } set { FErrorActions = value; } }

        /// <summary>
        /// When true and there is an error reading cells in the template or writing the cells in the report, the
        /// error message will be written in the corresponding cell on the generated report. No Exception will be thrown.
        /// <br/>You can use this property to <b>DEBUG</b> reports, as it provides an easy
        /// way to see all errors at once in the place they are produced. But is it recommended that you leave this property <b>FALSE</b> in production, 
        /// or you could create xls files with error messages inside. See also <see cref="DebugExpressions"/>
        /// </summary>
        /// <remarks>You can also set this property in the template, by writing &lt;#ErrorsInResultFile&gt; in the expressions column in the config sheet.</remarks>
        [Category("Behavior"),
        Description("Determines if FlexCel will throw Exceptions when finding errors or write those errors in the cells."),
        DefaultValue(false)]
        public bool ErrorsInResultFile { get { return FErrorsInResultFile; } set { FErrorsInResultFile = value; } }

        /// <summary>
        /// Set this value to true if you want to analize how FlexCel is evaluating the tags in a file. When true, a full stack
        /// trace will be written in the cell instead of the tag values. See the section on Debugging reports in the "Using FlexCel Reports (Designing Templates)" for
        /// information on how to use those stack traces.
        /// </summary>
        /// <remarks>You can also set this property in the template, by writing &lt;#Debug&gt; in the expressions column in the config sheet.
        /// Debug in the template will set both this property and <see cref="ErrorsInResultFile"/> to true.</remarks>
        [Category("Behavior"),
        Description("When true FlexCel will output the full stack of sub expressions used to calculate the tags instead of the calculated values."),
        DefaultValue(false)]
        public bool DebugExpressions { get { return FDebugExpressions; } set { FDebugExpressions = value; } }

        /// <summary>
        /// Format string for replacing the standard parameter names on DIRECT SQL commands. You can leave it empty for ODBC, OLEDB or SQLSERVER databases.
        /// See Also <see cref="SqlParametersType"/>
        /// </summary>
        /// <remarks>
        /// On ADO.NET sql parameters can be defined on multiple ways, depending on
        /// the database you use.
        /// For example, SQL Server will use "@ParamName", while OLEDB will use "?"
        /// as parameter id. 
        /// But, to keep templates database independent, you will always use 
        /// "@ParamName" notation on the Direct SQL commands you write on templates.
        /// When using OLEDB, ODBC or SqlSever, FlexCel knows how to 
        /// replace the correct parameters on the SQL on templates, but for more
        /// generic databases you might need to use this property.
        /// This is a standard .NET format string, where {0} will be replaced by the parameter name. For example, 
        /// on Oracle database you might need to set this value to "\:{0}" (without quotes)
        /// that will translate on ":ParamName"
        /// </remarks>
        /// <example>
        /// On oracle databases, you can set this property to "\:{0}" (without quotes)
        /// It will create parameters of type ":ParamName".
        /// </example>
        [Category("Direct SQL"),
        Description("Format string for replacing the standard parameter names on DIRECT SQL commands. You can leave it empty for ODBC, OLEDB or SQLSERVER databases."),
        DefaultValue(null)]
        public string SqlParameterReplace { get { return FSqlParameterReplace; } set { FSqlParameterReplace = value; } }

        /// <summary>
        /// Type of parameters for the database. Positional parameters are the ones where you write 
        /// "?" on the sql, and positional are when you write a name, like &quot;@employee&quot; or &quot;:orderid&quot;.
        /// See Also <see cref="SqlParameterReplace"/>
        /// </summary>
        [Category("Direct SQL"),
        Description("Type of parameters for the database."),
        DefaultValue(TSqlParametersType.Automatic)]
        public TSqlParametersType SqlParametersType { get { return FSqlParametersType; } set { FSqlParametersType = value; } }


        /// <summary>
        /// When this property is set to true, absolute references to cells inside bands being copied will be treated as relative.
        /// This way, if you have "=$A$1" inside a band and cell A1 is also inside the band, it will change to A2,A3..etc when the band is copied down. 
        /// This can be useful in a master-detail report, where you want the cells in the detail to point to a fixed cell inside every record of the master.
        /// See <see cref="ExcelFile.SemiAbsoluteReferences"/> for more information.
        /// </summary>
        [Category("Behavior"),
        Description("When this property is set to true, absolute references to cells inside bands being copied will be treated as relative."),
        DefaultValue(false)]
        public bool SemiAbsoluteReferences { get { return FSemiAbsoluteReferences; } set { FSemiAbsoluteReferences = value; } }

        #endregion

        #region Events
        /// <summary>
        /// Fires before starting to generate the report and before the template has been loaded.
        /// It allows to set the template password if you are using one.
        /// </summary>
        [Category("Generate"),
        Description("Fires before starting to generate the report and before the template has been loaded.")]
        public event GenerateEventHandler BeforeReadTemplate;

        /// <summary>
        /// Replace this event when creating a custom descendant of FlexCelReport. See also <see cref="BeforeReadTemplate"/>
        /// </summary>
        /// <param name="e"></param>
        protected virtual void OnBeforeReadTemplate(GenerateEventArgs e)
        {
            if (BeforeReadTemplate != null) BeforeReadTemplate(this, e);
        }

        /// <summary>
        /// Fires before starting to generate the report but after the template has been loaded.
        /// It allows to do some in-place modifications to the template before generating the report.
        /// </summary>
        [Category("Generate"),
        Description("Fires before starting to generate the report but after the template has been loaded.")]
        public event GenerateEventHandler BeforeGenerateWorkbook;

        /// <summary>
        /// Replace this event when creating a custom descendant of FlexCelReport. See also <see cref="BeforeGenerateWorkbook"/>
        /// </summary>
        /// <param name="e"></param>
        protected virtual void OnBeforeGenerateWorkbook(GenerateEventArgs e)
        {
            if (BeforeGenerateWorkbook != null) BeforeGenerateWorkbook(this, e);
        }


        /// <summary>
        /// Fires After the report has been fully generated but is not saved.
        /// Allows to do last clean up things before saving the report.
        /// </summary>
        [Category("Generate"),
        Description("Fires After the report has been fully generated but is not saved.")]
        public event GenerateEventHandler AfterGenerateWorkbook;

        /// <summary>
        /// Replace this event when creating a custom descendant of FlexCelReport. See also <see cref="AfterGenerateWorkbook"/>
        /// </summary>
        /// <param name="e"></param>
        protected virtual void OnAfterGenerateWorkbook(GenerateEventArgs e)
        {
            if (AfterGenerateWorkbook != null) AfterGenerateWorkbook(this, e);
        }


        /// <summary>
        /// Fires Before each sheet on the file is generated.
        /// </summary>
        [Category("Generate"),
        Description("Fires Before each sheet on the file is generated.")]
        public event GenerateEventHandler BeforeGenerateSheet;

        /// <summary>
        /// Replace this event when creating a custom descendant of FlexCelReport. See also <see cref="BeforeGenerateSheet"/>
        /// </summary>
        /// <param name="e"></param>
        protected virtual void OnBeforeGenerateSheet(GenerateEventArgs e)
        {
            if (BeforeGenerateSheet != null) BeforeGenerateSheet(this, e);
        }


        /// <summary>
        /// Fires After each sheet on the file is generated.
        /// </summary>
        [Category("Generate"),
        Description("Fires After each sheet on the file is generated.")]
        public event GenerateEventHandler AfterGenerateSheet;

        /// <summary>
        /// Replace this event when creating a custom descendant of FlexCelReport. See also <see cref="AfterGenerateSheet"/>
        /// </summary>
        /// <param name="e"></param>
        protected virtual void OnAfterGenerateSheet(GenerateEventArgs e)
        {
            if (AfterGenerateSheet != null) AfterGenerateSheet(this, e);
        }


        /// <summary>
        /// Fires before an image is saved to the report.
        /// Use it if the image is on a proprietary format on the database, to return a format FlexCel can understand.
        /// </summary>
        [Category("Data"),
        Description("Fires before an image is saved to the report. " +
            "Use it if the image is on a proprietary format on the database, to return a format FlexCel can understand.")]
        public event GetImageDataEventHandler GetImageData;

        /// <summary>
        /// Replace this event when creating a custom descendant of FlexCelReport. See also <see cref="GetImageData"/>
        /// </summary>
        /// <param name="e"></param>
        protected virtual void OnGetImageData(GetImageDataEventArgs e)
        {
            if (GetImageData != null) GetImageData(this, e);
        }

        /// <summary>
        /// Fires before including a file with &lt;#include&gt;.
        /// Use it if you want to provide an alternative path for the file, of if you want to read the include
        /// file from a different place, for example a database or an embedded resource.
        /// </summary>
        /// <remarks>
        /// If the including file is a real file (not an stream) and FileName is relative, it will be relative to the
        /// including file path.
        /// </remarks>
        [Category("Data"),
        Description("Fires before including a file with <#include>;. " +
            "Use it if you want to provide an alternative path for the file, of if you want to read the include" +
            "file from a different place, for example a database or an embedded resource.")]
        public event GetIncludeEventHandler GetInclude;

        /// <summary>
        /// Replace this event when creating a custom descendant of FlexCelReport. See also <see cref="GetInclude"/>
        /// </summary>
        /// <param name="e"></param>
        protected internal virtual void OnGetInclude(GetIncludeEventArgs e)
        {
            if (GetInclude != null) GetInclude(this, e);
        }

        /// <summary>
        /// Fires on each &lt;#USER TABLE&gt; tag in the config sheet, allowing to add your own datasets to the report.
        /// </summary>
        [Category("Data"),
        Description("Fires on each <#USER TABLE> tag in the config sheet, allowing to add your own datasets to the report.")]
        public event UserTableEventHandler UserTable;

        /// <summary>
        /// Replace this event when creating a custom descendant of FlexCelReport. See also <see cref="UserTable"/>
        /// </summary>
        /// <param name="e"></param>
        protected virtual void OnUserTable(UserTableEventArgs e)
        {
            if (UserTable == null) FlxMessages.ThrowException(FlxErr.ErrUserTableEventNotAssigned);
            UserTable(this, e);
        }

        /// <summary>
        /// Fires whenever an undefined table is called, allowing to load your own datasets in demand to the report. For more control, you might use User Tables. Look at the example for more information.
        /// </summary>
        /// <example>
        /// If you are running a report and don't know beforehand wich tables it uses, you can use the following event:
        /// <code>
        ///    FlexCelReport fr = NewFlexCelReport();
        ///    fr.LoadTable += new LoadTableEventHandler(fr_LoadTable);
        ///    fr.Run(...);
        /// </code>
        /// and define the event "fr_LoadTable" as:
        /// <code>
        /// void fr_LoadTable(object sender, LoadTableEventArgs e)
        /// {
        ///    ((FlexCelReport)sender).AddTable(e.TableName, GetTable(e.TableName), TDisposeMode.DisposeAfterRun);
        /// }
        /// </code>
        /// 
        /// Instead of using fr.AddTable for all used tables before running the report. If you need tables on demand, you might also look at User Tables or Direct SQL.
        /// </example>
        [Category("Data"),
        Description("Fires whenever an undefined table is called, allowing to load your own datasets in demand to the report. For more control, you might use User Tables. Look at the example in the Help")]
        public event LoadTableEventHandler LoadTable;

        /// <summary>
        /// Replace this event when creating a custom descendant of FlexCelReport. See also <see cref="LoadTable"/>
        /// </summary>
        /// <param name="e"></param>
        protected virtual void OnLoadTable(LoadTableEventArgs e)
        {
            if (LoadTable == null) return;
            LoadTable(this, e);
        }
        #endregion

        #region Utilities
        private static TExcelFileErrorActions ConvertErrorActions(TErrorActions ErrorActions)
        {
            return (TExcelFileErrorActions)ErrorActions;//in this case the correspondence is 1-1. If more error actions are added in the future, this method will have to be revised.
        }

        private void GotoFirstVisibleSheet()
        {
            Workbook.ActiveSheet = 1;
            //This might select a no visible sheet if all sheets are not visible. 
            //It does not matter, since FlexCel will not allow to save this file.
            while (Workbook.ActiveSheet < Workbook.SheetCount && Workbook.SheetVisible != TXlsSheetVisible.Visible)
            {
                Workbook.ActiveSheet++;
            }
        }

        private TDataSourceInfo FindDataSource(string SourceName, bool InConfig, string AdditionalData)
        {
            bool TableExists;
            TDataSourceInfo Result = FindDataSource(SourceName, InConfig, AdditionalData, out TableExists);
            if (!TableExists)
            {
                if (InConfig)
                    FlxMessages.ThrowException(FlxErr.ErrDataSetNotFoundInConfig, SourceName, AdditionalData);
                else
                    FlxMessages.ThrowException(FlxErr.ErrDataSetNotFound, SourceName, AdditionalData);
            }

            return Result;
        }

        private TDataSourceInfo FindDataSource(string SourceName, bool InConfig, string AdditionalData, out bool TableExists)
        {
            TableExists = true;
            if (SourceName.StartsWith(ReportTag.StrExcludeSheet))
                return null;

            TDataSourceInfo di = TryGetDataTable(SourceName);

            if (di == null)
            {
                TableExists = false;
            }
            return di;
        }

        private void GetAnchors(List<TClientAnchor> Result, TShapeProperties sp)
        {
            if (sp.Anchor != null)
            {
                Result.Add(sp.Anchor);
                return;
            }

            //some shapes do not have anchors. Wee need to look inside them.
            for (int i = 1; i <= sp.ChildrenCount; i++)
            {
                GetAnchors(Result, sp.Children(i));
            }
        }

        private void CalcMaxUsed(ref int MaxRowCount, ref int MaxColCount)
        {
            MaxRowCount = Workbook.RowCount;
            MaxColCount = Workbook.ColCount;

            int aCount = Workbook.ObjectCount;
            for (int i = 1; i <= aCount; i++) //No need to search inside grouped images
            {
                TShapeProperties sp = Workbook.GetObjectProperties(i, false);
                List<TClientAnchor> Anchors = new List<TClientAnchor>();
                GetAnchors(Anchors, sp);
                foreach (TClientAnchor Anchor in Anchors)
                {
                    if (Anchor != null && Anchor.Row2 > MaxRowCount) MaxRowCount = Anchor.Row2;
                    if (Anchor != null && Anchor.Col2 > MaxColCount) MaxColCount = Anchor.Col2;
                }
            }
        }

        private void ProcessSheet(int i, TBandSheetList DataSheetList, ref TBand MainBand, List<TKeepTogether> KeepRows, List<TKeepTogether> KeepCols)
        {
            FProgress.SetSheet(i);
            if (Canceled) return;

            OnBeforeGenerateSheet(new GenerateEventArgs(Workbook));

            if (DataSheetList[i] != null)
                if (i <= 1 || DataSheetList[i - 1] != DataSheetList[i]) DataSheetList[i].MoveFirst(TBandMoveType.Alone);
                else DataSheetList[i].MoveNext(TBandMoveType.Alone);

            ReplaceHeadersAndFooters(DataSheetList[i]);

            if (Workbook.SheetType == TSheetType.Worksheet || Workbook.SheetType == TSheetType.Chart)
            {
                if (MainBand == null || MainBand.Preprocessed)  //When doing multiple sheet reports (MainBand != null), if preprocessing, we need to preprocess each sheet so we can't optimize.
                {
                    if (MainBand != null) MainBand.Dispose();
                    MainBand = NewMainBand(null, DataSheetList[i], "MAIN");
                    ReadTemplate(ref MainBand, false, KeepRows, KeepCols); if (Canceled) return;
                }

                KeepRowsAndCols(KeepRows, KeepCols);

                TSheetState SheetState = new TSheetState();

                ExportData(MainBand, SheetState); if (Canceled) return;

                AfterExportData(SheetState);

                FProgress.SetPhase(FlexCelReportProgressPhase.OrganizeData);
                OnAfterGenerateSheet(new GenerateEventArgs(Workbook));
            }
        }

        private void AfterExportData(TSheetState SheetState)
        {
            switch (SheetState.AutofitInfo.AutofitType)
            {
                case TAutofitType.Sheet: Workbook.AutofitRow(1, Workbook.RowCount, false, SheetState.AutofitInfo.KeepAutofit, SheetState.AutofitInfo.GlobalAdjustment, SheetState.AutofitInfo.GlobalAdjustmentFixed, 0, 0, SheetState.AutofitInfo.MergedMode);
                    break;
                case TAutofitType.OnlyMarked: Workbook.AutofitMarkedRowsAndCols(SheetState.AutofitInfo.KeepAutofit, false, SheetState.AutofitInfo.GlobalAdjustment, SheetState.AutofitInfo.GlobalAdjustmentFixed, 0, 0, 0, 0, SheetState.AutofitInfo.MergedMode);
                    break;
            }
            if (Canceled) return;

            if (SheetState.AutoPageBreaksPercent >= 0)
            {
                Workbook.AutoPageBreaks(SheetState.AutoPageBreaksPercent, SheetState.AutoPageBreaksPageScale);
            }
        }

        private void KeepRowsAndCols(List<TKeepTogether> KeepRows, List<TKeepTogether> KeepCols)
        {
            Workbook.ClearKeepRowsAndColsTogether();
            foreach (TKeepTogether keep in KeepRows)
            {
                Workbook.KeepRowsTogether(keep.R1, keep.R2, keep.Level, false);
            }
            foreach (TKeepTogether keep in KeepCols)
            {
                Workbook.KeepColsTogether(keep.R1, keep.R2, keep.Level, false);
            }
        }

        private TBand NewMainBand(TXlsCellRange XlsRange, TBand aParentBand, string aBandName)
        {
            if (XlsRange == null)
            {
                int MaxRowCount = 0; int MaxColCount = 0;
                CalcMaxUsed(ref MaxRowCount, ref MaxColCount);
                XlsRange = new TXlsCellRange(1, 1, MaxRowCount, MaxColCount);
            }
            return new TBand(null, aParentBand, XlsRange, aBandName, TBandType.Static, false, String.Empty);
        }

        private bool RemoveSheet(int sheet)
        {
            bool SheetDeleted = false;
            if (sheet > 0)
            {
                SheetDeleted = !Workbook.HasMacros();  //Avoid the exception if we know it has macros, for performance reasons.
                Workbook.ActiveSheet = sheet;
                if (SheetDeleted) Workbook.DeleteSheet(1);
                if (!SheetDeleted) //Can happen if you have macros on the sheet.
                {
                    Workbook.ClearSheet();
                    Workbook.SheetVisible = TXlsSheetVisible.VeryHidden;
                }
            }
            return SheetDeleted;

        }

        internal TDataSourceInfoList GetDataTables()
        {
            TDataSourceInfoList Result = new TDataSourceInfoList(false, this);
            if (DataSourceList != null)
                foreach (TDataSourceInfo di in DataSourceList.Values)
                    Result.Add(di.Name, di);
            if (ConfigDataSourceList != null)
                foreach (TDataSourceInfo di in ConfigDataSourceList.Values)
                    Result.Add(di.Name, di);

            return Result;

        }

        TDataSourceInfo IDataTableFinder.TryGetDataTable(string DataTableName)
        {
            return this.TryGetDataTable(DataTableName);
        }

        internal TDataSourceInfo TryGetDataTable(string DataTableName)
        {
            if (DataSourceList != null)
            {
                TDataSourceInfo di = DataSourceList[DataTableName];
                if (di != null) return di;
            }
            if (ConfigDataSourceList != null)
            {
                TDataSourceInfo di = ConfigDataSourceList[DataTableName];
                if (di != null) return di;
            }
            
            OnLoadTable(new LoadTableEventArgs(DataTableName));
            return DataSourceList[DataTableName];
        }

        internal TDataSourceInfo GetDataTable(string DataTableName, string ErrorCause)
        {
            TDataSourceInfo Result = TryGetDataTable(DataTableName);
            if (Result == null) FlxMessages.ThrowException(FlxErr.ErrDataSetNotFound, DataTableName, ErrorCause);
            return Result;
        }

        #endregion

        #region Include runs
        /// <summary>
        /// Used on included reports. For performance, the report will be parsed only once.
        /// </summary>
        internal void PreLoad(ExcelFile aWorkbook, ref TBand startingBand, int sheetToLoad, ref byte[] ReportData, out List<TKeepTogether> KeepRows, out List<TKeepTogether> KeepCols)
        {
            Workbook = aWorkbook;
            //try
            {
                ConfigDataSourceList = new TDataSourceInfoList(true, this);
                FormatList = new TFormatList();
                //ExpressionList=new TExpressionList; ExpressionList comes from the parent.

                FindAndLoadConfigSheet();

                Workbook.ActiveSheet = sheetToLoad;
                KeepRows = new List<TKeepTogether>();
                KeepCols = new List<TKeepTogether>();
                ReadTemplate(ref startingBand, true, KeepRows, KeepCols);

                //KeepRowsAndCols(KeepRows, KeepCols);
                //KeepRowsAndCols has no effect here, since those values are not saved with the template.

                if (startingBand.Preprocessed) //Re save the preprocessed report. It will be used more than once when including master details.
                {
                    using (MemoryStream Ms = new MemoryStream())
                    {
                        Workbook.Save(Ms);
                        ReportData = Ms.ToArray();
                    }
                }
            }
            //We do not really need to clean up here, as the whole report will be freed.
            // And if there is an exception it can slow down things.
            /* catch (Exception)
             {
                 ConfigDataSourceList=null;
                 FormatList=null;
                 ExpressionList=null;
                 throw;
             }*/
        }

        internal void RunPreloaded(ExcelFile aWorkbook, TBand mainBand, List<TKeepTogether> KeepRows, List<TKeepTogether> KeepCols)
        {
            Workbook = aWorkbook;
            KeepRowsAndCols(KeepRows, KeepCols); //must be called each time, since the workbook has a newly loaded file, and loaded files don't have that info.
            TSheetState SheetState = new TSheetState();
            ExportData(mainBand, SheetState); if (Canceled) return; //TSheetState will not pass from include to master or vice versa.
            AfterExportData(SheetState);
        }

        #endregion

        #region Parser
        /// <summary>
        /// This is the method that does the parsing. Could be made virtual and override it on a descendant class to support 
        /// self defined Tags.
        /// </summary>
        /// <param name="s">String to parse.</param>
        /// <param name="XF">Original XF of the cell. The value returned might change it, if for example there is a #FormatCell tag.</param>
        /// <param name="CurrentBand">The band we are currently in.</param>
        /// <returns>A parsed class, with values replaced by tags found on s.</returns>
        internal TOneCellValue ResolveString(object s, int XF, TBand CurrentBand)
        {
            return ResolveString(s, XF, CurrentBand, false);
        }

        internal TOneCellValue ResolveString(object s, int XF, TBand CurrentBand, bool CanAddDataSets)
        {
            return TCellParser.GetCellValue(s, Workbook, new TStackData(new TUsedRefs(), null), XF, CurrentBand, this, CanAddDataSets);
        }

        private TBand FindSheetDataSet(TOneCellValue sourceValue)
        {
            if (sourceValue == null) return null;
            for (int i = 0; i < sourceValue.Count; i++)
            {
                if (sourceValue[i].ValueType == TValueType.DataSet)
                {
                    return sourceValue[i].DataBand();
                }
                TBand Result = FindSheetDataSet(sourceValue[i].Resolve(0, 0, 0, 0, null, 0));
                if (Result != null) return Result;
            }
            return null;
        }

        #endregion

        #region Report Sheets

        private static void ValidateSQL(string sql)
        {
            sql = sql.Trim();
            if (sql.IndexOf(";") >= 0)
                FlxMessages.ThrowException(FlxErr.ErrInvalidSqlString, sql);
            if (sql.IndexOf("--") >= 0)
                FlxMessages.ThrowException(FlxErr.ErrInvalidSqlString, sql);

            string StrSelect = "SELECT ";
            if (sql.Length < StrSelect.Length || String.Compare(sql, 0, StrSelect, 0, StrSelect.Length, StringComparison.InvariantCultureIgnoreCase) != 0)
                FlxMessages.ThrowException(FlxErr.ErrInvalidSqlString, sql);
        }

        private string GetParamStr(IDbCommand SelectCommand, string ParamName)
        {
            //First see if there is a personalized format string.
            if (SqlParameterReplace != null && SqlParameterReplace.Trim().Length > 0)
            {
                return String.Format(CultureInfo.InvariantCulture, SqlParameterReplace, ParamName);
            }

            //Now, hard test the most usual cases.
#if (!COMPACTFRAMEWORK && !MONOTOUCH)
            if (SelectCommand is System.Data.Odbc.OdbcCommand ||
                SelectCommand is System.Data.OleDb.OleDbCommand)
            {
                return "?";
            }
#endif
            //return the "@" param.
            return "@" + ParamName;

        }

        private bool SqlParamsPositional(IDbDataParameter Param)
        {
#if (!COMPACTFRAMEWORK && !MONOTOUCH)
            if (SqlParametersType == TSqlParametersType.Automatic)
            {
                return
                    Param is System.Data.Odbc.OdbcParameter
                    ||
                    Param is System.Data.OleDb.OleDbParameter;
            }
#endif
            if (SqlParametersType == TSqlParametersType.Positional)
            {
                return true;
            }

            return false;
        }

        private void AssignSQLCommand(IDbCommand SelectCommand, string SelectText)
        {
            SelectCommand.Parameters.Clear();
            StringBuilder FinalSelect = new StringBuilder();

            string ParamStr = ReportTag.ConfigTag(ConfigTagEnum.SQLParam);

            int ParsePos = SelectText.IndexOf(ParamStr);
            int LastSelectPos = 0;
            while (ParsePos > 0)
            {
                int StartPos = ParsePos;
                while (ParsePos < SelectText.Length && SelectText[ParsePos] > (char)32 && "(),-+*".IndexOf(SelectText[ParsePos]) < 0)
                    ParsePos++;

                string ParamName = SelectText.Substring(StartPos + 1, ParsePos - StartPos - 1);

                IDbDataParameter Prm = SqlParameterList[ParamName];
                if (Prm.ParameterName != null && SelectCommand.Parameters.Contains(Prm.ParameterName))
                {
                    if (SqlParamsPositional(Prm))
                    {
                        //We need to insert a new parameter with the different name.
                        IDbDataParameter NewParam = (IDbDataParameter)((ICloneable)Prm).Clone();
                        NewParam.ParameterName = null;
                        SelectCommand.Parameters.Add(NewParam);
                    }
                }
                else
                {
                    SelectCommand.Parameters.Add(Prm);
                }

                FinalSelect.Append(SelectText, LastSelectPos, StartPos - LastSelectPos);
                LastSelectPos = ParsePos;

                FinalSelect.Append(GetParamStr(SelectCommand, ParamName));

                ParsePos = SelectText.IndexOf(ParamStr, ParsePos);
            }
            if (LastSelectPos < SelectText.Length)
                FinalSelect.Append(SelectText, LastSelectPos, SelectText.Length - LastSelectPos);

            SelectCommand.CommandText = FinalSelect.ToString();
        }

        private VirtualDataTable CreateDataSetFromSplit(string TableName, string Params, string CellAddress)
        {
            TDataSourceInfo SplitSource = null;
            string[] p = Params.Split(ReportTag.ParamDelim);
            if (p.Length != 2)
                FlxMessages.ThrowException(FlxErr.ErrInvalidSplit2Params, Params);

            SplitSource = FindDataSource(p[0], true, CellAddress);

            TValueAndXF val = new TValueAndXF();
            val.Workbook = Workbook;
            ResolveString(Convert.ToString(p[1]), -1, null).Evaluate(0, 0, 0, 0, val);
            string sSplitCount = Convert.ToString(val.Value);

            double dSplitCount = 0;
            if (!TCompactFramework.ConvertToNumber(sSplitCount, CultureInfo.InvariantCulture, out dSplitCount))
                FlxMessages.ThrowException(FlxErr.ErrInvalidSplitCount, p[1]);

            if (dSplitCount < 1 || dSplitCount > int.MaxValue)
                FlxMessages.ThrowException(FlxErr.ErrInvalidSplitCount, p[1]);

            int SplitCount = (int)dSplitCount;

            VirtualDataTable Result = new TMasterSplitDataTable(TableName, null, SplitSource.Name, SplitCount);
            return Result;
        }

        private VirtualDataTable CreateDataSetFromTopN(string TableName, string Params, string CellAddress)
        {
            TDataSourceInfo TopSource = null;
            string[] p = Params.Split(ReportTag.ParamDelim);
            if (p.Length != 2)
                FlxMessages.ThrowException(FlxErr.ErrInvalidTop2Params, Params);

            TopSource = FindDataSource(p[0], true, CellAddress);

            TValueAndXF val = new TValueAndXF();
            val.Workbook = Workbook;
#if (!COMPACTFRAMEWORK || FRAMEWORK20)
            ResolveString(Convert.ToString(p[1], CultureInfo.InvariantCulture), -1, null).Evaluate(0, 0, 0, 0, val);
            string sTopCount = Convert.ToString(val.Value, CultureInfo.InvariantCulture);
#else
            ResolveString(Convert.ToString(p[1]), -1, null).Evaluate(0,0,0,0,val);
            string sTopCount= Convert.ToString(val.Value);
#endif

            double dTopCount = 0;
            if (!TCompactFramework.ConvertToNumber(sTopCount, CultureInfo.InvariantCulture, out dTopCount))
                FlxMessages.ThrowException(FlxErr.ErrInvalidTopCount, p[1]);

            if (dTopCount < 1 || dTopCount > int.MaxValue)
                FlxMessages.ThrowException(FlxErr.ErrInvalidTopCount, p[1]);

            int TopCount = (int)dTopCount;

            VirtualDataTable Result = new TTopDataTable(TableName, TopSource.Table, TopSource.Table, TopCount);
            return Result;
        }


        private VirtualDataTable CreateDataSetFromNRows(string TableName, string Params)
        {
            string[] p = Params.Split(ReportTag.ParamDelim);
            if (p.Length != 1)
                FlxMessages.ThrowException(FlxErr.ErrInvalidNRows1Param, Params);

            TValueAndXF val = new TValueAndXF();
            val.Workbook = Workbook;
            ResolveString(Convert.ToString(p[0], CultureInfo.InvariantCulture), -1, null).Evaluate(0, 0, 0, 0, val);
            string rCount = Convert.ToString(val.Value, CultureInfo.InvariantCulture);

            double drCount = 0;
            if (!TCompactFramework.ConvertToNumber(rCount, CultureInfo.InvariantCulture, out drCount))
                FlxMessages.ThrowException(FlxErr.ErrInvalidNRowsCount, p[0]);

            if (drCount < 0 || drCount > int.MaxValue)
                FlxMessages.ThrowException(FlxErr.ErrInvalidNRowsCount, p[0]);

            int TopCount = (int)drCount;

            VirtualDataTable Result = new TNRowsDataTable(TableName, null, TopCount);
            return Result;
        }

        private VirtualDataTable CreateDataSetFromColumns(string TableName, string Params)
        {
            string[] p = Params.Split(ReportTag.ParamDelim);
            if (p.Length != 1)
                FlxMessages.ThrowException(FlxErr.ErrInvalidColumns1Param, Params);

            TValueAndXF val = new TValueAndXF();
            val.Workbook = Workbook;
            ResolveString(Convert.ToString(p[0], CultureInfo.InvariantCulture), -1, null).Evaluate(0, 0, 0, 0, val);
            string MasterName = Convert.ToString(val.Value, CultureInfo.InvariantCulture);

            VirtualDataTable Result = new TColumnsDataTable(TableName, null, GetTable(MasterName));
            return Result;
        }

        
        private DataTable CreateSQLDataSet(string TableName, string Params)
        {
            string[] p = Params.Split(ReportTag.ParamDelim);
            if (p.Length != 2)
                FlxMessages.ThrowException(FlxErr.ErrInvalidSql2Params, Params);

            string AdapterStr = p[0].Trim();

            TValueAndXF val = new TValueAndXF();
            ResolveString(AdapterStr, -1, null).Evaluate(0, 0, 0, 0, val);
            AdapterStr = Convert.ToString(val.Value).Trim();

            TAdapterData AdapterData = AdapterList[AdapterStr];
            if (AdapterData == null)
                FlxMessages.ThrowException(FlxErr.ErrInvalidSqlAdapterNotFound, AdapterStr);

            IDbDataAdapter Adapter = AdapterData.Adapter;

            if (Adapter.SelectCommand == null)
                FlxMessages.ThrowException(FlxErr.ErrInvalidSqlAdapterNoSelect, AdapterStr, Params);

            string SqlStr = p[1].Trim();
            ValidateSQL(SqlStr);
            AssignSQLCommand(Adapter.SelectCommand, SqlStr);
            //AdapterData.Adapter.TableMappings.Clear();
            //AdapterData.Adapter.TableMappings.Add("Table", TableName); //We need this because IDataAdapter does not support Fill(Dataset, tablename)
            DataTable Temp = new DataTable();
            try
            {
                Temp.Locale = AdapterData.Locale;
                Temp.CaseSensitive = AdapterData.CaseSensitive;
                ((DbDataAdapter)Adapter).Fill(Temp); //small hack, but the interfaces do not allow to fill directly a table, and we cannot use the dataset overload and dispose the dataset but not the table.

                Temp.TableName = TableName;
                return Temp;
            }
            catch
            {
                TCompactFramework.DisposeDataTable(Temp);
                throw;
            }
        }

        private void CreateConfigDataSet(string TableName, string SourceName, string RowFilter, string Sort, string CellAddress)
        {
            VirtualDataTable dt = null;
            bool NormalDataset = true;
            SourceName = SourceName.Trim();
            int fPos = SourceName.IndexOf(ReportTag.StrOpenParen);
            if (fPos > 2)
            {
                string SourceTableName = SourceName.Substring(0, fPos).Trim();
                if (String.Equals(SourceTableName, ReportTag.ConfigTag(ConfigTagEnum.SQL), StringComparison.InvariantCultureIgnoreCase))
                {
                    if (SourceName[SourceName.Length - 1] != ReportTag.StrCloseParen)
                        FlxMessages.ThrowException(FlxErr.ErrInvalidSqlParen, SourceName);

                    dt = new TAdoDotNetDataTable(TableName, null, CreateSQLDataSet(TableName, SourceName.Substring(fPos + 1, SourceName.Length - fPos - 2)), true);
                    NormalDataset = false;
                }
                else
                    if (String.Equals(SourceTableName, ReportTag.ConfigTag(ConfigTagEnum.Split), StringComparison.InvariantCultureIgnoreCase))
                    {
                        if (SourceName[SourceName.Length - 1] != ReportTag.StrCloseParen)
                            FlxMessages.ThrowException(FlxErr.ErrInvalidSplitParen, SourceName);

                        dt = CreateDataSetFromSplit(TableName, SourceName.Substring(fPos + 1, SourceName.Length - fPos - 2), CellAddress);
                        NormalDataset = false;
                    }
                    else
                        if (String.Equals(SourceTableName, ReportTag.ConfigTag(ConfigTagEnum.Top), StringComparison.InvariantCultureIgnoreCase))
                        {
                            if (SourceName[SourceName.Length - 1] != ReportTag.StrCloseParen)
                                FlxMessages.ThrowException(FlxErr.ErrInvalidTopParen, SourceName);

                            dt = CreateDataSetFromTopN(TableName, SourceName.Substring(fPos + 1, SourceName.Length - fPos - 2), CellAddress);
                            NormalDataset = false;
                        }
                        else
                            if (String.Equals(SourceTableName, ReportTag.ConfigTag(ConfigTagEnum.UserTable), StringComparison.InvariantCultureIgnoreCase))
                            {
                                if (SourceName[SourceName.Length - 1] != ReportTag.StrCloseParen)
                                    FlxMessages.ThrowException(FlxErr.ErrInvalidUserTableParen, SourceName);

                                string Param = SourceName.Substring(fPos + 1, SourceName.Length - fPos - 2);
                                TValueAndXF val = new TValueAndXF();
                                ResolveString(Param, -1, null).Evaluate(0, 0, 0, 0, val);
                                Param = FlxConvert.ToString(val.Value);

                                OnUserTable(new UserTableEventArgs(TableName, Param));
                                return;
                            }
                            else
                                if (String.Equals(SourceTableName, ReportTag.ConfigTag(ConfigTagEnum.NRows), StringComparison.InvariantCultureIgnoreCase))
                                {
                                    if (SourceName[SourceName.Length - 1] != ReportTag.StrCloseParen)
                                        FlxMessages.ThrowException(FlxErr.ErrInvalidNRowsParen, SourceName);

                                    dt = CreateDataSetFromNRows(TableName, SourceName.Substring(fPos + 1, SourceName.Length - fPos - 2));
                                    NormalDataset = false;
                                }
                                else
                                    if (String.Equals(SourceTableName, ReportTag.ConfigTag(ConfigTagEnum.Columns), StringComparison.InvariantCultureIgnoreCase))
                                    {
                                        if (SourceName[SourceName.Length - 1] != ReportTag.StrCloseParen)
                                            FlxMessages.ThrowException(FlxErr.ErrInvalidColumnsParen, SourceName);

                                        dt = CreateDataSetFromColumns(TableName, SourceName.Substring(fPos + 1, SourceName.Length - fPos - 2));
                                        NormalDataset = false;
                                    }
            }

            if (NormalDataset)
            {
                TDataSourceInfo dsi = FindDataSource(SourceName, true, CellAddress);
                if (dsi == null) return;
                dt = dsi.Table;
            }

            ConfigDataSourceList.Add(TableName, new TDataSourceInfo(TableName, dt, RowFilter, Sort, !NormalDataset, !NormalDataset, this));
        }


        private void CreateRelationship(string RelatedTables, string RelatedFields, string CellAddress)
        {
            if (RelatedTables == null || RelatedTables.Length == 0
                || RelatedFields == null || RelatedFields.Length == 0)
                FlxMessages.ThrowException(FlxErr.ErrInvalidRelationshipNullValues, CellAddress);

            string[] Tables = StrUtils.Split(RelatedTables, ReportTag.RelationshipSeparator);
            if (Tables.Length != 2)
                FlxMessages.ThrowException(FlxErr.ErrInvalidRelationship2Tables, CellAddress, RelatedTables);

            string[] Fields = RelatedFields.Split(ReportTag.ParamDelim);
            if (Fields.Length <= 0)
                FlxMessages.ThrowException(FlxErr.ErrInvalidRelationshipFields, CellAddress, RelatedFields);

            TDataSourceInfo dSource = FindDataSource(Tables[0], true, CellAddress);
            if (dSource == null) return;
            TDataSourceInfo dDest = FindDataSource(Tables[1], true, CellAddress);
            if (dDest == null) return;

            if (dSource.Table == null)
                FlxMessages.ThrowException(FlxErr.ErrInvalidRelationshipDatasetNull, CellAddress, Tables[0]);
            if (dDest.Table == null)
                FlxMessages.ThrowException(FlxErr.ErrInvalidRelationshipDatasetNull, CellAddress, Tables[1]);

            string[] ParentColumns = new string[Fields.Length];
            string[] ChildColumns = new string[Fields.Length];

            for (int i = 0; i < Fields.Length; i++)
            {
                string FieldRel = Fields[i];
                string[] FieldRels = StrUtils.Split(FieldRel, ReportTag.RelationshipSeparator);
                if (FieldRels.Length != 2)
                    FlxMessages.ThrowException(FlxErr.ErrInvalidRelationship2Fields, CellAddress, FieldRel);

                int cs = dSource.Table.GetColumn(FieldRels[0]);
                if (cs < 0)
                    FlxMessages.ThrowException(FlxErr.ErrColumNotFound, FieldRels[0], dSource.Name);

                int cd = dDest.Table.GetColumn(FieldRels[1]);
                if (cd < 0)
                    FlxMessages.ThrowException(FlxErr.ErrColumNotFound, FieldRels[1], dDest.Name);

                ParentColumns[i] = FieldRels[0];
                ChildColumns[i] = FieldRels[1];
            }

            ExtraRelations.Add(new TRelationship(dSource.Name, dDest.Name, ParentColumns, ChildColumns));
        }

        private void CreateConfigDataSets(int Row)
        {
            string TableName = Convert.ToString(Workbook.GetCellValue(Row, ConfigColTableName));
            if (TableName.Length > 0)
            {
                try
                {
                    string SourceName = Convert.ToString(Workbook.GetCellValue(Row, ConfigColSourceName));
                    TValueAndXF val = new TValueAndXF();
                    ResolveString(Convert.ToString(Workbook.GetCellValue(Row, ConfigColFilter)), -1, null).Evaluate(0, 0, 0, 0, val);
                    string RowFilter = Convert.ToString(val.Value);
                    val.Clear();
                    ResolveString(Convert.ToString(Workbook.GetCellValue(Row, ConfigColSort)), -1, null).Evaluate(0, 0, 0, 0, val);
                    string Sort = Convert.ToString(val.Value);

                    if (String.Equals(TableName, ReportTag.ConfigTag(ConfigTagEnum.Relationship), StringComparison.InvariantCultureIgnoreCase))
                    {
                        CreateRelationship(SourceName, RowFilter, new TCellAddress(Row, ConfigColSourceName).CellRef);
                    }
                    else
                    {
                        CreateConfigDataSet(TableName, SourceName, RowFilter, Sort, new TCellAddress(Row, ConfigColSourceName).CellRef);
                    }
                }
                catch (FlexCelException)
                {
                    throw;
                }
                catch (IOException ex)
                {
                    FlxMessages.ThrowException(FlxErr.ErrOnConfigTables, Row, ex.Message);
                }
            }
        }

        private bool CheckExpressionTag(string Tag)
        {
            Tag = Tag.Trim();
            if (String.Equals(Tag, ReportTag.StrOpen + ReportTag.StrDebug + ReportTag.StrClose, StringComparison.InvariantCultureIgnoreCase))
            {
                IntDebugExpressions = true;
                IntErrorsInResultFile = true;
                return true;
            }
            if (String.Equals(Tag, ReportTag.StrOpen + ReportTag.StrErrorsInResultFile + ReportTag.StrClose, StringComparison.InvariantCultureIgnoreCase))
            {
                IntErrorsInResultFile = true;
                return true;
            }
            return false;
        }

        internal void LoadConfigSheet()
        {
            if (StaticExpressionList != null)
            {
#if(FRAMEWORK20)
                foreach (KeyValuePair<string, TExpression> di in StaticExpressionList)
                {
                    ExpressionList.Add(di.Key, di.Value);
                }
#else
                foreach (DictionaryEntry di in StaticExpressionList)
                {
                    ExpressionList.Add(di.Key, di.Value);
                }
#endif
            }

            for (int i = FirstConfigRow; i <= Workbook.RowCount; i++)
            {
                //First Create the dataviews.
                //CreateDataViews(i);

                //Read the formats
                string FormatName = Convert.ToString(Workbook.GetCellValue(i, ConfigColFormatName));
                if (FormatName != null && FormatName.Length > 0)
                {
                    string FormatNameOnly = null;
                    TRichString RFormatName = new TRichString(FormatName);
                    TRichString Params = null;
                    TCellParser.ParseTag(RFormatName, out FormatNameOnly, out Params);

                    List<TRichString> Sections = new List<TRichString>();
                    TCellParser.ParseParams(RFormatName, Params, Sections);
                    TFlxApplyFormat Apply = null;
                    bool exteriorBorders = false;
                    foreach (TRichString rs in Sections)
                    {
                        if (rs == null) continue;
                        string s = Convert.ToString(rs);
                        if (s.Trim().Length <= 0) continue;

                        if (Apply == null) Apply = new TFlxApplyFormat();
                        ReportTag.ApplyFormatTag(s, FormatName, Apply, ref exteriorBorders);
                    }


                    FormatList.Add(FormatNameOnly, new TConfigFormat(Workbook.GetCellVisibleFormat(i, ConfigColFormatDef), Apply, exteriorBorders));
                }

                //Read the Expressions
                string ExpName = Convert.ToString(Workbook.GetCellValue(i, ConfigColExpName));
                if (ExpName.Length > 0)
                {
                    if (!CheckExpressionTag(ExpName))
                    {
                        string ExpBaseName = ExpName;
                        string[] ParamArray = null;
                        int ParamStart = ExpBaseName.IndexOf(ReportTag.StrOpenParen);
                        if (ParamStart > 0)
                        {
                            if (!ExpName.EndsWith(ReportTag.StrCloseParen.ToString()))
                            {
                                FlxMessages.ThrowException(FlxErr.ErrMissingParen, ExpName);
                            }
                            ExpBaseName = ExpName.Substring(0, ParamStart);
                            string ParamStr = ExpName.Substring(ParamStart + 1, ExpName.Length - ParamStart - 2);
                            ParamArray = ParamStr.Split(ReportTag.ParamDelim);
                        }
                        ExpressionList.Add(ExpBaseName, new TExpression(ParamArray, Workbook.GetCellValue(i, ConfigColExpDef)));
                    }
                }

                //Create the dataviews after the expressions have been defined, so we can use expressions in filters.
                CreateConfigDataSets(i);
            }
        }

        /// <summary>
        /// Creates a list of bands with all available datasets, to be used with sheet datasets.
        /// </summary>
        /// <param name="SheetBands"></param>
        private void FillSheetBands(ref TBand[] SheetBands)
        {
            if (SheetBands != null)
            {
                foreach (TBand b in SheetBands) if (b != null) b.Dispose();
            }

            int iLen = DataSourceList.Count + ConfigDataSourceList.Count + 1;
            SheetBands = new TBand[iLen];
            int i = 1;
            foreach (TDataSourceInfo ds in DataSourceList.Values)
            {
                SheetBands[i] = new TBand(ds.CreateDataSource(null, ExtraRelations, StaticRelations), null, SheetBands[i - 1], TXlsCellRange.FullRange(), ds.Name, TBandType.Static, false, ds.Name);
                i++;
            }
            foreach (TDataSourceInfo ds in ConfigDataSourceList.Values)
            {
                SheetBands[i] = new TBand(ds.CreateDataSource(null, ExtraRelations, StaticRelations), null, SheetBands[i - 1], TXlsCellRange.FullRange(), ds.Name, TBandType.Static, false, ds.Name);
                i++;
            }

            if (SheetBands.Length == 1)
            {
                SheetBands[0] = new TBand(null, null, null, null, TBandType.Static, false, null);
            }

        }

        internal void InsertSheets(TBandSheetList DataSheetList, TBoolArray UsePreviousTemplate, ref int ConfigSheet)
        {
            TBand[] SheetBands = null;
            try
            {
                FillSheetBands(ref SheetBands);
                bool ConfigSheetLoaded = GetStaticConfigSheet(ref SheetBands);

                int i = 1;
                while (i <= Workbook.SheetCount)
                {
                    try
                    {
                        Workbook.ActiveSheet = i;
                        using (TOneCellValue v = ResolveString(Workbook.SheetName, -1, SheetBands[SheetBands.Length - 1], true))
                        {
                            TBand SheetBand = null;
                            if (v.Count == 1 && v[0].ValueType == TValueType.Const) continue; //optimize the most common case.

                            SheetBand = FindSheetDataSet(v);
                            {
                                if (SheetBand != null && SheetBand.RecordCount == 0) //Allow empty datasets, just delete the sheet.
                                {
                                    if (RemoveSheet(i)) i--;
                                    continue;
                                }

                                TValueAndXF val = new TValueAndXF();
                                v.Evaluate(0, 0, 0, 0, val);

                                if (val.Action == TValueType.ConfigSheet)
                                {
                                    if (!ConfigSheetLoaded)
                                    {
                                        LoadConfigSheet();
                                        if (i < Workbook.SheetCount) FillSheetBands(ref SheetBands);  //So new sheets can use the datasets defined on the config.
                                    }
                                    ConfigSheet = i;
                                }
                                else
                                    if (val.Action == TValueType.DeleteSheet)
                                    {
                                        if (RemoveSheet(i)) i--;
                                    }
                                    else
                                    {
                                        int rc = 0;
                                        if (SheetBand != null) rc = SheetBand.RecordCount - 1;
                                        if (rc > 0)
                                            Workbook.InsertAndCopySheets(i, i + 1, rc);

                                        if (SheetBand != null) SheetBand.MoveFirst(TBandMoveType.Alone);
                                        for (int k = 0; k < rc + 1; k++)
                                        {
                                            Workbook.ActiveSheet = i + k;
                                            TValueAndXF VXf = new TValueAndXF();
                                            v.Evaluate(0, 0, 0, 0, VXf);
                                            Workbook.SheetName = Convert.ToString(VXf.Value);
                                            DataSheetList[i + k] = SheetBand;
                                            if (SheetBand != null) SheetBand.Refs++;

                                            if (k > 0) UsePreviousTemplate[i + k] = true;
                                            if (SheetBand != null) SheetBand.MoveNext(TBandMoveType.Alone);
                                        }
                                        i += rc;
                                        Workbook.ActiveSheet = i;
                                    }
                            }
                        }
                    }
                    finally
                    {
                        i++;
                    }
                }
            }
            finally
            {
                if (SheetBands != null)
                {
                    foreach (TBand b in SheetBands)
                        if (b != null)
                        {
                            //When using the band in the sheet report, each band has a "searchband" of other of the available bands,
                            //creating a chained list that allows FlexCel find the correct band.
                            //But when using this band inside a sheet, its SearchBand must be null, as it is the top level band.
                            b.SearchBand = null;

                            b.Dispose();
                        }
                }
            }

        }

        private bool GetStaticConfigSheet(ref TBand[] SheetBands)
        {
            //This method allows to use the config sheet even if it isn't the first, when it is a fixed text.(99% of the times)
            int asheet = Workbook.GetSheetIndex(ReportTag.StrOpen + ReportTag.StrConfigSheet + ReportTag.StrClose, false);
            if (asheet < 1) return false;
            Workbook.ActiveSheet = asheet;
            LoadConfigSheet();
            FillSheetBands(ref SheetBands);  //So new sheets can use the datasets defined on the config.
            return true;
        }

        private void FindAndLoadConfigSheet()
        {
            TBand[] SheetBands = null;
            try
            {
                FillSheetBands(ref SheetBands);

                if (GetStaticConfigSheet(ref SheetBands)) return;

                for (int i = 1; i <= Workbook.SheetCount; i++)
                {
                    Workbook.ActiveSheet = i;
                    using (TOneCellValue v = ResolveString(Workbook.SheetName, -1, SheetBands[SheetBands.Length - 1]))
                    {
                        TValueAndXF val = new TValueAndXF();
                        v.Evaluate(0, 0, 0, 0, val);
                        if (val.Action == TValueType.ConfigSheet)
                        {
                            LoadConfigSheet();
                            return;
                        }
                    }
                }
            }
            finally
            {
                if (SheetBands != null)
                {
                    foreach (TBand b in SheetBands) if (b != null) b.Dispose();
                }
            }
        }


        #endregion

        #region Read Template
        private void ReadTemplate(ref TBand MainBand, bool IsFromInclude, List<TKeepTogether> KeepRows, List<TKeepTogether> KeepCols)
        {
            FProgress.SetPhase(FlexCelReportProgressPhase.ReadTemplate);
            FindBands(MainBand, KeepRows, KeepCols);
            ReadTemplateValues(ref MainBand, IsFromInclude, KeepRows, KeepCols);
            ReadBandImages(MainBand);
        }

        private void ReadTemplateValues(ref TBand MainBand, bool IsFromInclude, List<TKeepTogether> KeepRows, List<TKeepTogether> KeepCols)
        {
            bool NeedsPreprocess = false;
            TWaitingRangeList WaitingRanges = new TWaitingRangeList();
            ReadBandValues(MainBand, true, ref NeedsPreprocess, WaitingRanges);
            if (NeedsPreprocess)
            {
                Preprocess(MainBand, WaitingRanges);

                TBand MasterBand = MainBand.MasterBand;
                string BandName = MainBand.Name;
                TXlsCellRange XlsRange = null;
                if (IsFromInclude)  //Deal with includes. In this case the band range is given by the calling template.
                {
                    XlsRange = Workbook.GetNamedRange(BandName, -1);
                    if (XlsRange == null) FlxMessages.ThrowException(FlxErr.ErrCantFindNamedRange, BandName);
                }

                MainBand.Dispose();
                MainBand = NewMainBand(XlsRange, MasterBand, BandName);
                FindBands(MainBand, KeepRows, KeepCols);
                NeedsPreprocess = false;
                ReadBandValues(MainBand, true, ref NeedsPreprocess, null);
                if (NeedsPreprocess)
                {
                    FlxMessages.ThrowException(FlxErr.ErrNoDuplicatedPreprocess);
                }
                MainBand.Preprocessed = true;
            }
        }

        private void Preprocess(TBand MainBand, TWaitingRangeList WaitingRanges)
        {
            for (int i = WaitingRanges.Count - 1; i >= 0; i--)
            {
                TWaitingCoords Coords = new TWaitingCoords(0, 0, 0, 0, MainBand.CellRange.Bottom, MainBand.CellRange.Right);
                TWaitingRange wr = WaitingRanges[i];
                if (wr != null) wr.Execute(Workbook, Coords, MainBand);
            }
        }

        private static bool IsTag(string Name, string LeftTag, string RightTag, ref string BandName, out bool DeleteLastRow, out int FixedRange)
        {
            FixedRange = 0;
            DeleteLastRow = false;

            int rt = Name.LastIndexOf(RightTag);
            if (
                (Name.Length > LeftTag.Length + RightTag.Length)
                &&
                String.Equals(Name.Substring(0, LeftTag.Length), LeftTag, StringComparison.InvariantCulture)
                &&
                rt > 0
                )
            {
                BandName = Name.Substring(LeftTag.Length, rt - LeftTag.Length);
                if (rt + RightTag.Length == Name.Length)
                {
                    return true;
                }
                string Modifier = Name.Substring(rt + RightTag.Length);
                if (String.Equals(Modifier, ReportTag.StrDeleteLastRow, StringComparison.CurrentCultureIgnoreCase))
                {
                    DeleteLastRow = true;
                    return true;
                }
                if (Modifier.StartsWith(ReportTag.StrDontInsertRanges, StringComparison.CurrentCultureIgnoreCase))
                {
                    FixedRange = -1;
                    if (Modifier.Length > ReportTag.StrDontInsertRanges.Length)
                    {
                        double dFix = 1;
                        string FixedCount = Modifier.Substring(ReportTag.StrDontInsertRanges.Length);
                        if (!string.IsNullOrEmpty(FixedCount) && TCompactFramework.ConvertToNumber(FixedCount, CultureInfo.InvariantCulture, out dFix))
                        {
                            FixedRange = (int)dFix;
                        }

                    }
                    return true;
                }
                return false;
            }
            return false;
        }

        private static bool IsDBName(string Name, ref string BandName, ref TBandType BandType, out int FixedRange, out bool DeleteLastRow, out bool IsFullRange)
        {
            FixedRange = 0;
            IsFullRange = false;

            //Ranges containing '!' (like "'Sheet1'!Name") are special in Excel, they are local and refer to just one page
            if (Name.IndexOf("!") >= 0) Name = Name.Remove(0, Name.IndexOf("!"));

            if (IsTag(Name, ReportTag.ColFull1, ReportTag.ColFull2, ref BandName, out DeleteLastRow, out FixedRange))
            {
                if (FixedRange < 0) BandType = TBandType.FixedCol; else BandType = TBandType.ColFull;
                IsFullRange = true;
                return true;
            }
            if (IsTag(Name, ReportTag.RowFull1, ReportTag.RowFull2, ref BandName, out DeleteLastRow, out FixedRange))
            {
                if (FixedRange < 0) BandType = TBandType.FixedRow; else BandType = TBandType.RowFull;
                IsFullRange = true;
                return true;
            }
            if (IsTag(Name, ReportTag.ColRange1, ReportTag.ColRange2, ref BandName, out DeleteLastRow, out FixedRange))
            {
                if (FixedRange < 0) BandType = TBandType.FixedCol; else BandType = TBandType.ColRange;
                return true;
            }
            if (IsTag(Name, ReportTag.RowRange1, ReportTag.RowRange2, ref BandName, out DeleteLastRow, out FixedRange))
            {
                if (FixedRange< 0) BandType = TBandType.FixedRow; else BandType = TBandType.RowRange;
                return true;
            }

            DeleteLastRow = false;
            return false;
        }

        private static bool IsKeepTogetherName(string Name, string Tag, out int Level)
        {
            Level = 0;

            //Ranges containing '!' (like "'Sheet1'!Name") are special in Excel, they are local and refer to just one page
            if (Name.IndexOf("!") >= 0) Name = Name.Remove(0, Name.IndexOf("!"));

#if (FRAMEWORK20)
            if (!Name.StartsWith(Tag, StringComparison.InvariantCultureIgnoreCase)) return false;
#else
            if (!Name.ToUpper(CultureInfo.InvariantCulture).StartsWith(Tag.ToUpper(CultureInfo.InvariantCulture))) return false;
#endif

            int TagEnd = Tag.Length;
            int LevelEnd = Name.IndexOf(ReportTag.StrEndKeepTogether, TagEnd);
            if (LevelEnd <= TagEnd) return false;

            string LevelStr = Name.Substring(TagEnd, LevelEnd - TagEnd);
            if (LevelStr.Length < 1 || LevelStr.Length > 8) return false;
            foreach (char ch in LevelStr)
            {
                if (!char.IsDigit(ch)) return false;
                Level = Level * 10 + (int)ch - (int)'0';
            }

            return true;
        }

        private void FindBands(TBand ParentBand, List<TKeepTogether> KeepRows, List<TKeepTogether> KeepCols)
        {
            //Locate bands on the sheet.
            int MasterCount = Workbook.NamedRangeCount;
            for (int i = 1; i <= MasterCount; i++)
            {
                TXlsNamedRange n = Workbook.GetNamedRange(i);
                string BandName = String.Empty;
                bool DeleteLastRow = false;
                bool IsFullRange; int FixedRange;
                TBandType BandType = TBandType.Static;
                if (n.SheetIndex == Workbook.ActiveSheet)
                {
                    if (IsDBName(n.Name, ref BandName, ref BandType, out FixedRange, out DeleteLastRow, out IsFullRange))
                    {
                        if (IsFullRange)
                        {
                            if (BandType == TBandType.ColFull || BandType == TBandType.FixedCol) { n.Top = 1; n.Bottom = FlxConsts.Max_Rows + 1; }
                            if (BandType == TBandType.RowFull || BandType == TBandType.FixedRow) { n.Left = 1; n.Right = FlxConsts.Max_Columns + 1; }
                        }

                        TBand NewBand = new TBand(null, ParentBand, n, n.Name, BandType, DeleteLastRow, BandName);
                        NewBand.FixedOfs = FixedRange < 0 ? 0 : FixedRange;
                        ParentBand.DetailBands.Add(NewBand);
                    }
                    else
                    {
                        int Level;

                        if (IsKeepTogetherName(n.Name, ReportTag.KeepRowsTogether, out Level))
                        {
                            KeepRows.Add(new TKeepTogether(n.Top, n.Bottom, Level));
                        }
                        else
                            if (IsKeepTogetherName(n.Name, ReportTag.KeepColsTogether, out Level))
                            {
                                KeepCols.Add(new TKeepTogether(n.Left, n.Right, Level));
                            }
                    }
                }
            }

            ParentBand.DetailBands.Sort();

            int DetailCount = ParentBand.DetailBands.Count;
            for (int i = DetailCount - 1; i > 0; i--)  //Bands are sorted in ascending order.
            {
                TBand CurrentBand = ParentBand.DetailBands[i];

                //Order the ranges one inside the other.
                for (int k = i - 1; k >= 0; k--)
                {
                    TBand BiggerBand = ParentBand.DetailBands[k];
                    TRectPos RectPos = RectUtils.TestInRect(CurrentBand.CellRange, BiggerBand.CellRange);
                    switch (RectPos)
                    {
                        case TRectPos.Inside:
                            CurrentBand.MasterBand = BiggerBand; CurrentBand.SearchBand = CurrentBand.MasterBand;
                            BiggerBand.DetailBands.Add(CurrentBand); //the details will be unsorted, but we don't care. To keep them sorted, it should be .Insert(0,CurrentBand)
                            ParentBand.DetailBands.Delete(i);
                            k = -1;
                            break;

                        case TRectPos.Separated:
                            break;

                        default: FlxMessages.ThrowException(FlxErr.ErrIntersectingRanges, BiggerBand.Name, CurrentBand.Name);
                            break;
                    } //Switch
                }
            }

            //Now that we have loaded the bands and they are ordered, load the datasources
            LoadDataSources(ParentBand);
        }

        private void LoadDataSources(TBand Band)
        {
            TDataSourceInfo di = null;

            if (Band.DataSourceName.Length > 0)
            {
                TBand MasterBand = Band.MasterBand;

                bool TableExists;
                di = FindDataSource(Band.DataSourceName, false, Band.Name, out TableExists);
                if (!TableExists)
                {
                    di = AddChildDataSource(Band);
                }

                if (di == null)
                    Band.BandType = TBandType.Ignore;
                else
                    Band.DataSource = di.CreateDataSource(MasterBand, ExtraRelations, StaticRelations);
            }

            int ChildCount = 0;
            TMasterSplitDataTable SplitMaster = null;
            if (di != null) SplitMaster = Band.DataSource.SplitMaster;
            for (int i = 0; i < Band.DetailBands.Count; i++)
            {
                LoadDataSources(Band.DetailBands[i]);

                if (SplitMaster != null && SplitMaster.DetailName == Band.DetailBands[i].DataSourceName)
                {
                    TFlexCelDataSource SplitDetail = Band.DetailBands[i].DataSource;
                    SplitMaster.DetailData = SplitDetail;
                    ChildCount++;
                }
            }

            if (SplitMaster != null)
            {
                if (ChildCount != 1) FlxMessages.ThrowException(FlxErr.ErrSplitNeedsOneAndOnlyOneDetail, di.Name, SplitMaster.DetailName);
            }

        }

        private TDataSourceInfo AddChildDataSource(TBand Band)
        {
            if (Band.MasterBand != null)
            {
                //Search for tables that are children.
                VirtualDataTable vt = TDataSourceInfo.FindLinkedTable(Band.MasterBand, Band.DataSourceName, null);
                if (vt != null)
                {
                    ConfigDataSourceList.Add(vt.TableName, new TDataSourceInfo(vt.TableName, vt, string.Empty, string.Empty, true, true, this));
                }
            }

            return FindDataSource(Band.DataSourceName, false, Band.Name);
        }

        private static bool IsInRange(TBandList BandList, int row1, int col1, int row2, int col2)
        {
            int aCount = BandList.Count;
            for (int i = 0; i < aCount; i++)
            {
                if (BandList[i].CellRange.HasRow(row1) && BandList[i].CellRange.HasCol(col1) &&
                    BandList[i].CellRange.HasRow(row2) && BandList[i].CellRange.HasCol(col2)) return true;
            }
            return false;
        }

        private static bool IsInRange(TBandList BandList, int row1, int col1)
        {
            int aCount = BandList.Count;
            for (int i = 0; i < aCount; i++)
            {
                if (BandList[i].CellRange.HasRow(row1) && BandList[i].CellRange.HasCol(col1)) return true;
            }
            return false;
        }

        internal void ReadBandValues(TBand ParentBand, bool IsMain, ref bool NeedsPreprocess, TWaitingRangeList WaitingRanges)
        {
            if (ParentBand.BandType == TBandType.Ignore)
            {
                ParentBand.Rows = new TOneRowValue[0];
                return;
            }

            int DetailCount = ParentBand.DetailBands.Count;
            for (int i = 0; i < DetailCount; i++)
            {
                ReadBandValues(ParentBand.DetailBands[i], false, ref NeedsPreprocess, WaitingRanges);
            }

            TSheetState TmpSheetState = new TSheetState(); //no need to know the real sheet state when reading the template, only when evaluating it. So we will use a dummy here.

            {
                //Read the Cells.
                int aRowCount = ParentBand.CellRange.RowCount;
                if (aRowCount < 0 || ParentBand.CellRange.ColCount < 0) FlxMessages.ThrowException(FlxErr.ErrInvalidRangeRef, ParentBand.Name);
                ParentBand.Rows = new TOneRowValue[aRowCount];

                for (int r0 = 0; r0 < aRowCount; r0++)
                {
                    int r = ParentBand.CellRange.Top + r0;
                    int aColCount = Workbook.ColCountInRow(r);
                    ParentBand.Rows[r0] = new TOneRowValue();
                    TOneRowValue CurrentRow = ParentBand.Rows[r0];
                    CurrentRow.Cols = new TOneCellValue[aColCount];
                    int c0Ofs = 0;
                    for (int c0 = 1; c0 <= aColCount; c0++)
                    {
                        int rc0 = c0 - c0Ofs; //we are updating the template on the fly, so columns can dissapear. We prefer to do this 'hack' instead of reading the template right to left, because other parts of the code rely on the correct reading order. Anyway, this only happens when there are <#preprocess> tags.
                        int c = Workbook.ColFromIndex(r, rc0);
                        if (c < ParentBand.Left) continue;
                        if (c > ParentBand.CellRange.Right) break;
                        if (IsInRange(ParentBand.DetailBands, r, c)) continue;
                        int XF = -1;

                        try
                        {
                            object v = Workbook.GetCellValueIndexed(r, rc0, ref XF);
                            TOneCellValue Cv = ResolveString(v, XF, ParentBand);

                            if (Cv.IsPreprocess)
                            {
                                NeedsPreprocess = true;

                                TValueAndXF VXf;
                                TFormatRangeList FormatRangeList;
                                TFormatRangeList FormatCellList;
                                CreateValueAndXF(WaitingRanges, TmpSheetState, out VXf, out FormatRangeList, out FormatCellList);

                                Cv.Evaluate(r - 1, c - 1, 0, 0, VXf);
                                Workbook.SetCellValue(r, c, VXf.Value);
                                c0Ofs = aColCount - Workbook.ColCountInRow(r);

                                ApplyFormats(0, 0, FormatRangeList, FormatCellList);


                                Cv.Dispose();
                            }
                            else
                            {
                                if (!NeedsPreprocess && !(Cv.Count == 1 && Cv[0].ValueType == TValueType.Const && ((TSectionConst)Cv[0]).Value is TFormula))
                                {
                                    Cv.Col = c;
                                    CurrentRow.Cols[CurrentRow.ColCount] = Cv;
                                    CurrentRow.ColCount++;
                                }
                                else
                                {
                                    Cv.Dispose(); //Formulas will not be stored, as they have to be copied. Also, if we are preprocessing, it makes no sense to continue here.
                                }
                            }

                        }
                        catch (Exception ex)
                        {
                            DoErrorInFile(r, c, ex);
                        }
                    }

                    if (!NeedsPreprocess)  //no need to keep reading the template, we will read it again.
                    {
                        //Comments.
                        CurrentRow.Comments = new TOneCellValue[Workbook.CommentCountRow(r)];
                        for (int cIndex = 1; cIndex <= CurrentRow.Comments.Length; cIndex++)
                        {
                            int cmCol = Workbook.GetCommentRowCol(r, cIndex);
                            if (cmCol < ParentBand.Left || cmCol > ParentBand.CellRange.Right) continue;
                            if (IsInRange(ParentBand.DetailBands, r, cmCol)) continue;
                            TOneCellValue cm = ResolveString(Workbook.GetCommentRow(r, cIndex), -1, ParentBand);
                            cm.Col = cmCol;
                            CurrentRow.Comments[cIndex - 1] = cm;
                        }
                    }
                }
            }
        }


        internal void ReadBandImages(TBand ParentBand)
        {
            ParentBand.SetAllHasObjects(false);

            if (ParentBand.BandType == TBandType.Ignore)
            {
                ParentBand.SetAllHasObjects(true); //we need to copy static images anyway. HasObjects is just an optimization, on this case we cannot use it.
            }

            int aCount = Workbook.ObjectCount;
            for (int i = 1; i <= aCount; i++)
            {
                TShapeProperties ShProp = Workbook.GetObjectProperties(i, true);
                FindBandImage(ParentBand, i, ShProp);
            }
        }

        private static int GetImageRow1(TShapeProperties ShProp)
        {
            if (ShProp.Anchor != null) return ShProp.Anchor.Row1;
            //The first shape governs the others.
            if (ShProp.ChildrenCount > 1 && ShProp.Children(1).ShapeType == TShapeType.NotPrimitive)
            {
                TClientAnchor Anchor = ShProp.Children(1).Anchor;
                if (Anchor != null)
                    return Anchor.Row1;
            }
            return -1;
        }

        private static int GetImageCol1(TShapeProperties ShProp)
        {
            if (ShProp.Anchor != null) return ShProp.Anchor.Col1;
            //The first shape governs the others.
            if (ShProp.ChildrenCount > 1 && ShProp.Children(1).ShapeType == TShapeType.NotPrimitive)
            {
                TClientAnchor Anchor = ShProp.Children(1).Anchor;
                if (Anchor != null)
                    return Anchor.Col1;
            }
            return -1;
        }

        private static int GetImageRow2(TShapeProperties ShProp)
        {
            if (ShProp.Anchor != null) return ShProp.Anchor.Row2;
            //The first shape governs the others.
            if (ShProp.ChildrenCount > 1 && ShProp.Children(1).ShapeType == TShapeType.NotPrimitive)
            {
                TClientAnchor Anchor = ShProp.Children(1).Anchor;
                if (Anchor != null)
                    return Anchor.Row2;
            }
            return -1;
        }

        private static int GetImageCol2(TShapeProperties ShProp)
        {
            if (ShProp.Anchor != null) return ShProp.Anchor.Col2;
            //The first shape governs the others.
            if (ShProp.ChildrenCount > 1 && ShProp.Children(1).ShapeType == TShapeType.NotPrimitive)
            {
                TClientAnchor Anchor = ShProp.Children(1).Anchor;
                if (Anchor != null)
                    return Anchor.Col2;
            }
            return -1;
        }

        private bool FindBandImage(TBand ParentBand, int ShPos, TShapeProperties ShProp)
        {
            TXlsCellRange BandRange = ParentBand.CellRange;

            if (BandRange.HasCol(GetImageCol1(ShProp)) && BandRange.HasRow(GetImageRow1(ShProp))
                && BandRange.HasCol(GetImageCol2(ShProp)) && BandRange.HasRow(GetImageRow2(ShProp)))
            {
                if (ParentBand.BandType == TBandType.Ignore) //findbands is called recursively. At any level, if it is an ignore band we leave.
                {
                    ParentBand.SetAllHasObjects(true); //we need to copy static images anyway. HasObjects is just an optimization, on this case we cannot use it.
                    return true;
                }

                ParentBand.HasObjects = true;
                int DetailCount = ParentBand.DetailBands.Count;
                for (int i = 0; i < DetailCount; i++)
                {
                    if (FindBandImage(ParentBand.DetailBands[i], ShPos, ShProp)) return true;
                }

                if (ShProp.ObjectType == TObjectType.Comment) return true;

                AddBandImage(ParentBand, ShPos, ShProp);
                return true;
            }
            return false;
        }

        private void AddBandImage(TBand ParentBand, int ShPos, TShapeProperties ShProp)
        {
            AddSubBandImage(ParentBand, ShProp);

            for (int i = 1; i <= ShProp.ChildrenCount; i++)
            {
                AddBandImage(ParentBand, ShPos, ShProp.Children(i));
            }

            if (ShProp.ShapeType == TShapeType.HostControl && ShProp.ObjectType == TObjectType.Chart)
            {
                ExcelChart Chart = Workbook.GetChart(ShPos, ShProp.ObjectPath);
                if (Chart != null)
                {
                    for (int i = 1; i <= Chart.ObjectCount; i++)
                    {
                        AddBandImage(ParentBand, ShPos, Chart.GetObjectProperties(i, true));
                    }
                }
            }
        }

        private void AddSubBandImage(TBand ParentBand, TShapeProperties ShProp)
        {
            AddBandImageText(ParentBand, ShProp.Text);

            if (ShProp.ShapeOptions != null)
            {
                string WordArtText = ShProp.ShapeOptions.AsUnicodeString(TShapeOption.gtextUNICODE, null);
                AddBandImageText(ParentBand, new TRichString(WordArtText));

                string AlternateText = ShProp.ShapeOptions.AsUnicodeString(TShapeOption.wzDescription, null);
                AddBandImageText(ParentBand, new TRichString(AlternateText));

                string AlternateText2 = ShProp.ShapeOptions.AsUnicodeString(TShapeOption.gtextRTF, null);
                AddBandImageText(ParentBand, new TRichString(AlternateText2));
            }

            //if (ShProp.ObjectType == TObjectType.Picture)  //We can apply this now to any object.
            {
                string ImageName = ShProp.ShapeName;
                if (ImageName != null && ImageName.IndexOf(ReportTag.StrOpen) == 0)
                {
                    TOneCellValue cm = ResolveString(ImageName, -1, ParentBand);
                    if (cm != null)
                    {
                        if (ParentBand.Images == null) ParentBand.Images = new TBandImages();
                        ParentBand.Images[new TRichString(ImageName)] = ResolveString(new TRichString(ImageName), -1, ParentBand);
                    }
                }
            }
        }

        private void AddBandImageText(TBand ParentBand, TRichString Text)
        {
            if (Text != null && Text.ToString().IndexOf(ReportTag.StrOpen) >= 0)
            {
                if (ParentBand.Images == null) ParentBand.Images = new TBandImages();
                ParentBand.Images[Text] = ResolveString(Text, -1, ParentBand);
            }
        }

        #endregion

        #region ExportData
        private void ExportData(TBand Band, TSheetState SheetState)
        {
            FProgress.SetPhase(FlexCelReportProgressPhase.FillData);
            FillBand(Band, 0, 0, SheetState);
        }

        private void FillBand(TBand Band, int RowOfs, int ColOfs, TSheetState SheetState)
        {
            TCopiedImageData CopiedImageData = new TCopiedImageData(0);
            InsertAndCopyBand(Band, RowOfs, ColOfs, CopiedImageData.OrigObjects);
            FillBandData(Band, RowOfs, ColOfs, ref CopiedImageData, SheetState);
        }

        private void InsertAndCopyBand(TBand Band, int RowOfs, int ColOfs, TExcelObjectList ObjectsInBand)
        {
            Band.TmpExpandedRows = 0;
            Band.TmpExpandedCols = 0;
            Band.TmpPartialRows = 0;
            Band.TmpPartialCols = 0;

            Band.ChildTmpExpandedRows = new TAddedRowColList(Band.CellRange.Left);
            Band.ChildTmpExpandedCols = new TAddedRowColList(Band.CellRange.Top);
            TXlsCellRange cr = Band.CellRange.Offset(Band.CellRange.Top + RowOfs, Band.CellRange.Left + ColOfs);

            if (Band.RecordCount == 0) //use this instead of eof since anyway we are going to use RecordCount below, and this way we avoid opening a cursor for eof.
            {
                if (IsRowRange(Band.BandType))
                {
                    if (FDeleteEmptyRanges)
                    {
                        //Cross refs will be erased by rows.
                        Workbook.DeleteRange(cr, TFlxInsertMode.ShiftRangeDown);
                        Band.AddTmpExpandedRows(-cr.RowCount, cr.Left, cr.Right);
                    }
                    else
                    {
                        Workbook.DeleteRange(cr, TFlxInsertMode.NoneDown);
                    }
                }
                else if (IsColRange(Band.BandType))
                {
                    if (FDeleteEmptyRanges)
                    {
                        Workbook.DeleteRange(cr, TFlxInsertMode.ShiftRangeRight);
                        Band.AddTmpExpandedCols(-cr.ColCount, cr.Top, cr.Bottom);
                    }
                    else
                    {
                        Workbook.DeleteRange(cr, TFlxInsertMode.NoneRight);
                    }
                }
            }


            TRangeCopyMode CopyMode = TRangeCopyMode.OnlyFormulas;
            if (!Band.HasObjects) CopyMode = TRangeCopyMode.OnlyFormulasAndNoObjects; //To speed up. there are no objects to copy anyway, and we will not have to compare them all against the range being copied.

            if (IsRowRange(Band.BandType))
            {
                if (Band.RealRecordCount > 1)
                {
                    Workbook.InsertAndCopyRange(cr, cr.Bottom + 1, cr.Left, Band.RealRecordCount - 1, (TFlxInsertMode)Band.BandType, CopyMode, null, 0, ObjectsInBand);
                    Band.AddTmpExpandedRows((Band.RealRecordCount - 1) * cr.RowCount, cr.Left, cr.Right);
                }
            }
            else
                if (IsColRange(Band.BandType))
                {
                    if (Band.RealRecordCount > 1)
                    {
                        Workbook.InsertAndCopyRange(cr, cr.Top, cr.Right + 1, Band.RealRecordCount - 1, (TFlxInsertMode)Band.BandType, CopyMode, null, 0, ObjectsInBand);
                        Band.AddTmpExpandedCols((Band.RealRecordCount - 1) * cr.ColCount, cr.Top, cr.Bottom);
                    }
                }

            if (Band.RecordCount == 1) Workbook.GetObjectsInRange(cr, ObjectsInBand);


            // "X" datasets
            if (Band.DeleteLastRow)
            {
                if (IsRowRange(Band.BandType))
                {
                    Workbook.DeleteRange(new TXlsCellRange(cr.Bottom + Band.TmpExpandedRows + 1, cr.Left, cr.Bottom + Band.TmpExpandedRows + 1, cr.Right), TFlxInsertMode.ShiftRangeDown);
                    Band.AddTmpExpandedRows(-1, cr.Left, cr.Right);
                }
                else
                {
                    Workbook.DeleteRange(new TXlsCellRange(cr.Top, cr.Right + Band.TmpExpandedCols + 1, cr.Bottom, cr.Right + Band.TmpExpandedCols + 1), TFlxInsertMode.ShiftRangeRight);
                    Band.AddTmpExpandedCols(-1, cr.Top, cr.Bottom);
                }
            }

        }


        private void FillBandData(TBand Band, int RowOfs, int ColOfs, ref TCopiedImageData CopiedImageData, TSheetState SheetState)
        {
            Band.MoveFirst(TBandMoveType.DirectChildren);

            int roInc = 0;
            int coInc = 0;
            if (IsRowRange(Band.BandType))
                roInc = Band.CellRange.RowCount;
            else
                coInc = Band.CellRange.ColCount;

            int ro = RowOfs;
            int co = ColOfs;
            int ro1 = ro;
            int co1 = co;

            bool EmptyDs = true;
            int RecordCount = Band.RecordCount; //must be done after moving the record so it has the filtered data.
            for (int k = 0; k < RecordCount; k++)
            {
                EmptyDs = false;
                TWaitingRangeList WaitingRanges = new TWaitingRangeList();
                FillOneBandData(Band, ro, co, WaitingRanges, ref CopiedImageData, SheetState);
                FillBandDetails(Band, ro, co, WaitingRanges, SheetState);
                if (IsRowRange(Band.BandType) || IsColRange(Band.BandType)) BalanceBand(Band, ro, co);
                if (k < RecordCount - 1) Band.MoveNext(TBandMoveType.DirectChildren);


                if (IsRowRange(Band.BandType))
                {
                    Band.ChildTmpExpandedRows = new TAddedRowColList(Band.Left);
                    ro1 += roInc;
                    ro = ro1 + Band.TmpPartialRows;
                }
                else if (IsColRange(Band.BandType))
                {
                    Band.ChildTmpExpandedCols = new TAddedRowColList(Band.Top);
                    co1 += coInc;
                    co = co1 + Band.TmpPartialCols;
                }
            }

            //if (Band.DataSource != null && !Band.DataSource.AllRecordsUsed()) FlxMessages.ThrowException(FlxErr.ErrInvalidReportRowCount);
            if (EmptyDs) BalanceBand(Band, 0, 0);
        }

        private void BalanceBand(TBand Band, int ro, int co)
        {
            int InsRows = Band.ChildTmpExpandedRows.Max(Band.CellRange.Right);
            int InsCols = Band.ChildTmpExpandedCols.Max(Band.CellRange.Bottom);
            Band.TmpPartialRows += InsRows;
            Band.TmpPartialCols += InsCols;

            Band.AddTmpExpandedRows(InsRows, Band.CellRange.Left + co, Band.CellRange.Right + co);
            Band.AddTmpExpandedCols(InsCols, Band.CellRange.Top + ro, Band.CellRange.Bottom + ro);

            BalanceOneBand(true, InsRows, Band.ChildTmpExpandedRows, ro, Band.CellRange.Right, Band.CellRange.Bottom, FlxConsts.Max_Rows);
            BalanceOneBand(false, InsCols, Band.ChildTmpExpandedCols, co, Band.CellRange.Bottom, Band.CellRange.Right, FlxConsts.Max_Columns);

        }

        private void BalanceOneBand(bool IsRow, int InsRows, TAddedRowColList ChildExpanded, int ro, int Right, int Bottom, int MaxRows)
        {
            for (int i = 0; i < ChildExpanded.Count; i++)
            {
                if (ChildExpanded.Cells[i] > Right) break;

                if (ChildExpanded.InsertedCount[i] < InsRows)
                {
                    int ToIns = InsRows - ChildExpanded.InsertedCount[i];
                    int LastCell = i + 1 < ChildExpanded.Count ? ChildExpanded.Cells[i + 1] - 1 : Right;
                    LastCell = Math.Min(LastCell, Right);
                    int ro2 = ro + ChildExpanded.InsertedCount[i];

                    if (Bottom + ro2 < MaxRows)
                    {
                        if (IsRow)
                        {
                            Workbook.InsertAndCopyRange(new TXlsCellRange(Bottom + ro2, ChildExpanded.Cells[i], Bottom + ro2, LastCell), Bottom + ro2 + 1, ChildExpanded.Cells[i], ToIns, TFlxInsertMode.ShiftRangeDown, TRangeCopyMode.Formats);
                        }
                        else
                        {
                            Workbook.InsertAndCopyRange(new TXlsCellRange(ChildExpanded.Cells[i], Bottom + ro2, LastCell, Bottom + ro2), ChildExpanded.Cells[i], Bottom + ro2 + 1, ToIns, TFlxInsertMode.ShiftRangeRight, TRangeCopyMode.Formats);
                        }
                    }
                }

            }
        }

        private static void CalcImageSize(byte[] imgData, double Zoom, double AspectRatio, bool BoundImage, GetImageDataEventArgs ev)
        {
            byte[] data = ImageUtils.StripOLEHeader(imgData);
            TXlsImgType imgType = ImageUtils.GetImageType(data);

            using (MemoryStream MemSt = new MemoryStream(data))  //MemSt must be OPEN during all lifetime of image objects.
            {
                Image Img = null;
                try
                {
                    if (imgType == TXlsImgType.Bmp || imgType == TXlsImgType.Jpeg || imgType == TXlsImgType.Png)
                    {
#if (!MONOTOUCH)
                        Img = new Bitmap(MemSt);
#endif
                    }
                    if (Img == null)
                        Img = TCompactFramework.GetImage(MemSt);

                    if (Img == null) return;
#if (COMPACTFRAMEWORK || (!FRAMEWORK30 && !MONOTOUCH))
                    int ImgHeight = Img.Height;
                    int ImgWidth = Img.Width;
#else
                    int ImgHeight = Img.Height();
                    int ImgWidth = Img.Width();
#endif

                    if (BoundImage)
                    {
                        Zoom = -1;
                        if (ev.Height * ImgWidth > ev.Width * ImgHeight) AspectRatio = 1; else AspectRatio = -1;
                    }


                    if (Zoom > 0)
                    {
                        ev.Height = ImgHeight * Zoom / 100.0;
                        ev.Width = ImgWidth * Zoom / 100.0;
                    }
                    if (AspectRatio > 0)
                    {
                        if (ImgWidth > 0)
                            ev.Height = ImgHeight * ev.Width / (ImgWidth * 1.0);
                    }
                    else
                        if (AspectRatio < 0)
                        {
                            if (ImgHeight > 0)
                                ev.Width = ImgWidth * ev.Height / (ImgHeight * 1.0);
                        }

                }
                finally
                {
                    if (Img != null) Img.Dispose();
                }
            }
        }

        private static void CalcImageFit(ExcelFile Workbook, TShapeProperties ImgProps, TImageFitParams p, GetImageDataEventArgs ev)
        {
            if (p.FitInRows != TAutofitGrow.None)
            {
                double RowMargin = TImageFitParams.Eval(Workbook, p.RowMargin);
                int NewHeight = Convert.ToInt32((ev.Height + RowMargin) * FlxConsts.RowMult);
                switch (p.FitInRows)
                {
                    case TAutofitGrow.All:
                        ev.File.SetRowHeight(ImgProps.Anchor.Row1, NewHeight);
                        break;

                    case TAutofitGrow.DontShrink:
                        {
                            int OldHeight = ev.File.GetRowHeight(ImgProps.Anchor.Row1);
                            if (NewHeight > OldHeight)
                            {
                                ev.File.SetRowHeight(ImgProps.Anchor.Row1, NewHeight);
                            }
                            break;
                        }

                    case TAutofitGrow.DontGrow:
                        {
                            int OldHeight = ev.File.GetRowHeight(ImgProps.Anchor.Row1);
                            if (NewHeight < OldHeight)
                            {
                                ev.File.SetRowHeight(ImgProps.Anchor.Row1, NewHeight);
                            }
                            break;
                        }

                }
            }

            if (p.FitInCols != TAutofitGrow.None)
            {
                double ColMargin = TImageFitParams.Eval(Workbook, p.ColMargin);
                int NewWidth = Convert.ToInt32((ev.Width + ColMargin) * ExcelMetrics.ColMult(ev.File));
                switch (p.FitInCols)
                {
                    case TAutofitGrow.All:
                        ev.File.SetColWidth(ImgProps.Anchor.Col1, NewWidth);
                        break;

                    case TAutofitGrow.DontShrink:
                        {
                            int OldWidth = ev.File.GetColWidth(ImgProps.Anchor.Col1);
                            if (NewWidth > OldWidth)
                            {
                                ev.File.SetColWidth(ImgProps.Anchor.Col1, NewWidth);
                            }
                            break;
                        }

                    case TAutofitGrow.DontGrow:
                        {
                            int OldWidth = ev.File.GetColWidth(ImgProps.Anchor.Col1);
                            if (NewWidth < OldWidth)
                            {
                                ev.File.SetColWidth(ImgProps.Anchor.Col1, NewWidth);
                            }
                            break;
                        }

                }
            }
        }

        private static void CalcImagePos(ExcelFile Workbook, TShapeProperties ImgProps, TImagePosParams p, GetImageDataEventArgs ev, ref int Dy1Pix, ref int Dx1Pix)
        {
            double RowOffs = TImagePosParams.Eval(Workbook, p.RowOffs);
            double ColOffs = TImagePosParams.Eval(Workbook, p.ColOffs);

            switch (p.RowAlign)
            {
                case TImageVAlign.None:
                    break;
                case TImageVAlign.Top:
                    Dy1Pix = Convert.ToInt32(RowOffs);
                    break;
                case TImageVAlign.Center:
                    {
                        double rh = ev.File.GetRowHeight(ImgProps.Anchor.Row1) / FlxConsts.RowMult;
                        Dy1Pix = Convert.ToInt32((rh - ev.Height) / 2.0 + RowOffs);
                        break;
                    }
                case TImageVAlign.Bottom:
                    {
                        double rh = ev.File.GetRowHeight(ImgProps.Anchor.Row1) / FlxConsts.RowMult;
                        Dy1Pix = Convert.ToInt32((rh - ev.Height) + RowOffs);
                        break;
                    }
            }

            switch (p.ColAlign)
            {
                case TImageHAlign.None:
                    break;
                case TImageHAlign.Left:
                    Dx1Pix = Convert.ToInt32(ColOffs);
                    break;
                case TImageHAlign.Center:
                    {
                        double cw = ev.File.GetColWidth(ImgProps.Anchor.Col1) / ExcelMetrics.ColMult(ev.File);
                        Dx1Pix = Convert.ToInt32((cw - ev.Width) / 2.0 + ColOffs);
                        break;
                    }
                case TImageHAlign.Right:
                    {
                        double cw = ev.File.GetColWidth(ImgProps.Anchor.Col1) / ExcelMetrics.ColMult(ev.File);
                        Dx1Pix = Convert.ToInt32((cw - ev.Width) + ColOffs);
                        break;
                    }
            }
        }

        private void ProcessShapeName(string ShapeName, TShapeProperties ShProp, TBand ParentBand, int RowOfs, int ColOfs, out bool ImageDeleted)
        {
            ImageDeleted = false;
            if (ShapeName == null || ShapeName.IndexOf(ReportTag.StrOpen) != 0) return;//Optimize most usual case.

            TOneCellValue cm;
            if (!ParentBand.Images.TryGetValue(new TRichString(ShapeName), out cm)) return;

            TValueAndXF Cvxf = new TValueAndXF();
            cm.Evaluate(0, 0, RowOfs, ColOfs, Cvxf);

            if (Cvxf.ImageDelete)
            {
                Workbook.DeleteObject(0, ShProp.ObjectPathShapeId);
                ImageDeleted = true;
                return;
            }

            if (ShProp.ObjectType == TObjectType.Picture)
            {
                ReplaceOneImage(ParentBand, ShapeName, Cvxf, ShProp);
            }

        }

        private void ReplaceOneImage(TBand ParentBand, string ImageName, TValueAndXF Cvxf, TShapeProperties ShProp)
        {
            byte[] imgData = (Cvxf.Value as byte[]);
            if (imgData == null) return;

            double w = 0;
            double h = 0;
            ShProp.Anchor.CalcImageCoords(ref h, ref w, Workbook);

            GetImageDataEventArgs ev = new GetImageDataEventArgs(Workbook, ImageName, imgData, h, w);

            if (Cvxf.ImageSize != null)
            {
                CalcImageSize(imgData, Cvxf.ImageSize.Zoom, Cvxf.ImageSize.AspectRatio, Cvxf.ImageSize.BoundImage, ev);
            }


            OnGetImageData(ev);
            imgData = ev.ImageData;

            if (Cvxf.ImageFit != null) //imagefit should be before imagepos, and after imagesize.
            {
                CalcImageFit(Workbook, ShProp, Cvxf.ImageFit, ev);
            }

            int Dy1Pix0 = ShProp.Anchor.Dy1Pix(Workbook);
            int Dx1Pix0 = ShProp.Anchor.Dx1Pix(Workbook);
            int Dy1Pix = Dy1Pix0;
            int Dx1Pix = Dx1Pix0;

            if (Cvxf.ImagePos != null)
            {
                CalcImagePos(Workbook, ShProp, Cvxf.ImagePos, ev, ref Dy1Pix, ref Dx1Pix);
            }

            if (ev.Height != h || ev.Width != w || Dy1Pix != Dy1Pix0 || Dx1Pix != Dx1Pix0)
            {
                ShProp.Anchor = new TClientAnchor(ShProp.Anchor.AnchorType,
                    ShProp.Anchor.Row1, Dy1Pix, ShProp.Anchor.Col1, Dx1Pix, (int)Math.Round(ev.Height), (int)Math.Round(ev.Width), Workbook);
                Workbook.SetObjectAnchor(0, ShProp.ObjectPathShapeId, ShProp.Anchor);
            }
            Workbook.SetImage(0, imgData, true, ShProp.ObjectPathShapeId);
        }


        private bool ReplaceShapeText(TRichString Text, int RowOfs, int ColOfs, TBand ParentBand, out object ReplacedText)
        {
            ReplacedText = String.Empty;
            if (Text == null || Text.ToString().IndexOf(ReportTag.StrOpen) < 0) return false;//Optimize most usual case.

            TOneCellValue cm;
            if (!ParentBand.Images.TryGetValue(Text, out cm)) return false;
            TValueAndXF Cvxf = new TValueAndXF();
            cm.Evaluate(0, 0, RowOfs, ColOfs, Cvxf, false);

            if (Cvxf == null) return false;

            if ((Cvxf.IncludeHtml == TIncludeHtml.Undefined && HtmlMode) || Cvxf.IncludeHtml == TIncludeHtml.Yes)
            {
                ReplacedText = TRichString.FromHtml(Convert.ToString(Cvxf.Value), Workbook.GetDefaultFormat, Workbook);
            }
            else
            {
                ReplacedText = Cvxf.Value;
            }
            return true;
        }

        private void ReplaceOneAutoShape(TShapeProperties ShProp, int RowOfs, int ColOfs, TBand ParentBand, IEmbeddedObjects EmbeddedObj)
        {
            for (int i = 1; i <= ShProp.ChildrenCount; i++)
                ReplaceOneAutoShape(ShProp.Children(i), RowOfs, ColOfs, ParentBand, EmbeddedObj);

            if (ShProp.ShapeType == TShapeType.HostControl && ShProp.ObjectType == TObjectType.Chart)
            {
                ExcelChart Chart = Workbook.GetChart(0, ShProp.ObjectPathShapeId);
                if (Chart != null)
                {
                    for (int i = 1; i <= Chart.ObjectCount; i++)
                    {
                        ReplaceOneAutoShape(Chart.GetObjectProperties(i, true), RowOfs, ColOfs, ParentBand, Chart);
                    }
                }
            }


            bool ImageDeleted;
            ProcessShapeName(ShProp.ShapeName, ShProp, ParentBand, RowOfs, ColOfs, out ImageDeleted);
            if (ImageDeleted) return;


            object Text = null;
            if (ReplaceShapeText(ShProp.Text, RowOfs, ColOfs, ParentBand, out Text))
            {
                TRichString rs = (Text as TRichString);
                if (rs != null)
                    EmbeddedObj.SetObjectText(0, ShProp.ObjectPathShapeId, rs);
                else
                    EmbeddedObj.SetObjectText(0, ShProp.ObjectPathShapeId, new TRichString(Convert.ToString(Text)));
            }

            if (ShProp.ShapeOptions != null)
            {
                ReplaceStringProp(ShProp, RowOfs, ColOfs, ParentBand, TShapeOption.gtextUNICODE); //WordArt
                ReplaceStringProp(ShProp, RowOfs, ColOfs, ParentBand, TShapeOption.wzDescription); //Alternate text.
                ReplaceStringProp(ShProp, RowOfs, ColOfs, ParentBand, TShapeOption.gtextRTF); //Alternate text.

                THyperLink HLink = ShProp.ShapeOptions.AsHyperLink(TShapeOption.pihlShape, null);
                if (HLink != null && HLink.Text != null && HLink.Text.Length > 0)
                {
                    string HText = GetHLinkTag(HLink.Text);
                    string HHint = GetHLinkTag(HLink.Hint);
                    bool ConstantText = HText.IndexOf(ReportTag.StrOpen) < 0;
                    if (ConstantText && HHint.IndexOf(ReportTag.StrOpen) < 0)
                    { }//Optimize most usual case.
                    else
                    {
                        ReplaceHLinkText(ParentBand, RowOfs, ColOfs, HLink, HText);
                        ReplaceHLinkHint(ParentBand, RowOfs, ColOfs, HLink, HHint);

                        if (HLink.Text == null || HLink.Text.Length == 0)
                        {
                            HLink = null; //to remove it.
                            Workbook.SetObjectProperty(0, ShProp.ObjectPathShapeId, TShapeOption.fPrint, 3, false); //fisButton
                        }
                        Workbook.SetObjectProperty(0, ShProp.ObjectPathShapeId, TShapeOption.pihlShape, HLink);

                    }
                }
            }

        }

        private void ReplaceStringProp(TShapeProperties ShProp, int RowOfs, int ColOfs, TBand ParentBand, TShapeOption ShapeOption)
        {
            string CurrentText = ShProp.ShapeOptions.AsUnicodeString(ShapeOption, null);
            object Text;
            if (CurrentText != null && CurrentText.Length > 0
                && ReplaceShapeText(new TRichString(CurrentText), RowOfs, ColOfs, ParentBand, out Text))
            {
                string NewText = FlxConvert.ToString(Text);
                if (NewText.Length == 0) NewText = null;
                Workbook.SetObjectProperty(0, ShProp.ObjectPathShapeId, ShapeOption, NewText);
            }

        }

        private void ReplaceAutoShapes(TBand ParentBand, int RowOfs, int ColOfs, ref TCopiedImageData CopiedImageData)
        {
            if (ParentBand.Images == null || ParentBand.Images.Count <= 0) return;
            TXlsCellRange BandRange = ParentBand.CellRange.Offset(ParentBand.CellRange.Top + RowOfs, ParentBand.CellRange.Left + ColOfs);

            //We need to replace the first record (the original objects).
            long[] ShapeIds = CopiedImageData.GetObjects();
            CopiedImageData.RecordPos++;
            foreach (long ShapeId in ShapeIds)
            {
                TShapeProperties ShProp = Workbook.GetObjectPropertiesByShapeId(ShapeId, true);
                if (ShProp == null) continue; //object has been erased.
                ProcessAutoShape(ParentBand, RowOfs, ColOfs, BandRange, ShProp);

            }
        }

        private void ProcessAutoShape(TBand ParentBand, int RowOfs, int ColOfs, TXlsCellRange BandRange,
            TShapeProperties ShProp)
        {
            if (ShProp.ObjectType == TObjectType.Comment) return;
            ReplaceOneAutoShape(ShProp, RowOfs, ColOfs, ParentBand, Workbook);
        }

        private static bool IsUrlEmpty(string s)
        {
            const string UrlSep = "://";
            int i = s.IndexOf(UrlSep);
            if (i < 0) return false;
            for (int k = i + UrlSep.Length; k < s.Length; k++)
                if (s[k] != '/') return false;
            return true;
        }

        private static string GetHLinkTag(string s)
        {
            string s1 = GetHLinkTag(s, ReportTag.StrOpenHLink2, ReportTag.StrOpen, ReportTag.StrCloseHLink, ReportTag.StrClose);
            return GetHLinkTag(s1, ReportTag.StrOpenHLink, ReportTag.StrOpen, ReportTag.StrClose, ReportTag.StrClose);
        }

        private static string GetHLinkTag(string s, string OldOpenTag, string NewOpenTag, string OldCloseTag, string NewCloseTag)
        {
            StringBuilder Res = new StringBuilder();

            bool InTag = false;
            int i = 0;
            while (i < s.Length)
            {
                if (InTag)
                {
                    DoChar(s, Res, ref InTag, ref i, OldCloseTag, NewCloseTag);
                }
                else
                {
                    DoChar(s, Res, ref InTag, ref i, OldOpenTag, NewOpenTag);
                }
            }

            return Res.ToString();
        }

        private static void DoChar(string s, StringBuilder Res, ref bool InTag, ref int i, string OldTag, string NewTag)
        {
            bool Found = true;
            for (int k = 0; k < OldTag.Length; k++)
            {
                if (s[i + k] != OldTag[k]) { Found = false; break; }
            }
            if (Found)
            {
                InTag = !InTag;
                Res.Append(NewTag);
                i += OldTag.Length;
            }
            else
            {
                Res.Append(s[i]);
                i++;
            }
        }

        private void ReplaceHyperLinks(TBand ParentBand, int RowOfs, int ColOfs)
        {
            TXlsCellRange BandRange = ParentBand.CellRange.Offset(ParentBand.CellRange.Top + RowOfs, ParentBand.CellRange.Left + ColOfs);

            int aCount = Workbook.HyperLinkCount;
            for (int i = aCount; i > 0; i--)
            {
                THyperLink HLink = Workbook.GetHyperLink(i);
                if (HLink == null) continue;//Optimize most usual case.
                string HText = GetHLinkTag(HLink.Text);
                string HHint = GetHLinkTag(HLink.Hint);
                string HTextMark = GetHLinkTag(HLink.TextMark);

                bool ConstantText = HText.IndexOf(ReportTag.StrOpen) < 0;
                bool ConstantHint = HHint.IndexOf(ReportTag.StrOpen) < 0;
                bool ConstantTextMark = HTextMark.IndexOf(ReportTag.StrOpen) < 0;
                if (ConstantText && ConstantHint && ConstantTextMark) continue;//Optimize most usual case.
                TXlsCellRange HLinkPos = Workbook.GetHyperLinkCellRange(i);
                if (BandRange.HasCol(HLinkPos.Left) && BandRange.HasRow(HLinkPos.Top))
                {
                    if (IsInRange(ParentBand.DetailBands, HLinkPos.Top - RowOfs, HLinkPos.Left - ColOfs)) continue;
                    ReplaceHLinkText(ParentBand, RowOfs, ColOfs, HLink, HText);
                    ReplaceHLinkHint(ParentBand, RowOfs, ColOfs, HLink, HHint);
                    ReplaceHLinkTextMark(ParentBand, RowOfs, ColOfs, HLink, HTextMark);

                    if (EmptyHLink(HLink, ConstantText, ConstantTextMark))
                    {
                        Workbook.DeleteHyperLink(i);
                    }
                    else
                        Workbook.SetHyperLink(i, HLink);

                }
            }
        }

        private static bool EmptyHLink(THyperLink HLink, bool ConstantText, bool ConstantTextMark)
        {
            if (ConstantText && ConstantTextMark) return false;
            if (!String.IsNullOrEmpty(HLink.Text) && !IsUrlEmpty(HLink.Text)) return false;
            if (!String.IsNullOrEmpty(HLink.TextMark)) return false;
            return true;
        }

        private void ReplaceHLinkHint(TBand ParentBand, int RowOfs, int ColOfs, THyperLink HLink, string HHint)
        {
            using (TOneCellValue cm = ResolveString(HHint, -1, ParentBand))
            {

                TValueAndXF Cvxf = new TValueAndXF();
                if (cm != null && (cm.Count != 1 || cm[0].ValueType != TValueType.Const)) //Optimize most usual case. 
                {
                    cm.Evaluate(0, 0, RowOfs, ColOfs, Cvxf);
                    HLink.Hint = Convert.ToString(Cvxf.Value);
                }
            }
        }

        private void ReplaceHLinkTextMark(TBand ParentBand, int RowOfs, int ColOfs, THyperLink HLink, string HTextMark)
        {
            using (TOneCellValue cm = ResolveString(HTextMark, -1, ParentBand))
            {

                TValueAndXF Cvxf = new TValueAndXF();
                if (cm != null && (cm.Count != 1 || cm[0].ValueType != TValueType.Const)) //Optimize most usual case. 
                {
                    cm.Evaluate(0, 0, RowOfs, ColOfs, Cvxf);
                    HLink.TextMark = Convert.ToString(Cvxf.Value);
                }
            }
        }

        private void ReplaceHLinkText(TBand ParentBand, int RowOfs, int ColOfs, THyperLink HLink, string HText)
        {
            using (TOneCellValue cm = ResolveString(HText, -1, ParentBand))
            {

                TValueAndXF Cvxf = new TValueAndXF();
                if (cm != null && (cm.Count != 1 || cm[0].ValueType != TValueType.Const)) //Optimize most usual case. 
                {
                    cm.Evaluate(0, 0, RowOfs, ColOfs, Cvxf);
                    HLink.Text = Convert.ToString(Cvxf.Value);
                    string Text = HLink.Text.ToUpper(CultureInfo.InvariantCulture);
                    if (HLink.LinkType == THyperLinkType.LocalFile &&
                        (Text.StartsWith("HTTP:") || Text.StartsWith("HTTPS:") || Text.StartsWith("FTP:")))
                    {
                        HLink.LinkType = THyperLinkType.URL;  //Convert local files to URLS.
                    }
                    if
                        (Text.StartsWith("FILE://"))
                    {
                        HLink.LinkType = THyperLinkType.LocalFile;  //Convert local files to URLS.
                    }
                    if
                        (Text.StartsWith("UNC://") || Text.StartsWith(@"\\"))
                    {
                        HLink.LinkType = THyperLinkType.UNC;  //Convert local files to URLS.
                    }
                }
            }
        }

        private int CalcDateFormat(DateTime dt, TFlxFormat XfDef)
        {
            XfDef.Format = FlexCel.XlsAdapter.TFormatRecordList.GetInternalFormat(0x14);  //We are going to define 3 formats, date, datetime, time.

            if (dt.TimeOfDay == TimeSpan.Zero)
            {
                XfDef.Format = FlexCel.XlsAdapter.TFormatRecordList.GetInternalFormat(0x0E);  //Date
                return Workbook.AddFormat(XfDef);
            }
            else
                if (dt.Date == DateTime.MinValue)
                {
                    XfDef.Format = FlexCel.XlsAdapter.TFormatRecordList.GetInternalFormat(0x14);  //Time
                    return Workbook.AddFormat(XfDef);
                }
                else
                {
                    XfDef.Format = FlexCel.XlsAdapter.TFormatRecordList.GetInternalFormat(0x16);  //DateTime.
                    return Workbook.AddFormat(XfDef);
                }

        }

        private void SetCellValue(int Row, int Col, object Value, int XF, bool AddFormula, TIncludeHtml UseHtml)
        {
            string sValue = FlxConvert.ToString(Value);
            if ((AddFormula && Value != null)
                || (EnterFormulas && sValue.Length > 0 && sValue.StartsWith(TFormulaMessages.TokenString(TFormulaToken.fmStartFormula)))
                )
            {
                Workbook.SetCellValue(Row, Col, new TFormula(FlxConvert.ToString(Value)), XF);
                return;
            }

            if ((UseHtml == TIncludeHtml.Undefined && HtmlMode) || UseHtml == TIncludeHtml.Yes)
            {
                string s = Value as string;
                if (s != null)
                {
                    Workbook.SetCellFromHtml(Row, Col, s, XF);
                    return;
                }
            }

            if (FTryToConvertStrings)
                Workbook.SetCellFromString(Row, Col, FlxConvert.ToString(Value), XF);
            else
                Workbook.SetCellValue(Row, Col, Value, XF);
        }

        private void DoErrorInFile(int r, int c, Exception ex)
        {
            if (IntErrorsInResultFile)
            {
                TFlxFormat fmt = Workbook.GetCellVisibleFormatDef(r, c);
                fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
                fmt.FillPattern.FgColor = Colors.Yellow;
                fmt.Font.Color = Colors.Red;

                Workbook.SetCellValue(r, c, FlxMessages.GetString(FlxErr.CellError, ex.Message), Workbook.AddFormat(fmt));
            }
            else
            {
                TCellAddress adr = new TCellAddress(Workbook.SheetName, r, c, false, false);
                string ErrMsg = FlxMessages.GetString(FlxErr.ErrAtCell, adr.CellRef, ex.Message);
                throw new FlexCelCoreException(ErrMsg, FlxErr.ErrAtCell, ex);
            }
        }

        private void FillOneBandData(TBand Band, int RowOfs, int ColOfs, TWaitingRangeList WaitingRanges, ref TCopiedImageData CopiedImageData, TSheetState SheetState)
        {
            if (Band.Rows == null) return;
            int RowLen = Band.Rows.Length;
            TValueAndXF VXf;
            TFormatRangeList FormatRangeList;
            TFormatRangeList FormatCellList;
            CreateValueAndXF(WaitingRanges, SheetState, out VXf, out FormatRangeList, out FormatCellList);

            //we do this first so images are replaced before comments or ranges are deleted, but in the next loop iteration it will be the same anyway.
            ReplaceAutoShapes(Band, RowOfs, ColOfs, ref CopiedImageData);

            for (int r = 0; r < RowLen; r++)
            {
                int ColLen = Band.Rows[r].ColCount;
                bool ColumnDone = false;
                VXf.Clear(); //In case it doesn't enter the column loop.
                VXf.DebugStack = null;
                TConfigFormat RowVXF = null;
                for (int c0 = 0; c0 < ColLen; c0++)
                {
                    VXf.Clear();
                    if (IntDebugExpressions) VXf.DebugStack = new TDebugStack();
                    TOneCellValue Cell = Band.Rows[r].Cols[c0];

                    try
                    {
                        Cell.Evaluate(Band.CellRange.Top + r - 1, Cell.Col - 1, RowOfs, ColOfs, VXf);
                        if (VXf.AutoPageBreaksPercent >= 0)
                        {
                            SheetState.AutoPageBreaksPercent = VXf.AutoPageBreaksPercent;
                            SheetState.AutoPageBreaksPageScale = VXf.AutoPageBreaksPageScale;
                        }

                        FormatCellList.SetCurrentCell(Band.CellRange.Top + r, Cell.Col);

                        SetCellValueWithDebug(Band, RowOfs, ColOfs, VXf, r, Cell);
                    }
                    catch (Exception ex)
                    {
                        DoErrorInFile(Band.CellRange.Top + r + RowOfs, Cell.Col + ColOfs, ex);
                    }

                    if (VXf.XFRow != null) RowVXF = VXf.XFRow;

                    if (VXf.XFCol != null)
                    {
                        if (VXf.XFCol.ApplyFmt == null)
                        {
                            Workbook.SetColFormat(Cell.Col + ColOfs, VXf.XFCol.XF);
                        }
                        else
                        {
                            Workbook.SetColFormat(Cell.Col + ColOfs, Workbook.GetFormat(VXf.XFCol.XF), VXf.XFCol.ApplyFmt, true);
                        }
                    }

                    if (VXf.FullDataSetColumnCount > 0)
                    {
                        int CCount = VXf.FullDataSetColumnCount;
                        TValueAndXF VXf1;
                        CreateValueAndXF(WaitingRanges, SheetState, out VXf1);
                        VXf1.FormatCellList = FormatCellList;
                        VXf1.FormatRangeList = FormatRangeList;

                        for (int i = 0; i < CCount; i++) //i=0 has already been filled, but we will again because of datetimes.
                        {
                            try
                            {
                                VXf1.Clear();
                                if (IntDebugExpressions) VXf1.DebugStack = new TDebugStack();
                                VXf1.FullDataSetColumnIndex = i;
                                Cell.Evaluate(Band.CellRange.Top + r - 1, Cell.Col - 1 + i, RowOfs, ColOfs, VXf1);
                                FormatCellList.SetCurrentCell(Band.CellRange.Top + r, Cell.Col + i);

                                object rv = VXf1.Value;
                                if (rv is DateTime)
                                {
                                    TFlxFormat XfDef = Workbook.GetFormat(VXf1.XF);
                                    VXf1.XF = CalcDateFormat(Convert.ToDateTime(rv, CultureInfo.CurrentCulture), XfDef);
                                }

                                SetCellValueWithDebug(Band, RowOfs, ColOfs + i, VXf1, r, Cell);
                            }
                            catch (Exception ex)
                            {
                                DoErrorInFile(Band.CellRange.Top + r + RowOfs, Cell.Col + ColOfs + i, ex);
                            }

                        }

                        ApplyFormats(RowOfs, ColOfs, FormatRangeList, FormatCellList);

                        TXlsCellRange AutoFilter = Workbook.GetAutoFilterRange();
                        if (AutoFilter != null && CCount > 1)
                        {
                            if (AutoFilter.Top == Band.CellRange.Top + r &&
                                AutoFilter.Left <= Cell.Col && AutoFilter.Right >= Cell.Col)
                            {
                                Workbook.SetAutoFilter(AutoFilter.Top, Math.Min(AutoFilter.Left, Cell.Col), Math.Max(AutoFilter.Right, Cell.Col + CCount - 1));
                            }
                        }
                        ColumnDone = true;
                    } // Full datasets.

                    if (ColumnDone) break;
                }
                if (RowVXF != null)
                {
                    if (RowVXF.ApplyFmt == null)
                    {
                        Workbook.SetRowFormat(Band.CellRange.Top + r + RowOfs, RowVXF.XF);
                    }
                    else
                    {
                        Workbook.SetRowFormat(Band.CellRange.Top + r + RowOfs, Workbook.GetFormat(RowVXF.XF), RowVXF.ApplyFmt, true);
                    }
                }

                //Comments
                for (int cIndex = 1; cIndex <= Band.Rows[r].Comments.Length; cIndex++)
                {
                    TValueAndXF Cvxf = new TValueAndXF();
                    TOneCellValue cm = Band.Rows[r].Comments[cIndex - 1];

                    if (cm == null || (cm.Count == 1 && cm[0].ValueType == TValueType.Const)) continue; //Optimize most usual case.

                    cm.Evaluate(0, 0, RowOfs, ColOfs, Cvxf);
                    TRichString rs = (Cvxf.Value as TRichString);
                    if (rs != null)
                        Workbook.SetComment(Band.CellRange.Top + r + RowOfs, cm.Col + ColOfs, rs);
                    else
                        Workbook.SetComment(Band.CellRange.Top + r + RowOfs, cm.Col + ColOfs, Convert.ToString(Cvxf.Value));
                }
            }

            ApplyFormats(RowOfs, ColOfs, FormatRangeList, FormatCellList);

            ReplaceHyperLinks(Band, RowOfs, ColOfs);
        }

        private void SetCellValueWithDebug(TBand Band, int RowOfs, int ColOfs, TValueAndXF VXf, int r, TOneCellValue Cell)
        {
            if (IntDebugExpressions)
                Workbook.SetCellValue(Band.CellRange.Top + r + RowOfs, Cell.Col + ColOfs, VXf.DebugStack.ToRichString(Workbook, VXf.XF), VXf.XF);
            else
                SetCellValue(Band.CellRange.Top + r + RowOfs, Cell.Col + ColOfs, VXf.Value, VXf.XF, VXf.IsFormula, VXf.IncludeHtml);
        }

        private void CreateValueAndXF(TWaitingRangeList WaitingRanges, TSheetState SheetState, out TValueAndXF VXf, out TFormatRangeList FormatRangeList, out TFormatRangeList FormatCellList)
        {
            FormatRangeList = new TFormatRangeList();
            FormatCellList = new TFormatRangeList();
            CreateValueAndXF(WaitingRanges, SheetState, out VXf);
            VXf.FormatRangeList = FormatRangeList; //So the formats get included there.
            VXf.FormatCellList = FormatCellList; //We will have to rewrite them at the end.
        }

        private void CreateValueAndXF(TWaitingRangeList WaitingRanges, TSheetState SheetState, out TValueAndXF VXf)
        {
            VXf = new TValueAndXF();
            VXf.WaitingRanges = WaitingRanges;
            VXf.Workbook = Workbook;
            VXf.AutofitInfo = SheetState.AutofitInfo;
        }

        private void ApplyFormats(int RowOfs, int ColOfs, TFormatRangeList FormatRangeList, TFormatRangeList FormatCellList)
        {
            foreach (TFormatRange fr in FormatRangeList)
            {
                fr.ApplyFormat(Workbook, RowOfs, ColOfs);
            }

            //Cell formats have higher priority than Range Formats.
            foreach (TFormatRange fr in FormatCellList)
            {
                fr.ApplyFormat(Workbook, RowOfs, ColOfs);
            }
        }

        internal static bool IsRowRange(TBandType BandType)
        {
            return BandType == TBandType.RowFull || BandType == TBandType.RowRange || BandType == TBandType.FixedRow;
        }

        internal static bool IsColRange(TBandType BandType)
        {
            return BandType == TBandType.ColFull || BandType == TBandType.ColRange || BandType == TBandType.FixedCol;
        }

        internal static void FillAllBandList(List<ITopLeft> AllRanges, TBandList DetailBands, TWaitingRangeList WaitingRanges)
        {
            for (int i = DetailBands.Count - 1; i >= 0; i--)
            {
                AllRanges.Add(DetailBands[i]);
            }

            for (int i = WaitingRanges.Count - 1; i >= 0; i--)
            {
                AllRanges.Add(WaitingRanges[i]);
            }

            AllRanges.Sort(TBandList.RowColComparerMethod);
        }

        private void DoRanges(List<ITopLeft> AllRanges, int RowOfs, int ColOfs, TBand MainBand, TSheetState SheetState)
        {
            for (int i = AllRanges.Count - 1; i >= 0; i--)
            {
                TBand detBand = (AllRanges[i] as TBand);
                if (detBand != null)
                {
                    FillBand(detBand, RowOfs, ColOfs, SheetState);
                    continue;
                }

                TWaitingRange wr = (AllRanges[i] as TWaitingRange);
                if (wr != null)
                {
                    TWaitingCoords Coords = new TWaitingCoords(RowOfs, ColOfs, MainBand.TmpExpandedRows + MainBand.ChildTmpExpandedRows.Max(MainBand.CellRange.Right),
                        MainBand.TmpExpandedCols + MainBand.ChildTmpExpandedCols.Max(MainBand.CellRange.Bottom),
                        MainBand.CellRange.Bottom + RowOfs, MainBand.CellRange.Right + ColOfs);

                    wr.Execute(Workbook, Coords, MainBand);
                    continue;
                }
            }
        }

        internal void FillBandDetails(TBand MainBand, int RowOfs, int ColOfs, TWaitingRangeList WaitingRanges, TSheetState SheetState)
        {
            TBandList DetailBands = MainBand.DetailBands;
            //Remember that a range can grow both vertically and horizontally, because it can have ranges inside.

            List<ITopLeft> AllRanges1 = new List<ITopLeft>(DetailBands.Count + WaitingRanges.Count); //Mix all real ranges with pending includes/deletes
            FillAllBandList(AllRanges1, DetailBands, WaitingRanges);

            DoRanges(AllRanges1, RowOfs, ColOfs, MainBand, SheetState);
        }
        #endregion

        #region Headers and footers
        private string ReplaceOneHeaderOrFooter(string s, TBand SheetBand)
        {
            using (TOneCellValue v = ResolveString(s, -1, SheetBand))
            {
                if (!(v == null || (v.Count == 1 && v[0].ValueType == TValueType.Const)))
                {
                    TValueAndXF val = new TValueAndXF();
                    v.Evaluate(0, 0, 0, 0, val);
                    return Convert.ToString(val.Value);
                }
            }
            return s;

        }
        private void ReplaceHeadersAndFooters(TBand SheetBand)
        {
            THeaderAndFooter HeadFoot = Workbook.GetPageHeaderAndFooter();

            HeadFoot.DefaultHeader = ReplaceOneHeaderOrFooter(HeadFoot.DefaultHeader, SheetBand);
            HeadFoot.DefaultFooter = ReplaceOneHeaderOrFooter(HeadFoot.DefaultFooter, SheetBand);
            HeadFoot.EvenHeader = ReplaceOneHeaderOrFooter(HeadFoot.EvenHeader, SheetBand);
            HeadFoot.EvenFooter = ReplaceOneHeaderOrFooter(HeadFoot.EvenFooter, SheetBand);
            HeadFoot.FirstHeader = ReplaceOneHeaderOrFooter(HeadFoot.FirstHeader, SheetBand);
            HeadFoot.FirstFooter = ReplaceOneHeaderOrFooter(HeadFoot.FirstFooter, SheetBand);

            Workbook.SetPageHeaderAndFooter(HeadFoot);
        }
        #endregion

        #region CleanUp
        /// <summary>
        /// This cleans the resources allocated by preload.
        /// </summary>
        internal void Unload()
        {
            if (ConfigDataSourceList != null)
            {
                ConfigDataSourceList.Dispose();
                ConfigDataSourceList = null;
            }
        }

        /// <summary>
        /// Override this instance when creating descendants of FlexCelReport.
        /// Remember to always call base.Dispose(disposing).
        /// </summary>
        /// <param name="disposing"></param>
        protected override void Dispose(bool disposing)
        {
            try
            {
                if (disposing)
                {
                    if (DataSourceList != null) { DataSourceList.Dispose(); DataSourceList = null; }
                    if (ConfigDataSourceList != null) { ConfigDataSourceList.Dispose(); ConfigDataSourceList = null; }  //Just to make sure, it should never come here.
                }

            }
            finally
            {
                base.Dispose(disposing);
            }

        }


        #endregion
    }
}



