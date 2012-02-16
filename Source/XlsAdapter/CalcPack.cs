using System;
using FlexCel.Core;
using System.Diagnostics;
using System.Collections.Generic;
using System.Globalization;

namespace FlexCel.AddinFunctions
{
    #region Utils
    internal static class FinancialUtils
    {

        internal static object TryGetBasis(ExcelFile Xls, object[] parameters, int pPos, out TDayCountBasis Basis)
        {
            Basis = TDayCountBasis.UsPsa30_360;
            double dBasis = 0;
            if (parameters.Length > pPos)
            {
                TFlxFormulaErrorValue Err;
                if (!TUserDefinedFunction.TryGetDouble(Xls, parameters[pPos], out dBasis, out Err)) return Err;
                dBasis = Math.Floor(dBasis);
            }
            if (dBasis < 0 || dBasis > 4) return TFlxFormulaErrorValue.ErrNum;

            Basis = (TDayCountBasis)(int)dBasis;
            return null;
        }
    }
    #endregion

    #region EDate
    /// <summary>
    /// Implements the EDate addin function.
    /// Returns the serial number that represents the date that is the indicated number of months before or after a specified date (the start_date). 
    /// Use EDATE to calculate maturity dates or due dates that fall on the same day of the month as the date of issue.
    /// </summary>
    public class EDateImpl : TUserDefinedFunction
    {
        /// <summary>
        /// Creates a new EDate implementation.
        /// </summary>
        public EDateImpl()
            : base("EDATE")
        {
        }

        /// <summary>
        /// Evaluates the EDate function.
        /// </summary>
        /// <param name="arguments"></param>
        /// <param name="parameters"></param>
        /// <returns></returns>
        public override object Evaluate(TUdfEventArgs arguments, object[] parameters)
        {
            #region Get Parameters
            TFlxFormulaErrorValue Err;
            if (!CheckParameters(parameters, 2, out Err)) return Err;

            //The first parameter is starting date.
            DateTime StartDate;
            if (!TryGetDate(arguments.Xls, parameters[0], true, out StartDate, out Err)) return Err;

            //The second parameter is the number of months.
            double dMonths;
            if (!TryGetDouble(arguments.Xls, parameters[1], out dMonths, out Err)) return Err;
            #endregion

            DateTime EndDate = StartDate.AddMonths(Convert.ToInt32(dMonths));

            double dEndDate;
            if (!FlxDateTime.TryToOADate(EndDate, arguments.Xls.OptionsDates1904, out dEndDate)) return TFlxFormulaErrorValue.ErrValue;
            return dEndDate; // won't work for dates between 1.1.1900 and 29.1.1900
        }
    }

    #endregion

    #region DayCount

    /// <summary>
    /// Enumerates different Basis parameters for financial functions like Duration.
    /// </summary>
    public enum TDayCountBasis
    {
        /// <summary>
        /// US method (NASD), 12 months of 30 days each
        /// </summary>
        UsPsa30_360 = 0,

        /// <summary>
        ///  Actual number of days in months, actual number of days in year
        /// </summary>
        ActualActual = 1,

        /// <summary>
        /// Actual number of days in month, year has 360 days
        /// </summary>
        Actual360 = 2,

        /// <summary>
        /// Actual number of days in month, year has 365 days
        /// </summary>
        Actual365 = 3,

        /// <summary>
        /// European method, 12 months of 30 days each
        /// </summary>
        Europ30_360 = 4
    }
    #endregion

    #region Bonds

    /// <summary>
    /// A date that can have feb 30, for 360 day years.
    /// </summary>
    internal struct UncheckedDate
    {
        internal int Year;
        internal int Month;
        internal int Day;

        internal UncheckedDate(int aYear, int aMonth, int aDay)
        {
            Year = aYear;
            Month = aMonth;
            Day = aDay;
        }

        internal UncheckedDate(DateTime dt)
        {
            Year = dt.Year;
            Month = dt.Month;
            Day = dt.Day;
        }
        public static implicit operator DateTime(UncheckedDate dt)
        {
            return new DateTime(dt.Year, dt.Month, dt.Day);
        }

        public static implicit operator UncheckedDate(DateTime dt)
        {
            return new UncheckedDate(dt.Year, dt.Month, dt.Day);
        }

        public double Diff360(UncheckedDate StartDate)
        {
            return (Year - StartDate.Year) * 360 + (Month - StartDate.Month) * 30 + (Day - StartDate.Day);
        }


    }

    /// <summary>
    /// Implements the basics for Bond functions, like CoupDays, CoupDaysNC, CoupDayBS, CoupNCD, CoupNum, CoupPCD.
    /// </summary>
    public abstract class BaseBondsImpl : TUserDefinedFunction
    {
        /// <summary>
        /// Creates a new implementation.
        /// </summary>
        protected BaseBondsImpl(string aName)
            : base(aName)
        {
        }

        /// <summary>
        /// Evaluates the function. Look At Excel docs for parameters.
        /// </summary>
        /// <param name="arguments"></param>
        /// <param name="parameters"></param>
        /// <returns></returns>
        public override object Evaluate(TUdfEventArgs arguments, object[] parameters)
        {
            #region Get Parameters
            TFlxFormulaErrorValue Err;
            if (!CheckParameters(parameters, -1, out Err)) return Err;
            if (parameters.Length < 3) return TFlxFormulaErrorValue.ErrNA;


            //Settlement date.
            DateTime SettlementDate;
            if (!TryGetDate(arguments.Xls, parameters[0], true, out SettlementDate, out Err)) return Err;

            //Maturity date.
            DateTime MaturityDate;
            if (!TryGetDate(arguments.Xls, parameters[1], true, out MaturityDate, out Err)) return Err;

            if (SettlementDate >= MaturityDate) return TFlxFormulaErrorValue.ErrNum;

            //Frequency.
            double Frequency;
            if (!TryGetDouble(arguments.Xls, parameters[2], out Frequency, out Err)) return Err;
            Frequency = Math.Floor(Frequency);
            if (Frequency != 1 && Frequency != 2 && Frequency != 4) return TFlxFormulaErrorValue.ErrNum;

            //Basis.
            TDayCountBasis Basis;
            object res = FinancialUtils.TryGetBasis(arguments.Xls, parameters, 3, out Basis); if (res != null) return res;

            #endregion


            return Calc(SettlementDate, MaturityDate, Frequency, Basis);

        }

        /// <summary>
        /// Calculates the result depending on the specific function.
        /// </summary>
        /// <param name="SettlementDate"></param>
        /// <param name="MaturityDate"></param>
        /// <param name="Frequency"></param>
        /// <param name="Basis"></param>
        /// <returns></returns>
        protected abstract object Calc(DateTime SettlementDate, DateTime MaturityDate, double Frequency, TDayCountBasis Basis);

        #region Utilities
        private static double DiffDays360Us(DateTime EndDate, DateTime StartDate, bool ChangeFeb)
        {

            //from http://en.wikipedia.org/wiki/Day_count_convention#30.2F360_methods
            // 30/360 US
            // Date adjustment rules (more than one may take effect; apply them in order, and if a date is changed in one rule the changed value is used in the following rules):
            // If the investment is EOM and (Date1 is the last day of February) and (Date2 is the last day of February), then change D2 to 30.
            // If the investment is EOM and (Date1 is the last day of February), then change D1 to 30.
            // If D2 is 31 and D1 is 30 or 31, then change D2 to 30.
            // If D1 is 31, then change D1 to 30.

            UncheckedDate d1 = StartDate;
            UncheckedDate d2 = EndDate;


            //Real standard algorithm
            /*
            if (AtFebEnd(d1) && AtFebEnd(d2)) d2.Day = 30;
            if (AtFebEnd(d1)) d1.Day = 30;
            if (d2.Day == 31 && (d1.Day >= 30)) d2.Day = 30;
            if (d1.Day == 31) d1.Day = 30;
            */

            //See http://www.dwheeler.com/yearfrac/
            //Excel Implementation
            if (d2.Day == 31 && (d1.Day >= 30)) d2.Day = 30;
            if (d1.Day == 31) d1.Day = 30;
            if (AtFebEnd(d1) && AtFebEnd(d2)) d2.Day = 30;
            if (AtFebEnd(d1) && ChangeFeb) d1.Day = 30;


            return d2.Diff360(d1);
        }

        private static double DiffDays360Eu(DateTime EndDate, DateTime StartDate)
        {
            UncheckedDate sd = StartDate;
            if (StartDate.Day == 31) sd.Day = 30;
            UncheckedDate ed = EndDate;
            if (EndDate.Day == 31) ed.Day = 30;

            return ed.Diff360(sd);
        }

        internal static bool AtFebEnd(UncheckedDate EndDate)
        {
            return EndDate.Month == 2 && (EndDate.Day == 29 || (EndDate.Day == 28 && !DateTime.IsLeapYear(EndDate.Year)));
        }

        internal static double DiffDate(DateTime EndDate, DateTime StartDate, TDayCountBasis Basis)
        {
            return DiffDate(EndDate, StartDate, Basis, true);
        }

        internal static double DiffDate(DateTime EndDate, DateTime StartDate, TDayCountBasis Basis, bool ChangeFeb)
        {
            switch (Basis)
            {
                case TDayCountBasis.UsPsa30_360:
                    return DiffDays360Us(EndDate, StartDate, ChangeFeb);

                case TDayCountBasis.ActualActual:
                case TDayCountBasis.Actual360:
                case TDayCountBasis.Actual365:
                    return (EndDate - StartDate).Days;

                case TDayCountBasis.Europ30_360:
                    return DiffDays360Eu(EndDate, StartDate);

            }

            FlxMessages.ThrowException(FlxErr.ErrInternal);
            return 0;
        }

        /// <summary>
        /// Returns the previous coupon for a given coupon.
        /// </summary>
        /// <param name="endDate">Date of the coupon we want to find.</param>
        /// <param name="MaturityDate">Maturity date.</param>
        /// <param name="Frequency">Frequency</param>
        /// <returns></returns>
        protected static DateTime PrevCoupon(DateTime endDate, DateTime MaturityDate, double Frequency)
        {
            DateTime dt = endDate.AddMonths(-12 / (int)Frequency);
            if (IsEOM(MaturityDate)) return LastDayOfMonth(dt);
            if (MaturityDate.Day > dt.Day) return LevelDayOfMonth(dt, MaturityDate);
            return dt;
        }

        /// <summary>
        /// Returns the next coupon after settlement date.
        /// </summary>
        /// <param name="SettlementDate"></param>
        /// <param name="MaturityDate"></param>
        /// <param name="Frequency"></param>
        /// <returns></returns>
        protected static DateTime NextCoupon(DateTime SettlementDate, DateTime MaturityDate, double Frequency)
        {
            DateTime BondDate0 = MaturityDate;
            DateTime BondDate = BondDate0;
            int i = (int)((MaturityDate.Year - SettlementDate.Year) * Frequency);
            while (BondDate > SettlementDate)
            {
                i++;
                BondDate = BondDate0.AddMonths(i * -12 / (int)Frequency); //we can't use Bondate=Bondate.Addmonts(- ) because when you arrive to feb, it will stay at 28 days from then on.
                if (IsEOM(MaturityDate)) BondDate = LastDayOfMonth(BondDate);
            }

            while (BondDate <= SettlementDate)
            {
                i--;
                BondDate = BondDate0.AddMonths(i * -12 / (int)Frequency);
                if (IsEOM(MaturityDate)) BondDate = LastDayOfMonth(BondDate);
            }

            return BondDate;
        }

        /// <summary>
        /// Makes the day in bonddate be as big Maturity date as possible.
        /// </summary>
        /// <param name="BondDate"></param>
        /// <param name="MaturityDate"></param>
        /// <returns></returns>
        protected static DateTime LevelDayOfMonth(DateTime BondDate, DateTime MaturityDate)
        {
            return new DateTime(BondDate.Year, BondDate.Month, Math.Min(MaturityDate.Day, DateTime.DaysInMonth(BondDate.Year, BondDate.Month)));
        }

        /// <summary>
        /// Returns the last day of a month.
        /// </summary>
        /// <param name="BondDate"></param>
        /// <returns></returns>
        protected internal static DateTime LastDayOfMonth(DateTime BondDate)
        {
            return new DateTime(BondDate.Year, BondDate.Month, DateTime.DaysInMonth(BondDate.Year, BondDate.Month));

        }

        /// <summary>
        /// Returns true if the date is the end of month.
        /// </summary>
        /// <param name="SettlementDate"></param>
        /// <returns></returns>
        protected static bool IsEOM(DateTime SettlementDate)
        {
            return SettlementDate.AddDays(1).Month != SettlementDate.Month;
        }

        #endregion
    }

    /// <summary>
    /// Implements the CoupPCD Excel function. Look at Excel documentation for more information.
    /// </summary>
    public class CoupPCDImpl : BaseBondsImpl
    {
        /// <summary>
        /// Creates a new implementation.
        /// </summary>
        public CoupPCDImpl()
            : base("COUPPCD")
        {
        }

        /// <summary>
        /// Override this method to specialize BaseBondsImpl
        /// </summary>
        /// <param name="SettlementDate"></param>
        /// <param name="MaturityDate"></param>
        /// <param name="Frequency"></param>
        /// <param name="Basis"></param>
        /// <returns></returns>
        protected override object Calc(DateTime SettlementDate, DateTime MaturityDate, double Frequency, TDayCountBasis Basis)
        {
            return CalcCoupPCD(SettlementDate, MaturityDate, Frequency, Basis);
        }

        /// <summary>
        /// See Excel docs for description of this function.
        /// </summary>
        /// <param name="SettlementDate">See Excel.</param>
        /// <param name="MaturityDate">See Excel.</param>
        /// <param name="Frequency">See Excel.</param>
        /// <param name="Basis">See Excel.</param>
        /// <returns></returns>
        public static object CalcCoupPCD(DateTime SettlementDate, DateTime MaturityDate, double Frequency, TDayCountBasis Basis)
        {
            DateTime LastCoupon = PrevCoupon(NextCoupon(SettlementDate, MaturityDate, Frequency), MaturityDate, Frequency);
            return LastCoupon;
        }
    }

    /// <summary>
    /// Implements the CoupDaysNC Excel function. Look at Excel documentation for more information.
    /// </summary>
    public class CoupDaysNCImpl : BaseBondsImpl
    {
        /// <summary>
        /// Creates a new implementation.
        /// </summary>
        public CoupDaysNCImpl()
            : base("COUPDAYSNC")
        {
        }

        /// <summary>
        /// Override this method to specialize BaseBondsImpl
        /// </summary>
        /// <param name="SettlementDate"></param>
        /// <param name="MaturityDate"></param>
        /// <param name="Frequency"></param>
        /// <param name="Basis"></param>
        /// <returns></returns>
        protected override object Calc(DateTime SettlementDate, DateTime MaturityDate, double Frequency, TDayCountBasis Basis)
        {
            return CalcCoupDaysNC(SettlementDate, MaturityDate, Frequency, Basis);
        }

        /// <summary>
        /// See Excel docs for description of this function.
        /// </summary>
        /// <param name="SettlementDate">See Excel.</param>
        /// <param name="MaturityDate">See Excel.</param>
        /// <param name="Frequency">See Excel.</param>
        /// <param name="Basis">See Excel.</param>
        /// <returns></returns>
        public static object CalcCoupDaysNC(DateTime SettlementDate, DateTime MaturityDate, double Frequency, TDayCountBasis Basis)
        {
            //This methods behaves very weird in Excel. Id doesn't work 100% exactly as Excel, but it does give the same resultas as openoffice or other financial packages.
            DateTime NxCoupon = NextCoupon(SettlementDate, MaturityDate, Frequency);
            //see http://sc.openoffice.org/source/browse/sc/scaddins/source/analysis/analysishelper.cxx?view=markup
            if (Basis == TDayCountBasis.UsPsa30_360)
            {
                DateTime PrCoupon = PrevCoupon(NxCoupon, MaturityDate, Frequency);
                double TotalDays = CoupDaysImpl.CalcCoupDays(SettlementDate, MaturityDate, Frequency, Basis);
                return TotalDays - DiffDate(SettlementDate, PrCoupon, Basis, true); //sometimes this is false, but when is not clear.
            }
            return DiffDate(NxCoupon, SettlementDate, Basis);
        }

    }

    /// <summary>
    /// Implements the CoupNCD Excel function. Look at Excel documentation for more information.
    /// </summary>
    public class CoupNCDImpl : BaseBondsImpl
    {
        /// <summary>
        /// Creates a new implementation.
        /// </summary>
        public CoupNCDImpl()
            : base("COUPNCD")
        {
        }

        /// <summary>
        /// Override this method to specialize BaseBondsImpl
        /// </summary>
        /// <param name="SettlementDate"></param>
        /// <param name="MaturityDate"></param>
        /// <param name="Frequency"></param>
        /// <param name="Basis"></param>
        /// <returns></returns>
        protected override object Calc(DateTime SettlementDate, DateTime MaturityDate, double Frequency, TDayCountBasis Basis)
        {
            return CalcCoupNCD(SettlementDate, MaturityDate, Frequency, Basis);
        }

        /// <summary>
        /// See Excel docs for description of this function.
        /// </summary>
        /// <param name="SettlementDate">See Excel.</param>
        /// <param name="MaturityDate">See Excel.</param>
        /// <param name="Frequency">See Excel.</param>
        /// <param name="Basis">See Excel.</param>
        /// <returns></returns>
        public static object CalcCoupNCD(DateTime SettlementDate, DateTime MaturityDate, double Frequency, TDayCountBasis Basis)
        {
            DateTime NxCoupon = NextCoupon(SettlementDate, MaturityDate, Frequency);
            return NxCoupon;
        }
    }

    /// <summary>
    /// Implements the CoupDaysBS Excel function. Look at Excel documentation for more information.
    /// </summary>
    public class CoupDaysBSImpl : BaseBondsImpl
    {
        /// <summary>
        /// Creates a new implementation.
        /// </summary>
        public CoupDaysBSImpl()
            : base("COUPDAYBS")
        {
        }

        /// <summary>
        /// Override this method to specialize BaseBondsImpl
        /// </summary>
        /// <param name="SettlementDate"></param>
        /// <param name="MaturityDate"></param>
        /// <param name="Frequency"></param>
        /// <param name="Basis"></param>
        /// <returns></returns>
        protected override object Calc(DateTime SettlementDate, DateTime MaturityDate, double Frequency, TDayCountBasis Basis)
        {
            return CalcCoupDaysBS(SettlementDate, MaturityDate, Frequency, Basis);
        }

        /// <summary>
        /// See Excel docs for description of this function.
        /// </summary>
        /// <param name="SettlementDate">See Excel.</param>
        /// <param name="MaturityDate">See Excel.</param>
        /// <param name="Frequency">See Excel.</param>
        /// <param name="Basis">See Excel.</param>
        /// <returns></returns>
        public static double CalcCoupDaysBS(DateTime SettlementDate, DateTime MaturityDate, double Frequency, TDayCountBasis Basis)
        {
            DateTime LastCoupon = PrevCoupon(NextCoupon(SettlementDate, MaturityDate, Frequency), MaturityDate, Frequency);
            return DiffDate(SettlementDate, LastCoupon, Basis);
        }
    }

    /// <summary>
    /// Implements the CoupDays Excel function. Look at Excel documentation for more information.
    /// </summary>
    public class CoupDaysImpl : BaseBondsImpl
    {
        /// <summary>
        /// Creates a new implementation.
        /// </summary>
        public CoupDaysImpl()
            : base("COUPDAYS")
        {
        }

        /// <summary>
        /// Override this method to specialize BaseBondsImpl
        /// </summary>
        /// <param name="SettlementDate"></param>
        /// <param name="MaturityDate"></param>
        /// <param name="Frequency"></param>
        /// <param name="Basis"></param>
        /// <returns></returns>
        protected override object Calc(DateTime SettlementDate, DateTime MaturityDate, double Frequency, TDayCountBasis Basis)
        {
            return CalcCoupDays(SettlementDate, MaturityDate, Frequency, Basis);
        }

        /// <summary>
        /// See Excel docs for description of this function.
        /// </summary>
        /// <param name="SettlementDate">See Excel.</param>
        /// <param name="MaturityDate">See Excel.</param>
        /// <param name="Frequency">See Excel.</param>
        /// <param name="Basis">See Excel.</param>
        /// <returns></returns>
        public static double CalcCoupDays(DateTime SettlementDate, DateTime MaturityDate, double Frequency, TDayCountBasis Basis)
        {
            switch (Basis)
            {
                case TDayCountBasis.UsPsa30_360:
                    return 360 / Frequency;

                case TDayCountBasis.ActualActual:
                    //This is the days between date before and after settlement date. We use maturity to know when that is.

                    DateTime endDate = NextCoupon(SettlementDate, MaturityDate, Frequency);
                    return (endDate - PrevCoupon(endDate, MaturityDate, Frequency)).Days;

                case TDayCountBasis.Actual360:
                    return 360 / Frequency;

                case TDayCountBasis.Actual365:
                    return 365 / Frequency;

                case TDayCountBasis.Europ30_360:
                    return 360 / Frequency;
            }

            FlxMessages.ThrowException(FlxErr.ErrInternal);
            return 0;
        }
    }

    /// <summary>
    /// Implements the CoupNum Excel function. Look at Excel documentation for more information.
    /// </summary>
    public class CoupNumImpl : BaseBondsImpl
    {
        /// <summary>
        /// Creates a new implementation.
        /// </summary>
        public CoupNumImpl()
            : base("COUPNUM")
        {
        }

        /// <summary>
        /// Override this method to specialize BaseBondsImpl
        /// </summary>
        /// <param name="SettlementDate"></param>
        /// <param name="MaturityDate"></param>
        /// <param name="Frequency"></param>
        /// <param name="Basis"></param>
        /// <returns></returns>
        protected override object Calc(DateTime SettlementDate, DateTime MaturityDate, double Frequency, TDayCountBasis Basis)
        {
            return CalcCoupNum(SettlementDate, MaturityDate, Frequency, Basis);
        }

        /// <summary>
        /// See Excel docs for description of this function.
        /// </summary>
        /// <param name="SettlementDate">See Excel.</param>
        /// <param name="MaturityDate">See Excel.</param>
        /// <param name="Frequency">See Excel.</param>
        /// <param name="Basis">See Excel.</param>
        /// <returns></returns>
        public static double CalcCoupNum(DateTime SettlementDate, DateTime MaturityDate, double Frequency, TDayCountBasis Basis)
        {
            DateTime BondDate = NextCoupon(SettlementDate, MaturityDate, Frequency);
            int Months = (MaturityDate.Year - BondDate.Year) * 12 + (MaturityDate.Month - BondDate.Month);
            return 1 + Months / 12.0 * Frequency;
        }
    }
    #endregion

    #region YearFrac
    /// <summary>
    /// Implements the YearFrac Excel function. Look at Excel documentation for more information.
    /// </summary>
    public class YearFracImpl : TUserDefinedFunction
    {
        /// <summary>
        /// Creates a new implementation.
        /// </summary>
        public YearFracImpl()
            : base("YEARFRAC")
        {
        }

        /// <summary>
        /// Evaluates the function. Look At Excel docs for parameters.
        /// </summary>
        /// <param name="arguments"></param>
        /// <param name="parameters"></param>
        /// <returns></returns>
        public override object Evaluate(TUdfEventArgs arguments, object[] parameters)
        {
            #region Get Parameters
            TFlxFormulaErrorValue Err;
            if (!CheckParameters(parameters, -1, out Err)) return Err;
            if (parameters.Length < 2) return TFlxFormulaErrorValue.ErrNA;


            //date1.
            DateTime Date1;
            if (!TryGetDate(arguments.Xls, parameters[0], true, out Date1, out Err)) return Err;

            //date2.
            DateTime Date2;
            if (!TryGetDate(arguments.Xls, parameters[1], true, out Date2, out Err)) return Err;

            if (Date1 == Date2) return 0;
            if (Date1 > Date2)
            {
                DateTime tmp = Date1;
                Date1 = Date2;
                Date2 = tmp;
            };

            //Basis.
            TDayCountBasis Basis;
            object res = FinancialUtils.TryGetBasis(arguments.Xls, parameters, 2, out Basis); if (res != null) return res;

            #endregion

            return CalcYearFrac(Date1, Date2, Basis);
        }

        /// <summary>
        /// See Excel docs for description of this function.
        /// </summary>
        /// <param name="Date1"></param>
        /// <param name="Date2"></param>
        /// <param name="Basis"></param>
        public static double CalcYearFrac(DateTime Date1, DateTime Date2, TDayCountBasis Basis)
        {
            double days = BaseBondsImpl.DiffDate(Date2, Date1, Basis);
            switch (Basis)
            {
                case TDayCountBasis.UsPsa30_360:
                    return days / 360.0;

                case TDayCountBasis.ActualActual:
                    return days / ActualYearLen(Date1, Date2);

                case TDayCountBasis.Actual360:
                    return days / 360.0;

                case TDayCountBasis.Actual365:
                    return days / 365.0;

                case TDayCountBasis.Europ30_360:
                    return days / 360.0;

            }

            FlxMessages.ThrowException(FlxErr.ErrInternal);
            return 0;
        }

        private static double ActualYearLen(DateTime Date1, DateTime Date2)
        {
            if (AppearsLess1Year(Date1, Date2))
            {
                if (Date1.Year == Date2.Year && DateTime.IsLeapYear(Date1.Year))
                    return 366;


                if (Feb29Between(Date1, Date2) || (Date2.Month == 2 && Date2.Day == 29)) return 366;

                return 365;
            }

            double NumYears = (Date2.Year - Date1.Year) + 1;
            double DaysInYears = (new DateTime(Date2.Year + 1, 1, 1) - new DateTime(Date1.Year, 1, 1)).Days;
            return DaysInYears / NumYears;
        }

        private static bool Feb29Between(DateTime Date1, DateTime Date2)
        {
            for (int y = Date1.Year + 1; y < Date2.Year; y++)
            {
                if (DateTime.IsLeapYear(y)) return true;
            }

            if (DateTime.IsLeapYear(Date1.Year) && (Date1.Month < 2 || (Date1.Month == 2 && Date1.Day == 29))) return true;
            if (DateTime.IsLeapYear(Date2.Year) && (Date2.Month > 2 || (Date2.Month == 2 && Date2.Day == 29))) return true;
            return false;
        }

        private static bool AppearsLess1Year(DateTime Date1, DateTime Date2)
        {
            //Returns True if date1 and date2 "appear" to be 1 year or less apart.
            // This compares the values of year, month, and day directly to each other.
            // Requires date1 <= date2; returns boolean.  Used by basis 1.
            if (Date1.Year == Date2.Year) return true;
            if (((Date1.Year + 1) == Date2.Year) &&
                   ((Date1.Month > Date2.Month) ||
                   ((Date1.Month == Date2.Month) && (Date1.Day >= Date2.Day)))) return true;
            return false;
        }
    }
    #endregion

    #region Duration - MDuration
    /// <summary>
    /// Implements the Duration and MDuration addin functions.
    /// Returns the [modified] Macauley duration for a security with an assumed par value of $100.
    /// </summary>
    public class DurationImpl : TUserDefinedFunction
    {
        bool Modified;

        /// <summary>
        /// Creates a new Duration/MDuration implementation.
        /// </summary>
        public DurationImpl(bool aModified)
            : base(aModified ? "MDURATION" : "DURATION")
        {
            Modified = aModified;
        }

        /// <summary>
        /// Evaluates the DURATION/MDURATION function. Look At Excel docs for parameters.
        /// </summary>
        /// <param name="arguments"></param>
        /// <param name="parameters"></param>
        /// <returns></returns>
        public override object Evaluate(TUdfEventArgs arguments, object[] parameters)
        {
            #region Get Parameters
            TFlxFormulaErrorValue Err;
            if (!CheckParameters(parameters, -1, out Err)) return Err;
            if (parameters.Length < 5) return TFlxFormulaErrorValue.ErrNA;


            //Settlement date.
            DateTime SettlementDate;
            if (!TryGetDate(arguments.Xls, parameters[0], true, out SettlementDate, out Err)) return Err;

            //Maturity date.
            DateTime MaturityDate;
            if (!TryGetDate(arguments.Xls, parameters[1], true, out MaturityDate, out Err)) return Err;

            if (SettlementDate >= MaturityDate) return TFlxFormulaErrorValue.ErrNum;

            //Coupon Rate.
            double CouponRate;
            if (!TryGetDouble(arguments.Xls, parameters[2], out CouponRate, out Err)) return Err;
            if (CouponRate < 0) return TFlxFormulaErrorValue.ErrNum;

            //Annual Yield.
            double Yield;
            if (!TryGetDouble(arguments.Xls, parameters[3], out Yield, out Err)) return Err;
            if (Yield < 0) return TFlxFormulaErrorValue.ErrNum;

            //Frequency.
            double Frequency;
            if (!TryGetDouble(arguments.Xls, parameters[4], out Frequency, out Err)) return Err;
            Frequency = Math.Floor(Frequency);
            if (Frequency != 1 && Frequency != 2 && Frequency != 4) return TFlxFormulaErrorValue.ErrNum;

            //Basis.
            TDayCountBasis Basis;
            object res = FinancialUtils.TryGetBasis(arguments.Xls, parameters, 5, out Basis); if (res != null) return res;

            #endregion

            return Calc(Modified, SettlementDate, MaturityDate, CouponRate, Yield, Frequency, Basis);
        }

        /// <summary>
        /// Duration/MDuration implementation. You can call this method on its own.
        /// </summary>
        /// <param name="Modified">If true, we return MDuration, if not, Duration.</param>
        /// <param name="SettlementDate">See Excel.</param>
        /// <param name="MaturityDate">See Excel.</param>
        /// <param name="CouponRate">See Excel.</param>
        /// <param name="Yield">See Excel.</param>
        /// <param name="Frequency">See Excel.</param>
        /// <param name="Basis">See Excel.</param>
        /// <returns></returns>
        public static object Calc(bool Modified, DateTime SettlementDate, DateTime MaturityDate, double CouponRate, double Yield, double Frequency, TDayCountBasis Basis)
        {
            // f = (-(f r_c)-y^2+f r_c ((f+y)/f)^n+y (y+r_c (-1+((f+y)/f)^n)) alpha+y (-r_c+y) n)/(y (y+r_c (-1+((f+y)/f)^n)))  
            // y | annual yield
            // r_c | annual coupon rate
            // f | coupon frequency
            // n | number of whole coupon periods
            // alpha | fraction of year until next coupon

            double dbc = CoupDaysBSImpl.CalcCoupDaysBS(SettlementDate, MaturityDate, Frequency, Basis);
            double e = CoupDaysImpl.CalcCoupDays(SettlementDate, MaturityDate, Frequency, Basis);
            double n = CoupNumImpl.CalcCoupNum(SettlementDate, MaturityDate, Frequency, Basis);
            double dsc = e - dbc;
            double x1 = dsc / e;
            double x2 = x1 + n - 1;
            double x3 = Yield / Frequency + 1;
            double x4 = Math.Pow(x3, x2);
            if (x4 == 0) return TFlxFormulaErrorValue.ErrDiv0;
            double term1 = x2 * 100.0 / x4;
            double term3 = 100.0 / x4;

            double term2 = 0;
            double term4 = 0;
            for (int i = 1; i <= (int)n; i++)
            {
                double x5 = i - 1 + x1;
                double x6 = Math.Pow(x3, x5);
                if (x6 == 0) return TFlxFormulaErrorValue.ErrDiv0;
                double x7 = (100.0 * CouponRate / Frequency) / x6;
                term2 += x7 * x5;
                term4 += x7;
            }

            double term5 = term1 + term2;
            double term6 = term3 + term4;
            if (term6 == 0) return TFlxFormulaErrorValue.ErrDiv0;

            double Result = (term5 / term6) / Frequency;
            if (Modified) return Result / x3; else return Result;
        }

    }

    #endregion

    #region Conversion Bin / Hex / Dec / Oct
    /// <summary>
    /// Base conversion class for implementing Bin2Hex, Dec2Bin, etc.
    /// </summary>
    public abstract class BaseBinHexImpl : TUserDefinedFunction
    {
        bool HasPlaces;
        bool NeedsStrings;

        static Dictionary<char, int>[] CharsInBase = InitCharsInBase();

        private static Dictionary<char, int>[] InitCharsInBase()
        {
            Dictionary<char, int>[] Result = new Dictionary<char, int>[5];
            Result[1] = new Dictionary<char, int>(2);
            Result[1]['0'] = 0;
            Result[1]['1'] = 1;

            Result[3] = new Dictionary<char, int>(8);
            for (int i = 0; i < 8; i++)
            {
                Result[3][(char)('0' + i)] = i;
            }

            Result[4] = new Dictionary<char, int>(16 + 6);
            for (int i = 0; i < 10; i++)
            {
                Result[4][(char)('0' + i)] = i;
            }
            for (int i = 0; i < 6; i++)
            {
                Result[4][(char)('A' + i)] = 10 + i;
            }
            for (int i = 0; i < 6; i++)
            {
                Result[4][(char)('a' + i)] = 10 + i;
            }

            return Result;
        }


        /// <summary>
        /// Creates a new object.
        /// </summary>
        protected BaseBinHexImpl(string aName, bool aHasPlaces, bool aNeedsStrings)
            : base(aName)
        {
            HasPlaces = aHasPlaces;
            NeedsStrings = aNeedsStrings;
        }

        /// <summary>
        /// Evaluates the function.
        /// </summary>
        /// <param name="arguments"></param>
        /// <param name="parameters"></param>
        /// <returns></returns>
        public override object Evaluate(TUdfEventArgs arguments, object[] parameters)
        {
            #region Get Parameters
            TFlxFormulaErrorValue Err;
            if (!CheckParameters(parameters, -1, out Err)) return Err;

            if (parameters.Length < 1 || parameters.Length > 2) return TFlxFormulaErrorValue.ErrValue;
            if (!HasPlaces && parameters.Length == 2) return TFlxFormulaErrorValue.ErrValue;

            //The first parameter is the string with the number.
            string num = null;
            double dnum = 0;
            if (NeedsStrings)
            {
                if (!TryGetString(arguments.Xls, parameters[0], out num, out Err)) return Err;
            }
            else
            {
                if (!TryGetDouble(arguments.Xls, parameters[0], out dnum, out Err)) return Err;
            }

            //The second parameter is places.
            int Places = -1;
            if (parameters.Length > 1)
            {
                double dPlaces;
                if (!TryGetDouble(arguments.Xls, parameters[1], out dPlaces, out Err)) return Err;
                if (dPlaces <= 0 || Math.Floor(dPlaces) > 10) return TFlxFormulaErrorValue.ErrNum;
                Places = (int)dPlaces;
            }
            #endregion

            return Calc(num, dnum, Places);

        }

        /// <summary>
        /// Calculates the result depending on the specific function.
        /// </summary>
        /// <returns></returns>
        protected abstract object Calc(string num, double dnum, int Places);

        #region Utilities
        internal object Pad(string s, int Places)
        {
            if (s == null) return TFlxFormulaErrorValue.ErrNum;

            if (s.Length > 10) return TFlxFormulaErrorValue.ErrNum;

            if (Places >= 0)
            {
                if (s.Length > Places) return TFlxFormulaErrorValue.ErrNum;
                return s.PadLeft(Places, '0');
            }

            return s;
        }

        internal string ConvertBase(string num, int b1, int b2, ref int Places) //doesn't handle negatives or decimals in the input string, not needed. only use to convert to/from b2, b8, b16.
        {
            if (num == null) return "0";
            long Value;
            if (!TryFromString(num, b1, out Value)) return null;

            bool IsNegative;
            string Result = ToString(Value, b1, b2, out IsNegative);
            if (IsNegative) Places = 10;
            return Result;
        }

        internal bool TryFromString(string num, int b1, out long Value)
        {
            Value = 0;
            int shf = 0;
            int shfInc = GetShift(b1);
            if (num.Length > 10) return false;
            for (int i = num.Length - 1; i >= 0; i--)
            {
                int rx;
                if (!CharsInBase[shfInc].TryGetValue(num[i], out rx)) return false;
                Value += ((long)rx) << shf;
                shf += shfInc;
            }

            return true;
        }

        internal string ToString(long Value, int b1, int b2, out bool IsNegative)
        {
            int shfInc = GetShift(b1);
            long NegBit = 0x1L << (shfInc * 10 - 1);
            IsNegative = IsNegative10(Value, b1);
            int shfInc2 = GetShift(b2);
            long NegBit2 = 0x1L << (shfInc2 * 10);
            if (IsNegative) //complement 2 in 10*shfinc2 bits.
            {
                Value = NegBit2 - (NegBit - (Value & ~NegBit));
                if (Value < (NegBit2 >> 1)) return null;
            }
            else
            {
                if (Value >= (NegBit2 >> 1)) return null;
            }

#if (COMPACTFRAMEWORK)
            return Convert.ToString(Value, b2).ToUpper(CultureInfo.InvariantCulture);
#else
            return Convert.ToString(Value, b2).ToUpperInvariant();
#endif
        }

        internal static bool IsNegative10(long Value, int b1)
        {
            int shfInc = GetShift(b1);
            long NegBit = 0x1L << (shfInc * 10 - 1);
            return (Value & NegBit) != 0;
        }

        internal static long ReverseDec10(long Value, int b1)
        {
            int shfInc = GetShift(b1);
            long NegBit = 0x1L << (shfInc * 10 - 1);
            return -(NegBit - (Value & ~NegBit));
        }

        internal object ConvertFromDec(double dnum, int b1, ref int Places)
        {
            int shfInc = GetShift(b1);
            long NegBit = 0x1L << (shfInc * 10 - 1);
            long MaxVal = NegBit << 1;
            if (dnum >= NegBit || dnum < -NegBit) return TFlxFormulaErrorValue.ErrNum;
            long LNum = (long)dnum;
            if (LNum < 0)
            {
                Places = 10;
                LNum = MaxVal + LNum;
            }

#if (COMPACTFRAMEWORK)
            return Pad(Convert.ToString(LNum, b1).ToUpper(CultureInfo.InvariantCulture), Places);
#else
            return Pad(Convert.ToString(LNum, b1).ToUpperInvariant(), Places);
#endif
        }


        private static int GetShift(int b1)
        {
            switch (b1)
            {
                case 8: return 3;
                case 16: return 4;
            }
            return 1;
        }
        #endregion
    }

    #region Bin2
    /// <summary>
    /// Implements the Bin2Hex Excel function. Look at Excel documentation for more information.
    /// </summary>
    public class Bin2HexImpl : BaseBinHexImpl
    {
        /// <summary>
        /// Creates a new implementation.
        /// </summary>
        public Bin2HexImpl()
            : base("BIN2HEX", true, true)
        {
        }

        /// <summary>
        /// Implement this method to specialize BaseBinHexImpl class.
        /// </summary>
        /// <returns></returns>
        protected override object Calc(string num, double dnum, int Places)
        {
            return Pad(ConvertBase(num, 2, 16, ref Places), Places);
        }
    }

    /// <summary>
    /// Implements the Bin2Oct Excel function. Look at Excel documentation for more information.
    /// </summary>
    public class Bin2OctImpl : BaseBinHexImpl
    {
        /// <summary>
        /// Creates a new implementation.
        /// </summary>
        public Bin2OctImpl()
            : base("BIN2OCT", true, true)
        {
        }

        /// <summary>
        /// Implement this method to specialize BaseBinHexImpl class.
        /// </summary>
        /// <returns></returns>
        protected override object Calc(string num, double dnum, int Places)
        {
            return Pad(ConvertBase(num, 2, 8, ref Places), Places);
        }
    }

    /// <summary>
    /// Implements the Bin2Dec Excel function. Look at Excel documentation for more information.
    /// </summary>
    public class Bin2DecImpl : BaseBinHexImpl
    {
        /// <summary>
        /// Creates a new implementation.
        /// </summary>
        public Bin2DecImpl()
            : base("BIN2DEC", false, true)
        {
        }

        /// <summary>
        /// Implement this method to specialize BaseBinHexImpl class.
        /// </summary>
        /// <returns></returns>
        protected override object Calc(string num, double dnum, int Places)
        {
            long Value;
            if (!TryFromString(num, 2, out Value)) return TFlxFormulaErrorValue.ErrNum;
            if (IsNegative10(Value, 2))
            {
                Value = ReverseDec10(Value, 2);
            }
            return (double)Value;
        }
    }
    #endregion

    #region Dec2
    /// <summary>
    /// Implements the Dec2Hex Excel function. Look at Excel documentation for more information.
    /// </summary>
    public class Dec2HexImpl : BaseBinHexImpl
    {
        /// <summary>
        /// Creates a new implementation.
        /// </summary>
        public Dec2HexImpl()
            : base("DEC2HEX", true, false)
        {
        }

        /// <summary>
        /// Implement this method to specialize BaseBinHexImpl class.
        /// </summary>
        /// <returns></returns>
        protected override object Calc(string num, double dnum, int Places)
        {
            return ConvertFromDec(dnum, 16, ref Places);
        }
    }

    /// <summary>
    /// Implements the Dec2Oct Excel function. Look at Excel documentation for more information.
    /// </summary>
    public class Dec2OctImpl : BaseBinHexImpl
    {
        /// <summary>
        /// Creates a new implementation.
        /// </summary>
        public Dec2OctImpl()
            : base("DEC2OCT", true, false)
        {
        }

        /// <summary>
        /// Implement this method to specialize BaseBinHexImpl class.
        /// </summary>
        /// <returns></returns>
        protected override object Calc(string num, double dnum, int Places)
        {
            return ConvertFromDec(dnum, 8, ref Places);
        }
    }

    /// <summary>
    /// Implements the Dec2Bin Excel function. Look at Excel documentation for more information.
    /// </summary>
    public class Dec2BinImpl : BaseBinHexImpl
    {
        /// <summary>
        /// Creates a new implementation.
        /// </summary>
        public Dec2BinImpl()
            : base("DEC2BIN", true, false)
        {
        }

        /// <summary>
        /// Implement this method to specialize BaseBinHexImpl class.
        /// </summary>
        /// <returns></returns>
        protected override object Calc(string num, double dnum, int Places)
        {
            return ConvertFromDec(dnum, 2, ref Places);
        }
    }
    #endregion

    #region Oct2
    /// <summary>
    /// Implements the Oct2Hex Excel function. Look at Excel documentation for more information.
    /// </summary>
    public class Oct2HexImpl : BaseBinHexImpl
    {
        /// <summary>
        /// Creates a new implementation.
        /// </summary>
        public Oct2HexImpl()
            : base("OCT2HEX", true, true)
        {
        }

        /// <summary>
        /// Implement this method to specialize BaseBinHexImpl class.
        /// </summary>
        /// <returns></returns>
        protected override object Calc(string num, double dnum, int Places)
        {
            return Pad(ConvertBase(num, 8, 16, ref Places), Places);
        }
    }

    /// <summary>
    /// Implements the Oct2Dec Excel function. Look at Excel documentation for more information.
    /// </summary>
    public class Oct2DecImpl : BaseBinHexImpl
    {
        /// <summary>
        /// Creates a new implementation.
        /// </summary>
        public Oct2DecImpl()
            : base("OCT2DEC", false, true)
        {
        }

        /// <summary>
        /// Implement this method to specialize BaseBinHexImpl class.
        /// </summary>
        /// <returns></returns>
        protected override object Calc(string num, double dnum, int Places)
        {
            long Value;
            if (!TryFromString(num, 8, out Value)) return TFlxFormulaErrorValue.ErrNum;
            if (IsNegative10(Value, 8))
            {
                Value = ReverseDec10(Value, 8);
            }
            return (double)Value;
        }
    }

    /// <summary>
    /// Implements the Oct2Bin Excel function. Look at Excel documentation for more information.
    /// </summary>
    public class Oct2BinImpl : BaseBinHexImpl
    {
        /// <summary>
        /// Creates a new implementation.
        /// </summary>
        public Oct2BinImpl()
            : base("OCT2BIN", true, true)
        {
        }

        /// <summary>
        /// Implement this method to specialize BaseBinHexImpl class.
        /// </summary>
        /// <returns></returns>
        protected override object Calc(string num, double dnum, int Places)
        {
            return Pad(ConvertBase(num, 8, 2, ref Places), Places);
        }
    }
    #endregion

    #region Hex2
    /// <summary>
    /// Implements the Hex2Oct Excel function. Look at Excel documentation for more information.
    /// </summary>
    public class Hex2OctImpl : BaseBinHexImpl
    {
        /// <summary>
        /// Creates a new implementation.
        /// </summary>
        public Hex2OctImpl()
            : base("HEX2OCT", true, true)
        {
        }

        /// <summary>
        /// Implement this method to specialize BaseBinHexImpl class.
        /// </summary>
        /// <returns></returns>
        protected override object Calc(string num, double dnum, int Places)
        {
            return Pad(ConvertBase(num, 16, 8, ref Places), Places);
        }
    }

    /// <summary>
    /// Implements the Hex2Dec Excel function. Look at Excel documentation for more information.
    /// </summary>
    public class Hex2DecImpl : BaseBinHexImpl
    {
        /// <summary>
        /// Creates a new implementation.
        /// </summary>
        public Hex2DecImpl()
            : base("HEX2DEC", false, true)
        {
        }

        /// <summary>
        /// Implement this method to specialize BaseBinHexImpl class.
        /// </summary>
        /// <returns></returns>
        protected override object Calc(string num, double dnum, int Places)
        {
            long Value;
            if (!TryFromString(num, 16, out Value)) return TFlxFormulaErrorValue.ErrNum;
            if (IsNegative10(Value, 16))
            {
                Value = ReverseDec10(Value, 16);
            }
            return (double)Value;
        }
    }

    /// <summary>
    /// Implements the Hex2Bin Excel function. Look at Excel documentation for more information.
    /// </summary>
    public class Hex2BinImpl : BaseBinHexImpl
    {
        /// <summary>
        /// Creates a new implementation.
        /// </summary>
        public Hex2BinImpl()
            : base("HEX2BIN", true, true)
        {
        }

        /// <summary>
        /// Implement this method to specialize BaseBinHexImpl class.
        /// </summary>
        /// <returns></returns>
        protected override object Calc(string num, double dnum, int Places)
        {
            return Pad(ConvertBase(num, 16, 2, ref Places), Places);
        }
    }
    #endregion

    #endregion

    #region IsOdd / Even
    /// <summary>
    /// Implements the IsOdd Excel function. Look at Excel documentation for more information.
    /// </summary>
    public class IsOddImpl : TUserDefinedFunction
    {
        /// <summary>
        /// Creates a new implementation.
        /// </summary>
        public IsOddImpl()
            : base("ISODD")
        {
        }

        /// <summary>
        /// Evaluates the function. Look At Excel docs for parameters.
        /// </summary>
        /// <param name="arguments"></param>
        /// <param name="parameters"></param>
        /// <returns></returns>
        public override object Evaluate(TUdfEventArgs arguments, object[] parameters)
        {
            TFlxFormulaErrorValue Err;
            if (!CheckParameters(parameters, 1, out Err)) return Err;

            //value.
            double value;
            if (!TryGetDouble(arguments.Xls, parameters[0], out value, out Err)) return Err;
#if (COMPACTFRAMEWORK)
            value = value > 0 ? Math.Floor(value) : Math.Ceiling(value);
#else
            value = Math.Truncate(value);
#endif
            return value % 2 != 0;

        }
    }

    /// <summary>
    /// Implements the IsEven Excel function. Look at Excel documentation for more information.
    /// </summary>
    public class IsEvenImpl : TUserDefinedFunction
    {
        /// <summary>
        /// Creates a new implementation.
        /// </summary>
        public IsEvenImpl()
            : base("ISEVEN")
        {
        }

        /// <summary>
        /// Evaluates the function. Look At Excel docs for parameters.
        /// </summary>
        /// <param name="arguments"></param>
        /// <param name="parameters"></param>
        /// <returns></returns>
        public override object Evaluate(TUdfEventArgs arguments, object[] parameters)
        {
            TFlxFormulaErrorValue Err;
            if (!CheckParameters(parameters, 1, out Err)) return Err;

            //value.
            double value;
            if (!TryGetDouble(arguments.Xls, parameters[0], out value, out Err)) return Err;
#if (COMPACTFRAMEWORK)
            value = value > 0 ? Math.Floor(value) : Math.Ceiling(value);
#else
            value = Math.Truncate(value);
#endif

            return value % 2 == 0;

        }
    }
    #endregion

    #region Delta
    /// <summary>
    /// Implements the Delta Excel function. Look at Excel documentation for more information.
    /// </summary>
    public class DeltaImpl : TUserDefinedFunction
    {
        /// <summary>
        /// Creates a new implementation.
        /// </summary>
        public DeltaImpl()
            : base("DELTA")
        {
        }

        /// <summary>
        /// Evaluates the function. Look At Excel docs for parameters.
        /// </summary>
        /// <param name="arguments"></param>
        /// <param name="parameters"></param>
        /// <returns></returns>
        public override object Evaluate(TUdfEventArgs arguments, object[] parameters)
        {
            TFlxFormulaErrorValue Err;
            if (!CheckParameters(parameters, -1, out Err)) return Err;
            if (parameters.Length < 1 || parameters.Length > 2) return TFlxFormulaErrorValue.ErrValue;

            //value1.
            double value1;
            if (!TryGetDouble(arguments.Xls, parameters[0], out value1, out Err)) return Err;
            //value2.

            double value2 = 0;
            if (parameters.Length > 1)
            {
                if (!TryGetDouble(arguments.Xls, parameters[1], out value2, out Err)) return Err;
            }

            return value1 == value2 ? 1 : 0;

        }
    }
    #endregion

    #region DollarDe / DollarFr
    /// <summary>
    /// Implements the DollarDe Excel function. Look at Excel documentation for more information.
    /// </summary>
    public class DollarDeImpl : TUserDefinedFunction
    {
        /// <summary>
        /// Creates a new implementation.
        /// </summary>
        public DollarDeImpl()
            : base("DOLLARDE")
        {
        }

        /// <summary>
        /// Evaluates the function. Look At Excel docs for parameters.
        /// </summary>
        /// <param name="arguments"></param>
        /// <param name="parameters"></param>
        /// <returns></returns>
        public override object Evaluate(TUdfEventArgs arguments, object[] parameters)
        {
            TFlxFormulaErrorValue Err;
            if (!CheckParameters(parameters, 2, out Err)) return Err;

            //value1.
            double value1;
            if (!TryGetDouble(arguments.Xls, parameters[0], out value1, out Err)) return Err;
            //value2.

            double fraction = 0;
            if (!TryGetDouble(arguments.Xls, parameters[1], out fraction, out Err)) return Err;
            fraction = Math.Floor(fraction);
            if (fraction < 0) return TFlxFormulaErrorValue.ErrNum;
            if (fraction == 0) return TFlxFormulaErrorValue.ErrDiv0;

#if (COMPACTFRAMEWORK)
            double tr = value1 < 0? Math.Ceiling(value1): Math.Round(value1);
#else
            double tr = Math.Truncate(value1);
#endif
            double Fact = Math.Pow(10, (Math.Ceiling(Math.Log10(fraction))));
            return tr + Fact * (value1 - tr) / fraction;

        }
    }

    /// <summary>
    /// Implements the DollarDe Excel function. Look at Excel documentation for more information.
    /// </summary>
    public class DollarFrImpl : TUserDefinedFunction
    {
        /// <summary>
        /// Creates a new implementation.
        /// </summary>
        public DollarFrImpl()
            : base("DOLLARFR")
        {
        }

        /// <summary>
        /// Evaluates the function. Look At Excel docs for parameters.
        /// </summary>
        /// <param name="arguments"></param>
        /// <param name="parameters"></param>
        /// <returns></returns>
        public override object Evaluate(TUdfEventArgs arguments, object[] parameters)
        {
            TFlxFormulaErrorValue Err;
            if (!CheckParameters(parameters, 2, out Err)) return Err;

            //value1.
            double value1;
            if (!TryGetDouble(arguments.Xls, parameters[0], out value1, out Err)) return Err;
            //value2.

            double fraction = 0;
            if (!TryGetDouble(arguments.Xls, parameters[1], out fraction, out Err)) return Err;
            fraction = Math.Floor(fraction);
            if (fraction < 0) return TFlxFormulaErrorValue.ErrNum;
            if (fraction == 0) return TFlxFormulaErrorValue.ErrDiv0;

#if (COMPACTFRAMEWORK)
            double tr = value1 < 0 ? Math.Ceiling(value1) : Math.Round(value1);
#else
            double tr = Math.Truncate(value1);
#endif
            double Fact = Math.Pow(10, (Math.Ceiling(Math.Log10(fraction))));
            return tr + fraction * (value1 - tr) / Fact;

        }
    }

    #endregion

    #region Effect
    /// <summary>
    /// Implements the Effect Excel function. Look at Excel documentation for more information.
    /// </summary>
    public class EffectImpl : TUserDefinedFunction
    {
        /// <summary>
        /// Creates a new implementation.
        /// </summary>
        public EffectImpl()
            : base("EFFECT")
        {
        }

        /// <summary>
        /// Evaluates the function. Look At Excel docs for parameters.
        /// </summary>
        /// <param name="arguments"></param>
        /// <param name="parameters"></param>
        /// <returns></returns>
        public override object Evaluate(TUdfEventArgs arguments, object[] parameters)
        {
            TFlxFormulaErrorValue Err;
            if (!CheckParameters(parameters, 2, out Err)) return Err;

            double nominal_rate;
            if (!TryGetDouble(arguments.Xls, parameters[0], out nominal_rate, out Err)) return Err;

            double npery = 0;
            if (!TryGetDouble(arguments.Xls, parameters[1], out npery, out Err)) return Err;
            npery = Math.Floor(npery);

            //Check after both args have been checked.
            if (nominal_rate <= 0) return TFlxFormulaErrorValue.ErrNum;
            if (npery < 1) return TFlxFormulaErrorValue.ErrNum;


            return Math.Pow(1 + nominal_rate / npery, npery) - 1;

        }
    }
    #endregion

    #region EOMonth
    /// <summary>
    /// Implements the EOMonth Excel function. Look at Excel documentation for more information.
    /// </summary>
    public class EOMonthImpl : TUserDefinedFunction
    {
        /// <summary>
        /// Creates a new implementation.
        /// </summary>
        public EOMonthImpl()
            : base("EOMONTH")
        {
        }

        /// <summary>
        /// Evaluates the function. Look At Excel docs for parameters.
        /// </summary>
        /// <param name="arguments"></param>
        /// <param name="parameters"></param>
        /// <returns></returns>
        public override object Evaluate(TUdfEventArgs arguments, object[] parameters)
        {
            TFlxFormulaErrorValue Err;
            if (!CheckParameters(parameters, 2, out Err)) return Err;

            DateTime StartDate;
            if (!TryGetDate(arguments.Xls, parameters[0], true, out StartDate, out Err)) return Err;

            double months;
            if (!TryGetDouble(arguments.Xls, parameters[1], out months, out Err)) return Err;
#if (COMPACTFRAMEWORK)
            months = months > 0 ? Math.Floor(months) : Math.Ceiling(months);
#else
            months = Math.Truncate(months);
#endif

            //Check after both args have been checked.
            DateTime ResultDate = StartDate.AddMonths((int)months);
            ResultDate = BaseBondsImpl.LastDayOfMonth(ResultDate);

            double Result;
            if (!FlxDateTime.TryToOADate(ResultDate, arguments.Xls.OptionsDates1904, out Result)) return TFlxFormulaErrorValue.ErrNum;
            return Result;

        }
    }
    #endregion

    #region FactDouble
    /// <summary>
    /// Implements the FactDouble Excel function. Look at Excel documentation for more information.
    /// </summary>
    public class FactDoubleImpl : TUserDefinedFunction
    {
        internal const int MaxDoubleFact = 300;

        /// <summary>
        /// Creates a new implementation.
        /// </summary>
        public FactDoubleImpl()
            : base("FACTDOUBLE")
        {
        }

        /// <summary>
        /// Evaluates the function. Look At Excel docs for parameters.
        /// </summary>
        /// <param name="arguments"></param>
        /// <param name="parameters"></param>
        /// <returns></returns>
        public override object Evaluate(TUdfEventArgs arguments, object[] parameters)
        {
            TFlxFormulaErrorValue Err;
            if (!CheckParameters(parameters, 1, out Err)) return Err;

            double n;
            if (!TryGetDouble(arguments.Xls, parameters[0], out n, out Err)) return Err;

            n = Math.Floor(n);
            if (n < -1) return TFlxFormulaErrorValue.ErrNum;
            if (n == -1) return 1d; //see http://mathworld.wolfram.com/DoubleFactorial.html
            if (n > MaxDoubleFact) return TFlxFormulaErrorValue.ErrNum;

            return CalcDoubleFact((int)n);
        }

        private object CalcDoubleFact(int n)
        {
            double Result = 1;
            for (int i = n; i > 1; i -= 2)
            {
                Result *= i;
            }

            return Result;
        }
    }
    #endregion

    #region GCD / LCM

    abstract class TGcdLcmAgg : IUserDefinedFunctionAggregator
    {
        double Result;
        bool HasData;
        bool HasError;
        public bool Process(double value, out TFlxFormulaErrorValue error)
        {
            error = TFlxFormulaErrorValue.ErrNA; //will not be used, we will process all values always. (Or we wouldn't catch errors in cells)
            if (value < 0)
            {
                HasError = true;
            }

            if (value < 0 || value > Math.Pow(2, 53)) { HasError = true;}
            if (HasError) return true;

            if (HasData)
            {
                Result = Calc(Result, Math.Floor(value));
            }
            else
            {
                Result = Math.Floor(value);
                HasData = true;
            }

            return true;
        }

        protected abstract double Calc(double a1, double a2);

        public object Value
        {
            get
            {
                if (HasError || !HasData) return TFlxFormulaErrorValue.ErrNum;
                return Result;
            }
        }

    }

    class TGcdAgg : TGcdLcmAgg, IUserDefinedFunctionAggregator
    {
        protected override double Calc(double a1, double a2)
        {
            return CalcGcd(a1, a2);
        }

        internal static double CalcGcd(double a1, double a2)
        {
            while (a1 != 0 && a2 != 0)
            {
                if (a1 > a2) a1 %= a2; else a2 %= a1;
            }
            if (a1 == 0) return a2;
            return a1;

        }
    }

    class TLcmAgg : TGcdLcmAgg, IUserDefinedFunctionAggregator
    {
        protected override double Calc(double a1, double a2)
        {
            double gcd = TGcdAgg.CalcGcd(a1, a2);
            if (gcd == 0) return 0; //a1 or a2 must be 0.
            return a1 * a2 / gcd;
        }
    }

    /// <summary>
    /// A base implementation for both GCD and LCM
    /// </summary>
    public abstract class BaseGCDLCM : TUserDefinedFunction
    {
        /// <summary>
        /// Creates a new instance of the class.
        /// </summary>
        /// <param name="aName"></param>
        protected BaseGCDLCM(string aName) : base(aName) { }

        internal abstract TGcdLcmAgg CreateAgg();

        /// <summary>
        /// Evaluates the function. Look At Excel docs for parameters.
        /// </summary>
        /// <param name="arguments"></param>
        /// <param name="parameters"></param>
        /// <returns></returns>
        public override object Evaluate(TUdfEventArgs arguments, object[] parameters)
        {
            TFlxFormulaErrorValue Err;
            if (!CheckParameters(parameters, -1, out Err)) return Err;
            if (parameters.Length == 0) return TFlxFormulaErrorValue.ErrNA;

            TGcdLcmAgg Agg = CreateAgg();
            if (!TryGetDoubleList(arguments.Xls, parameters, 0, -1, Agg, out Err)) return Err;

            return Agg.Value;
        }
    }

    /// <summary>
    /// Implements the GCD Excel function. Look at Excel documentation for more information.
    /// </summary>
    public class GcdImpl : BaseGCDLCM
    {
        /// <summary>
        /// Creates a new implementation.
        /// </summary>
        public GcdImpl()
            : base("GCD")
        {
        }

        internal override TGcdLcmAgg CreateAgg()
        {
            return new TGcdAgg();
        }
    }

    /// <summary>
    /// Implements the LCM Excel function. Look at Excel documentation for more information.
    /// </summary>
    public class LcmImpl : BaseGCDLCM
    {
        /// <summary>
        /// Creates a new implementation.
        /// </summary>
        public LcmImpl()
            : base("LCM")
        {
        }

        internal override TGcdLcmAgg CreateAgg()
        {
            return new TLcmAgg();
        }
    }

    #endregion

    #region GeStep
    /// <summary>
    /// Implements the GeStep Excel function. Look at Excel documentation for more information.
    /// </summary>
    public class GeStepImpl : TUserDefinedFunction
    {
        /// <summary>
        /// Creates a new implementation.
        /// </summary>
        public GeStepImpl()
            : base("GESTEP")
        {
        }

        /// <summary>
        /// Evaluates the function. Look At Excel docs for parameters.
        /// </summary>
        /// <param name="arguments"></param>
        /// <param name="parameters"></param>
        /// <returns></returns>
        public override object Evaluate(TUdfEventArgs arguments, object[] parameters)
        {
            TFlxFormulaErrorValue Err;
            if (!CheckParameters(parameters, -1, out Err)) return Err;
            if (parameters.Length < 1 || parameters.Length > 2) return TFlxFormulaErrorValue.ErrValue;

            //value1.
            double value1;
            if (!TryGetDouble(arguments.Xls, parameters[0], out value1, out Err)) return Err;
            //value2.

            double value2 = 0;
            if (parameters.Length > 1)
            {
                if (!TryGetDouble(arguments.Xls, parameters[1], out value2, out Err)) return Err;
            }

            return value1 >= value2 ? 1 : 0;

        }
    }
    #endregion

    #region MRound
    /// <summary>
    /// Implements the MRound Excel function. Look at Excel documentation for more information.
    /// </summary>
    public class MRoundImpl : TUserDefinedFunction
    {
        /// <summary>
        /// Creates a new implementation.
        /// </summary>
        public MRoundImpl()
            : base("MROUND")
        {
        }

        /// <summary>
        /// Evaluates the function. Look At Excel docs for parameters.
        /// </summary>
        /// <param name="arguments"></param>
        /// <param name="parameters"></param>
        /// <returns></returns>
        public override object Evaluate(TUdfEventArgs arguments, object[] parameters)
        {
            TFlxFormulaErrorValue Err;
            if (!CheckParameters(parameters, 2, out Err)) return Err;

            double x;
            if (!TryGetDouble(arguments.Xls, parameters[0], out x, out Err)) return Err;

            double multiple;
            if (!TryGetDouble(arguments.Xls, parameters[1], out multiple, out Err)) return Err;

            //Check after both args have been checked.
            if (x == 0 || multiple == 0) return 0;
            if (Math.Sign(x) != Math.Sign(multiple)) return TFlxFormulaErrorValue.ErrNum;

#if (COMPACTFRAMEWORK)
            return multiple * Math.Round(x / multiple);
#else
            return multiple * Math.Round(x / multiple, MidpointRounding.AwayFromZero);
#endif


        }
    }
    #endregion

    #region Multinomial

    class TMultinomialAgg : IUserDefinedFunctionAggregator
    {
        double ResultNum = 0;
        double ResultDen = 1;
        bool HasError;
        public bool Process(double value, out TFlxFormulaErrorValue error)
        {
            error = TFlxFormulaErrorValue.ErrNA; //will not be used, we will process all values always. (Or we wouldn't catch errors in cells)
            if (value < 0 || value > TFactToken.MaxFact)
            {
                HasError = true;
            }
            if (HasError) return true;

            double fact = TFactToken.Factorial((int)value);
            ResultNum += Math.Floor(value);
            ResultDen *= fact;
            return true;
        }

        public object Value
        {
            get
            {
                if (HasError) return TFlxFormulaErrorValue.ErrNum;
                if (ResultDen == 0) return TFlxFormulaErrorValue.ErrDiv0;
                if (ResultNum > TFactToken.MaxFact) return TFlxFormulaErrorValue.ErrNum;
                return TFactToken.Factorial((int)ResultNum) / ResultDen;
            }
        }

    }

    /// <summary>
    /// Implements the Multinomial Excel function. Look at Excel documentation for more information.
    /// </summary>
    public class MultinomialImpl : TUserDefinedFunction
    {
        /// <summary>
        /// Creates a new implementation.
        /// </summary>
        public MultinomialImpl()
            : base("MULTINOMIAL")
        {
        }

        /// <summary>
        /// Evaluates the function. Look At Excel docs for parameters.
        /// </summary>
        /// <param name="arguments"></param>
        /// <param name="parameters"></param>
        /// <returns></returns>
        public override object Evaluate(TUdfEventArgs arguments, object[] parameters)
        {
            TFlxFormulaErrorValue Err;
            if (!CheckParameters(parameters, -1, out Err)) return Err;
            if (parameters.Length == 0) return TFlxFormulaErrorValue.ErrNA;

            TMultinomialAgg Agg = new TMultinomialAgg();
            if (!TryGetDoubleList(arguments.Xls, parameters, 0, -1, Agg, out Err)) return Err;

            return Agg.Value;
        }
    }

    #endregion

    #region NetWorkDays
    class TNetWorkDaysAgg : IUserDefinedFunctionAggregator
    {
        Dictionary<DateTime, DateTime> Holidays = new Dictionary<DateTime, DateTime>();
        Dictionary<DateTime, DateTime> OutsideHolidays = new Dictionary<DateTime, DateTime>();
        ExcelFile xls;
        DateTime StartDate;
        DateTime EndDate;
        internal bool[] NonWorkingDays;

        public TNetWorkDaysAgg(ExcelFile axls, DateTime aStartDate, DateTime aEndDate, bool[] aNonWorkingDays)
        {
            xls = axls;
            if (aStartDate <= aEndDate)
            {
                StartDate = aStartDate;
                EndDate = aEndDate;
            }
            else
            {
                StartDate = aEndDate;
                EndDate = aStartDate;
            }

            NonWorkingDays = aNonWorkingDays;
        }

        public bool Process(double value, out TFlxFormulaErrorValue error)
        {
            error = TFlxFormulaErrorValue.ErrNA;
            DateTime Date0;
            if (value < 0 || !FlxDateTime.TryFromOADate(Math.Floor(value), xls.OptionsDates1904, out Date0))
            {
                error = TFlxFormulaErrorValue.ErrNum; //here Excel doesn't follow its pattern of checking everything first.
                return false;
            }

            if (Date0 < StartDate || Date0 > EndDate)
            {
                OutsideHolidays.Add(Date0, Date0);
                return true;
            }
            if (IsWeekendOrHoliday(Date0)) return true;

            Holidays.Add(Date0, Date0);


            return true;
        }

        public bool IsWeekendOrHoliday(DateTime Date0)
        {
            return IsWeekend(Date0) || Holidays.ContainsKey(Date0);
        }

        public bool IsWeekend(DateTime Date0)
        {
            return NonWorkingDays[(int)Date0.DayOfWeek];
        }

        public bool IsOutsideHoliday(DateTime Date0)
        {
            return OutsideHolidays.ContainsKey(Date0);
        }

        public int HolidayCount
        {
            get
            {
                return Holidays.Count;
            }
        }
    }

    /// <summary>
    /// Implements the basis for WorkDays and NetWorkDays Excel functions. Look at Excel documentation for more information.
    /// </summary>
    public abstract class BaseWorkDaysImpl : TUserDefinedFunction
    {
        internal bool Intl;

        /// <summary>
        /// Creates a new implementation.
        /// </summary>
        protected BaseWorkDaysImpl(bool aIntl, string aName)
            : base(aIntl ? aName + ".INTL" : aName, aIntl ? "_xlfn." + aName + ".INTL" : aName)
        {
            Intl = aIntl;
        }

        internal static bool ReadWeekends(ExcelFile xls, object parameter, ref bool[] NonWorkingDays, out TFlxFormulaErrorValue Err)
        {
            Err = TFlxFormulaErrorValue.ErrValue;

            object Weekend = GetSingleParameter(xls, parameter);

            string WeekendStr = Weekend as string;
            if (WeekendStr != null)
            {
                if (WeekendStr.Length != 7) return false;
                for (int i = 0; i < 7; i++)
                {
                    if (WeekendStr[i] == '0') NonWorkingDays[(i + 1) % 7] = false;
                    else if (WeekendStr[i] == '1') NonWorkingDays[(i + 1) % 7] = true;
                    else return false;
                }

                return true;
            }

            if (Weekend is double)
            {
                NonWorkingDays = GetNonWorkingDays((double)Weekend);
                if (NonWorkingDays == null)
                {
                    Err = TFlxFormulaErrorValue.ErrNum;
                    return false;
                }

                return true;
            }

            if (parameter == null) return true; //missing arg
            if (Weekend == null) Err = TFlxFormulaErrorValue.ErrNum; //want to make it complex?
            return false;
        }

        private static bool[] GetNonWorkingDays(double Weekend)
        {
            switch ((int)Weekend)
            {
                case 1: return new bool[] { true, false, false, false, false, false, true };
                case 2: return new bool[] { true, true, false, false, false, false, false };
                case 3: return new bool[] { false, true, true, false, false, false, false };
                case 4: return new bool[] { false, false, true, true, false, false, false };
                case 5: return new bool[] { false, false, false, true, true, false, false };
                case 6: return new bool[] { false, false, false, false, true, true, false };
                case 7: return new bool[] { false, false, false, false, false, true, true };

                case 11: return new bool[] { true, false, false, false, false, false, false };
                case 12: return new bool[] { false, true, false, false, false, false, false };
                case 13: return new bool[] { false, false, true, false, false, false, false };
                case 14: return new bool[] { false, false, false, true, false, false, false };
                case 15: return new bool[] { false, false, false, false, true, false, false };
                case 16: return new bool[] { false, false, false, false, false, true, false };
                case 17: return new bool[] { false, false, false, false, false, false, true };

            }
            return null;
        }

        internal static int GetWorkDays(DateTime StartDate, DateTime EndDate, TNetWorkDaysAgg Agg)
        {
            if (AllDaysAreNonWorking(Agg.NonWorkingDays)) return 0;

            int Neg = 1;
            if (EndDate < StartDate)
            {
                Neg = -1;
                DateTime tmp = EndDate;
                EndDate = StartDate;
                StartDate = tmp;
            }

            while (Agg.IsWeekend(EndDate)) EndDate = EndDate.AddDays(-1);
            while (Agg.IsWeekend(StartDate)) StartDate = StartDate.AddDays(+1);
            if (EndDate < StartDate) return 0;

            int Weeks = 0;
            for (int day = 0; day < 7; day++)
            {
                if (Agg.NonWorkingDays[day]) Weeks += GetNonWorkingDaysOfDay(StartDate, EndDate, Agg, (DayOfWeek)day);
            }

            return Neg * ((EndDate - StartDate).Days + 1 - Weeks - Agg.HolidayCount);
        }

        internal static bool AllDaysAreNonWorking(bool[] NonWorkingDays)
        {
            for (int i = 0; i < NonWorkingDays.Length; i++)
            {
                if (!NonWorkingDays[i])
                {
                    return false;
                }
            }
            return true;
        }

        private static int GetNonWorkingDaysOfDay(DateTime StartDate, DateTime EndDate, TNetWorkDaysAgg Agg, DayOfWeek Day)
        {
            DateTime LastSun = EndDate.AddDays(-(7 + (int)EndDate.DayOfWeek - (int)Day) % 7);
            DateTime FirstSat = StartDate.AddDays((7 + (int)Day - (int)StartDate.DayOfWeek) % 7);

            return 1 + ((LastSun - FirstSat).Days) / 7;
        }

        internal static bool[] StandardNonWorkingDays()
        {
            return new bool[] { true, false, false, false, false, false, true };
        }
    }

    /// <summary>
    /// Implements the NetWorkDays Excel function. Look at Excel documentation for more information.
    /// </summary>
    public class NetWorkDaysImpl : BaseWorkDaysImpl
    {
        /// <summary>
        /// Creates a new implementation.
        /// </summary>
        public NetWorkDaysImpl(bool aIntl)
            : base(aIntl, "NETWORKDAYS")
        {
        }

        /// <summary>
        /// Evaluates the function. Look At Excel docs for parameters.
        /// </summary>
        /// <param name="arguments"></param>
        /// <param name="parameters"></param>
        /// <returns></returns>
        public override object Evaluate(TUdfEventArgs arguments, object[] parameters)
        {
            TFlxFormulaErrorValue Err;
            if (!CheckParameters(parameters, -1, out Err)) return Err;
            if (parameters.Length < 2 || parameters.Length > 4) return TFlxFormulaErrorValue.ErrNA;

            DateTime StartDate;
            if (!TryGetDate(arguments.Xls, parameters[0], true, out StartDate, out Err)) return Err;

            DateTime EndDate;
            if (!TryGetDate(arguments.Xls, parameters[1], true, out EndDate, out Err)) return Err;

            bool[] NonWorkingDays = StandardNonWorkingDays();

            int hl = 2;
            if (Intl && parameters.Length > 2)
            {
                hl++;
                if (!ReadWeekends(arguments.Xls, parameters[2], ref NonWorkingDays, out Err)) return Err;
            }

            TNetWorkDaysAgg Agg = new TNetWorkDaysAgg(arguments.Xls, StartDate, EndDate, NonWorkingDays);
            if (!TryGetDoubleList(arguments.Xls, parameters, hl, -1, Agg, out Err)) return Err;

            return GetWorkDays(StartDate, EndDate, Agg);
        }
    }

    /// <summary>
    /// Implements the WorkDay Excel function. Look at Excel documentation for more information.
    /// </summary>
    public class WorkDayImpl : BaseWorkDaysImpl
    {
        /// <summary>
        /// Creates a new implementation.
        /// </summary>
        public WorkDayImpl(bool aIntl)
            : base(aIntl, "WORKDAY")
        {
        }

        /// <summary>
        /// Evaluates the function. Look At Excel docs for parameters.
        /// </summary>
        /// <param name="arguments"></param>
        /// <param name="parameters"></param>
        /// <returns></returns>
        public override object Evaluate(TUdfEventArgs arguments, object[] parameters)
        {
            TFlxFormulaErrorValue Err;
            if (!CheckParameters(parameters, -1, out Err)) return Err;
            if (parameters.Length < 2 || parameters.Length > 4) return TFlxFormulaErrorValue.ErrNA;

            DateTime StartDate;
            if (!TryGetDate(arguments.Xls, parameters[0], true, out StartDate, out Err)) return Err;

            double DayCount;
            if (!TryGetDouble(arguments.Xls, parameters[1], out DayCount, out Err)) return Err;
            DayCount = Math.Floor(DayCount); //A weird one. Excel uses floor here, not truncate as in everywhere else.

            bool[] NonWorkingDays = StandardNonWorkingDays();

            int hl = 2;
            if (Intl && parameters.Length > 2)
            {
                hl++;
                if (!ReadWeekends(arguments.Xls, parameters[2], ref NonWorkingDays, out Err)) return Err;
            }

            if (AllDaysAreNonWorking(NonWorkingDays)) return TFlxFormulaErrorValue.ErrValue;
            if (DayCount == 0) return StartDate;
            DateTime EndDate = StartDate.AddDays(DayCount);

            TNetWorkDaysAgg Agg = new TNetWorkDaysAgg(arguments.Xls, StartDate, EndDate, NonWorkingDays);
            if (!TryGetDoubleList(arguments.Xls, parameters, hl, -1, Agg, out Err)) return Err;

            int WorkDayCount = Math.Abs(GetWorkDays(StartDate, EndDate, Agg));
            if (!Agg.IsWeekendOrHoliday(StartDate)) WorkDayCount--; //WorkDayCount includes the ending day.
            int MissingDays = (int)Math.Abs(DayCount) - WorkDayCount;
            Debug.Assert(MissingDays >= 0);
            int sg = Math.Sign(DayCount);

            while (MissingDays > 0)
            {
                EndDate = EndDate.AddDays(sg);
                if (Agg.IsWeekend(EndDate) || Agg.IsOutsideHoliday(EndDate)) continue;
                MissingDays--;
            }


            return EndDate;
        }
    }

    #endregion

    #region Nominal
    /// <summary>
    /// Implements the Nominal Excel function. Look at Excel documentation for more information.
    /// </summary>
    public class NominalImpl : TUserDefinedFunction
    {
        /// <summary>
        /// Creates a new implementation.
        /// </summary>
        public NominalImpl()
            : base("NOMINAL")
        {
        }

        /// <summary>
        /// Evaluates the function. Look At Excel docs for parameters.
        /// </summary>
        /// <param name="arguments"></param>
        /// <param name="parameters"></param>
        /// <returns></returns>
        public override object Evaluate(TUdfEventArgs arguments, object[] parameters)
        {
            TFlxFormulaErrorValue Err;
            if (!CheckParameters(parameters, 2, out Err)) return Err;

            double effect_rate;
            if (!TryGetDouble(arguments.Xls, parameters[0], out effect_rate, out Err)) return Err;

            double npery = 0;
            if (!TryGetDouble(arguments.Xls, parameters[1], out npery, out Err)) return Err;
            npery = Math.Floor(npery);

            //Check after both args have been checked.
            if (effect_rate <= 0) return TFlxFormulaErrorValue.ErrNum;
            if (npery < 1) return TFlxFormulaErrorValue.ErrNum;


            return (Math.Pow(1 + effect_rate, 1 / npery) - 1) * npery;

        }
    }
    #endregion

    #region Quotient
    /// <summary>
    /// Implements the Quotient Excel function. Look at Excel documentation for more information.
    /// </summary>
    public class QuotientImpl : TUserDefinedFunction
    {
        /// <summary>
        /// Creates a new implementation.
        /// </summary>
        public QuotientImpl()
            : base("QUOTIENT")
        {
        }

        /// <summary>
        /// Evaluates the function. Look At Excel docs for parameters.
        /// </summary>
        /// <param name="arguments"></param>
        /// <param name="parameters"></param>
        /// <returns></returns>
        public override object Evaluate(TUdfEventArgs arguments, object[] parameters)
        {
            TFlxFormulaErrorValue Err;
            if (!CheckParameters(parameters, 2, out Err)) return Err;

            double num;
            if (!TryGetDouble(arguments.Xls, parameters[0], out num, out Err)) return Err;

            double den;
            if (!TryGetDouble(arguments.Xls, parameters[1], out den, out Err)) return Err;

            if (den == 0) return TFlxFormulaErrorValue.ErrDiv0;

#if (COMPACTFRAMEWORK)
            double value = num /den;
            if (value > 0) return Math.Floor(value); else return Math.Ceiling(value);
#else
            return Math.Truncate(num / den);
#endif


        }
    }
    #endregion

    #region RandBetween
    /// <summary>
    /// Implements the Quotient Excel function. Look at Excel documentation for more information.
    /// </summary>
    public class RandBetweenImpl : TUserDefinedFunction
    {
        internal readonly Random rnd = new Random();  //initialized only once.

        /// <summary>
        /// Creates a new implementation.
        /// </summary>
        public RandBetweenImpl()
            : base("RANDBETWEEN")
        {
        }

        /// <summary>
        /// Evaluates the function. Look At Excel docs for parameters.
        /// </summary>
        /// <param name="arguments"></param>
        /// <param name="parameters"></param>
        /// <returns></returns>
        public override object Evaluate(TUdfEventArgs arguments, object[] parameters)
        {
            TFlxFormulaErrorValue Err;
            if (!CheckParameters(parameters, 2, out Err)) return Err;

            double bottom;
            if (!TryGetDouble(arguments.Xls, parameters[0], out bottom, out Err)) return Err;
#if (COMPACTFRAMEWORK)
            bottom = bottom > 0 ? Math.Floor(bottom) : Math.Ceiling(bottom);
#else
            bottom = Math.Truncate(bottom);
#endif

            double top;
            if (!TryGetDouble(arguments.Xls, parameters[1], out top, out Err)) return Err;

#if (COMPACTFRAMEWORK)
            top = top > 0 ? Math.Floor(top) : Math.Ceiling(top);
#else
            top = Math.Truncate(top);
#endif

            if (top < bottom) return TFlxFormulaErrorValue.ErrNum;

            //rnd.next(bottom, top) looks perfect, but it only takes ints :(
            return bottom + rnd.NextDouble() * (top - bottom);
        }
    }
    #endregion

    #region SeriesSum
    class TSeriesSumAgg : IUserDefinedFunctionAggregator
    {
        double x;
        double n;
        double m;
        double Result = 0;
        double ResultCount = 0;

        public TSeriesSumAgg(double ax, double an, double am)
        {
            x = ax;
            n = an;
            m = am;
        }

        public bool Process(double value, out TFlxFormulaErrorValue error)
        {
            error = TFlxFormulaErrorValue.ErrNA; //will not be used, we will process all values always. (Or we wouldn't catch errors in cells)

            double term = value * Math.Pow(x, n + ResultCount * m);
            if (double.IsNaN(term)) term = 0; //might happen if the above is sqrt(-1), for example.
            Result += term;
            ResultCount++;
            return true;
        }

        public object Value
        {
            get
            {
                if (ResultCount == 0) return TFlxFormulaErrorValue.ErrNA;
                return Result;
            }
        }

    }


    /// <summary>
    /// Implements the SeriesSum Excel function. Look at Excel documentation for more information.
    /// </summary>
    public class SeriesSumImpl : TUserDefinedFunction
    {
        /// <summary>
        /// Creates a new implementation.
        /// </summary>
        public SeriesSumImpl()
            : base("SERIESSUM")
        {
        }

        /// <summary>
        /// Evaluates the function. Look At Excel docs for parameters.
        /// </summary>
        /// <param name="arguments"></param>
        /// <param name="parameters"></param>
        /// <returns></returns>
        public override object Evaluate(TUdfEventArgs arguments, object[] parameters)
        {
            TFlxFormulaErrorValue Err;
            if (!CheckParameters(parameters, 4, out Err)) return Err;

            double x;
            if (!TryGetDouble(arguments.Xls, parameters[0], out x, out Err)) return Err;

            double n;
            if (!TryGetDouble(arguments.Xls, parameters[1], out n, out Err)) return Err;

            double M;
            if (!TryGetDouble(arguments.Xls, parameters[2], out M, out Err)) return Err;

            TSeriesSumAgg Agg = new TSeriesSumAgg(x, n, M);
            if (!TryGetDoubleList(arguments.Xls, parameters, 3, -1, Agg, out Err)) return Err;

            return Agg.Value;


        }
    }
    #endregion

    #region SqrtPi
    /// <summary>
    /// Implements the Quotient Excel function. Look at Excel documentation for more information.
    /// </summary>
    public class SqrtPiImpl : TUserDefinedFunction
    {
        /// <summary>
        /// Creates a new implementation.
        /// </summary>
        public SqrtPiImpl()
            : base("SQRTPI")
        {
        }

        /// <summary>
        /// Evaluates the function. Look At Excel docs for parameters.
        /// </summary>
        /// <param name="arguments"></param>
        /// <param name="parameters"></param>
        /// <returns></returns>
        public override object Evaluate(TUdfEventArgs arguments, object[] parameters)
        {
            TFlxFormulaErrorValue Err;
            if (!CheckParameters(parameters, 1, out Err)) return Err;

            double num;
            if (!TryGetDouble(arguments.Xls, parameters[0], out num, out Err)) return Err;
            if (num < 0) return TFlxFormulaErrorValue.ErrNum;


            return Math.Sqrt(num * Math.PI);

        }
    }
    #endregion

    #region Weeknum
    /// <summary>
    /// Implements the Weeknum Excel function. Look at Excel documentation for more information.
    /// </summary>
    public class WeekNumImpl : TUserDefinedFunction
    {
        /// <summary>
        /// Creates a new implementation.
        /// </summary>
        public WeekNumImpl()
            : base("WEEKNUM")
        {
        }

        /// <summary>
        /// Evaluates the function. Look At Excel docs for parameters.
        /// </summary>
        /// <param name="arguments"></param>
        /// <param name="parameters"></param>
        /// <returns></returns>
        public override object Evaluate(TUdfEventArgs arguments, object[] parameters)
        {
            TFlxFormulaErrorValue Err;
            if (!CheckParameters(parameters, -1, out Err)) return Err;
            if (parameters.Length < 1 || parameters.Length > 2) return TFlxFormulaErrorValue.ErrValue;

            DateTime Date0;
            if (!TryGetDate(arguments.Xls, parameters[0], true, out Date0, out Err)) return Err;

            double DateSystem = 1;
            if (parameters.Length > 1)
            {
                if (!TryGetDouble(arguments.Xls, parameters[1], out DateSystem, out Err)) return Err;
            }

            if (DateSystem <= 0 || DateSystem > 21) return TFlxFormulaErrorValue.ErrNum;

            return GetWeekNum(Date0, (int)DateSystem);

        }

        private object GetWeekNum(DateTime Date0, int DateSystem)
        {
            CalendarWeekRule rule = CalendarWeekRule.FirstDay;
            DayOfWeek firstDayOfWeek = DayOfWeek.Sunday;

            switch (DateSystem)
            {
                case 1: break;
                case 2:
                case 11: firstDayOfWeek = DayOfWeek.Monday; break;
                case 12: firstDayOfWeek = DayOfWeek.Tuesday; break;
                case 13: firstDayOfWeek = DayOfWeek.Wednesday; break;
                case 14: firstDayOfWeek = DayOfWeek.Thursday; break;
                case 15: firstDayOfWeek = DayOfWeek.Friday; break;
                case 16: firstDayOfWeek = DayOfWeek.Saturday; break;
                case 17: firstDayOfWeek = DayOfWeek.Sunday; break;

                case 21:
                    firstDayOfWeek = DayOfWeek.Monday;
                    rule = CalendarWeekRule.FirstFourDayWeek;
                    break;

                default: return TFlxFormulaErrorValue.ErrNum;

            }

            Calendar cal = CultureInfo.CurrentCulture.Calendar;
            return cal.GetWeekOfYear(Date0, rule, firstDayOfWeek);

        }
    }
    #endregion

    #region Convert
    /// <summary>
    /// Implements the Convert Excel function. Look at Excel documentation for more information.
    /// </summary>
    public class ConvertImpl : TUserDefinedFunction
    {
        /// <summary>
        /// Creates a new implementation.
        /// </summary>
        public ConvertImpl()
            : base("CONVERT")
        {
        }

        /// <summary>
        /// Evaluates the function. Look At Excel docs for parameters.
        /// </summary>
        /// <param name="arguments"></param>
        /// <param name="parameters"></param>
        /// <returns></returns>
        public override object Evaluate(TUdfEventArgs arguments, object[] parameters)
        {
            TFlxFormulaErrorValue Err;
            if (!CheckParameters(parameters, 3, out Err)) return Err;

            double value;
            if (!TryGetDouble(arguments.Xls, parameters[0], out value, out Err)) return Err;

            string FromUnit;
            if (!TryGetString(arguments.Xls, parameters[1], out FromUnit, out Err)) return Err;

            string ToUnit;
            if (!TryGetString(arguments.Xls, parameters[2], out ToUnit, out Err)) return Err;

            return DoConversion(value, FromUnit, ToUnit);

        }

        private object DoConversion(double value, string FromUnit, string ToUnit)
        {
            FromUnit = GetRealUnit(FromUnit, ref value, true);
            ToUnit = GetRealUnit(ToUnit, ref value, false);
            double v1; double s1;
            double v2; double s2;

            if (ToMass(FromUnit, out v1))
            {
                if (!ToMass(ToUnit, out v2)) return TFlxFormulaErrorValue.ErrNA;
                return value * v1 / v2;
            }

            if (ToDistance(FromUnit, out v1))
            {
                if (!ToDistance(ToUnit, out v2)) return TFlxFormulaErrorValue.ErrNA;
                return value * v1 / v2;
            }

            if (ToTime(FromUnit, out v1))
            {
                if (!ToTime(ToUnit, out v2)) return TFlxFormulaErrorValue.ErrNA;
                return value * v1 / v2;
            }

            if (ToPressure(FromUnit, out v1))
            {
                if (!ToPressure(ToUnit, out v2)) return TFlxFormulaErrorValue.ErrNA;
                return value * v1 / v2;
            }

            if (ToForce(FromUnit, out v1))
            {
                if (!ToForce(ToUnit, out v2)) return TFlxFormulaErrorValue.ErrNA;
                return value * v1 / v2;
            }

            if (ToEnergy(FromUnit, out v1))
            {
                if (!ToEnergy(ToUnit, out v2)) return TFlxFormulaErrorValue.ErrNA;
                return value * v1 / v2;
            }

            if (ToPower(FromUnit, out v1))
            {
                if (!ToPower(ToUnit, out v2)) return TFlxFormulaErrorValue.ErrNA;
                return value * v1 / v2;
            }

            if (ToMagnetism(FromUnit, out v1))
            {
                if (!ToMagnetism(ToUnit, out v2)) return TFlxFormulaErrorValue.ErrNA;
                return value * v1 / v2;
            }

            if (ToTemperature(FromUnit, out v1, out s1))
            {
                if (!ToTemperature(ToUnit, out v2, out s2)) return TFlxFormulaErrorValue.ErrNA;
                return (value + s1) * v1 / v2 - s2;
            }

            if (ToVolume(FromUnit, out v1))
            {
                if (!ToVolume(ToUnit, out v2)) return TFlxFormulaErrorValue.ErrNA;
                return value * v1 / v2;
            }

            return TFlxFormulaErrorValue.ErrNA;
        }

        private bool ToMass(string s, out double val)
        {
            switch (s)
            {
                case "g":
                    val = 1;
                    return true;

                case "sg":
                    val = 14593.9029372064;
                    return true;

                case "lbm":
                    val = 453.59237;
                    return true;

                case "u":
                    val = 1.66053100460465E-24;
                    return true;

                case "ozm":
                    val = 28.349523125;
                    return true;
            }
            val = 0;
            return false;
        }

        private bool ToDistance(string s, out double val)
        {
            switch (s)
            {
                case "m": val = 1; return true;
                case "mi": val = 1609.344; return true;
                case "Nmi": val = 1852; return true;
                case "in": val = 0.0254; return true;
                case "ft": val = 0.3048; return true;
                case "yd": val = 0.9144; return true;
                case "ang": val = 1E-10; return true;
                case "pica": val = 0.00423333333333333; return true;
            }
            val = 0;
            return false;
        }

        private bool ToTime(string s, out double val)
        {
            switch (s)
            {
                case "yr": val = 31557600; return true;
                case "day": val = 86400; return true;
                case "hr": val = 3600; return true;
                case "mn": val = 60; return true;
                case "sec": val = 1; return true;
            }
            val = 0;
            return false;
        }

        private bool ToPressure(string s, out double val)
        {
            switch (s)
            {
                case "Pa": val = 1; return true;
                case "p": val = 1; return true;
                case "atm":
                case "at": val = 101325; return true;
                case "mmHg": val = 133.322368421053; return true;
            }
            val = 0;
            return false;
        }

        private bool ToForce(string s, out double val)
        {
            switch (s)
            {
                case "N": val = 1; return true;
                case "dyn":
                case "dy": val = 0.00001; return true;
                case "lbf": val = 4.4482216152605; return true;
            }
            val = 0;
            return false;
        }

        private bool ToEnergy(string s, out double val)
        {
            switch (s)
            {
                case "J": val = 1; return true;
                case "e": val = 0.0000001; return true;
                case "c": val = 4.184; return true;
                case "cal": val = 4.1868; return true;
                case "eV":
                case "ev": val = 1.60219000146921E-19; return true;
                case "HPh":
                case "hh": val = 2684519.53769617; return true;
                case "Wh":
                case "wh": val = 3600; return true;
                case "flb": val = 1.3558179483314; return true;
                case "BTU":
                case "btu": val = 1055.05585262; return true;

            }
            val = 0;
            return false;
        }

        private bool ToPower(string s, out double val)
        {
            switch (s)
            {
                case "HP":
                case "h": val = 1; return true;
                case "W":
                case "w": val = 0.00134102208959503; return true;
            }
            val = 0;
            return false;
        }

        private bool ToMagnetism(string s, out double val)
        {
            switch (s)
            {
                case "T": val = 1; return true;
                case "ga": val = 0.0001; return true;
            }
            val = 0;
            return false;
        }

        private bool ToTemperature(string s, out double val, out double sum)
        {
            switch (s)
            {
                case "C":
                case "cel": val = 1; sum = 0; return true;
                case "F":
                case "fah": val = 5.0 / 9.0; sum = -32; return true;
                case "K":
                case "kel": val = 1; sum = -273.15; return true;
            }
            val = 0; sum = 0;
            return false;
        }

        private bool ToVolume(string s, out double val)
        {
            switch (s)
            {
                case "tsp": val = 1;return true;
                case "tbs": val = 3;return true;
                case "oz": val = 6;return true;
                case "cup": val = 48;return true;
                case "pt": 
                case "us_pt": val = 96;return true;
                case "uk_pt": val = 115.291192848466; return true;
                case "qt": val = 192;return true;
                case "gal": val = 768;return true;
                case "l":
                case "lt": val = 202.884136211058; return true;
            }
            val = 0;
            return false;
        }

        private bool CanHavePrefix(string s)
        {
            if (s.Length < 2) return false;
            switch (s.Substring(1))
            {
                case "g":
                case "u":
                case "m":
                case "ang":
                case "sec":
                case "Pa":
                case "p":
                case "atm":
                case "at":
                case "mmHg":
                case "N":
                case "dyn":
                case "dy":
                case "J":
                case "e":
                case "c":
                case "cal":
                case "eV":
                case "ev":
                case "Wh":
                case "wh":
                case "W":
                case "w":
                case "T":
                case "ga":
                case "K":
                case "kel":
                case "l":
                case "lt":
                    return true;
            }

            return false;
        }

        private string GetRealUnit(string ToUnit, ref double value, bool SourcePrefix)
        {
            if (!CanHavePrefix(ToUnit)) return ToUnit;
            double mult = TryWithPrefix(ToUnit);
            if (mult == 0) return ToUnit;

            if (SourcePrefix) value *= mult; else value /= mult;
            return ToUnit.Substring(1);
        }

        private double TryWithPrefix(string s)
        {
            if (s.Length < 2) return 0;

            return GetPrefix(s[0]);
        }

        private double GetPrefix(char p)
        {
            switch (p)
            {
                case 'E': return 1E+18; //exa 
                case 'P': return 1E+15; //peta 
                case 'T': return 1E+12; //tera 
                case 'G': return 1E+09; //giga 
                case 'M': return 1E+06; //mega 
                case 'k': return 1E+03; //kilo 
                case 'h': return 1E+02; //hecto 
                case 'e': return 1E+01; //deka 
                case 'd': return 1E-01; //deci 
                case 'c': return 1E-02; //centi 
                case 'm': return 1E-03; //milli 
                case 'u': return 1E-06; //micro 
                case 'n': return 1E-09; //nano 
                case 'p': return 1E-12; //pico 
                case 'f': return 1E-15; //femto 
                case 'a': return 1E-18; //atto 
                default:
                    return 0;
            }
        }

    }
    #endregion

}


