using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;


namespace FlexCel.Core
{

    #region Virtual Read
    /// <summary>
    /// Delegate for virtual cell reads.
    /// </summary>
    /// <param name="sender">Excel file from where we are reading. Remember that the file is not fully loaded when this event is called, so you are not likely to be able to use this parameter.</param>
    /// <param name="e">Contains the value of the cell we are reading.</param>
    public delegate void VirtualCellReadEventHandler(object sender, VirtualCellReadEventArgs e);

    /// <summary>
    /// Delegate for virtual cell reads. It is called before we start reading the file, but after sheet names are known.
    /// </summary>
    /// <param name="sender">Excel file from where we are reading. Remember that the file is not fully loaded when this event is called, so you are not likely to be able to use this parameter.</param>
    /// <param name="e">Arguments for the event.</param>
    public delegate void VirtualCellStartReadingEventHandler(object sender, VirtualCellStartReadingEventArgs e);

    /// <summary>
    /// Delegate for virtual cell reads. This delegate is called after the full file has been read.
    /// </summary>
    /// <param name="sender">Excel file that has been loaded. At the time this event is called, it is fully loaded.</param>
    /// <param name="e">Arguments for the event.</param>
    public delegate void VirtualCellEndReadingEventHandler(object sender, VirtualCellEndReadingEventArgs e);

    #endregion

    /*
    /// <summary>
    /// Use this class to write a spreadsheet "as you go", without loading the full thing in memory.
    /// </summary>
    public sealed class VirtualCellWriter: IDisposable
    {
        #region IDisposable Members

        /// <summary>
        /// Finishes writing the file and frees the resources. Make sure you always call this method once you are done writing.
        /// </summary>
        public void Dispose()
        {
            throw new NotImplementedException();
        }

        #endregion
    }
    */

    /// <summary>
    /// Represents a cell, including the row, column and sheet where it was read.
    /// </summary>
    public sealed class CellValue
    {
        int FSheet;
        int FRow;
        int FCol;
        object FValue;
        int FXF;

        /// <summary>
        /// Creates a new CellValue instance.
        /// </summary>
        public CellValue(int aSheet, int aRow, int aCol, object aValue, int aXF)
        {
            FSheet = aSheet;
            FRow = aRow;
            FCol = aCol;
            FValue = aValue;
            XF = aXF;
        }

        /// <summary>
        /// Sheet where the cell was read. (1 based)
        /// </summary>
        public int Sheet { get { return FSheet; } set { FSheet = value; } }

        /// <summary>
        /// Row where the cell was read. (1 based)
        /// </summary>
        public int Row { get { return FRow; } set { FRow = value; } }

        /// <summary>
        /// Column where the cell was read. (1 based)
        /// </summary>
        public int Col { get { return FCol; } set { FCol = value; } }

        /// <summary>
        /// Value of the cell. The possible objects here are the same as the returned by <see cref="ExcelFile.GetCellValue(int, int, ref int)"/>
        /// </summary>
        public object Value { get { return FValue; } set { FValue = value; } }

        /// <summary>
        /// Format of the cell.
        /// </summary>
        public int XF { get { return FXF; } set { FXF = value; } }

    }

    /// <summary>
    /// Arguments passed in the event.
    /// </summary>
    public class VirtualCellStartReadingEventArgs: EventArgs
    {
        string[] FSheetNames;

        /// <summary>
        /// Creates a new instance.
        /// </summary>
        /// <param name="aSheetNames">Array with all the sheet names available in the file.</param>
        public VirtualCellStartReadingEventArgs(string[] aSheetNames)
        {
            FSheetNames = aSheetNames;
        }

        internal VirtualCellStartReadingEventArgs(ExcelFile Xls)
        {
            FSheetNames = new string[Xls.SheetCount];
            for (int i = 0; i < Xls.SheetCount; i++)
            {
                SheetNames[i] = Xls.GetSheetName(i + 1);
            }
        }

        /// <summary>
        /// A list with all the sheets available in the file.
        /// </summary>
        public string[] SheetNames { get { return FSheetNames; } }
    }

    /// <summary>
    /// Arguments passed in the event.
    /// </summary>
    public class VirtualCellEndReadingEventArgs : EventArgs
    {
    }   
    
    /// <summary>
    /// Arguments passed in the event.
    /// </summary>
    public class VirtualCellReadEventArgs : EventArgs
    {
        CellValue FCell;

        /// <summary>
        /// Creates a new instance.
        /// </summary>
        /// <param name="aCell">Cell value.</param>
        public VirtualCellReadEventArgs(CellValue aCell)
        {
            FCell = aCell;
        }

        /// <summary>
        /// Value and position of a cell.
        /// </summary>
        public CellValue Cell { get { return FCell; }}

    }
}
