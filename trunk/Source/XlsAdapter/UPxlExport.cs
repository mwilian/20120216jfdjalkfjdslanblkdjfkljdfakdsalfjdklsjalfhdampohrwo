using System;
using System.Text;
using System.IO;
using FlexCel.Core;

namespace FlexCel.XlsAdapter
{
    /// <summary>
    /// A stream abstraction to write Pxl files.
    /// </summary>
    internal class TPxlStream
    {
        internal Stream St;

        internal TPxlStream(Stream aSt)
        {
            St = aSt;
        }

        internal void WriteByte(byte data)
        {
            St.WriteByte(data);
        }

        internal void Write(byte[] data, int Start, int Count)
        {
            St.Write(data, Start, Count);
        }

        internal void Write16(UInt16 data)
        {
            St.WriteByte((byte)(data & 0xFF));
            St.WriteByte((byte)((data >> 8) & 0xFF));
        }

        internal void WriteString8  (string s)
        {
            byte[] Data = Encoding.Unicode.GetBytes(s);
            St.WriteByte((byte)(Data.Length / 2));
            St.Write(Data, 0, Data.Length);            
        }

        internal void WriteString16(string s)
        {
            byte[] Data = Encoding.Unicode.GetBytes(s);
            Write16((UInt16)(Data.Length / 2));
            St.Write(Data, 0, Data.Length);           
        }

    }

    /// <summary>
    /// Extra data needed to save a pxl file.
    /// </summary>
    internal class TPxlSaveData
    {
        internal TWorkbookGlobals Globals;
        internal IFlexCelPalette Palette { get { return Globals == null ? null : Globals.Workbook; } }
        internal TPxlSaveData(TWorkbookGlobals aGlobals)
        {
            Globals = aGlobals;
        }

        internal ushort GetBiff8FromCellXF(int XF)
        {
            if (XF < 0 || XF > Globals.CellXF.Count) XF = 0;
            return (UInt16) (XF); //no styles here.
        }
    }
}
