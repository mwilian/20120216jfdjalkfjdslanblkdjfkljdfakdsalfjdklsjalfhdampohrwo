using System;
using System.Text;

using FlexCel.Core;
using System.Collections.Generic;



namespace FlexCel.XlsAdapter
{
    /// <summary>
    /// One property and its offset in the file.
    /// </summary>
    internal struct TPropIdOffset: IComparable
    {
        internal UInt32 Id;
        internal UInt32 Offset;

        #region IComparable Members

        public int CompareTo(object obj)
        {
            return Offset.CompareTo(((TPropIdOffset)obj).Offset);
        }

        #endregion
    }

    /// <summary>
    /// One property
    /// </summary>
    internal class TUnconvertedOlePropertyName
    {
        internal UInt32 Id;
        internal TUnconvertedString Name;

        internal TOlePropertyName ConvertToPropertyName(Encoding CodePage)
        {
            return new TOlePropertyName(Id, Name.ConvertToString(CodePage));
        }
    }

    /// <summary>
    /// Used to read the strings when we do not know the actual codepage.
    /// </summary>
    internal class TUnconvertedString
    {
        internal byte[] Value;
        internal bool UseUnicode;

        internal TUnconvertedString(byte[] aValue, bool aUseUnicode)
        {
            Value = aValue;
            UseUnicode = aUseUnicode;
        }

        internal string ConvertToString(Encoding CodePage)
        {
            int li = Value.Length;
            while ((li >0) && Value[li - 1] == (char)0) li--;  //remove the empty #0 at the end.
            if (li == 0) return String.Empty;

            if (UseUnicode) 
                return Encoding.Unicode.GetString(Value, 0, li);
            else
                return CodePage.GetString(Value,0, li);
        }
    }

#if(FRAMEWORK20)
    internal sealed class TPropertyList : Dictionary<UInt32, object>
    {
        public TPropertyList(): base()
        {
        }
    }
#else
    internal sealed class TPropertyList: Hashtable
    {
        public bool TryGetValue(UInt32 key, out object Result)
        {
            Result = this[key];
            if (Result == null) return false;  
            return true;
        }
    }
#endif

    /// <summary>
    /// Manages the properties of an OLE2 file.
    /// </summary>
    internal class TOle2Properties
    {
        #region Privates
        private TPropertyList PropertyList;
        byte[] PropHeader;
        byte[][] FmtSection;
    #endregion
		
        #region Constructors
        internal TOle2Properties()
        {
            PropertyList = new TPropertyList();
        }
        #endregion

        #region Read Properties
        private void CheckHeader(TOle2File OleFile)
        {
            PropHeader = new byte[2 + 2 + 4 + 16 + 4];
            OleFile.Read(PropHeader, PropHeader.Length);
            byte[] ExpectedHeader = {0xFE, 0xFF, 0x00, 0x00};
            for (int i = 0; i < ExpectedHeader.Length; i++)
            {
                if (ExpectedHeader[i] != PropHeader[i])
                    XlsMessages.ThrowException(XlsErr.ErrInvalidPropertySector);
            }

            if (BitOps.GetCardinal(PropHeader, 24) < 1) //There should be at least one property section.
                XlsMessages.ThrowException(XlsErr.ErrInvalidPropertySector);
        }

        private void ReadPropHeader(TOle2File OleFile)
        {
            int SectCount = (int)BitOps.GetCardinal(PropHeader, 24);
            FmtSection = new byte[SectCount][];
            for (int i = 0; i < SectCount; i++)
            {
                FmtSection[i] = new byte[16 + 4];
                OleFile.Read(FmtSection[i], FmtSection[i].Length);
            }

            long SectorOffset = BitOps.GetCardinal(FmtSection[0], 16);

            OleFile.SeekForward(SectorOffset);
        }

        private object ConvertStrings(object o, Encoding CodePage)
        {
            object[] oArr = o as object[];
            if (oArr != null)
            {
                for (int i = 0; i < oArr.Length; i++)
                {
                    oArr[i] = ConvertStrings(oArr[i], CodePage);
                }
                return oArr;
            }

            TUnconvertedString uc = o as TUnconvertedString;
            if (uc != null)
            {
                return uc.ConvertToString(CodePage);
            }
            
            TUnconvertedOlePropertyName pn = o as TUnconvertedOlePropertyName;
            if (pn != null)
            {
                return pn.ConvertToPropertyName(CodePage);
            }

            return o;
        }

        private void ReadPropSectionHeader(TOle2File OleFile)
        {
            long PropStart = OleFile.Position;
            byte[] PropSectionHeader = new byte[8];
            OleFile.Read(PropSectionHeader, PropSectionHeader.Length);

            int PropCount = (int) BitOps.GetWord(PropSectionHeader, 4);
            TPropIdOffset[] PropOffsets = new TPropIdOffset[PropCount];
            for (int i = 0; i < PropOffsets.Length; i++)
            {
                OleFile.Read(PropSectionHeader, 8); //Reuse the array to avoid allocate more memory
                PropOffsets[i].Id = (UInt32) BitOps.GetInt32(PropSectionHeader, 0);
                PropOffsets[i].Offset = (UInt32) BitOps.GetInt32(PropSectionHeader, 4);
            }

            Array.Sort(PropOffsets);

            for (int i = 0; i < PropOffsets.Length; i++)
            {
                OleFile.SeekForward(PropOffsets[i].Offset + PropStart);
                ReadProperty(OleFile, PropOffsets[i].Id);
            }

            //Get the actual codepage and convert the strings
            int Cp = 1252;  //windows western encoding.
            object Cpo;
            if (PropertyList.TryGetValue(1, out Cpo)) Cp = Convert.ToInt32(Cpo);

            Encoding CodePage = Encoding.GetEncoding(Cp);

            UInt32[] Keys = new UInt32[PropertyList.Count];
            PropertyList.Keys.CopyTo(Keys, 0);
            foreach(UInt32 key in Keys)
            {
                PropertyList[key] = ConvertStrings(PropertyList[key], CodePage);
            }
        }

        /*  Extracted from WTypes.h
         * * [V] - may appear in a VARIANT
         * * [T] - may appear in a TYPEDESC
         * * [P] - may appear in an OLE property set
         * * [S] - may appear in a Safe Array
         * */
        [Flags]
            internal enum TPropertyTypes
        {
            VT_EMPTY           = 0,   // [V]   [P]  nothing                    
            VT_NULL            = 1,   // [V]        SQL style Null             
            VT_I2              = 2,   // [V][T][P]  2 byte signed int          
            VT_I4              = 3,   // [V][T][P]  4 byte signed int          
            VT_R4              = 4,   // [V][T][P]  4 byte real                
            VT_R8              = 5,   // [V][T][P]  8 byte real                
            VT_CY              = 6,   // [V][T][P]  currency                   
            VT_DATE            = 7,   // [V][T][P]  date                       
            VT_BSTR            = 8,   // [V][T][P]  binary string              
            VT_DISPATCH        = 9,   // [V][T]     IDispatch FAR*             
            VT_ERROR           = 10,  // [V][T]     SCODE                      
            VT_BOOL            = 11,  // [V][T][P]  True=-1, False=0           
            VT_VARIANT         = 12,  // [V][T][P]  VARIANT FAR*               
            VT_UNKNOWN         = 13,  // [V][T]     IUnknown FAR*              
            VT_DECIMAL         = 14,  // [V][T]   [S]  16 byte fixed point     

            VT_I1              = 16,  //    [T]     signed char                
            VT_UI1             = 17,  //    [T]     unsigned char              
            VT_UI2             = 18,  //    [T]     unsigned short             
            VT_UI4             = 19,  //    [T]     unsigned long              
            VT_I8              = 20,  //    [T][P]  signed 64-bit int          
            VT_UI8             = 21,  //    [T]     unsigned 64-bit int        
            VT_INT             = 22,  //    [T]     signed machine int         
            VT_UINT            = 23,  //    [T]     unsigned machine int       
            VT_VOID            = 24,  //    [T]     C style void               
            VT_HRESULT         = 25,  //    [T]                                
            VT_PTR             = 26,  //    [T]     pointer type               
            VT_SAFEARRAY       = 27,  //    [T]     (use VT_ARRAY in VARIANT)  
            VT_CARRAY          = 28,  //    [T]     C style array              
            VT_USERDEFINED     = 29,  //    [T]     user defined type          
            VT_LPSTR           = 30,  //    [T][P]  null terminated string     
            VT_LPWSTR          = 31,  //    [T][P]  wide null terminated string

            VT_FILETIME        = 64,  //       [P]  FILETIME                   
            VT_BLOB            = 65,  //       [P]  Length prefixed bytes      
            VT_STREAM          = 66,  //       [P]  Name of the stream follows 
            VT_STORAGE         = 67,  //       [P]  Name of the storage follows
            VT_STREAMED_OBJECT = 68,  //       [P]  Stream contains an object  
            VT_STORED_OBJECT   = 69,  //       [P]  Storage contains an object 
            VT_BLOB_OBJECT     = 70,  //       [P]  Blob contains an object    
            VT_CF              = 71,  //       [P]  Clipboard format           
            VT_CLSID           = 72,  //       [P]  A Class ID                 

            VT_VECTOR        = 0x1000, //       [P]  simple counted array      
            VT_ARRAY         = 0x2000, // [V]        SAFEARRAY*                
            VT_BYREF         = 0x4000, // [V]                                  
            VT_RESERVED      = 0x8000,
            VT_ILLEGAL       = 0xffff,
            VT_ILLEGALMASKED = 0x0fff,
            VT_TYPEMASK      = 0x0fff
        }

        private void ReadProperty(TOle2File Ole2File, uint Id)
        {
            if (Id == 0) //This is the name list.
            {
                PropertyList.Add(Id, GetNameDictionary(Ole2File));
                return;
            }

            byte[] PropTypeArray = new byte[4]; 
            Ole2File.Read(PropTypeArray, PropTypeArray.Length);
            Int32 PropType = BitOps.GetInt32(PropTypeArray, 0);

            if (Id == 1 && PropType == (int)TPropertyTypes.VT_I2) //H a c k to get the correct codepage. It should not be negative.
            {
                PropType = (int)TPropertyTypes.VT_UI2;
            }

            object Value = GetOneProperty(Ole2File, PropType);
            PropertyList.Add(Id, Value);
        }

        private object GetOneProperty(TOle2File Ole2File, int PropType)
        {
            if ((PropType & (int)TPropertyTypes.VT_VECTOR) != 0)
            {
                byte[] i4 = new byte[4];
                Ole2File.Read(i4, i4.Length);

                object[] Vector = new object[BitOps.GetInt32(i4, 0)];
                for (int i = 0; i<Vector.Length; i++)
                {
                    Vector[i] = GetOneProperty(Ole2File, PropType & ~(int)TPropertyTypes.VT_VECTOR);
                }
                return Vector;
            }

            switch ((TPropertyTypes)(PropType & 0xFF))
            {
                case TPropertyTypes.VT_EMPTY:
                    return null;

                case TPropertyTypes.VT_I2:
                    byte[] i2 = new byte[2];
                    Ole2File.Read(i2, i2.Length);
                    return BitConverter.ToInt16(i2, 0);

                case TPropertyTypes.VT_UI2:  //This is not really suported, but we need to convert the CodePage to a Signed int.
                    byte[] ui2 = new byte[2];
                    Ole2File.Read(ui2, ui2.Length);
                    return (Int32)BitConverter.ToUInt16(ui2, 0);

                case TPropertyTypes.VT_I4:
                    byte[] i4 = new byte[4];
                    Ole2File.Read(i4, i4.Length);
                    return BitOps.GetInt32(i4, 0);

                case TPropertyTypes.VT_R4:
                    byte[] d4 = new byte[4];
                    Ole2File.Read(d4, d4.Length);
                    return BitConverter.ToSingle(d4, 0);

                case TPropertyTypes.VT_R8:
                    byte[] d8 = new byte[8];
                    Ole2File.Read(d8, d8.Length);
                    return BitConverter.ToDouble(d8, 0);

                case TPropertyTypes.VT_CY:
                    byte[] cy = new byte[8];
                    Ole2File.Read(cy, cy.Length);

                    return TCompactFramework.DecimalFromOACurrency(BitConverter.ToInt64(cy, 0));

                case TPropertyTypes.VT_DATE:
                    byte[] dd = new byte[8];
                    Ole2File.Read(dd, dd.Length);
                    DateTime Dt;
                    if (!FlxDateTime.TryFromOADate(BitConverter.ToDouble(dd, 0), false, out Dt)) return DateTime.MinValue;
                    return Dt.Date;

                case TPropertyTypes.VT_BSTR:
                    byte[] sl = new byte[4];
                    Ole2File.Read(sl, sl.Length);
                    UInt32 StrLen = BitConverter.ToUInt32(sl, 0);
                    if (StrLen <= 1) return String.Empty;  //StrLen includes the trailing #0
                    byte[] Str = new byte[StrLen - 1];
                    Ole2File.Read(Str, Str.Length);
                    Ole2File.SeekForward(Ole2File.Position + 1); //go over the 0 byte. This is needed for vectors/arrays.
                    return new TUnconvertedString(Str, false);

                case TPropertyTypes.VT_BOOL:
                    byte[] bl = new byte[2];
                    Ole2File.Read(bl, bl.Length);
                    return BitConverter.ToInt16(bl, 0) == 0? false: true;

                case TPropertyTypes.VT_VARIANT:
                    byte[] VariantTypeArray = new byte[4]; 
                    Ole2File.Read(VariantTypeArray, VariantTypeArray.Length);
                    Int32 VariantType = BitOps.GetInt32(VariantTypeArray, 0);
                    return GetOneProperty(Ole2File, VariantType);


                case TPropertyTypes.VT_I8:
                    byte[] i8 = new byte[8];
                    Ole2File.Read(i8, i8.Length);
                    return BitConverter.ToInt64(i8, 0);

                case TPropertyTypes.VT_LPSTR:
                    byte[] sl2 = new byte[4];
                    Ole2File.Read(sl2, sl2.Length);
                    UInt32 StrLen2 = BitConverter.ToUInt32(sl2, 0);
                    if (StrLen2 <= 1) return String.Empty;  //StrLen includes the trailing #0
                    byte[] Str2 = new byte[StrLen2 - 1];
                    Ole2File.Read(Str2, Str2.Length);
                    Ole2File.SeekForward(Ole2File.Position + 1); //go over the 0 byte. This is needed for vectors/arrays.
                    return new TUnconvertedString(Str2, false);

                case TPropertyTypes.VT_LPWSTR:
                    byte[] sl3 = new byte[4];
                    Ole2File.Read(sl3, sl3.Length);
                    UInt32 StrLen3 = BitConverter.ToUInt32(sl3, 0);
                    if (StrLen3 <= 1) return String.Empty;  //StrLen includes the trailing #0
                    byte[] Str3 = new byte[(StrLen3 - 1)*2];
                    Ole2File.SeekForward(Ole2File.Position + 2); //go over the 0 byte. This is needed for vectors/arrays.
                    Ole2File.Read(Str3, Str3.Length);
                    return new TUnconvertedString(Str3, true);

                case TPropertyTypes.VT_FILETIME:
                    byte[] ft = new byte[8];
                    Ole2File.Read(ft, ft.Length);
                    return DateTime.FromFileTime(BitConverter.ToInt64(ft, 0));

                case TPropertyTypes.VT_BLOB:
                    byte[] blb = new byte[4];
                    Ole2File.Read(blb, blb.Length);
                    UInt32 BlobLen = BitConverter.ToUInt32(blb, 0);
                    if (BlobLen <= 0) return new byte[0];  //BlobLen does not includes trailing #0
                    byte[] Blob = new byte[BlobLen];
                    Ole2File.Read(Blob, Blob.Length);
                    return Blob;
            }

            return null;  //Not a supported type.
        }


        private static object[] GetNameDictionary(TOle2File Ole2File)
        {
            byte[] PropCountArray = new byte[4]; 
            Ole2File.Read(PropCountArray, PropCountArray.Length);
            Int32 PropCount = BitOps.GetInt32(PropCountArray, 0);

            object[] Properties = new object[PropCount];
            for (int i = 0; i < Properties.Length; i++)
            {
                TUnconvertedOlePropertyName PropName = new TUnconvertedOlePropertyName();
                Ole2File.Read(PropCountArray, PropCountArray.Length);
                PropName.Id = BitConverter.ToUInt32(PropCountArray, 0);
                Ole2File.Read(PropCountArray, PropCountArray.Length);
                int StrLen = BitOps.GetInt32(PropCountArray, 0);

                if (StrLen <= 1) 
                {
                    PropName.Name = new TUnconvertedString(new byte[0], false);  //StrLen includes the trailing #0
                }
                else
                {
                    byte[] Str = new byte[StrLen - 1];
                    Ole2File.Read(Str, Str.Length);
                    Ole2File.SeekForward(Ole2File.Position + 1); //go over the 0 byte. This is needed for vectors/arrays.
                    PropName.Name = new TUnconvertedString(Str, false);
                }
                Properties[i] = PropName;
            }

            return Properties;


        }

        internal void Load(TOle2File OleFile)
        {
            PropertyList.Clear();

            CheckHeader(OleFile);
            ReadPropHeader(OleFile);
            ReadPropSectionHeader(OleFile);
        }
        #endregion

        #region Other
        internal object GetValue(UInt32 PropertyId)
        {
            object Result;
            if (PropertyList.TryGetValue(PropertyId, out Result)) return Result;
            return null;
        }
        #endregion

        #region Write Properties
        internal void Write(TOle2File OleFile)
        {
            WriteHeader(OleFile);
        }

        internal void WriteHeader(TOle2File OleFile)
        {
            OleFile.WriteRaw(PropHeader, PropHeader.Length);
            int SectCount = (int) BitOps.GetCardinal(PropHeader, 24);
            for (int i = 0; i < SectCount; i++)
            {
                OleFile.WriteRaw(FmtSection[i], FmtSection[i].Length);
            }
        }

        #endregion

    }

    internal class TFileProps
    {
        internal string Core;
        internal string App;
        internal string Custom;

        internal void Clear()
        {
            Core = null;
            App = null;
            Custom = null;
        }
    }
}
