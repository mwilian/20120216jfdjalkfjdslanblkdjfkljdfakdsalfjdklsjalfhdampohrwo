using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;

using System.Diagnostics;

namespace QueryBuilder
{
    //  Class Delegation
    // following class was VB module
    static public class Delegation 
    { 
        public delegate void EvaluateCell( string CellAddress, ref string ParaIndicator, ref string CellValue );
        public delegate bool IsCellAddress( string CellAddress, ref string cellValue );
        public delegate string[] GetNames();
        
        // Public Sub TestEval(ByVal CellAddress As String, ByRef ParaIndicator As String, ByRef CellValue As String)
        //     If CellAddress.Length = 2 Then
        //         'is address
        //         CellValue = CellAddress
        //         ParaIndicator = "{P}"
        //     Else
        //         'isnot address
        //         CellValue = CellAddress
        //         ParaIndicator = String.Empty
        //     End If
        
        // End Sub
    } 
    
    
} 

