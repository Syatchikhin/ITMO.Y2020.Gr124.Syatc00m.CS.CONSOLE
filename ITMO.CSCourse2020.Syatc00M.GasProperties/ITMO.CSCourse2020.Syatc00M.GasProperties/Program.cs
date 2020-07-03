using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ITMO.CSCourse2020.Syatc00M.GasProperties
{
    static class Program
    {
        public static void Main()
        {
            //--Air----------------------------------
            //string path1 = @"c:\temp\source1.xlsx";
            //GasComposition myGas1 = new GasComposition();
            //GasComposition myGas1Composition = GasComposition.ReadExcelFile(ref path1, ref myGas1);
            //GasComposition myGas1NormalizedComposition = GasComposition.Normalize(ref myGas1Composition);
            //GasComposition myGas1Calculated = GasComposition.CalculateProperties(ref myGas1NormalizedComposition);
            //GasComposition.OutputGasComposition(myGas1Composition);
            //GasComposition.PrintResults(myGas1Calculated);
            //GasComposition.SaveResultsToExcel(ref path1, ref myGas1Calculated);

            //--Natural gas------------------------
            string path2 = @"c:\temp\source2.xlsx";
            GasComposition myGas2 = new GasComposition();
            GasComposition myGas2Composition = GasComposition.ReadExcelFile(ref path2, ref myGas2);
            GasComposition myGas2NormalizedComposition = GasComposition.Normalize(ref myGas2Composition);
            GasComposition myGas2Calculated = GasComposition.CalculateProperties(ref myGas2NormalizedComposition);
            GasComposition.OutputGasComposition(myGas2Composition);
            GasComposition.PrintResults(myGas2Calculated);
            GasComposition.SaveResultsToExcel(ref path2, ref myGas2Calculated);
        }
    }
}