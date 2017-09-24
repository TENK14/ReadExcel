using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;       //Microsoft Excel 14 object in references-> COM tab

namespace ReadExcel
{
    class Program
    {
        /// <summary>
        /// https://coderwall.com/p/app3ya/read-excel-file-in-c
        /// http://csharp.net-informations.com/excel/csharp-read-excel.htm
        /// </summary>
        /// <param name="args"></param>
        static void Main(string[] args)
        {
            int rowCount = 3;
            int colCount = 3;

            //var location = $@"{AppDomain.CurrentDomain.BaseDirectory}Data\excelSheet.xlsx";
            var location = $@"{AppDomain.CurrentDomain.BaseDirectory}excelSheet.xlsx";
            //Create COM Objects. Create a COM object for everything that is referenced
            Excel.Application xlApp = new Excel.Application(); 
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(location); // D:\_tmp\excelSheet.xlsx

            //Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];//Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@"D:\_tmp\excelSheet.xlsx"); // D:\_tmp\excelSheet.xlsx

            foreach (Excel._Worksheet xlWorksheet in xlWorkbook.Sheets)
            {
                Console.WriteLine(new String('-', 20));
                Excel.Range xlRange = xlWorksheet.UsedRange;

                Console.WriteLine(xlWorksheet.Name);

                //iterate over the rows and columns and print to the console as it appears in the file
                //excel is not zero based!!

                rowCount = xlRange.Rows.Count;
                colCount = xlRange.Columns.Count;

                for (int i = 1; i <= rowCount; i++)
                {
                    for (int j = 1; j <= colCount; j++)
                    {
                        //new line
                        if (j == 1)
                        {
                            Console.Write("\r\n");
                        }

                        //write the value to the console
                        if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                        {
                            Console.Write(xlRange.Cells[i, j].Value2.ToString() + "\t");
                        }
                        else
                        {
                            Console.Write("\t");
                        }

                        //add useful things here!   
                    }
                }
                Console.WriteLine();

                Marshal.ReleaseComObject(xlRange);
                Marshal.ReleaseComObject(xlWorksheet);
            }
            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            //rule of thumb for releasing com objects:
            //  never use two dots, all COM objects must be referenced and released individually
            //  ex: [somthing].[something].[something] is bad

            //release com objects to fully kill excel process from running in the background
            //Marshal.ReleaseComObject(xlRange);
            //Marshal.ReleaseComObject(xlWorksheet);

            //close and release
            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);

            //quit and release
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);

            Console.ReadKey();

        }
    }
}
