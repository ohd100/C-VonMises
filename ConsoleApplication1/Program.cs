using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace ConsoleApplication1
{
    public class radtanStress
    {
        public static double radStress(double OD,double ID,double Pi,double Po,bool inwall)
        {
                double Ro; double Ri; double r;
                double a; double b;
    
                /*'Using Lame correlations
                'OD, ID in inches
                'Pi and Po in psi
                ' inwalloutwall of 1 measures stress at inner wall; value of 0 measures stress at outer wall
                */
                
                Ro = OD / 2;
                Ri = ID / 2;
    
                if (inwall == true)
                    r = Ri;
                else 
                    r = Ro;
                
                
    

                a = ((-(Ro * Ro) * (Ri * Ri) * (Pi - Po))) / ((((Ro * Ro) - (Ri * Ri)) * (r * r)));
                b = ((Math.Pow(Ri,2) * Pi) - (Math.Pow(Ro,2) * Po)) / (Math.Pow(Ro,2) - Math.Pow(Ri,2));

                double radStress = a + b;

                return radStress;
                //'Print #1, "sigR" & vbTab & radStress
        }

        public static double tanStress(double OD, double ID, double Pi, double Po, bool inwall)
        {
            double Ro; double Ri; double r;
            double a; double b;

            /*'Using Lame correlations
            'OD, ID in inches
            'Pi and Po in psi
            ' inwalloutwall of 1 measures stress at inner wall; value of 0 measures stress at outer wall
            */

            Ro = OD / 2;
            Ri = ID / 2;
    
            if (inwall==true)
                r = Ri;
            else 
                r = Ro;
            
    
    

            a = (((Ro * Ro) * (Ri * Ri) * (Pi - Po))) / ((((Ro * Ro) - (Ri * Ri)) * (r * r)));
            b = ((Math.Pow(Ri, 2) * Pi) - (Math.Pow(Ro, 2) * Po)) / (Math.Pow(Ro, 2) - Math.Pow(Ri, 2));

            double tanStress = a + b;


            return tanStress;
            //'Print #1, "sigR" & vbTab & radStress
        }

    }
    public class interpolate
    {
        public double[,] pullCasingSpecsExcel()
        {
            //Excel = Microsoft.Office.Interop.Excel;
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@"D:\Completions\CasingSpecs.xlsm");
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets["CasingInputs"];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            //int usedCols = xlRange.Columns;

            //xLrange = (Excel.Range)(xlWorkSheet.UsedRange.Columns

            int numRows = xlWorksheet.UsedRange.Rows.Count;
            int numCols = xlWorksheet.UsedRange.Columns.Count;

            double[,] xLarray = new double[numRows-1, numCols];

            for (int i = 2; i <= numRows; i++)
            {
                for (int j = 1; j <= numCols; j++)
                {
                    xLarray[i - 2, j - 1] = Convert.ToDouble(xlWorksheet.Cells[i, j].Value);
                }
            }

            GC.Collect();
            GC.WaitForPendingFinalizers();

            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);

            //close and release
            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);

            //quit and release
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);



            return xLarray;
        }

        public double[,] pullDesignLinesExcel()
        {
            //Excel = Microsoft.Office.Interop.Excel;
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@"D:\Completions\CasingSpecs.xlsm");
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets["BurstCollapseDesign"];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            //int usedCols = xlRange.Columns;

            //xLrange = (Excel.Range)(xlWorkSheet.UsedRange.Columns

            int numRows = xlWorksheet.UsedRange.Rows.Count;
            int numCols = xlWorksheet.UsedRange.Columns.Count;

            double[,] xLarray = new double[numRows - 1, numCols];

            for (int i = 2; i <= numRows; i++)
            {
                for (int j = 1; j <= numCols; j++)
                {
                    xLarray[i - 2, j - 1] = Convert.ToDouble(xlWorksheet.Cells[i, j].Value);
                }
            }

            GC.Collect();
            GC.WaitForPendingFinalizers();

            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);

            //close and release
            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);

            //quit and release
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);



            return xLarray;
        }

        public static int interpolateBinary(double[] inp,double x,int column)
        {
            int searchIndex; double upperVal; double lowerVal; 
            int index = Array.BinarySearch(inp, x);
            double interpolatedValue;
            

                if(index>0)
                {
                    searchIndex=index;
                }
                else
                {
                    searchIndex=~index;
                }
            
            
            return searchIndex;
        }

        public List<double> arrayToList(double[,] inp, int column)
        {
            var retList = new List<double>();

            int numRow = inp.GetLength(0);
            //int numCol = inp.GetLength(1);
            for (int i = 0; i <= numRow; i++)
            {
                retList.Add(inp[i, column-1]);
            }

            return retList;
        }
    }
    class Program
    {
        static void Main(string[] args)
        {
            interpolate AAA = new interpolate();
            double[,] casingArr = AAA.pullCasingSpecsExcel();
            double[,] designLinesArr = AAA.pullDesignLinesExcel();

            List<double> colDepths = AAA.arrayToList(designLinesArr, 1);

            ////2d to 1d list (
            //int numRow = designLinesArr.GetLength(0);
            //int numCol = designLinesArr.GetLength(1);
            //for (int i = 0; i < numRow; i++)
            //{
            //    for (int j = 0; j < numCol; j++)
            //        list.Add(arr1[i, j]);
            //}

            ////Test casing read from excel
            //for (int i = 0; i < casingArr.GetLength(0); i++)
            //{
            //    for (int j = 0; j < casingArr.GetLength(1); j++)
            //    {
            //        Console.WriteLine(casingArr[i, j]);
            //    }
            //}
            //Console.ReadLine();

            ////test design lines read from excel
            //for (int i = 0; i < designLinesArr.GetLength(0); i++)
            //{
            //    for (int j = 0; j < designLinesArr.GetLength(1); j++)
            //    {
            //        Console.WriteLine(designLinesArr[i, j]);
            //    }
            //}
            //Console.ReadLine();
            
            //Console.WriteLine("Hello, world!");
            //Console.ReadLine();





        }
    }
}
