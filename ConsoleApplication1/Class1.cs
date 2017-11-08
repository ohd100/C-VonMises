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
        public static double radStress(double OD, double ID, double Pi, double Po, bool inwall)
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
            b = ((Math.Pow(Ri, 2) * Pi) - (Math.Pow(Ro, 2) * Po)) / (Math.Pow(Ro, 2) - Math.Pow(Ri, 2));

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

            if (inwall == true)
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

        public static int interpolateBinary(double[] inp, double x, int column)
        {
            int searchIndex; double upperVal; double lowerVal;
            int index1 = Array.BinarySearch(inp, x);
            double interpolatedValue;

            searchIndex = index1;


            return searchIndex;
        }

        public double[] arrayTo1DArray(double[,] inp, int column)
        {
            int numRow = inp.GetLength(0);
            var retList = new double[numRow];

            
            //int numCol = inp.GetLength(1);
            for (int i = 0; i < numRow ; i++)
            {
                retList[i]=inp[i,column];
            }
            numRow = 0;
            
            return retList;
        }
    }

    public class trajectory
    {
        public double[,] buildTrajectory(double KOP, double DLS, double buildsection, double finalinclination, double MD, double intervalstep)
        {
            double critpoint; double endbuild;
            int numsteps = Convert.ToInt32((MD / intervalstep));
            double[,] trajectory = new double[numsteps+1, 2];
            //    'col 1: MD
            //    'col 2: Inclination angles
            endbuild = MD - (MD - buildsection - KOP);

            //'Populating MDs in trajectory array
                for (int i = 1; i <= numsteps+1; i++)
                {
                    trajectory[i - 1, 0] = (i - 1) * intervalstep;
                }

            //'Populating inclination angles
                for (int j = 1; j <= numsteps+1; j++)
                {
                    if (trajectory[j-1, 0] < KOP)
                    {
                        trajectory[j-1, 1] = 0;
                    }
                    else if (trajectory[j-1, 0] > endbuild)
                    {
                        trajectory[j-1, 1] = finalinclination;
                    }
                    else
                    {
                        trajectory[j-1, 1] = trajectory[j - 2, 1] + (DLS / 100) * intervalstep;
                    }   
                }
                    


                //i = 1
                //j = 1

                //'Populating MDs in trajectory array
                //For i = 1 To numsteps + 1
                //    trajectory(i, 1) = (i - 1) * intervalstep
                //Next i


                //'Populating inclination angles
                //For j = 1 To numsteps + 1
                //    If trajectory(j, 1) < KOP Then
                //        trajectory(j, 2) = 0
                //    ElseIf trajectory(j, 1) > endbuild Then
                //        trajectory(j, 2) = finalinclination
                //    Else: trajectory(j, 2) = trajectory(j - 1, 2) + (DLS / 100) * intervalstep
                //    End If
                //Next j

                return trajectory;
        }
        
        public void poptrajincline(ref double[,] solArray, ref double[] trajMD, ref double[] trajAngles)
        {
            //ind 0:  Column 1:   MD
            for (int mdA = 0; mdA < trajMD.GetLength(0); mdA++)
            {
                solArray[mdA, 0] = trajMD[mdA];
            }
            //ind 1:  Column 2:   Inclination angle
            for (int trA = 0; trA < trajAngles.GetLength(0); trA++)
            {
                solArray[trA, 1] = trajAngles[trA];
            }
        }
    }

    public class checks
    {
        public double maxCasing(int currRow)
        {
            //Worksheets("Trajectory").Cells(currRow, 6) = WorksheetFunction.Max(Worksheets("Trajectory").Cells(currRow, 4), Worksheets("Trajectory").Cells(currRow, 5))
            return 1.0;
        }
    }
}
