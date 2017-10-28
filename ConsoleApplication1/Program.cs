using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace ConsoleApplication1
{
   
    class Program
    {
        static void Main(string[] args)
        {
            interpolate AAA = new interpolate();
            double[,] casingArr = AAA.pullCasingSpecsExcel();
            double[,] designLinesArr = AAA.pullDesignLinesExcel();

            List<double> colDepths = AAA.arrayToList(designLinesArr, 1);
            List<double> colDesign = AAA.arrayToList(designLinesArr, 2);
            List<double> burDepths = AAA.arrayToList(designLinesArr, 3);
            List<double> burDesign = AAA.arrayToList(designLinesArr, 4);
            

            // Inputs:
                            double KOP = 4500;
                            double DLS = 3.5;
                            double buildsection = 2571;
                            double finalinclination = 90;
                            double MD = 9071;
                            double intervalstep = 1;
                            double MW = 10.6;
                            double SFcollapse = 1.125;
                            double SFjoint = 1.5;
                            double SFyield = 1.5;
                            double SFinternalyield = 1; //burst
                            double Mu = 0.3;    //friction factor
            
            //  Build Trajectory:
                            trajectory BBB = new trajectory();
                            double[,] trajectoryarray = BBB.buildTrajectory(KOP,DLS,buildsection,finalinclination,MD,intervalstep);

                        //    'Set current row to bottom of sheet
                        //    currRow = Worksheets("Trajectory").UsedRange.Rows.Count
    
                        //    Open "D:\Completions\HW5\HW5Log.txt" For Output As #1
                        //    Print #1, "StartLog"
    
    
                        //    'for Last Row, set weight and max casing
                        //   For currRow = Worksheets("Trajectory").UsedRange.Rows.Count To 3 Step -1
                        //        Print #1, "MD= " & Worksheets("Trajectory").Cells(currRow, 1) & " FT"
                        //        Call maxCasing(currRow)
                        //        Call wBelow(currRow)
                        //        Call collapseCheck(currRow)
                        //        Call correctedcollapseCheck(currRow)
                        //        Call burstCheck(currRow)
                        //        Call bodyCheck(currRow)
                        //        Call jointCheck(currRow)
                        //        Call finalcollcheck(currRow)
                        //        Print #1, "---------------------------------------------------------"
                        //    Next currRow
    
                        //    Close #1

                        //End Sub



            ////2d to 1d list (
            //int numRow = designLinesArr.GetLength(0);
            //int numCol = designLinesArr.GetLength(1);
            //for (int i = 0; i < numRow; i++)
            //{
            //    for (int j = 0; j < numCol; j++)
            //        list.Add(arr1[i, j]);
            //}

                            ////Test casing read from excel
                            //for (int i = 0; i < trajectoryarray.GetLength(0); i++)
                            //{
                            //    for (int j = 0; j < trajectoryarray.GetLength(1); j++)
                            //    {
                            //        Console.WriteLine(trajectoryarray[i, j]);
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

            //////test lists
            //for (int i = 0; i < colDepths.Count; i++)
            //{
                
            //        Console.WriteLine(colDepths[i]);
                
            //}
            //Console.ReadLine();






        }
    }
}
