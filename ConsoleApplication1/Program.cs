using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace ConsoleApplication1
{
   
    class Program
    {
        static void Main(string[] args)
        {
            int atRow = 0;
            interpolate AAA = new interpolate();
            double[,] casingArr = AAA.pullCasingSpecsExcel();
            double[,] designLinesArr = AAA.pullDesignLinesExcel();

            double[] colDepths = AAA.arrayTo1DArray(designLinesArr, 0);
            double[] colDesign = AAA.arrayTo1DArray(designLinesArr, 1);
            double[] burDepths = AAA.arrayTo1DArray(designLinesArr, 2);
            double[] burDesign = AAA.arrayTo1DArray(designLinesArr, 3);
            

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

                            double[] trajMD = AAA.arrayTo1DArray(trajectoryarray, 0);
                            double[] trajAngles = AAA.arrayTo1DArray(trajectoryarray, 1);

                                //List<double> burDepths = AAA.arrayToList(trajectoryarray, 3);
                                //List<double> burDesign = AAA.arrayToList(trajectoryarray, 4);
                        
                        // Create new solution array
                            double[,] solArray = new double[trajMD.GetLength(0), 17];
                        
                        //Populate it with MD and inclination angles
                            BBB.poptrajincline(ref solArray,ref trajMD,ref trajAngles);
                            AAA.collapseP(ref solArray, ref colDepths, ref colDesign);
                            AAA.burstP(ref solArray, ref burDepths, ref burDesign);

                        //Write solArray to file

                            string textfilepath = @"D:\Completions\HW5path.txt";
                            
                            using (StreamWriter outfile = new StreamWriter(textfilepath))
                            {
                                for (int x = 0; x < solArray.GetLength(0); x++)
                                {
                                    string content = "";
                                    for (int y = 0; y < 17; y++)
                                    {
                                        content += solArray[x, y].ToString("0.000") + ";";
                                    }
                                    outfile.WriteLine(content);
                                }
                            }

                                //ind 2:  Column 3:   Weight at point
                                //ind 3:  Column 4:   CasingType_Burst
                                //ind 4:  Column 5:   CasingType_Collapse
                                //ind 5:  Column 6:   CasingType_MaxNeeded
                                //ind 6:  Column 7:   CollapseLine[psi]
                                //ind 7:  Column 8:   BurstLine[psi]
                                //ind 8:  Column 9:   Force[lbs]
                                //ind 9:  Column 10:   CollEqStress[psi]
                                //ind 10:  Column 11:   BurstEqStress[psi]
                                //ind 11:  Column 12:   YieldMax[psi]
                                //ind 12:  Column 13:   CorrCollResist[psi]
                                //ind 13:  Column 14:   DLSValue[deg/100ft]
                                //ind 14:  Column 15:   BendingStress[psi]
                                //ind 15:  Column 16:   BodyCheck[1 or 0]
                                //ind 16:  Column 17:   JointCheck[1 or 0]
                       
                        //Initial burst and collapse casing picks (equivalent to initCasingDesign in VBA)
                                    //int indexBinary;                    
                                    //for (int a = 0; a<=trajAngles.Count - 1; a++)
                                    //{

                                    //    indexBinary = colDepths.BinarySearch(solArray);
                                    //    if (indexBinary < 0)
                                    //    {
                                    //        solArray[a, 6]=((colDesign[~indexBinary]-colDesign[~indexBinary-1])/(colDepths[~indexBinary]-colDepths[~indexBinary-1]))*
                                    //    }
                                    //    else
                                    //    {
                                    //        solArray[a, 6]=colDesign[indexBinary];
                                    //    }
                                    
                                    //k = m
                                    //burlin = Worksheets("Trajectory").Cells(i, 8).Value
                                    //munge = Worksheets("CasingInputs").Cells(k, 3).Value / SFinternalyield - burlin
                                    //Do While munge >= 0 And k <= m And k > 1
                                    //    Worksheets("Trajectory").Cells(i, 4) = k - 1
                                    //    k = k - 1
                                    //    If (k = 1) Then Exit Do
                                    //    munge = CDbl(Worksheets("CasingInputs").Cells(k, 3).Value) / SFinternalyield - burlin
                
                                    //k = m


                                    //  For i = 3 To j  'collapse
                                    //   'Worksheets("Trajectory").Cells(i, 4) = 1
                                    //    Worksheets("Trajectory").Cells(i, 5) = 1
                                    //    Next i
                                    //}
                                    

    

  







                        //    'Set current row to bottom of sheet
                               atRow = trajMD.GetLength(0);
    
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

                        // End Excel Code

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
