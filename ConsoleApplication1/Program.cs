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

                            double[] casNum = AAA.arrayTo1DArray(casingArr, 0);
                            double[] casW = AAA.arrayTo1DArray(casingArr, 1);
                            double[] casPI = AAA.arrayTo1DArray(casingArr, 2);
                            double[] casPC = AAA.arrayTo1DArray(casingArr, 3);
                            double[] casFJ = AAA.arrayTo1DArray(casingArr, 4);
                            double[] casYM = AAA.arrayTo1DArray(casingArr, 5);
                            double[] casAJ = AAA.arrayTo1DArray(casingArr, 6);
                            double[] casK = AAA.arrayTo1DArray(casingArr, 7);
                            double[] casODPIPE = AAA.arrayTo1DArray(casingArr, 8);
                            double[] casODJOINT = AAA.arrayTo1DArray(casingArr, 9);
                            double[] casID = AAA.arrayTo1DArray(casingArr, 10);

            // Inputs: (make these read from input array)

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
                        
                        //Populate it with collapse and burst design lines
                            AAA.collapseP(ref solArray, ref colDepths, ref colDesign);
                            AAA.burstP(ref solArray, ref burDepths, ref burDesign);

                            AAA.initCasingPicks(ref solArray, ref burDepths, ref burDesign, ref casNum, ref casPI, ref SFinternalyield);

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






        }
    }
}
