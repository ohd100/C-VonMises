using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace ConsoleApplication1
{
    public class weight
    {
        public double spotWeight(double weightbelow, double casingtype, double MudWeight, double intervalstep, bool buoyed, double g, double BF)
        {
                   
    
                int i; double w; int j; int colnum;
        


                if(!buoyed)
                {
                    w = g * intervalstep;
                }
                else
                {
                    w = g * BF * intervalstep;
                }
                   
                double spotWeight1 = weightbelow + w;
                return spotWeight1;
        }
    }
    
    public class vmStress
    {
        public double radStress(double OD, double ID, double Pi, double Po, bool inwall)
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

        public double tanStress(double OD, double ID, double Pi, double Po, bool inwall)
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

        public double bendingStress(double odconn, double odtub, double idtub, double DLS2, double P, int iscase1)
        {
            double r; double cp; double c; double l; double Po; double ccheck; double inertia; double Rp; double E; double k; double bendingstress1;

            //    'OD, ID in inches
            //    'DLS in deg/100ft
            //    ' 1 for iscase1 returns value for case 1
    
            //    'half length of pipe[in] and Young's modulus [psi] set below...change only if necessary
                l = 180;
                E = 30 * Math.Pow(10,6);
                inertia = (Math.PI / 64) * ((Math.Pow(odtub,4) - Math.Pow(idtub,4)));

                r = 0.5 * (odconn - odtub);

                Rp = 68755 / DLS2;
                c = 1 / Rp;
    
                k = Math.Pow((P / (E * inertia)),0.5);
            
            //    'Test for case 2 or 3
                ccheck = (r / (l*l)) / (0.5 - (Math.Cosh(k * l) - 1) / (k * l * Math.Sinh(k * l)));
    
    
                if(iscase1 == 1)
                {
                    cp = c;
                }
                else if(c < ccheck)
                {
                    cp = c * k * l / Math.Tanh(k * l);
                }
                else if(c > ccheck)
                {
                    cp = (c * k * l * Math.Sinh(k * l) - (k * l) - (0.5 + r / ((l *l) * c)) * (k * l) * ((Math.Cosh(k * l) - 1))) / (2 * (Math.Cosh(k * l) - 1) - (k * l * Math.Sinh(k * l)));
                }       
                else
                {
                    cp = 0;
                }
                
                bendingstress1 = (E * odtub * cp) / 2;
                return bendingstress1;
            //    'Print #1, "K" & vbTab & K
            //    'Print #1, dblCosh
            //    'Print #1, dblCosh
        }

        public void wBelow(ref double[,] solArray, double MW, double intervalstep, ref double[] casNum, ref double[] casW, int currRow)
        {
            int i; int w; int j; int colnum; double casingtype; double BF; double g;
                 weight iii = new weight();
            
            //public double weight(double weightbelow, int casingtype, double MudWeight, double intervalstep, bool buoyed, double g)
                            casingtype=solArray[currRow,5];
                            BF = 1 - (MW / 65.5);
                            j=Array.BinarySearch(casNum,casingtype);
                            g=casW[j];

                    i = solArray.GetLength(0)-1;

                    if(currRow == i) 
                    {
                        solArray[currRow, 2] = 0;
                    }
                    else
                    {
                        solArray[currRow, 2] = iii.spotWeight(solArray[currRow + 1, 2], solArray[currRow, 5], MW, intervalstep, false, g, BF);
                    }
               
                        
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

        public void collapseP(ref double[,] solArray, ref double[] colDepths, ref double[] colDesign)
        {
            //Collapse design lines in solArray
            int indexBinary;
            double lessP;
            double greaterP;
            double lessD;
            double greaterD;
            double slope;
            double currDepth;
            int solarraylength = solArray.GetLength(0);

            for (int a = 0;a<solarraylength; a++)
            {
                currDepth=solArray[a,0];

                indexBinary = Array.BinarySearch(colDepths,currDepth);
                if (indexBinary < 0)
                {
                    lessP=colDesign[~indexBinary-1];
                    greaterP=colDesign[~indexBinary];
                    lessD=colDepths[~indexBinary-1];
                    greaterD=colDepths[~indexBinary];
                    
                    slope=((greaterP-lessP)/(greaterD-lessD));
                    solArray[a, 6] = slope * (currDepth - lessD) + lessP;
                }
                else
                {
                    solArray[a, 6]=colDesign[indexBinary];
                }
            }
        }

        public void burstP(ref double[,] solArray, ref double[] burDepths, ref double[] burDesign)
        {
            //Burst design lines in solArray
            int indexBinary;
            double lessP;
            double greaterP;
            double lessD;
            double greaterD;
            double slope;
            double currDepth;
            int solarraylength = solArray.GetLength(0);

            for (int a = 0; a < solarraylength; a++)
            {
                currDepth = solArray[a, 0];

                indexBinary = Array.BinarySearch(burDepths, currDepth);
                if (indexBinary < 0)
                {
                    lessP = burDesign[~indexBinary - 1];
                    greaterP = burDesign[~indexBinary];
                    lessD = burDepths[~indexBinary - 1];
                    greaterD = burDepths[~indexBinary];

                    slope = ((greaterP - lessP) / (greaterD - lessD));
                    solArray[a, 7] = slope * (currDepth - lessD) + lessP;
                }
                else
                {
                    solArray[a, 7] = burDesign[indexBinary];
                }
            }
        }

        public void initCasingPicks(ref double[,] solArray, ref double[] burDepths, ref double[] burDesign, ref double[] casNum, ref double[] casPI, ref double SFinternalyield)
        {
                //Initial burst and collapse casing picks (equivalent to initCasingDesign in VBA)
                int indexBinary; 
                int solarraylength = solArray.GetLength(0);
                double burlin;
                int lastrow=casNum.GetLength(0);
                double munge;
                int k=lastrow;
    
                    //    j = Worksheets("Trajectory").UsedRange.Rows.Count
                    //    m = Worksheets("CasingInputs").UsedRange.Rows.Count

    
                    //    k = m
                
                // Burst Casing Picks (initial)
                for(int a = 0; a < solarraylength; a++)
                {
                    k = lastrow;
                    burlin = solArray[a, 7];
                    munge = casPI[k-1]/SFinternalyield - burlin;
                    while (munge >= 0 && k <= lastrow && k > 1)
                    {
                        solArray[a, 3] = k;
                        k = k - 1;
                        if (k == 1)
                        {
                            if((casPI[k-1] / SFinternalyield) - burlin>0)
                            {
                                solArray[a, 3] = 1;
                            }
                            break;
                        }
                        munge = (casPI[k-1] / SFinternalyield) - burlin;
                        //Console.WriteLine(munge);
                        //Console.WriteLine(casPI[k - 1] / SFinternalyield);
                        //Console.ReadLine();
                    }
                    k = lastrow;            
                }

            // Initial Collapse casing picks (pick weakest)
                        for(int a = 0; a < solarraylength; a++)
                        {
                            solArray[a, 4] = 1;
                        }                        
        }

        public void maxCasing(ref double[,] solArray, int currRow)
        {
            solArray[currRow,5] = Math.Max(solArray[currRow,3], solArray[currRow,4]);
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
