using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Globalization;

namespace CoeffsFinder
{
    class Program
    {
        static List<CompParams> EthCompParams { get; set; }

        IEnumerable<string> TestTypes = new List<string> { "7ZIP_CODE_EXTRACTING", "7ZIP_VIDEO_EXTRACTING", "EXCEL_TEST_READ", "EXCEL_TEST_WRITE", "WORD_TEST_READ", "WORD_TEST_WRITE" };

        static void Main(string[] args)
        {
            string[] csvRefArray = File.ReadAllLines("REF_DATA.csv");
            EthCompParams = new List<CompParams>();
            for(int i=1;i<csvRefArray.Length;i++)
            {
                string[] ethParams = csvRefArray[i].Split(';');

                string EthName = ethParams[0];
                string EthTestType = ethParams[1];
                double EthAverageTime = ConvertStringToDouble(ethParams[2]);
                double EthNCpu = ConvertStringToDouble(ethParams[7]);
                double EthNCpuSingle = ConvertStringToDouble(ethParams[8]);
                double EthNRam = ConvertStringToDouble(ethParams[9]);
                double EthNDisk = ConvertStringToDouble(ethParams[10]);
                EthCompParams.Add(new CompParams(EthName, EthTestType, EthAverageTime, EthNCpu, EthNCpuSingle, EthNRam, EthNDisk));
                
            }
            
            foreach(CompParams RefComp in EthCompParams)
            {
                Experiment(RefComp.TestType, "BOTH");
            }

            Console.WriteLine("END!");
            Console.ReadKey();
        }

        static void Experiment(string testType, string threadsType)
        {
            List<double[]> abcWhichFit = new List<double[]>();
            double[] minimalDivRecord = new double[4];

            Console.WriteLine($"EXPERIMENT: {testType} {threadsType}");
            CompParams refComp = EthCompParams.Find(comp => comp.TestType == testType);
            List<CompParams> otherComputersParams = new List<CompParams>();
            string[] csvArray = File.ReadAllLines("ALL_DATA.csv");
            for (int i = 1; i < csvArray.Length; i++)
            {
                string[] otherCompParams = csvArray[i].Split(';');
                string Name = otherCompParams[0];
                string TestType = otherCompParams[1];
                double AverageTime = ConvertStringToDouble(otherCompParams[2]);
                double NCpu = ConvertStringToDouble(otherCompParams[7]);
                double NCpuSingle = ConvertStringToDouble(otherCompParams[8]);
                double NRam = ConvertStringToDouble(otherCompParams[9]);
                double NDisk = ConvertStringToDouble(otherCompParams[10]);

                if (Name != refComp.Name && TestType==testType)
                {
                    otherComputersParams.Add(new CompParams(Name, TestType, AverageTime, NCpu, NCpuSingle, NRam, NDisk));
                }
            }

            Console.WriteLine("GOT ALL DATA, STARTED TO COUNT");

            double minimalDiv = 100;

            int marker = 0;

            for (double a = 0; a < 1; a += 0.001)
            {

                for (double b = 0; b < 1; b += 0.001)
                {
                    for (double c = 0; c < 1; c += 0.001)
                    {
                        double result = a + b + c;
                        result = Math.Round(result, 3);
                        if (result == 1.0)
                        {
                            List<CompParams> compsParamsBuf = new List<CompParams>(otherComputersParams);

                            foreach (CompParams comp in compsParamsBuf)
                            {
                                if(threadsType=="MULTI")
                                {
                                    CountTimeMultiThread(refComp, comp, a, b, c);
                                }
                                else if (threadsType == "SINGLE")
                                {
                                    CountTimeSingleThread(refComp,comp, a, b, c);
                                }
                                else if (threadsType == "BOTH")
                                {
                                    CountTimeMultiAndSingleThread(refComp, comp, a, b, c);
                                }

                            }

                            double averageDivergence = 0;

                            double localMinDivergence = 100;
                            double localMaxDivergence = 0;

                            foreach (CompParams comp in compsParamsBuf)
                            {
                                double divergence = 0;
                                if (threadsType == "MULTI")
                                {
                                    divergence = (double)100 - (comp.CountedTimeMultiThread / comp.AverageTime) * 100;
                                }
                                else if (threadsType == "SINGLE")
                                {
                                    divergence = (double)100 - (comp.CountedTimeSingleThread / comp.AverageTime) * 100;
                                }
                                else if (threadsType == "BOTH")
                                {
                                    divergence = (double)100 - (comp.CountedTimeMultiAndSingleThread / comp.AverageTime) * 100;
                                }

                                divergence = Math.Abs(divergence);

                                if (localMinDivergence > divergence)
                                {
                                    localMinDivergence = divergence;
                                }

                                if (localMaxDivergence < divergence)
                                {
                                    localMaxDivergence = divergence;
                                }

                                averageDivergence += divergence;
                            }

                            averageDivergence = averageDivergence - localMaxDivergence - localMinDivergence;
                            averageDivergence = averageDivergence / (compsParamsBuf.Count - 2);

                            if (averageDivergence < 15)
                            {
                                abcWhichFit.Add(new double[4] { a, b, c, averageDivergence });
                                Console.WriteLine($"a = {a}, b = {b}, c = {c} div = {averageDivergence}");
                            }

                            if (minimalDiv > averageDivergence)
                            {
                                minimalDiv = averageDivergence;
                                minimalDivRecord[0] = a;
                                minimalDivRecord[1] = b;
                                minimalDivRecord[2] = c;
                                minimalDivRecord[3] = averageDivergence;
                            }

                        }
                    }
                }

                marker++;
                if (marker == 100)
                {
                    marker = 0;
                    Console.WriteLine("A 100 CIRCLE");
                }
            }

            List<string> outputExperimentalData = new List<string>();
            foreach (double[] coefsArray in abcWhichFit)
            {
                outputExperimentalData.Add($"a = {coefsArray[0]}, b = {coefsArray[1]}, c = {coefsArray[2]}, average_div = {coefsArray[3]}");
            }

            outputExperimentalData.Add("");
            outputExperimentalData.Add($"a = {minimalDivRecord[0]}, b = {minimalDivRecord[1]}, c = {minimalDivRecord[2]}, average_div = {minimalDivRecord[3]}"); ;

            File.WriteAllLines($"ExperimentOutput_{testType}_{threadsType}.txt", outputExperimentalData);
        }

        static double ConvertStringToDouble(string DoubleStr)
        {
            double output = 0;
            DoubleStr = DoubleStr.Replace(',', '.');
            if (double.TryParse(DoubleStr, NumberStyles.Any, CultureInfo.InvariantCulture, out output))
            {
                return output;
            }
            else
            {
                return -1;
            }
        }

        static void CountTimeMultiThread(CompParams refComp, CompParams compParams, double a, double b, double c)
        {
            compParams.CountedTimeMultiThread = refComp.AverageTime / (a * compParams.NCPU + b * compParams.NRam + c * compParams.NDisk);
        }

        static void CountTimeMultiAndSingleThread(CompParams refComp, CompParams compParams, double a, double b, double c)
        {
            compParams.CountedTimeMultiAndSingleThread = refComp.AverageTime / (a * ((compParams.NCPU+compParams.NCPUSingleThread)/2) + b * compParams.NRam + c * compParams.NDisk);
        }

        static void CountTimeSingleThread(CompParams refComp, CompParams compParams, double a, double b, double c)
        {
            compParams.CountedTimeSingleThread = refComp.AverageTime / (a * compParams.NCPUSingleThread + b * compParams.NRam + c * compParams.NDisk);
        }
    }
}
