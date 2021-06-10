using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CoeffsFinder
{
    class CompParams
    {
        public string Name { get; set; }
        public string TestType { get; set; }
        public double AverageTime { get; set; }
        public double NCPU { get; set; }
        public double NCPUSingleThread { get; set; }
        public double NRam { get; set; }
        public double NDisk { get; set; }
        public double CountedTimeMultiThread { get; set; }
        public double CountedTimeSingleThread { get; set; }
        public double CountedTimeMultiAndSingleThread { get; set; }

        public CompParams()
        {

        }

        public CompParams(string name, string testType, double averageTime, double nCPU, double nCpuSingle, double nRam, double nDisk)
        {
            Name = name;
            TestType = testType;
            AverageTime = averageTime;
            NCPU = nCPU;
            NCPUSingleThread = nCpuSingle;
            NRam = nRam;
            NDisk = nDisk;
        }

    }
}
