using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PCPerformanceTester
{
    class TestResult
    {
        public string MachineName { get; set; }
        public string TestType { get; set; }
        public string StartDirectory { get; set; }
        public long ElapsedTime { get; set; }
        public DateTime TestTimestamp { get; set; }

        public TestResult(string machineName, string testType, string startDir, long timeMs)
        {
            MachineName = machineName;
            TestType = testType;
            StartDirectory = startDir;
            ElapsedTime = timeMs;
            TestTimestamp = DateTime.Now;
        }

    }
}
