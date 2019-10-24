using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FOGTestPlatform
{
    class ParameterClass
    {
    }
    class FilePara
    {
        public static string DataDirectory = null;
        public static string DataFileName = null;
        public static string IMUDataBaseDirectory = System.AppDomain.CurrentDomain.BaseDirectory + "TESTDATA";
        public static string IMUDataCurrentDirectory;
        public static string ClearDirectory = null;
    }
    class serialParameter
    {
        public string comName = "COM1";
        public string baudRate = "38400";
        public string dataBit = "8";
        public string stopBit = "1";
        public string parityBit = "none";
        public bool serial_enable = true;
    }
    class serialData
    {
        public List<byte> buffer = new List<byte>(4096);
        public UInt16 index = 0;
    }
    class TestCfgPara
    {
        public int numOftestChannels;
     
    }
   
}
