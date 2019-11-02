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
        public static string BaseDirectory = System.AppDomain.CurrentDomain.BaseDirectory + "TESTDATA";
        public static string CurrentDirectory;
        public static string ConfigFilePath;
        public static string ClearDirectory = null;
    }
    class SerialParameter
    {
        public string comName = "COM1";
        public string baudRate = "38400";
        public string dataBit = "8";
        public string stopBit = "1";
        public string parityBit = "none";
        public bool serial_enable = true;
    }
    class SerialData
    {
        public List<byte> buffer = new List<byte>(4096);
        public UInt16 index = 0;
    }
    class TestCfgPara
    {
        public int numOftestChannels;
        public bool[] serialportEnable = new bool[7];
    }
    public class Serialdata
    {
        public List<byte> buffer = new List<byte>(4096);
        Fogdata fogdata = new Fogdata();
    }
    public class Fogdata
    {
        public int i_fdata;
        public int i_tdata;
        public byte[] arrayRCVData = new byte[10];
        public int Counter;
        public double d_fdata;
        public double d_tdata;
        public double d_fdata_1s;
        public double d_tdata_1s;
        public List<double> fdata_array = new List<double>();
        public List<double> tdata_array = new List<double>();
        public List<double> fdata_1s_array = new List<double>();
        public List<double> tdata_1s_array = new List<double>();
        public double ave_Fog_data;
        public double std_Fog_data;
        public double Fog_Bias_std;
        public string FOGID;
        public List<byte> buffer = new List<byte>(4096);

    }
    class TableData
    {
        public double table_rate;
        public double table_drate;
        public double table_angle;
        public byte[] arrayOriginData = new byte[12];
        public int Counter = 0;
    }
    class TimePara
    {
        public int total_time = 0;
        public int dt;
        public int sampleFreq = 200;
    }
}
