/********************************************************************
	创建日期:	19:10:2019   16:23
	文件名: 	    E:\MyCode\C#Code\FOGTestPlatform\Form1.cs
	文件路径:	E:\MyCode\C#Code\FOGTestPlatform
	文件基类:	Form1
	扩展名:	    cs
	编写人:		Zhu Lufeng
	
	用途:	主文件
*********************************************************************/
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Windows.Forms.DataVisualization.Charting;
using System.IO.Ports;
using NPOI.HSSF.UserModel;
using NPOI.HPSF;
using NPOI.POIFS.FileSystem;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.XSSF.Util;
using NPOI.XSSF.UserModel;
using System.Runtime.InteropServices;
using System.Text;

namespace FOGTestPlatform
{

    public partial class Form1 : Form
    {
        #region 全局变量声明
        //通道串口声明
        SerialPort table_serial = new SerialPort();
        SerialPort ch1_serial = new SerialPort();
        SerialPort ch2_serial = new SerialPort();
        SerialPort ch3_serial = new SerialPort();
        SerialPort ch4_serial = new SerialPort();
        SerialPort ch5_serial = new SerialPort();
        SerialPort ch6_serial = new SerialPort();
        
        //保存文件流声明
        StreamWriter CH1_HEX_SW;
        StreamWriter CH2_HEX_SW;
        StreamWriter CH3_HEX_SW;
        StreamWriter CH4_HEX_SW;
        StreamWriter CH5_HEX_SW;
        StreamWriter CH6_HEX_SW;

        StreamWriter CH1_data_SW;
        StreamWriter CH2_data_SW;
        StreamWriter CH3_data_SW;
        StreamWriter CH4_data_SW;
        StreamWriter CH5_data_SW;
        StreamWriter CH6_data_SW;

        SerialData table_serialData = new SerialData();
        SerialData CH1_serialData = new SerialData();
        SerialData CH2_serialData = new SerialData();
        SerialData CH3_serialData = new SerialData();
        SerialData CH4_serialData = new SerialData();
        SerialData CH5_serialData = new SerialData();
        SerialData CH6_serialData = new SerialData();


        Fogdata CH1_FogData = new Fogdata();
        Fogdata CH2_FogData = new Fogdata();
        Fogdata CH3_FogData = new Fogdata();
        Fogdata CH4_FogData = new Fogdata();
        Fogdata CH5_FogData = new Fogdata();
        Fogdata CH6_FogData = new Fogdata();

        //参数类对象声明
        TestCfgPara testCfgPara = new TestCfgPara();
        TableData tabledata = new TableData();
        TimePara timePara = new TimePara();
        //定义委托
        delegate void UpdateTableFrmEventHandle();
        delegate void UpdateDataFrmEventHandle(string portName);
        UpdateDataFrmEventHandle updateDataFrm;
        UpdateTableFrmEventHandle updateTableFrmdata;
        //定义联合体
        [StructLayout(LayoutKind.Explicit, Size = 4)]
                
        public struct Union
        {
            [FieldOffset(0)]
            public Byte b0;
            [FieldOffset(1)]
            public Byte b1;
            [FieldOffset(2)]
            public Byte b2;
            [FieldOffset(3)]
            public Byte b3;
            [FieldOffset(0)]
            public Int32 i;
            [FieldOffset(0)]
            public Single f;
        }
        #endregion
        public Form1()
        {
            InitializeComponent();
            InitializeConfigFlie();
            IntializeChart();
            updateTableFrmdata += new UpdateTableFrmEventHandle(showtabledata);
            updateDataFrm += new UpdateDataFrmEventHandle(showFogdata);
        }
        /*************************************
        函数名：InitializeConfigFlie
        创建日期：2019/10/25
        函数功能：初始化配置文件
        函数参数：
        返回值：void
        *************************************/
        public void InitializeConfigFlie()
        {
            string timedata = DateTime.Now.ToString("yyyyMMdd-HHmmss");
            FilePara.CurrentDirectory = FilePara.BaseDirectory + timedata;
            Directory.CreateDirectory(FilePara.CurrentDirectory);
            try
            {
                FilePara.ConfigFilePath = FilePara.CurrentDirectory + @"\配置文件.xlsx";
                FileStream file = new FileStream(FilePara.ConfigFilePath, FileMode.Create);               
                XSSFWorkbook hsswfworkbook = new XSSFWorkbook();
                ISheet sheet = hsswfworkbook.CreateSheet("通道串口配置");
                hsswfworkbook.CreateSheet("试验配置");
                sheet.CreateRow(0).CreateCell(1).SetCellValue("使能");
                sheet.GetRow(0).CreateCell(2).SetCellValue("串口号");
                sheet.GetRow(0).CreateCell(3).SetCellValue("波特率");
                sheet.GetRow(0).CreateCell(4).SetCellValue("数据位");
                sheet.GetRow(0).CreateCell(5).SetCellValue("停止位");
                sheet.GetRow(0).CreateCell(6).SetCellValue("校验位");
                sheet.GetRow(0).CreateCell(7).SetCellValue("型号");
                sheet.GetRow(0).CreateCell(7).SetCellValue("标度因数");

                sheet.CreateRow(1).CreateCell(0).SetCellValue("转台通道");
                sheet.CreateRow(2).CreateCell(0).SetCellValue("通道一");
                sheet.CreateRow(3).CreateCell(0).SetCellValue("通道二");
                sheet.CreateRow(4).CreateCell(0).SetCellValue("通道三");
                sheet.CreateRow(5).CreateCell(0).SetCellValue("通道四");
                sheet.CreateRow(6).CreateCell(0).SetCellValue("通道五");
                sheet.CreateRow(7).CreateCell(0).SetCellValue("通道六");
                sheet.CreateRow(8).CreateCell(0).SetCellValue("测试通道数");

                sheet.SetColumnWidth(0, 12 * 256);
                hsswfworkbook.Write(file);
                file.Close();
            }
            catch (Exception)
            {

                MessageBox.Show("配置文件生产异常！");
                //throw;
            }
                 
        }
        /*************************************
        函数名：IntializeChart
        创建日期：2019/10/19
        函数功能：初始化图表
        函数参数：
        返回值：void
        *************************************/
        public void IntializeChart()
        {
            //图标的背景色
            chart.BackColor = Color.FromArgb(255, 0, 24, 55);//Color.SkyBlue;
            //图表背景色的渐变方式
            chart.BackGradientStyle = GradientStyle.None;//GradientStyle.None;
            //图表的边框线条颜色
            chart.BorderlineColor = Color.Black;
            //图表的边框线条样式
            chart.BorderlineDashStyle = ChartDashStyle.Solid;
            //图表边框线条宽度
            chart.BorderlineWidth = 2;
            //图表边框的皮肤
            chart.BorderSkin.SkinStyle = BorderSkinStyle.None;
            //图表边框宽度
            chart.BorderSkin.BorderWidth = 0;
                     
            
        }

        /*************************************
        函数名：AddChartArea
        创建日期：2019/10/25
        函数功能：添加图框
        函数参数：
        	num
        返回值：void
        *************************************/
        public void AddChartArea(int num)
        {
            for (int i = 0; i < num; i++)
            {
                chart.ChartAreas.Add(SetChartArea(i));
            }
        }
        /*************************************
        函数名：AddSeries
        创建日期：2019/10/25
        函数功能：添加数据线
        函数参数：
        	num
        返回值：void
        *************************************/
        public void AddSeries(int num)
        {
            for (int i = 0; i < num * 2; i++)
            {
                chart.Series.Add(SetSeries(i));
                if (i % 2 == 0)
                {
                    chart.Series[i].YAxisType = AxisType.Primary;
                }
                else
                {
                    chart.Series[i].YAxisType = AxisType.Secondary;
                }
            }
        }

        /*************************************
        函数名：SetSeries
        创建日期：2019/10/25
        函数功能：设置数据线格式
        函数参数：
        	index
        返回值：System.Windows.Forms.DataVisualization.Charting.Series
        *************************************/
        public Series SetSeries(int index)
        {
            Series series = new Series();
            //Series 的类型
            series.ChartType = SeriesChartType.Line;
            if (index % 2 == 0)
            {
                series.Color = Color.FromArgb(0xff, 0x32, 0xc5, 0xe9);
            }
            else
            {
                series.Color = Color.FromArgb(0xff, 0xff, 0x9f, 0x7f);
            }
            //Series线条阴影颜色
            series.ShadowColor = Color.Green;
            //阴影宽度
            series.ShadowOffset = 0;
            //是否显示数据说明
            series.IsVisibleInLegend = false;
            //线条上数据点上是否有数据显示
            series.IsValueShownAsLabel = false;
            //线条上的数据点标志类型
            series.MarkerStyle = MarkerStyle.None;
            //线条数据点的大小
            series.MarkerSize = 2;
            //Series 的边框颜色
            series.BorderColor = Color.Tomato;
            //Series线条的宽度
            series.BorderWidth = 2;

            return series;
        }
        /*************************************
        函数名：SetChartArea
        创建日期：2019/10/19
        函数功能：设置绘图区
        函数参数：
        	index
        返回值：System.Windows.Forms.DataVisualization.Charting.ChartArea
        *************************************/
        public ChartArea SetChartArea(int index)
        {
            ChartArea chartArea = new ChartArea();

            switch (index+1)
            {
                case 1:
                    chartArea.Name = ("CH1");
                    break;
                case 2:
                    chartArea.Name = ("CH2");
                    break;
                case 3:
                    chartArea.Name = ("CH3");
                    break;
                case 4:
                    chartArea.Name = ("CH4");
                    break;
                case 5:
                    chartArea.Name = ("CH5");
                    break;
                case 6:
                    chartArea.Name = ("CH6");
                    break;
                default:
                    MessageBox.Show("图表数目参数错误！");
                    break;
            }
            //背景色
            chartArea.BackColor = Color.FromArgb(255, 4, 33, 65);
            //背景渐变方式
            chartArea.BackGradientStyle = GradientStyle.None;
            //边框颜色
            chartArea.BorderColor = Color.FromArgb(255, 4, 33, 65);
            //边框柱线条宽度
            chartArea.BorderWidth = 2;
            //边框线条样式
            chartArea.BorderDashStyle = ChartDashStyle.Solid;
            //阴影颜色
            chartArea.ShadowColor = Color.Transparent;


            //设置X轴和Y轴线条的颜色和宽度
            chartArea.AxisX.LineColor = Color.Black;//.FromArgb(64, 64, 64, 64);//
            chartArea.AxisX.LineWidth = 1;
            chartArea.AxisY.LineColor = Color.Black;//.FromArgb(64, 64, 64, 64);//
            chartArea.AxisY.LineWidth = 1;
            //设置x轴和Y轴的标题
            chartArea.AxisX.Title = "时间";
            chartArea.AxisY.Title = "陀螺数据";
            chartArea.AxisY2.Title = "温度";
            chartArea.AxisX.TitleFont = new System.Drawing.Font("Microsoft Sans Serif", 12, FontStyle.Regular);
            chartArea.AxisY.TitleFont = new System.Drawing.Font("Microsoft Sans Serif", 10, FontStyle.Regular);
            chartArea.AxisY2.TitleFont = new System.Drawing.Font("Microsoft Sans Serif", 12, FontStyle.Regular);
            chartArea.AxisX.TitleForeColor = Color.FromArgb(255, 245, 254, 252);
            chartArea.AxisY.TitleForeColor = Color.FromArgb(0xff, 0x32, 0xc5, 0xe9);
            chartArea.AxisY2.TitleForeColor = Color.FromArgb(0xff, 0xff, 0x9f, 0x7f);
            //设置图表区网格横纵线条的颜色和宽度
            chartArea.AxisX.MajorGrid.LineColor = Color.FromArgb(255, 114, 175, 207);
            chartArea.AxisX.MajorGrid.LineWidth = 1;
            chartArea.AxisY.MajorGrid.LineColor = Color.FromArgb(64, 64, 64, 64);
            chartArea.AxisY.MajorGrid.LineWidth = 1;

            //启用X游标，以支持局部区域选择放大
            chartArea.CursorX.IsUserEnabled = true;
            chartArea.CursorX.IsUserSelectionEnabled = true;
            chartArea.CursorX.LineColor = Color.Pink;
            chartArea.CursorX.IntervalType = DateTimeIntervalType.Auto;
            chartArea.AxisX.ScaleView.Zoomable = false;
            chartArea.AxisX.ScrollBar.ButtonStyle = ScrollBarButtonStyles.All;//启用X轴滚动条按钮
            chartArea.AxisY.ScaleView.Zoomable = false;

            chartArea.AxisY.LabelStyle.Format = "##########.0";
            chartArea.AxisY2.LabelStyle.Format = "###.0000";
            chartArea.AxisY.LabelStyle.ForeColor = Color.FromArgb(255, 146, 175, 207);
            chartArea.AxisY2.LabelStyle.ForeColor = Color.FromArgb(255, 146, 175, 207);
            chartArea.AxisX.LabelStyle.ForeColor = Color.FromArgb(255, 151, 167, 186);

            return chartArea;
        }

        /*************************************
        函数名：ToolStripMenuItem_SerialCfgByDialog_Click
        创建日期：2019/10/22
        函数功能：通过对话框来配置测试的相关参数
        函数参数：
        	sender
        	e
        返回值：void
        *************************************/
        private void ToolStripMenuItem_SerialCfgByDialog_Click(object sender, EventArgs e)
        {
            SerialCfgDlg serialCfgDlg = new SerialCfgDlg();
            List<string> portIDList = new List<string>();
            if (serialCfgDlg.ShowDialog() != DialogResult.Cancel)
            {
                //读入配置文件
                FileStream rfile = new FileStream(FilePara.ConfigFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                XSSFWorkbook workbook = new XSSFWorkbook(rfile);
                rfile.Close();
                ISheet sht = workbook.GetSheet("通道串口配置");
                if (sht.GetRow(1).GetCell(1).ToString() == "True")
                {
                    table_serial = SetSerialPara(0);
                    testCfgPara.serialportEnable[0] = true;
                    portIDList.Add(sht.GetRow(1).GetCell(2).ToString());
                }
                else
                {
                    testCfgPara.serialportEnable[0] = false;
                }

                if (sht.GetRow(2).GetCell(1).ToString() == "True")
                {
                    ch1_serial = SetSerialPara(1);
                    testCfgPara.serialportEnable[1] = true;
                    portIDList.Add(sht.GetRow(2).GetCell(2).ToString());
                    CH1_HEX_SW = new StreamWriter(FilePara.CurrentDirectory + @"\" + sht.GetRow(2).GetCell(7).ToString() +"HexData.dat");
                    CH1_data_SW = new StreamWriter(FilePara.CurrentDirectory + @"\" + sht.GetRow(2).GetCell(7).ToString() + "Data.dat");
                }
                else
                {
                    testCfgPara.serialportEnable[1] = false;
                }

                if (sht.GetRow(3).GetCell(1).ToString() == "True")
                {
                    ch2_serial = SetSerialPara(2);
                    testCfgPara.serialportEnable[2] = true;
                    portIDList.Add(sht.GetRow(3).GetCell(2).ToString());
                    CH2_HEX_SW  = new StreamWriter(FilePara.CurrentDirectory + @"\" + sht.GetRow(3).GetCell(7).ToString() + "HexData.dat");
                    CH2_data_SW = new StreamWriter(FilePara.CurrentDirectory + @"\" + sht.GetRow(3).GetCell(7).ToString() + "Data.dat");
                }
                else
                {
                    testCfgPara.serialportEnable[2] = false;
                }

                if (sht.GetRow(4).GetCell(1).ToString() == "True")
                {
                    ch3_serial = SetSerialPara(3);
                    testCfgPara.serialportEnable[3] = true;
                    portIDList.Add(sht.GetRow(4).GetCell(2).ToString());
                    CH3_HEX_SW  = new StreamWriter(FilePara.CurrentDirectory + @"\" + sht.GetRow(4).GetCell(7).ToString() + "HexData.dat");
                    CH3_data_SW = new StreamWriter(FilePara.CurrentDirectory + @"\" + sht.GetRow(4).GetCell(7).ToString() + "Data.dat");
                }
                else
                {
                    testCfgPara.serialportEnable[3] = false;
                }

                if (sht.GetRow(5).GetCell(1).ToString() == "True")
                {
                    ch4_serial = SetSerialPara(4);
                    testCfgPara.serialportEnable[4] = true;
                    portIDList.Add(sht.GetRow(5).GetCell(2).ToString());
                    CH4_HEX_SW  = new StreamWriter(FilePara.CurrentDirectory + @"\" + sht.GetRow(5).GetCell(7).ToString() + "HexData.dat");
                    CH4_data_SW = new StreamWriter(FilePara.CurrentDirectory + @"\" + sht.GetRow(5).GetCell(7).ToString() + "Data.dat");
                }
                else
                {
                    testCfgPara.serialportEnable[4] = false;
                }

                if (sht.GetRow(6).GetCell(1).ToString() == "True")
                {
                    ch5_serial = SetSerialPara(5);
                    testCfgPara.serialportEnable[5] = true;
                    portIDList.Add(sht.GetRow(6).GetCell(2).ToString());
                    CH5_HEX_SW  = new StreamWriter(FilePara.CurrentDirectory + @"\" + sht.GetRow(6).GetCell(7).ToString() + "HexData.dat");
                    CH5_data_SW = new StreamWriter(FilePara.CurrentDirectory + @"\" + sht.GetRow(6).GetCell(7).ToString() + "Data.dat");
                }
                else
                {
                    testCfgPara.serialportEnable[5] = false;
                }


                if (sht.GetRow(7).GetCell(1).ToString() == "True")
                {
                    ch6_serial = SetSerialPara(6);
                    testCfgPara.serialportEnable[6] = true;
                    portIDList.Add(sht.GetRow(6).GetCell(2).ToString());
                    CH6_HEX_SW  = new StreamWriter(FilePara.CurrentDirectory + @"\" + sht.GetRow(7).GetCell(7).ToString() + "HexData.dat");
                    CH6_data_SW = new StreamWriter(FilePara.CurrentDirectory + @"\" + sht.GetRow(7).GetCell(7).ToString() + "Data.dat");
                }
                else
                {
                    testCfgPara.serialportEnable[6] = false;
                }
                HashSet<string> PortIDHashset = new HashSet<string>(portIDList);
                if (portIDList.Count() != PortIDHashset.Count())
                {
                    MessageBox.Show("不同通道选用了相同的串口号，请重新配置串口！");
                    Btn_Start.Enabled = false;

                }
                else
                {
                    testCfgPara.numOftestChannels = Convert.ToInt32(sht.GetRow(8).GetCell(1).ToString());
                    AddChartArea(testCfgPara.numOftestChannels);
                    AddSeries(testCfgPara.numOftestChannels);
                    Btn_Start.Enabled = true;
                }
                
            }
        }
        public SerialPort SetSerialPara(int index)
        {
            SerialPort serial          = new SerialPort();
            SerialParameter serialpara = new SerialParameter();
            //读入配置文件
            FileStream rfile      = new FileStream(FilePara.ConfigFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            XSSFWorkbook workbook = new XSSFWorkbook(rfile);
            rfile.Close();

            ISheet sht           = workbook.GetSheet("通道串口配置");
            serialpara.comName   = sht.GetRow(index + 1).GetCell(2).ToString();
            serialpara.baudRate  = sht.GetRow(index + 1).GetCell(3).ToString();
            serialpara.dataBit   = sht.GetRow(index + 1).GetCell(4).ToString();
            serialpara.stopBit   = sht.GetRow(index + 1).GetCell(5).ToString();
            serialpara.parityBit = sht.GetRow(index + 1).GetCell(6).ToString();

            serial.PortName = serialpara.comName;
            serial.BaudRate = Convert.ToInt32(serialpara.baudRate);
            serial.DataBits = Convert.ToInt32(serialpara.dataBit);

            switch(serialpara.stopBit)
            {
                case "1":
                    serial.StopBits = StopBits.One;
                    break;
                case "1.5":
                    serial.StopBits = StopBits.OnePointFive;
                    break;
                case "2":
                    serial.StopBits = StopBits.Two;
                    break;
                default:
                    serial.StopBits = StopBits.One;
                    break;
            }
            switch(serialpara.parityBit)
            {
                case "odd":
                    serial.Parity = Parity.Odd;
                    break;
                case "even":
                    serial.Parity = Parity.Even;
                    break;
                case "none":
                    serial.Parity = Parity.None;
                    break;
                default:
                    serial.Parity = Parity.None;
                    break;
            }
            return serial;
        }

        /*************************************
        函数名：Btn_Start_Click
        创建日期：2019/10/25
        函数功能：打开串口，开始测试
        函数参数：
        	sender
        	e
        返回值：void
        *************************************/
        private void Btn_Start_Click(object sender, EventArgs e)
        {
            if(testCfgPara.serialportEnable[0])
            {
                if(table_serial.IsOpen)
                {
                    table_serial.Close();
                }
                table_serial.DataReceived += new SerialDataReceivedEventHandler(tabledata_decode);
                table_serial.Open();                
            }
            if (testCfgPara.serialportEnable[1])
            {
                if(ch1_serial.IsOpen)
                {
                    ch1_serial.Close();
                }
                ch1_serial.DataReceived += new SerialDataReceivedEventHandler(channeldata_decode);
                ch1_serial.Open();
            }
            if (testCfgPara.serialportEnable[23])
            {
                if (ch2_serial.IsOpen)
                {
                    ch2_serial.Close();
                }
                ch2_serial.DataReceived += new SerialDataReceivedEventHandler(channeldata_decode);
                ch2_serial.Open();
            }
            if (testCfgPara.serialportEnable[3])
            {
                if (ch3_serial.IsOpen)
                {
                    ch3_serial.Close();
                }
                ch3_serial.DataReceived += new SerialDataReceivedEventHandler(channeldata_decode);
                ch3_serial.Open();
            }
            if (testCfgPara.serialportEnable[4])
            {
                if (ch4_serial.IsOpen)
                {
                    ch4_serial.Close();
                }
                ch4_serial.DataReceived += new SerialDataReceivedEventHandler(channeldata_decode);
                ch4_serial.Open();
            }
            if (testCfgPara.serialportEnable[5])
            {
                if (ch5_serial.IsOpen)
                {
                    ch5_serial.Close();
                }
                ch5_serial.DataReceived += new SerialDataReceivedEventHandler(channeldata_decode);
                ch5_serial.Open();
            }
            if (testCfgPara.serialportEnable[6])
            {
                if (ch6_serial.IsOpen)
                {
                    ch6_serial.Close();
                }
                ch6_serial.DataReceived += new SerialDataReceivedEventHandler(channeldata_decode);
                ch6_serial.Open();
            }
        }
        /*************************************
        函数名：tabledata_decode
        创建日期：2019/10/31
        函数功能：转台数据解码
        函数参数：
        	sender
        	e
        返回值：void
        *************************************/
        private void tabledata_decode(Object sender, SerialDataReceivedEventArgs e)
        {
            int n = table_serial.BytesToRead;
            byte[] readBuffer = new byte[n];
            byte[] buf = new byte[n];
            table_serial.Read(readBuffer,0,n);
            table_serialData.buffer.AddRange(readBuffer);
            UInt32 CheckSumA = 0;
            UInt32 CheckSumB = 0;
            Union udata = new Union();
            while (table_serialData.buffer.Count > 10)//判断缓存总是否保存大于一帧的数据
            {
                if (table_serialData.buffer[0] == 0xAA && table_serialData.buffer[1] == 0xA5 && table_serialData.buffer[2] == 0x55)//判断帧头
                {
                    CheckSumA = 0;
                    CheckSumB = 0;

                    for (int i = 0; i <= 10; i++)
                    {
                        CheckSumA += table_serialData.buffer[i];
                    }
                    CheckSumB = table_serialData.buffer[11];
                    if ((CheckSumA & 0xFF) == CheckSumB)//校验通过开始解码
                    {
                        table_serialData.buffer.CopyTo(0, tabledata.arrayOriginData, 0, 12);
                        udata.b0 = tabledata.arrayOriginData[3];
                        udata.b1 = tabledata.arrayOriginData[4];
                        udata.b2 = tabledata.arrayOriginData[5];
                        udata.b3 = tabledata.arrayOriginData[6];
                        tabledata.table_angle = Convert.ToDouble(udata.i) / 10000.0;
                        udata.b0 = tabledata.arrayOriginData[7];
                        udata.b1 = tabledata.arrayOriginData[8];
                        udata.b2 = tabledata.arrayOriginData[9];
                        udata.b3 = tabledata.arrayOriginData[10];
                        tabledata.table_rate = Convert.ToDouble(udata.i) / 10000.0;
                        tabledata.Counter++;
                        if (tabledata.Counter % 10 == 0)
                        {
                            this.BeginInvoke(updateTableFrmdata);
                        }
                    }
                    else//校验不对，移去一个字节
                    {
                        table_serialData.buffer.RemoveRange(0, 1);
                    }
                }

                else//如果帧头不对，移去一个字节
                {
                    table_serialData.buffer.RemoveRange(0, 1);
                }
            }
        }
        /*************************************
        函数名：showtabledata
        创建日期：2019/11/01
        函数功能：转台模块显示
        函数参数：
        返回值：void
        *************************************/
        private void showtabledata()
        {
            tBox_current_angle.Text = tabledata.table_angle.ToString();
            tBox_current_rate.Text = tabledata.table_rate.ToString();
        }

        /*************************************
        函数名：channeldata_decode
        创建日期：2019/11/02
        函数功能：采集通道数据接收
        函数参数：
        	sender
        	e
        返回值：void
        *************************************/
        private void channeldata_decode(Object sender, SerialDataReceivedEventArgs e)
        {
            SerialPort serialPort = (SerialPort)sender;
            if (serialPort.PortName == "COM1")
            {
                decodeFogData(serialPort, CH1_FogData);
            }
            if (serialPort.PortName == "COM2")
            {
                decodeFogData(serialPort, CH2_FogData);
            }
            if (serialPort.PortName == "COM3")
            {
                decodeFogData(serialPort, CH3_FogData);
            }
            if (serialPort.PortName == "COM4")
            {
                decodeFogData(serialPort, CH4_FogData);
            }
            if (serialPort.PortName == "COM5")
            {
                decodeFogData(serialPort, CH5_FogData);
            }
            if (serialPort.PortName == "COM6")
            {
                decodeFogData(serialPort, CH6_FogData);
            }
           
        }


        /*************************************
        函数名：decodeFogData
        创建日期：2019/11/02
        函数功能：
        函数参数：
            serialPort
            fogdata
        返回值：void
        *************************************/
        private void decodeFogData(SerialPort serialPort, Fogdata fogdata)
        {
            int n = serialPort.BytesToRead;
            byte[] readBuffer = new byte[n];
            byte[] buf = new byte[n];
            serialPort.Read(readBuffer, 0, n);
            fogdata.buffer.AddRange(readBuffer);
            UInt32 CheckSumA = 0;
            UInt32 CheckSumB = 0;
            UInt32 CheckSumC = 0;
            UInt32 CheckSumD = 0;
            while (fogdata.buffer.Count >= 10)
            {
                if (fogdata.buffer[0] == 0x80)
                {
                    CheckSumA = 0;
                    CheckSumB = 0;
                    CheckSumC = 0;
                    CheckSumD = 0;
                    for (int i = 1; i <= 5; i++)
                    {
                        CheckSumA = CheckSumA ^ fogdata.buffer[i];
                    }
                    CheckSumB = fogdata.buffer[6];
                    for (int i = 1; i <= 8; i++)
                    {
                        CheckSumC = CheckSumC ^ fogdata.buffer[i];
                    }
                    CheckSumD = fogdata.buffer[9];
                    if ((CheckSumA & 0x7F) == CheckSumB && (CheckSumC & 0x7F) == CheckSumD)
                    {
                        fogdata.buffer.CopyTo(0, fogdata.arrayRCVData, 0, 10);
                        fogdata.i_fdata = (Convert.ToInt32(fogdata.arrayRCVData[5]) * 128 * 128 * 128 * 128 + Convert.ToInt32(fogdata.arrayRCVData[4]) * 128 * 128 * 128
                                                            + Convert.ToInt32(fogdata.arrayRCVData[3]) * 128 * 128 + Convert.ToInt32(fogdata.arrayRCVData[2]) * 128 + Convert.ToInt32(fogdata.arrayRCVData[1]));
                        fogdata.i_tdata = (fogdata.arrayRCVData[8] * 128 * 128 * 128 * 16 + fogdata.arrayRCVData[7] * 128 * 128 * 16) / (128 * 128 * 16);

                        fogdata.d_fdata = Convert.ToDouble(fogdata.i_fdata);
                        fogdata.d_tdata = Convert.ToDouble(fogdata.i_tdata) / 16.0;

                        fogdata.fdata_array.Add(fogdata.d_fdata);
                        fogdata.tdata_array.Add(fogdata.d_tdata);
                        fogdata.Counter++;
                        //savedata(serialPort.PortName);
                        if (fogdata.Counter % timePara.sampleFreq == 0)
                        {
                            fogdata.d_fdata_1s = fogdata.fdata_array.Average();
                            fogdata.d_tdata_1s = fogdata.tdata_array.Average();
          

                            fogdata.fdata_1s_array.Add(fogdata.d_fdata_1s);
                            fogdata.tdata_1s_array.Add(fogdata.d_tdata_1s);
                            fogdata.ave_Fog_data = fogdata.fdata_1s_array.Average();
                            fogdata.std_Fog_data = CalculateStdDev(fogdata.fdata_1s_array);
                            fogdata.fdata_array.Clear();
                            fogdata.tdata_array.Clear();
                            this.BeginInvoke(updateDataFrm, serialPort.PortName);
                        }
                        fogdata.buffer.RemoveRange(0, 10);
                    }
                    else
                    {
                        fogdata.buffer.RemoveRange(0, 1);
                    }
                }
                else
                {
                    fogdata.buffer.RemoveRange(0, 1);
                }
            }

        }
        /*************************************
        函数名：showFogdata
        创建日期：2019/11/02
        函数功能：显示数据
        函数参数：
        portName 串口号
        返回值：void
        *************************************/
        private void showFogdata(string portName)
        {
            switch(portName)
            {
                case "COM1":
                    {
                        tBox_ch1_currentdata.Text = CH1_FogData.d_fdata_1s.ToString();
                        tBox_ch1_Caltdata.Text = CH1_FogData.d_fdata_1s.ToString();
                        tBox_ch1_stddata.Text = CH1_FogData.d_fdata_1s.ToString();
                        tBox_ch1_temdata.Text = CH1_FogData.d_tdata_1s.ToString();
                        break;
                    }
            }
        }
        /*************************************
        函数名：CalculateStdDev
        创建日期：2019/11/02
        函数功能：计算数组标准差 std = sqrt(sum((value(i)-ave(value))^2))/(N-1)
        函数参数：value
        返回值：double 标准差结果
        *************************************/
        private double CalculateStdDev(List<double> value)
        {
            double std_data = 0.0;
            if (value.Count > 1)
            {
                double ave_data = value.Average();
                double sum_data = value.Sum(data => Math.Pow((data - ave_data), 2));
                std_data = Math.Sqrt(sum_data / (value.Count - 1));

            }


            return std_data;
        }
    }
 }
