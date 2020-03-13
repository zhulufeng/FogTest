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
        List<SerialPort> channels_serial_list = new List<SerialPort>();
        SerialData table_serialData = new SerialData();
        //保存文件流声明
        List<StreamWriter> Channels_Hex_SW_list = new List<StreamWriter>();
        List<StreamWriter> Channels_Data_SW_list = new List<StreamWriter>();
        List<StreamWriter> Channels_SFData_SW_list = new List<StreamWriter>();
        List<string> Channels_portName_list = new List<string>();
        List<Fogdata> Channels_FogData_list = new List<Fogdata>();
        List<StreamReader> Channel_DataFile_SR_list = new List<StreamReader>();
        


        //参数类对象声明
        TestCfgPara testCfgPara = new TestCfgPara();//测试配置参数
        TableData tabledata = new TableData();
        TimePara timePara = new TimePara();
        List<string> portIDList = new List<string>();
        List<string> ChannelList = new List<string>();
        ScaleFactorPara scaleFactorPara = new ScaleFactorPara();
        bool isScaleFactorTest = false;
        bool ScaleFactorTestStart = false;
        //定义委托
        delegate void UpdateTableFrmEventHandle(bool isinfotBox,string text);
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
            IntializeChart();
            timePara.testTimes = 0;
            timePara.drawCount = 0;
            timePara.sampleTime = 500;
            timePara.switchRateTime = 1000;
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
            FilePara.CurrentDirectory = FilePara.BaseDirectory + @"FogData" + timedata;
            Directory.CreateDirectory(FilePara.CurrentDirectory);
            try
            {
                FilePara.ConfigFilePath = FilePara.CurrentDirectory + @"\配置文件.xlsx";
                FileStream file = new FileStream(FilePara.ConfigFilePath, FileMode.Create);               
                XSSFWorkbook hsswfworkbook = new XSSFWorkbook();
                ISheet sheet = hsswfworkbook.CreateSheet("通道串口配置");
                hsswfworkbook.CreateSheet("标度因数参数");
                sheet.CreateRow(0).CreateCell(1).SetCellValue("使能");
                sheet.GetRow(0).CreateCell(2).SetCellValue("串口号");
                sheet.GetRow(0).CreateCell(3).SetCellValue("波特率");
                sheet.GetRow(0).CreateCell(4).SetCellValue("数据位");
                sheet.GetRow(0).CreateCell(5).SetCellValue("停止位");
                sheet.GetRow(0).CreateCell(6).SetCellValue("校验位");
                sheet.GetRow(0).CreateCell(7).SetCellValue("型号");
                sheet.GetRow(0).CreateCell(8).SetCellValue("标度因数");

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

                MessageBox.Show("配置文件产生异常！");
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
                series.Color = Color.FromArgb(0xff, 0x32, 0xc5, 0xe9);//设置数据曲线的颜色
            }
            else
            {
                series.Color = Color.FromArgb(0xff, 0xff, 0x9f, 0x7f);//设置温度曲线的颜色
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

            chartArea.Name = Channels_FogData_list[index].FOG_Channel;
            
            //背景色
            chartArea.BackColor         = Color.FromArgb(255, 4, 33, 65);
            //背景渐变方式
            chartArea.BackGradientStyle = GradientStyle.None;
            //边框颜色
            chartArea.BorderColor       = Color.FromArgb(255, 4, 33, 65);
            //边框柱线条宽度
            chartArea.BorderWidth       = 2;
            //边框线条样式
            chartArea.BorderDashStyle   = ChartDashStyle.Solid;
            //阴影颜色
            chartArea.ShadowColor       = Color.Transparent;


            //设置X轴和Y轴线条的颜色和宽度
            chartArea.AxisX.LineColor = Color.Black;//.FromArgb(64, 64, 64, 64);//
            chartArea.AxisX.LineWidth = 1;
            chartArea.AxisY.LineColor = Color.Black;//.FromArgb(64, 64, 64, 64);//
            chartArea.AxisY.LineWidth = 1;
            //设置x轴和Y轴的标题
            chartArea.AxisX.Title           = "时间";
            chartArea.AxisY.Title           = Channels_FogData_list[index].FOG_Channel + "_" + Channels_FogData_list[index].FOGID + "_陀螺数据";
            chartArea.AxisY2.Title          = "温度";
            chartArea.AxisX.TitleFont       = new System.Drawing.Font("Microsoft Sans Serif", 12, FontStyle.Regular);
            chartArea.AxisY.TitleFont       = new System.Drawing.Font("Microsoft Sans Serif", 10, FontStyle.Regular);
            chartArea.AxisY2.TitleFont      = new System.Drawing.Font("Microsoft Sans Serif", 12, FontStyle.Regular);
            chartArea.AxisX.TitleForeColor  = Color.FromArgb(255, 245, 254, 252);
            chartArea.AxisY.TitleForeColor  = Color.FromArgb(0xff, 0x32, 0xc5, 0xe9);
            chartArea.AxisY2.TitleForeColor = Color.FromArgb(0xff, 0xff, 0x9f, 0x7f);
            //设置图表区网格横纵线条的颜色和宽度
            chartArea.AxisX.MajorGrid.LineColor = Color.FromArgb(255, 114, 175, 207);
            chartArea.AxisX.MajorGrid.LineWidth = 1;
            chartArea.AxisY.MajorGrid.LineColor = Color.FromArgb(64, 64, 64, 64);
            chartArea.AxisY.MajorGrid.LineWidth = 1;

            //启用X游标，以支持局部区域选择放大
            chartArea.CursorX.IsUserEnabled          = true;
            chartArea.CursorX.IsUserSelectionEnabled = true;
            chartArea.CursorX.LineColor              = Color.Pink;
            chartArea.CursorX.IntervalType           = DateTimeIntervalType.Auto;
            chartArea.AxisX.ScaleView.Zoomable       = false;
            chartArea.AxisX.ScrollBar.ButtonStyle    = ScrollBarButtonStyles.All;//启用X轴滚动条按钮
            chartArea.AxisY.ScaleView.Zoomable       = false;

            chartArea.AxisY.LabelStyle.Format        = "##########.0";
            chartArea.AxisY2.LabelStyle.Format       = "###.0000";
            chartArea.AxisY.LabelStyle.ForeColor     = Color.FromArgb(255, 146, 175, 207);
            chartArea.AxisY2.LabelStyle.ForeColor    = Color.FromArgb(255, 146, 175, 207);
            chartArea.AxisX.LabelStyle.ForeColor     = Color.FromArgb(255, 151, 167, 186);

            return chartArea;
        }
        /*************************************
        函数名：DrawFogData
        创建日期：2019/11/08
        函数功能：实时采集画图
        函数参数：
        portName
        返回值：void
        *************************************/
        private void DrawFogData(string portName)
        {
            timePara.drawCount++;
            if (timePara.drawCount >= portIDList.Count)
            {
                int index = portIDList.IndexOf(portName);
                timePara.drawIndexTime[index]++;
                if (!Channels_FogData_list[index].zoomed_flag)
                {
                    chart.ChartAreas[index].AxisY.Maximum  = Channels_FogData_list[index].fdata_1s_array.Max() + 100;
                    chart.ChartAreas[index].AxisY.Minimum  = Channels_FogData_list[index].fdata_1s_array.Min() - 100;
                    chart.ChartAreas[index].AxisY2.Maximum = Channels_FogData_list[index].tdata_1s_array.Max() + 1;
                    chart.ChartAreas[index].AxisY2.Minimum = Channels_FogData_list[index].tdata_1s_array.Min() - 1;

                    chart.ChartAreas[index].AxisX.Interval           = (Channels_FogData_list[index].fdata_1s_array.Count / 20 + 1);
                    chart.ChartAreas[index].AxisX.ScaleView.Size     = Channels_FogData_list[index].fdata_1s_array.Count * 1.1;
                    chart.ChartAreas[index].AxisX.ScaleView.Position = 0.0;
                    chart.ChartAreas[index].CursorX.SelectionStart   = chart.ChartAreas[index].CursorX.SelectionEnd = 0.0;
                    chart.ChartAreas[index].CursorX.Position         = -1;

                }
                chart.Series[2 * index].ChartArea = chart.ChartAreas[index].Name;
                chart.Series[2 * index + 1].ChartArea = chart.ChartAreas[index].Name;
                chart.Series[2 * index].Points.AddXY(timePara.drawIndexTime[index], Channels_FogData_list[index].d_fdata_1s);
                chart.Series[2 * index+1].Points.AddXY(timePara.drawIndexTime[index], Channels_FogData_list[index].d_tdata_1s);
            }
        }

        private void DrawLoadFogdata(int index)
        {
//             chart.Series[2 * index].Points.Clear();
//             chart.Series[2 * index + 1].Points.Clear(); 
            if (!Channels_FogData_list[index].zoomed_flag)
            {
                chart.ChartAreas[index].AxisY.Maximum            = Channels_FogData_list[index].fdata_smooth_array.Max() + 100;
                chart.ChartAreas[index].AxisY.Minimum            = Channels_FogData_list[index].fdata_smooth_array.Min() - 100;
                chart.ChartAreas[index].AxisY2.Maximum           = Channels_FogData_list[index].tdata_smooth_array.Max() + 1;
                chart.ChartAreas[index].AxisY2.Minimum           = Channels_FogData_list[index].tdata_smooth_array.Min() - 1;

                chart.ChartAreas[index].AxisX.Interval           = (Channels_FogData_list[index].time_smooth_array.Max() / 20 + 1);
                chart.ChartAreas[index].AxisX.ScaleView.Size     = Channels_FogData_list[index].time_smooth_array.Max();
                chart.ChartAreas[index].AxisX.Minimum            = Channels_FogData_list[index].time_smooth_array.Min();
                chart.ChartAreas[index].AxisX.Maximum            = Channels_FogData_list[index].time_smooth_array.Max();
                chart.ChartAreas[index].AxisX.ScaleView.Position = 0.0;
                chart.ChartAreas[index].CursorX.SelectionStart   = chart.ChartAreas[index].CursorX.SelectionEnd = 0.0;
                chart.ChartAreas[index].CursorX.Position         = -1;

            }
            chart.Series[2 * index].ChartArea = chart.ChartAreas[index].Name;
            chart.Series[2 * index + 1].ChartArea = chart.ChartAreas[index].Name;
            chart.Series[2 * index].Points.DataBindXY(Channels_FogData_list[index].time_smooth_array, Channels_FogData_list[index].fdata_smooth_array);
            chart.Series[2 * index + 1].Points.DataBindXY(Channels_FogData_list[index].time_smooth_array, Channels_FogData_list[index].tdata_smooth_array);
        
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
            //判断是否有串口，没有则退出
            string[] ArrayPort = SerialPort.GetPortNames();//搜索
            if (ArrayPort.Length == 0)
            {
                MessageBox.Show("没有可用的串口！");
                return;
            }
            InitializeConfigFlie();
            SerialCfgDlg serialCfgDlg = new SerialCfgDlg();
            if (serialCfgDlg.ShowDialog() != DialogResult.Cancel)
            {
                ConfigSerialPort();
            }
            testCfgPara.isLoadData = false;
        }
        
        /*************************************
        函数名：ConfigSerialPort
        创建日期：2019/11/06
        函数功能：通过配置文件
        函数参数：
        返回值：void
        *************************************/
        private void ConfigSerialPort()
        {
            //读入配置文件
            FileStream rfile = new FileStream(FilePara.ConfigFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            XSSFWorkbook workbook = new XSSFWorkbook(rfile);
            rfile.Close();
            portIDList.Clear();
            ISheet sht = workbook.GetSheet("通道串口配置");
            //配置转台参数
            if (sht.GetRow(1).GetCell(1).ToString() == "True")
            {
                table_serial = SetSerialPara(0, sht);
                testCfgPara.serialportEnable[0] = true;
               // portIDList.Add(sht.GetRow(1).GetCell(2).ToString());
            }
            else
            {
                testCfgPara.serialportEnable[0] = false;
            }
            //先清空通道参数
            channels_serial_list.Clear();
            portIDList.Clear();
            Channels_FogData_list.Clear();
            //配置各通道参数
            for (int i = 1; i <= 6; i++)
            {
                if (sht.GetRow(i + 1).GetCell(1).ToString() == "True")
                {
                    //MessageBox.Show(sht.GetRow(i + 1).GetCell(2).ToString());
                    if (sht.GetRow(i + 1).GetCell(2).ToString() == "null")
                    {
                        MessageBox.Show("有串口没有串口名，请重新配置串口！");
                        Btn_Start.Enabled = false;
                        return;
                    }
                    SetChannelPara(i, sht, portIDList);
                }
                else
                {
                    testCfgPara.serialportEnable[i] = false;
                }
            }

            HashSet<string> PortIDHashset = new HashSet<string>(portIDList);
            if (portIDList.Count() != PortIDHashset.Count())
            {
                MessageBox.Show("不同通道选用了相同的串口号，请重新配置串口！");
                Btn_Start.Enabled = false;

            }
            else
            {
                chart.ChartAreas.Clear();
                chart.Series.Clear();
                testCfgPara.numOftestChannels = Convert.ToInt32(sht.GetRow(8).GetCell(1).ToString());
                AddChartArea(testCfgPara.numOftestChannels);
                AddSeries(testCfgPara.numOftestChannels);
                Btn_Start.Enabled = true;
                Btn_Stop.Enabled = false;
            }
        }
        /*************************************
        函数名：SetChannelPara
        创建日期：2019/11/06
        函数功能：配置通道参数
        函数参数：
        channelID
        sht
        portIDList
        返回值：void
        *************************************/
        private void SetChannelPara(int channelID, ISheet sht, List<string> portIDList)
        {
            Fogdata fogdata = new Fogdata();
            channels_serial_list.Add(SetSerialPara(channelID, sht));
            testCfgPara.serialportEnable[channelID] = true;
            portIDList.Add(sht.GetRow(channelID + 1).GetCell(2).ToString());
            fogdata.FOGID = sht.GetRow(channelID + 1).GetCell(7).ToString();
            fogdata.FOG_Channel = sht.GetRow(channelID + 1).GetCell(0).ToString();
            fogdata.scaleFactor = Convert.ToDouble(sht.GetRow(channelID + 1).GetCell(8).ToString());
            fogdata.Fog_PortName = sht.GetRow(channelID + 1).GetCell(2).ToString();
            Channels_FogData_list.Add(fogdata);
        }
        /*************************************
        函数名：SetSerialPara
        创建日期：2019/11/02
        函数功能：设置串口参数
        函数参数：
        index
        返回值：System.IO.Ports.SerialPort
        *************************************/
        public SerialPort SetSerialPara(int index,ISheet sht)
        {
            SerialPort serial          = new SerialPort();
            SerialParameter serialpara = new SerialParameter();
            
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
                Send_table_Connect();
            }
            Channels_Hex_SW_list.Clear();
            Channels_Data_SW_list.Clear();
            Channels_SFData_SW_list.Clear();
            timePara.drawIndexTime.Clear();
            foreach (var item in chart.Series)
            {
                item.Points.Clear();
            }
            foreach (var item in Channels_FogData_list)
            {
                item.tdata_1s_array.Clear();
                item.fdata_1s_array.Clear();
                item.fdata_array.Clear();
                item.tdata_array.Clear();

            }
            foreach (var item in channels_serial_list)
            {
                int index = portIDList.IndexOf(item.PortName);
                Channels_Hex_SW_list.Add(new StreamWriter(FilePara.CurrentDirectory + @"\" + Channels_FogData_list[index].FOGID + "_HexData_" + timePara.testTimes.ToString() + ".hex"));
                Channels_Data_SW_list.Add(new StreamWriter(FilePara.CurrentDirectory + @"\" + Channels_FogData_list[index].FOGID + "_Data_" + timePara.testTimes.ToString() + ".dat"));
                if (isScaleFactorTest)
                {
                    Channels_SFData_SW_list.Add(new StreamWriter(FilePara.CurrentDirectory + @"\" + Channels_FogData_list[index].FOGID + "_SFData_" + timePara.testTimes.ToString() + ".dat"));
                }
                timePara.drawIndexTime.Add(0);
                if (item.IsOpen)
                {
                    item.Close();
                }
                item.DataReceived += new SerialDataReceivedEventHandler(channeldata_decode);
                item.Open();//需要判断串口是否有
            }
            ScaleFactorTestStart = true;
            Btn_Start.Enabled = false;
            Btn_Stop.Enabled  = true;
            timePara.testTimes++;
            
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
            while (table_serialData.buffer.Count >= 13)//判断缓存总是否保存大于一帧的数据
            {
                if (table_serialData.buffer[0] == 0xAA && table_serialData.buffer[1] == 0xA5 && table_serialData.buffer[2] == 0x55)//判断帧头
                {
                    CheckSumA = 0;
                    CheckSumB = 0;

                    for (int i = 0; i <= 11; i++)
                    {
                        CheckSumA += table_serialData.buffer[i];
                    }
                    CheckSumB = table_serialData.buffer[12];
                    if ((CheckSumA & 0xFF) == CheckSumB)//校验通过开始解码
                    {
                        table_serialData.buffer.CopyTo(0, tabledata.arrayOriginData, 0, 13);
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
                        if (tabledata.Counter % 100 == 0)
                        {
                            this.BeginInvoke(updateTableFrmdata,false,"");
                        }
                        if(ScaleFactorTestStart)
                        {
                            if (tabledata.SF_Counter == (tabledata.SF_Para_index + 1) * timePara.sampleTime + (tabledata.SF_Para_index)*timePara.switchRateTime && tabledata.SF_Para_index < scaleFactorPara.paracount)//paracount  转速合计
                            {
                                int drate;
                                if (Math.Abs(scaleFactorPara.RatePara[tabledata.SF_Para_index]) < 50)
                                {
                                    drate = 25;
                                   // timePara.switchRateTime = 500;
                                }

                                else if (Math.Abs(scaleFactorPara.RatePara[tabledata.SF_Para_index]) >= 50 && Math.Abs(scaleFactorPara.RatePara[tabledata.SF_Para_index]) < 500)
                                {
                                    drate = 100;
                                    //timePara.switchRateTime = 500;
                                }
                                else
                                {
                                    drate = 150;
                                   // timePara.switchRateTime = 2000;

                                }
//                                 if (tabledata.SF_Para_index == scaleFactorPara.paracount - 1)
//                                 {
//                                     drate = 80;
//                                    // timePara.switchRateTime = 500;
//                                 }
                                  
                                //Send_table_rateCommand(scaleFactorPara.RatePara[tabledata.SF_Para_index], scaleFactorPara.RatePara[tabledata.SF_Para_index] / (timePara.switchRateTime / 10));
                                Send_table_rateCommand(scaleFactorPara.RatePara[tabledata.SF_Para_index], drate);
                                
                                tabledata.SF_Para_index++;
                                
                            }

                            tabledata.SF_Counter++;

                        }
                        table_serialData.buffer.RemoveRange(0, 13);
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
        private void showtabledata(bool isinfoTbox, string text="")
        {
            tBox_current_angle.Text = tabledata.table_angle.ToString();
            tBox_current_rate.Text = tabledata.table_rate.ToString();
            if (isinfoTbox)
            {
                tBox_info.AppendText(text);
                tBox_info.ScrollToCaret();
                //tBox_info.Text += text;
            }
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
            int index = portIDList.IndexOf(serialPort.PortName);

            decodeFogData(serialPort, Channels_FogData_list[index]);
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
                        SaveChannledata(serialPort.PortName);
                        if (fogdata.Counter % timePara.sampleFreq == 0)
                        {
                            fogdata.d_fdata_1s = fogdata.fdata_array.Average();
                            fogdata.d_tdata_1s = fogdata.tdata_array.Average();
                            fogdata.time_array.Add(Convert.ToDouble(fogdata.Counter) / timePara.sampleFreq);
                            fogdata.fdata_1s_array.Add(fogdata.d_fdata_1s);
                            fogdata.tdata_1s_array.Add(fogdata.d_tdata_1s);
                            fogdata.fdata_smooth_array = fogdata.fdata_1s_array;
                            fogdata.tdata_smooth_array = fogdata.tdata_1s_array;
                            fogdata.time_smooth_array.Add(fogdata.Counter / timePara.sampleFreq);
                            fogdata.ave_Fog_data = fogdata.fdata_1s_array.Average();
                            fogdata.std_Fog_data = CalculateStdDev(fogdata.fdata_1s_array);
                            fogdata.Fog_Bias_std = fogdata.std_Fog_data / fogdata.scaleFactor *3600;
                            fogdata.Fog_Comped_data = fogdata.d_fdata_1s / fogdata.scaleFactor * 3600;
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
        private void showFogdata(string PortName)
        {
            //port_Dic[PortName];
            int index = portIDList.IndexOf(PortName);


            switch(Channels_FogData_list[index].FOG_Channel)
            {
                case "通道一":
                    {
                        tBox_ch1_currentdata.Text = Channels_FogData_list[index].d_fdata_1s.ToString();
                        tBox_ch1_Caltdata.Text    = Channels_FogData_list[index].Fog_Comped_data.ToString();
                        tBox_ch1_stddata.Text     = Channels_FogData_list[index].Fog_Bias_std.ToString();
                        tBox_ch1_temdata.Text     = Channels_FogData_list[index].d_tdata_1s.ToString();
                        break;
                    }
                case "通道二":
                    {
                        tBox_ch2_currentdata.Text = Channels_FogData_list[index].d_fdata_1s.ToString();
                        tBox_ch2_Caltdata.Text    = Channels_FogData_list[index].Fog_Comped_data.ToString();
                        tBox_ch2_stddata.Text     = Channels_FogData_list[index].Fog_Bias_std.ToString();
                        tBox_ch2_temdata.Text     = Channels_FogData_list[index].d_tdata_1s.ToString();
                        break;
                    }
                case "通道三":
                    {
                        tBox_ch3_currentdata.Text = Channels_FogData_list[index].d_fdata_1s.ToString();
                        tBox_ch3_Caltdata.Text    = Channels_FogData_list[index].Fog_Comped_data.ToString();
                        tBox_ch3_stddata.Text     = Channels_FogData_list[index].Fog_Bias_std.ToString();
                        tBox_ch3_temdata.Text     = Channels_FogData_list[index].d_tdata_1s.ToString();
                        break;
                    }
                case "通道四":
                    {
                        tBox_ch4_currentdata.Text = Channels_FogData_list[index].d_fdata_1s.ToString();
                        tBox_ch4_Caltdata.Text    = Channels_FogData_list[index].Fog_Comped_data.ToString();
                        tBox_ch4_stddata.Text     = Channels_FogData_list[index].Fog_Bias_std.ToString();
                        tBox_ch4_temdata.Text     = Channels_FogData_list[index].d_tdata_1s.ToString();
                        break;
                    }
                case "通道五":
                    {
                        tBox_ch5_currentdata.Text = Channels_FogData_list[index].d_fdata_1s.ToString();
                        tBox_ch5_Caltdata.Text    = Channels_FogData_list[index].Fog_Comped_data.ToString();
                        tBox_ch5_stddata.Text     = Channels_FogData_list[index].Fog_Bias_std.ToString();
                        tBox_ch5_temdata.Text     = Channels_FogData_list[index].d_tdata_1s.ToString();
                        break;
                    }
                case "通道六":
                    {
                        tBox_ch6_currentdata.Text = Channels_FogData_list[index].d_fdata_1s.ToString();
                        tBox_ch6_Caltdata.Text    = Channels_FogData_list[index].Fog_Comped_data.ToString();
                        tBox_ch6_stddata.Text     = Channels_FogData_list[index].Fog_Bias_std.ToString();
                        tBox_ch6_temdata.Text     = Channels_FogData_list[index].d_tdata_1s.ToString();
                        break;
                    }
            }
            DrawFogData(PortName);
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

        /*************************************
        函数名：SaveChannledata
        创建日期：2019/11/04
        函数功能：保存通道参数
        函数参数：
        PortName
        返回值：void
        *************************************/
        private void SaveChannledata(string PortName)
        {
            int index = portIDList.IndexOf(PortName);
            for (int i = 0; i < Channels_FogData_list[index].arrayRCVData.Length; i++)
            {
                Channels_Hex_SW_list[index].Write(Channels_FogData_list[index].arrayRCVData[i].ToString("X2") + "\t");
            }
            Channels_Hex_SW_list[index].Write("\n");
            StringBuilder sb = new StringBuilder();
            sb.AppendFormat("{0:0000000}", Convert.ToDouble(Channels_FogData_list[index].Counter) / timePara.sampleFreq);
            sb.AppendFormat("\t{0:00000.00}", Channels_FogData_list[index].d_fdata);
            sb.AppendFormat("\t{0:000.000}", Channels_FogData_list[index].d_tdata);
            Channels_Data_SW_list[index].WriteLine(sb.ToString());
            sb.Clear();
            if(ScaleFactorTestStart && tabledata.SF_Counter >= 550)//tabledata.SF_Counter 转台一个数加1  10Hz
            {
                if (tabledata.SF_Counter >= (tabledata.SF_Para_index - 1) * timePara.sampleTime + (tabledata.SF_Para_index)*1.5)
                {
                    sb.AppendFormat("{0:0000000}", Convert.ToDouble(Channels_FogData_list[index].Counter) / timePara.sampleFreq);
                    sb.AppendFormat("\t{0:00000.00}", Channels_FogData_list[index].d_fdata);
                    sb.AppendFormat("\t{0:00000.00}", tabledata.table_rate);
                    sb.AppendFormat("\t{0:000}", tabledata.SF_Para_index );
                }
                Channels_SFData_SW_list[index].WriteLine(sb.ToString());
                sb.Clear();
            }
            
        }

        private void Btn_Stop_Click(object sender, EventArgs e)
        {
            foreach (var item in channels_serial_list)
            {
                item.Close();
            }
            foreach (var item in Channels_Hex_SW_list)
            {
                item.Close();
            }
            foreach (var item in Channels_Data_SW_list)
            {
                item.Close();
            }
            
            Btn_Start.Enabled = true;
            Btn_Stop.Enabled  = false;
        }

        private void chart_SelectionRangeChanged(object sender, CursorEventArgs e)
        {
            int index = 0;
            List<double> lst = new List<double>();
            //遍历陀螺对象，根据通道号确定索引号
            foreach (var item in Channels_FogData_list)
            {
                string str = item.FOG_Channel;
                if (str == e.ChartArea.Name)
                {
                    if (testCfgPara.isLoadData)
                    {
                        index = ChannelList.IndexOf(item.FOG_Channel);
                    }
                    else
                    {
                        index = portIDList.IndexOf(item.Fog_PortName);
                    }
                    
                }
                
            }
            if (chart.Series[2*index].Points.Count == 0 || e.NewSelectionEnd == e.NewSelectionStart)
            {
                return;
            }
            //确定缩放的起始点
            double startPosition = Math.Min(e.NewSelectionStart, e.NewSelectionEnd);
            double endPosition   = Math.Max(e.NewSelectionStart, e.NewSelectionEnd);
            double myInterval    = endPosition - startPosition;
            chart.ChartAreas[index].AxisX.ScaleView.Zoom(startPosition, endPosition);
            chart.ChartAreas[index].AxisX.ScaleView.Position = startPosition;
            chart.ChartAreas[index].AxisX.ScaleView.Size     = myInterval;
            if (myInterval < 11.0)
            {
                chart.ChartAreas[index].AxisX.Interval = 1;
            }
            else
            {
                chart.ChartAreas[index].AxisX.Interval = Math.Floor(myInterval / 10.0);
            }
            for (int i = Convert.ToInt32(startPosition); i <= Convert.ToInt32(endPosition); i++)
            {
                if (endPosition > chart.Series[index * 2].Points.Count)
                {
                    break;
                }
                lst.Add(chart.Series[index * 2].Points[i].YValues[0]);
            }
            double std = CalculateStdDev(lst);
            tBox_info.Text += Channels_FogData_list[index].FOG_Channel + "_" + Channels_FogData_list[index].FOGID + "选择区间零偏稳定性为：" + (std / Channels_FogData_list[index].scaleFactor * 3600).ToString() + "\r\n";
            Channels_FogData_list[index].zoomed_flag = true;
        }

        /*************************************
        函数名：chart_MouseClick
        创建日期：2020/01/02
        函数功能：鼠标右键恢复图
        函数参数：
        sender
        e
        返回值：void
        *************************************/
        private void chart_MouseClick(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                if(testCfgPara.isLoadData)
                {
                    for (int i = 0; i < ChannelList.Count; i++)
                    {
                        if (Channels_FogData_list[i].zoomed_flag)
                        {
                            Channels_FogData_list[i].zoomed_flag = false;
                            DrawLoadFogdata(i);
                        }
                        

                    }
                }
                else
                {
                    for (int i = 0; i < portIDList.Count; i++)
                    {
                        Channels_FogData_list[i].zoomed_flag = false;
                        DrawLoadFogdata(i);
                    }
                }
                
                
            }
        }

        /*************************************
        函数名：ToolStripMenuItem_SerialCfgByFile_Click
        创建日期：2019/11/06
        函数功能：
        函数参数：
        sender
        e
        返回值：void
        *************************************/
        private void ToolStripMenuItem_SerialCfgByFile_Click(object sender, EventArgs e)
        {
            //判断是否有串口，没有则退出
            string[] ArrayPort = SerialPort.GetPortNames();//搜索
            if (ArrayPort.Length == 0)
            {
                MessageBox.Show("没有可用的串口！");
                return;
            }
            InitializeConfigFlie();
            OpenFileDialog ConfigFileLoadDlg = new OpenFileDialog();
            ConfigFileLoadDlg.InitialDirectory = System.AppDomain.CurrentDomain.BaseDirectory;
            ConfigFileLoadDlg.DefaultExt = "xlsx";
            ConfigFileLoadDlg.Filter = "Excel File(.xlsx)|*.xlsx";
            if (ConfigFileLoadDlg.ShowDialog() == DialogResult.OK)
            {
                FilePara.ConfigFileLoadPath = ConfigFileLoadDlg.FileName;
            }
            //读取现有配置文件
            FileStream rfile = new FileStream(FilePara.ConfigFileLoadPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            XSSFWorkbook workbook = new XSSFWorkbook(rfile);
            rfile.Close();

            SerialCfgDlg serialCfgDlg = new SerialCfgDlg(workbook);//配置文件导入
            if (serialCfgDlg.ShowDialog() == DialogResult.OK)
            {
                ConfigSerialPort();
            }
            //标度因数试验参数设置
            DialogResult dr;
            ISheet shtCfg = workbook.GetSheet("通道串口配置");
            dr = MessageBox.Show("需要导入标度因数测试的参数吗？", "确认对话框", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2);
            if (dr == DialogResult.Yes)
            {
                while (!testCfgPara.serialportEnable[0])//判断转台串口是否已经配置好   如果没有配置好重新配置//shtCfg.GetRow(1).GetCell(1).ToString() != "True"
                {
                    MessageBox.Show("转台串口没有配置，请重新配置！");
                    if (serialCfgDlg.ShowDialog() == DialogResult.OK)
                    {
                        ConfigSerialPort();
                    }
                    
                }
                rfile = new FileStream(FilePara.ConfigFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                workbook = new XSSFWorkbook(rfile);
                rfile.Close();
                ISheet sht = workbook.GetSheet("标度因数参数");
                //配置转台通道
                IRow row = sht.GetRow(1);
                scaleFactorPara.paracount = Convert.ToInt32(GetCell(sht, 2, 1).ToString());
                int rateParaIndex = 0;
                if (GetCell(sht, 0, 1).ToString() != "")
                {
                    scaleFactorPara.rateParaID = Convert.ToInt32(GetCell(sht, 0, 1).ToString());
                }
                
                if (GetCell(sht, 2 * scaleFactorPara.rateParaID, 3).ToString() != "")
                {
                    timePara.sampleTime = Convert.ToInt32(GetCell(sht, 2, 3).ToString()) * 100;
                }

                if (GetCell(sht, 2 * scaleFactorPara.rateParaID, 5).ToString() != "")
                {
                    timePara.switchRateTime = Convert.ToInt32(GetCell(sht, 2, 5).ToString()) * 100;
                }
                while (GetCell(sht, 3, rateParaIndex).ToString().ToLower() != "end")
                {
                    scaleFactorPara.RatePara.Add(Convert.ToDouble(GetCell(sht, 3, rateParaIndex).ToString()));
                    rateParaIndex++;
                }
                if (Math.Abs(scaleFactorPara.RatePara.Sum()) <= 0.005)
                {
                    tBox_info.Text += "标度因数试验数据导入成功！\r\n";
                    tBox_info.Text += "数据编号是：" + scaleFactorPara.rateParaID.ToString() + "\t" + "旋转时间：" + (timePara.sampleTime / 100).ToString() + "秒\t" + "变速时间是：" + (timePara.switchRateTime / 100).ToString() + "秒" + "\r\n";
                    tBox_info.Text += "转速为：\r\n";
                    for (int i = 0; i < scaleFactorPara.RatePara.Count; i++)
                    {
                        tBox_info.Text += scaleFactorPara.RatePara[i].ToString() + ",";
                    }
                    tBox_info.Text += "\r\n";
                }
                else
                {
                    tBox_info.Text += "转速为：\r\n";
                    for (int i = 0; i < scaleFactorPara.RatePara.Count; i++)
                    {
                        tBox_info.Text += scaleFactorPara.RatePara[i].ToString() + ",";
                    }
                    tBox_info.Text += "\r\n";
                    DialogResult rateDr;
                    rateDr = MessageBox.Show("转速好像有点不对称,是否继续？", "确认对话框", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2);
                    if (dr == DialogResult.No)
                    {
                        MessageBox.Show("请确认配置文件后，重新导入！");
                        return;
                    }
                    else
                    {
                        tBox_info.Text += "标度因数试验数据导入成功！\r\n";
                        tBox_info.Text += "数据编号是：" + scaleFactorPara.rateParaID.ToString() + "\t" + "旋转时间：" + (timePara.sampleTime / 100).ToString() + "秒\t" + "变速时间是：" + (timePara.switchRateTime / 100).ToString() + "秒\r\n";
                        tBox_info.Text += "转速为：\r\n";
                        for (int i = 0; i < scaleFactorPara.RatePara.Count; i++)
                        {
                            tBox_info.Text += scaleFactorPara.RatePara[i].ToString() + ",";
                        }
                        tBox_info.Text += "\r\n";
                    }
                    //tBox_info.Text += "转速好像有点不对称？\n";
                }
                //                 for (int i = 0; i < scaleFactorPara.paracount; i++)
                //                 {
                //                     scaleFactorPara.RatePara.Add(Convert.ToDouble(GetCell(sht, 3, i).ToString()));
                //                 }
                isScaleFactorTest = true;
            }
            else
            {
                return;
            }

        }

        private void ToolStripMenuItem_LoadDataFiles_Click(object sender, EventArgs e)
        {
            LoadDataFileDialog loadDataFileDialog = new LoadDataFileDialog();
            string[] dataSplited;//分割后字符串数组

            char[] rnSplitChar = new char[] { '\r', '\n', '_' };//分割符号
            char[] trnSplitChar = new char[] { '\r', '\n', '\t', ' ','\\' };
            testCfgPara.isLoadData = true;
            Channels_FogData_list.Clear();
            chart.ChartAreas.Clear();
            chart.Series.Clear();
            FilePara.dataFileList.Clear();
            Channel_DataFile_SR_list.Clear();
            ChannelList.Clear();
            if (loadDataFileDialog.ShowDialog() == DialogResult.OK)
            {
                if (FilePara.dataFileList.Count() > 0)
                {
                    foreach (var item in FilePara.dataFileList)
                    {
                        Fogdata fogdata = new Fogdata();
                        string filenameStr;
                        string dataLine;
                        StreamReader SR = new StreamReader(item);
                        int time = 0;
                        Channel_DataFile_SR_list.Add(SR);
                        dataSplited = item.Split(trnSplitChar, StringSplitOptions.RemoveEmptyEntries);
                        filenameStr = dataSplited.Last();
                        dataSplited = filenameStr.Split(rnSplitChar, StringSplitOptions.RemoveEmptyEntries);
                        fogdata.FOGID = dataSplited[dataSplited.Length - 1 - 2];
                        fogdata.FOG_Channel = "CH" + (FilePara.dataFileList.IndexOf(item) + 1);
                        fogdata.scaleFactor = loadDataFileDialog.scaleFactorList[FilePara.dataFileList.IndexOf(item)];
                        ChannelList.Add(fogdata.FOG_Channel);
                        
                        while (!SR.EndOfStream)
                        {
                            dataLine = SR.ReadLine();
                            dataSplited = dataLine.Split(trnSplitChar, StringSplitOptions.RemoveEmptyEntries);
                            //fogdata.time_array.Add(Convert.ToDouble(dataSplited[0]));
                            fogdata.fdata_array.Add(Convert.ToDouble(dataSplited[1]));
                            fogdata.tdata_array.Add(Convert.ToDouble(dataSplited[2]));
                            if (Convert.ToInt32(dataSplited[0]) % timePara.sampleFreq == 0)
                            {
                                time++;
                                fogdata.time_array.Add(time);
                                fogdata.fdata_1s_array.Add(fogdata.fdata_array.Sum() / timePara.sampleFreq);
                                fogdata.tdata_1s_array.Add(fogdata.tdata_array.Sum() / timePara.sampleFreq);
                                fogdata.fdata_array.Clear();
                                fogdata.tdata_array.Clear();
                                
                            }
                        }
                        fogdata.fdata_smooth_array = fogdata.fdata_1s_array;
                        fogdata.tdata_smooth_array = fogdata.tdata_1s_array;
                        Channels_FogData_list.Add(fogdata);

                    }
                    AddChartArea(Channel_DataFile_SR_list.Count());
                    AddSeries(Channel_DataFile_SR_list.Count());
                    for (int i = 0; i < Channels_FogData_list.Count; i++)
                    {
                        DrawLoadFogdata(i);
                    }
                }
                else
                {
                    MessageBox.Show("没有选择数据文件啊!");
                }
            }

            
        }

        private void Btn_Smooth_Click(object sender, EventArgs e)
        {
            double smoothTime = Convert.ToDouble(tBox_smoothTime.Text);
            for (int i = 0; i < Channels_FogData_list.Count; i++)
            {
                int index = Convert.ToInt32(Channels_FogData_list[i].fdata_1s_array.Count / Convert.ToInt32(smoothTime));
                Channels_FogData_list[i].time_smooth_array.Clear();
                //Channels_FogData_list[i].fdata_smooth_array.Clear();
                //Channels_FogData_list[i].tdata_smooth_array.Clear();

                
                Channels_FogData_list[i].fdata_smooth_array = SmoothData(Channels_FogData_list[i].fdata_1s_array, Convert.ToInt32(smoothTime));
                Channels_FogData_list[i].tdata_smooth_array = SmoothData(Channels_FogData_list[i].tdata_1s_array, Convert.ToInt32(smoothTime));
                for (int j = 0; j < index; j++)
                {
                    Channels_FogData_list[i].time_smooth_array.Add(j + 1);
                }
//                 for (int j = 0; j < index; j++)
//                 {
//                     double[] data_array = new double[Convert.ToInt32(smoothTime)];
//                     Channels_FogData_list[i].fdata_1s_array.CopyTo(j* Convert.ToInt32(smoothTime), data_array,0, Convert.ToInt32(smoothTime));
//                     Channels_FogData_list[i].fdata_smooth_array.Add(data_array.Average());
//                     Channels_FogData_list[i].tdata_1s_array.CopyTo(j * Convert.ToInt32(smoothTime), data_array, 0, Convert.ToInt32(smoothTime));
//                     Channels_FogData_list[i].tdata_smooth_array.Add(data_array.Average());
//                     Channels_FogData_list[i].time_array.Add(j + 1);
//                     Channels_FogData_list[i].Fog_Bias_std = CalculateStdDev(Channels_FogData_list[i].fdata_smooth_array) / Channels_FogData_list[i].scaleFactor * 3600;
//                 }
            }
            foreach (var item in chart.Series)
            {
                item.Points.Clear();
            }
            for (int i = 0; i < Channels_FogData_list.Count; i++)
            {
                DrawLoadFogdata(i);
            }
        }

        private List<double> SmoothData(List<double> data_list, int smoothTime)
        {
            List<double> smoothData_list = new List<double>();
            int index = data_list.Count / smoothTime;
            double[] data_array = new double[smoothTime];
            for (int i = 0; i < index; i++)
            {
                data_list.CopyTo(i * smoothTime, data_array, 0, smoothTime);
                smoothData_list.Add(data_array.Average());
            }



            return smoothData_list;
        }
        public void Send_table_Enable()
        {
            byte[] Sendbuff = new byte[15];
            Sendbuff[0] = 0xAA;
            Sendbuff[1] = 0xA5;
            Sendbuff[2] = 0x55;
            Sendbuff[3] = 0x1E;
            int checksum = 0;
            for (int i = 0; i < 4; i++)
            {
                Sendbuff[i + 4] = 0;
                Sendbuff[i + 8] = 0;
            }
            Sendbuff[8] = 0x00;

            for (int i = 0; i < 13; i++)
            {
                checksum += Sendbuff[i];
            }
            Sendbuff[13] = Convert.ToByte(checksum & 0xff);
            table_serial.Write(Sendbuff, 0, 14);
            tBox_info.Text += DateTime.Now.ToString("yyyyMMdd-HHmmss") +  ":  发送转台使能指令！" + "\r\n";
        }
        public void Send_table_Stop()
        {
            byte[] Sendbuff = new byte[15];
            Sendbuff[0] = 0xAA;
            Sendbuff[1] = 0xA5;
            Sendbuff[2] = 0x55;
            Sendbuff[3] = 0x1F;
            int checksum = 0;
            for (int i = 0; i < 4; i++)
            {
                Sendbuff[i + 4] = 0;
                Sendbuff[i + 8] = 0;
            }
            Sendbuff[8] = 0x00;

            for (int i = 0; i < 13; i++)
            {
                checksum += Sendbuff[i];
            }
            Sendbuff[13] = Convert.ToByte(checksum & 0xff);
            table_serial.Write(Sendbuff, 0, 14);
            tBox_info.Text += DateTime.Now.ToString("yyyyMMdd-HHmmss") + ":  发送转台停止指令！" + "\r\n";
        }
        public void Send_table_Connect()
        {
            byte[] Sendbuff = new byte[15];
            Sendbuff[0] = 0xAA;
            Sendbuff[1] = 0xA5;
            Sendbuff[2] = 0x55;
            Sendbuff[3] = 0x0F;
            int checksum = 0;
            for (int i = 0; i < 4; i++)
            {
                Sendbuff[i + 4] = 0;
                Sendbuff[i + 8] = 0;
            }
            Sendbuff[8] = 0x00;

            for (int i = 0; i < 13; i++)
            {
                checksum += Sendbuff[i];
            }
            Sendbuff[13] = Convert.ToByte(checksum & 0xff);
            table_serial.Write(Sendbuff, 0, 14);
            tBox_info.Text += DateTime.Now.ToString("yyyyMMdd-HHmmss") + ":  发送转台通讯指令！" + "\r\n";
        }
        public void Send_table_rateCommand(double rate,double drate)
        {
            byte[] Sendbuff = new byte[15];
            
            int irate;
            int idrate;
            int checksum = 0;

            irate = Convert.ToInt32(rate * 10000);
            idrate = Convert.ToInt32(drate * 10000);
            Sendbuff[0] = 0xAA;
            Sendbuff[1] = 0xA5;
            Sendbuff[2] = 0x55;
            Sendbuff[3] = 0x13;
            Sendbuff[4] = Convert.ToByte(irate & 0xff);
            Sendbuff[5] = Convert.ToByte((irate >> 8) & 0xff);
            Sendbuff[6] = Convert.ToByte((irate >> 16) & 0xff);
            Sendbuff[7] = Convert.ToByte((irate >> 24) & 0xff);
            Sendbuff[8] = 0x00;
            Sendbuff[9] = Convert.ToByte(idrate & 0xff);
            Sendbuff[10] = Convert.ToByte((idrate >> 8) & 0xff);
            Sendbuff[11] = Convert.ToByte((idrate >> 16) & 0xff);
            Sendbuff[12] = Convert.ToByte((idrate >> 24) & 0xff);
            for (int i = 0; i < 13; i++)
            {
                checksum += Sendbuff[i];
            }
            Sendbuff[13] = Convert.ToByte(checksum & 0xff);
            table_serial.Write(Sendbuff, 0, 14);
            string info_text = DateTime.Now.ToString("yyyyMMdd-HHmmss") + ":  发送转台速率指令！速率为：" + rate.ToString() + ", 加速率为："+ drate.ToString() + "\r\n";
            this.BeginInvoke(updateTableFrmdata, true,info_text);

        }
        /*************************************
       函数名：GetCell
       创建日期：2019/10/25
       函数功能：判断EXCEL中单元格是否创建，没有则创建
       函数参数：
           sheet
           row_num
           cell_num
       返回值：NPOI.SS.UserModel.ICell
       *************************************/
        public ICell GetCell(ISheet sheet, int row_num, int cell_num)
        {
            IRow row = sheet.GetRow(row_num) == null ? sheet.CreateRow(row_num) : sheet.GetRow(row_num);
            ICell cell = row.GetCell(cell_num) == null ? row.CreateCell(cell_num) : row.GetCell(cell_num);

            return cell;
        }

        private void Btn_Table_Enable_Click(object sender, EventArgs e)
        {
            Send_table_Connect();
            Send_table_Enable();
        }

        private void Brn_Table_stop_Click(object sender, EventArgs e)
        {
            Send_table_Stop();
        }

        private void Btn_AngleSet_Send_Click(object sender, EventArgs e)
        {
            byte[] Sendbuff = new byte[15];
            int angle;
            int rate;
            int checksum = 0;
            angle = Convert.ToInt32(Convert.ToDouble(tBox_AngleSet_angle.Text) * 10000);
            rate = Convert.ToInt32(Convert.ToDouble(tBox_AngleSet_rate.Text) * 10000);
            Sendbuff[0] = 0xAA;
            Sendbuff[1] = 0xA5;
            Sendbuff[2] = 0x55;

            Sendbuff[3] = 0x11;
            Sendbuff[4] = Convert.ToByte(angle & 0xff);
            Sendbuff[5] = Convert.ToByte((angle >> 8) & 0xff);
            Sendbuff[6] = Convert.ToByte((angle >> 16) & 0xff);
            Sendbuff[7] = Convert.ToByte((angle >> 24) & 0xff);
            Sendbuff[8] = 0xA0;
            Sendbuff[9] = Convert.ToByte(rate & 0xff);
            Sendbuff[10] = Convert.ToByte((rate >> 8) & 0xff);
            Sendbuff[11] = Convert.ToByte((rate >> 16) & 0xff);
            Sendbuff[12] = Convert.ToByte((rate >> 24) & 0xff);

            for (int i = 0; i < 13; i++)
            {
                checksum += Sendbuff[i];
            }
            Sendbuff[13] = Convert.ToByte(checksum & 0xff);
            table_serial.Write(Sendbuff, 0, 14);
            tBox_info.Text += DateTime.Now.ToString("yyyyMMdd-HHmmss") + ":  发送转台角度指令！角度为：" + tBox_AngleSet_angle.Text + ", 速率为：" + tBox_AngleSet_rate.Text + "\r\n";
        }

        private void Btn_RateSet_Send_Click(object sender, EventArgs e)
        {
            byte[] Sendbuff = new byte[15];
            int drate;
            int rate;
            int checksum = 0;
            rate = Convert.ToInt32(Convert.ToDouble(tBox_RateSet_rate.Text) * 10000);
            drate = Convert.ToInt32(Convert.ToDouble(tBox_RateSet_drate .Text) * 10000);
            Sendbuff[0] = 0xAA;
            Sendbuff[1] = 0xA5;
            Sendbuff[2] = 0x55;

            Sendbuff[3] = 0x13;
            Sendbuff[4] = Convert.ToByte(rate & 0xff);
            Sendbuff[5] = Convert.ToByte((rate >> 8) & 0xff);
            Sendbuff[6] = Convert.ToByte((rate >> 16) & 0xff);
            Sendbuff[7] = Convert.ToByte((rate >> 24) & 0xff);
            Sendbuff[8] = 0x00;
            Sendbuff[9] = Convert.ToByte(drate & 0xff);
            Sendbuff[10] = Convert.ToByte((drate >> 8) & 0xff);
            Sendbuff[11] = Convert.ToByte((drate >> 16) & 0xff);
            Sendbuff[12] = Convert.ToByte((drate >> 24) & 0xff);

            for (int i = 0; i < 13; i++)
            {
                checksum += Sendbuff[i];
            }
            Sendbuff[13] = Convert.ToByte(checksum & 0xff);
            table_serial.Write(Sendbuff, 0, 14);
            tBox_info.Text += DateTime.Now.ToString("yyyyMMdd-HHmmss") + ":  发送转台速率指令！速率为：" + tBox_RateSet_rate.Text + ", 加速率为：" + tBox_RateSet_drate.Text + "\r\n";
        }

        private void Btn_Step_Send_Click(object sender, EventArgs e)
        {
            byte[] Sendbuff = new byte[15];
            int angle;
            int rate;
            int checksum = 0;
            angle = Convert.ToInt32(Convert.ToDouble(tBox_SetpRate_angle.Text) * 10000);
            rate = Convert.ToInt32(Convert.ToDouble(tBox_SetpRate_rate.Text) * 10000);
            Sendbuff[0] = 0xAA;
            Sendbuff[1] = 0xA5;
            Sendbuff[2] = 0x55;

            Sendbuff[3] = 0x12;
            Sendbuff[4] = Convert.ToByte(rate & 0xff);
            Sendbuff[5] = Convert.ToByte((rate >> 8) & 0xff);
            Sendbuff[6] = Convert.ToByte((rate >> 16) & 0xff);
            Sendbuff[7] = Convert.ToByte((rate >> 24) & 0xff);
            Sendbuff[8] = 0x00;
            Sendbuff[9] = Convert.ToByte(angle & 0xff);
            Sendbuff[10] = Convert.ToByte((angle >> 8) & 0xff);
            Sendbuff[11] = Convert.ToByte((angle >> 16) & 0xff);
            Sendbuff[12] = Convert.ToByte((angle >> 24) & 0xff);

            for (int i = 0; i < 13; i++)
            {
                checksum += Sendbuff[i];
            }
            Sendbuff[13] = Convert.ToByte(checksum & 0xff);
            table_serial.Write(Sendbuff, 0, 14);
            tBox_info.Text += DateTime.Now.ToString("yyyyMMdd-HHmmss") + ":  发送转台点动指令！角度为：" + tBox_SetpRate_angle.Text + ", 速率为：" + tBox_SetpRate_rate.Text + "\r\n";
        }

        private void Btn_Swing_Send_Click(object sender, EventArgs e)
        {
            byte[] Sendbuff = new byte[15];
            int amp;
            int freq;
            int checksum = 0;
            amp = Convert.ToInt32(Convert.ToDouble(tBox_Swing_Amp.Text) * 10000);
            freq = Convert.ToInt32(Convert.ToDouble(tBox_Swing_Amp.Text) * 10000);

            Sendbuff[3] = 0x14;
            Sendbuff[4] = Convert.ToByte(amp & 0xff);
            Sendbuff[5] = Convert.ToByte((amp >> 8) & 0xff);
            Sendbuff[6] = Convert.ToByte((amp >> 16) & 0xff);
            Sendbuff[7] = Convert.ToByte((amp >> 24) & 0xff);
            Sendbuff[8] = 0x00;
            Sendbuff[9] = Convert.ToByte(freq & 0xff);
            Sendbuff[10] = Convert.ToByte((freq >> 8) & 0xff);
            Sendbuff[11] = Convert.ToByte((freq >> 16) & 0xff);
            Sendbuff[12] = Convert.ToByte((freq >> 24) & 0xff);

            for (int i = 0; i < 13; i++)
            {
                checksum += Sendbuff[i];
            }
            Sendbuff[13] = Convert.ToByte(checksum & 0xff);
            table_serial.Write(Sendbuff, 0, 14);
            tBox_info.Text += DateTime.Now.ToString("yyyyMMdd-HHmmss") + ":  发送转台摇摆指令！角度为：" + tBox_SetpRate_angle.Text + ", 频率为：" + tBox_SetpRate_rate.Text + "\r\n";
        }
    }
    
 }
