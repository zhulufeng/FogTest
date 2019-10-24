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

namespace FOGTestPlatform
{

    public partial class Form1 : Form
    {
        /*************************************
        全局变量声明
        **************************************/
        
        public Form1()
        {
            InitializeComponent();
            InitializeConfigFlie();
            IntializeChart();
        }
        public void InitializeConfigFlie()
        {
            Directory.CreateDirectory(FilePara.IMUDataBaseDirectory + DateTime.Now.ToString("yyyyMMdd-HHmmss"));
            // File.Create(FilePara.IMUDataBaseDirectory+DateTime.Now.ToString("yyyyMMdd-HHmmss")+@"\配置文件.xlsx");
            try
            {
                FileStream file = new FileStream(FilePara.IMUDataBaseDirectory + DateTime.Now.ToString("yyyyMMdd-HHmmss") + @"\配置文件.xlsx", FileMode.Create);
                XSSFWorkbook hsswfworkbook = new XSSFWorkbook();
                ISheet sheet = hsswfworkbook.CreateSheet("通道串口配置");
                hsswfworkbook.CreateSheet("试验配置");

                hsswfworkbook.Write(file);
                file.Close();
            }
            catch (Exception)
            {

                MessageBox.Show("配置文件生产异常！");
                throw;
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

            //添加图框
            chart.ChartAreas.Add(SetChartArea(1));
            chart.ChartAreas.Add(SetChartArea(2));
            chart.ChartAreas.Add(SetChartArea(3));
            chart.ChartAreas.Add(SetChartArea(4));
            chart.ChartAreas.Add(SetChartArea(5));


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

            switch (index)
            {
                case 1:
                    chartArea.Name = ("AxisX");
                    break;
                case 2:
                    chartArea.Name = ("AxisY");
                    break;
                case 3:
                    chartArea.Name = ("AxisZ");
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
            chartArea.AxisY.Title = "加速度数据";
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
            if (serialCfgDlg.ShowDialog() != DialogResult.Cancel)
            {
                //testCfgPara.numOftestChannels = 
            }
        }
    }
}
