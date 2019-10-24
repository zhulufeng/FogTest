/********************************************************************
	创建日期:	29:9:2019   14:49
	文件名: 	    E:\MyCode\C#Code\TestChart\TestChart\SerialCfgDlg.cs
	文件路径:	E:\MyCode\C#Code\TestChart\TestChart
	文件基类:	SerialCfgDlg
	扩展名:	    cs
	编写人:		Zhu Lufeng
	
	用途:	
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
    
    public partial class SerialCfgDlg : Form
    {
        public SerialCfgDlg()
        {
            InitializeComponent();
            SetConfigFile();
            setComBox();
            
        }

        SerialPort table_serial = new SerialPort();
        SerialPort ch1_serial   = new SerialPort();
        SerialPort ch2_serial   = new SerialPort();
        SerialPort ch3_serial   = new SerialPort();
        SerialPort ch4_serial   = new SerialPort();
        SerialPort ch5_serial   = new SerialPort();

        serialParameter tableSerialPara = new serialParameter();
        serialParameter ch1SerialPara   = new serialParameter();
        serialParameter ch2SerialPara   = new serialParameter();
        serialParameter ch3SerialPara   = new serialParameter();
        serialParameter ch4SerialPara   = new serialParameter();
        serialParameter ch5SerialPara   = new serialParameter();
        serialParameter ch6SerialPara = new serialParameter();
        /*************************************
        函数名：InitializeConfigFile
        创建日期：2019/10/24
        函数功能：初始化配置文件中串口部分
        函数参数：
        返回值：void
        *************************************/
        public void SetConfigFile()
        {
            
        }
        /*************************************
        函数名：setComBox
        创建日期：2019/10/22
        函数功能：设置串口号下来菜单为可用用串口
        函数参数：
        返回值：void
        *************************************/
        private void setComBox()
        {
            string[] ArryPort = SerialPort.GetPortNames();//搜索
            cBox_Table_COMID.Items.Clear();
            cBox_CH1_COMID.Items.Clear();
            cBox_CH2_COMID.Items.Clear();
            cBox_CH3_COMID.Items.Clear();
            cBox_CH4_COMID.Items.Clear();
            cBox_CH5_COMID.Items.Clear();
            cBox_CH6_COMID.Items.Clear();

            for (int i = 0; i < ArryPort.Length; i++)
            {
                cBox_Table_COMID.Items.Add(ArryPort[i]);
                cBox_CH1_COMID.Items.Add(ArryPort[i]);
                cBox_CH2_COMID.Items.Add(ArryPort[i]);
                cBox_CH3_COMID.Items.Add(ArryPort[i]);
                cBox_CH4_COMID.Items.Add(ArryPort[i]);
                cBox_CH5_COMID.Items.Add(ArryPort[i]);
                cBox_CH6_COMID.Items.Add(ArryPort[i]);

            }
            try
            {
                cBox_Table_COMID.SelectedIndex = 0;
            }
            catch (Exception)
            {
                MessageBox.Show("未接入串口！");
                //throw;
            }
            

        }
    
      
        /*************************************
        函数名：Btn_OK_Click
        创建日期：2019/09/29
        函数功能：将选好的串口参数配置给相应选通的串口
        函数参数：
        	sender
        	e
        返回值：void
        *************************************/
        private void Btn_OK_Click(object sender, EventArgs e)
        {
         
            if (checkedListBox_Channel.GetItemChecked(0))
            {
                tableSerialPara.comName   = cBox_Table_COMID.SelectedItem.ToString();
                tableSerialPara.baudRate  = CBox_Table_BaudRate.SelectedItem.ToString();
                tableSerialPara.dataBit   = cBox_Table_DataBit.SelectedItem.ToString();
                tableSerialPara.stopBit   = cBox_Table_StopBit.SelectedItem.ToString();
                tableSerialPara.parityBit = cBox_Table_CheckBit.SelectedItem.ToString();
            }

            if (ch1SerialPara.serial_enable)
            {
                ch1SerialPara.comName   = cBox_CH1_COMID.SelectedItem.ToString();
                ch1SerialPara.baudRate  = cBox_CH1_BaudRate.SelectedItem.ToString();
                ch1SerialPara.dataBit   = cBox_CH1_DataBit.SelectedItem.ToString();
                ch1SerialPara.stopBit   = cBox_CH1_StopBit.SelectedItem.ToString();
                ch1SerialPara.parityBit = cBox_CH1_CheckBit.SelectedItem.ToString();
            }

            if (ch2SerialPara.serial_enable)
            {
                ch2SerialPara.comName   = cBox_CH2_COMID.SelectedItem.ToString();
                ch2SerialPara.baudRate  = cBox_CH2_BaudRate.SelectedItem.ToString();
                ch2SerialPara.dataBit   = cBox_CH2_DataBit.SelectedItem.ToString();
                ch2SerialPara.stopBit   = cBox_CH2_StopBit.SelectedItem.ToString();
                ch2SerialPara.parityBit = cBox_CH2_CheckBit.SelectedItem.ToString();
            }

            if (ch3SerialPara.serial_enable)
            {
                ch3SerialPara.comName   = cBox_CH3_COMID.SelectedItem.ToString();
                ch3SerialPara.baudRate  = cBox_CH3_BaudRate.SelectedItem.ToString();
                ch3SerialPara.dataBit   = cBox_CH3_DataBit.SelectedItem.ToString();
                ch3SerialPara.stopBit   = cBox_CH3_StopBit.SelectedItem.ToString();
                ch3SerialPara.parityBit = cBox_CH3_CheckBit.SelectedItem.ToString();
            }

            if (ch4SerialPara.serial_enable)
            {
                ch4SerialPara.comName   = cBox_CH4_COMID.SelectedItem.ToString();
                ch4SerialPara.baudRate  = cBox_CH4_BaudRate.SelectedItem.ToString();
                ch4SerialPara.dataBit   = cBox_CH4_DataBit.SelectedItem.ToString();
                ch4SerialPara.stopBit   = cBox_CH4_StopBit.SelectedItem.ToString();
                ch4SerialPara.parityBit = cBox_CH4_CheckBit.SelectedItem.ToString();
            }

            if (ch5SerialPara.serial_enable)
            {
                ch5SerialPara.comName   = cBox_CH5_COMID.SelectedItem.ToString();
                ch5SerialPara.baudRate  = cBox_CH5_BaudRate.SelectedItem.ToString();
                ch5SerialPara.dataBit   = cBox_CH5_DataBit.SelectedItem.ToString();
                ch5SerialPara.stopBit   = cBox_CH5_StopBit.SelectedItem.ToString();
                ch5SerialPara.parityBit = cBox_CH5_CheckBit.SelectedItem.ToString();
            }
            
        }

        private void checkedListBox1_SelectedValueChanged(object sender, EventArgs e)
        {
            
        }

        /*************************************
        函数名：checkedListBox_Channel_SelectedIndexChanged
        创建日期：2019/09/29
        函数功能：根据选通的通道，使能串口编辑
        函数参数：
        	sender
        	e
        返回值：void
        *************************************/
        private void checkedListBox_Channel_SelectedIndexChanged(object sender, EventArgs e)
        {
            //转台通道串口配置使能
            if(checkedListBox_Channel.GetItemChecked(0))
            {
                groupBox_Table.Enabled = true;
                tableSerialPara.serial_enable = true;
            }
            else
            {
                groupBox_Table.Enabled = false;
                tableSerialPara.serial_enable = false;
            }
            //通道1串口配置使能
            if (checkedListBox_Channel.GetItemChecked(1))
            {
                groupBox_channel_1.Enabled = true;
                ch1SerialPara.serial_enable = true;
            }
            else
            {
                groupBox_channel_1.Enabled = false;
                ch1SerialPara.serial_enable = false;
            }
            //通道2串口配置使能
            if (checkedListBox_Channel.GetItemChecked(2))
            {
                groupBox_channel_2.Enabled = true;
                ch2SerialPara.serial_enable = true;
            }
            else
            {
                groupBox_channel_2.Enabled = false;
                ch2SerialPara.serial_enable = false;
            }
            //通道3串口配置使能
            if (checkedListBox_Channel.GetItemChecked(3))
            {
                groupBox_channel_3.Enabled = true;
                ch3SerialPara.serial_enable = true;
            }                    
            else                 
            {                    
                groupBox_channel_3.Enabled = false;
                ch3SerialPara.serial_enable = false;
            }
            //通道4串口配置使能
            if (checkedListBox_Channel.GetItemChecked(4))
            {
                groupBox_channel_4.Enabled = true;
                ch4SerialPara.serial_enable = true;
            }
            else
            {
                groupBox_channel_4.Enabled = false;
                ch4SerialPara.serial_enable = false;
            }
            //通道5串口配置使能
            if (checkedListBox_Channel.GetItemChecked(5))
            {
                groupBox_channel_5.Enabled = true;
                ch5SerialPara.serial_enable = true;
            }
            else
            {
                groupBox_channel_5.Enabled = false;
                ch5SerialPara.serial_enable = false;
            }
            //通道6串口配置使能
            if (checkedListBox_Channel.GetItemChecked(6))
            {
                groupBox_channel_6.Enabled = true;
                ch6SerialPara.serial_enable = true;
            }
            else
            {
                groupBox_channel_6.Enabled = false;
                ch6SerialPara.serial_enable = false;
            }
        }

        
    }
}
