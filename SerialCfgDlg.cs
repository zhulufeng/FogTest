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
using System.IO;
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
        bool isByFile = false;
        public SerialCfgDlg()
        {
            InitializeComponent();
            //SetConfigFile();
            setComBox();
            checkedListBox_Channel.SetItemChecked(0, true);
            CBox_Table_BaudRate.SelectedIndex = 6;
            cBox_Table_DataBit.SelectedIndex  = 1;
            cBox_Table_StopBit.SelectedIndex  = 1;
            cBox_Table_ParityBit.SelectedIndex = 1;
            isByFile = false;
        }
        public SerialCfgDlg(XSSFWorkbook workbook)
        {
            InitializeComponent();
            //SetConfigFile();
            isByFile = true;
            setComBox();
            SetFormByConfigFile(workbook);
            
            
        }
        /*************************************
        函数名：SetFormByConfigFile
        创建日期：2019/11/06
        函数功能：
        函数参数：
        workbook
        返回值：void
        *************************************/
        public void SetFormByConfigFile(XSSFWorkbook workbook)
        {
            ISheet sht = workbook.GetSheet("通道串口配置");
            this.groupBox_Table.Enabled = false;
            this.groupBox_channel_1.Enabled = false;
            this.groupBox_channel_2.Enabled = false;
            this.groupBox_channel_3.Enabled = false;
            this.groupBox_channel_4.Enabled = false;
            this.groupBox_channel_5.Enabled = false;
            this.groupBox_channel_6.Enabled = false;
            for (int i = 0; i < sht.LastRowNum; i++)
            {
                if (sht.GetRow(i).GetCell(1).ToString().Equals("true",StringComparison.OrdinalIgnoreCase))//判断通道是否使能
                {
                    SetGroupBoxByFile(sht.GetRow(i).GetCell(0).ToString(), sht);
                }
                
            }

            
        }
        /*************************************
        函数名：SetGroupBoxByFile
        创建日期：2019/11/06
        函数功能：
        函数参数：
        ChannelID
        返回值：void
        *************************************/
        public void SetGroupBoxByFile(string ChannelID,ISheet sht)
        {
            string[] ArryPort = SerialPort.GetPortNames();//搜索
            switch (ChannelID)
            {
                case "转台通道":
                    {
                        groupBox_Table.Enabled    = true;
                        checkedListBox_Channel.SetItemChecked(0, true);
                        tBox_table_ID.Text        = sht.GetRow(1).GetCell(7).ToString();
                        CBox_Table_BaudRate.Text  = sht.GetRow(1).GetCell(3).ToString();
                        cBox_Table_DataBit.Text   = sht.GetRow(1).GetCell(4).ToString();
                        cBox_Table_StopBit.Text   = sht.GetRow(1).GetCell(5).ToString();
                        cBox_Table_ParityBit.Text = sht.GetRow(1).GetCell(6).ToString();
                        if (ArryPort.Contains(sht.GetRow(1).GetCell(2).ToString().ToUpper()))
                        {
                            cBox_Table_COMID.Text = sht.GetRow(1).GetCell(2).ToString().ToUpper();
                        }
                        else
                        {
                            MessageBox.Show("串口号不对，请在列表中选择！");
                        }
                        break;
                    }
                case "通道一":
                    {
                        groupBox_channel_1.Enabled = true;
                        checkedListBox_Channel.SetItemChecked(1, true);
                        tBox_CH1_FOGID.Text     = sht.GetRow(2).GetCell(7).ToString();
                        cBox_CH1_BaudRate.Text  = sht.GetRow(2).GetCell(3).ToString();
                        cBox_CH1_DataBit.Text   = sht.GetRow(2).GetCell(4).ToString();
                        cBox_CH1_StopBit.Text   = sht.GetRow(2).GetCell(5).ToString();
                        cBox_CH1_ParityBit.Text = sht.GetRow(2).GetCell(6).ToString();
                        if (ArryPort.Contains(sht.GetRow(2).GetCell(2).ToString().ToUpper()))
                        {
                            cBox_CH1_COMID.Text = sht.GetRow(2).GetCell(2).ToString().ToUpper();
                        }
                        else
                        {
                            MessageBox.Show("串口号不对，请在列表中选择！");
                        }
                        break;
                    }
                case "通道二":
                    {
                        groupBox_channel_2.Enabled = true;
                        checkedListBox_Channel.SetItemChecked(2, true);
                        tBox_CH2_FOGID.Text     = sht.GetRow(3).GetCell(7).ToString();
                        cBox_CH2_BaudRate.Text  = sht.GetRow(3).GetCell(3).ToString();
                        cBox_CH2_DataBit.Text   = sht.GetRow(3).GetCell(4).ToString();
                        cBox_CH2_StopBit.Text   = sht.GetRow(3).GetCell(5).ToString();
                        cBox_CH2_ParityBit.Text = sht.GetRow(3).GetCell(6).ToString();
                        if (ArryPort.Contains(sht.GetRow(3).GetCell(2).ToString().ToUpper()))
                        {
                            cBox_CH2_COMID.Text = sht.GetRow(3).GetCell(2).ToString().ToUpper();
                        }
                        else
                        {
                            MessageBox.Show("串口号不对，请在列表中选择！");
                        }
                        break;
                    }
                case "通道三":
                    {
                        groupBox_channel_3.Enabled = true;
                        checkedListBox_Channel.SetItemChecked(3, true);
                        tBox_CH3_FOGID.Text     = sht.GetRow(4).GetCell(7).ToString();
                        cBox_CH3_BaudRate.Text  = sht.GetRow(4).GetCell(3).ToString();
                        cBox_CH3_DataBit.Text   = sht.GetRow(4).GetCell(4).ToString();
                        cBox_CH3_StopBit.Text   = sht.GetRow(4).GetCell(5).ToString();
                        cBox_CH3_ParityBit.Text = sht.GetRow(4).GetCell(6).ToString();
                        if (ArryPort.Contains(sht.GetRow(4).GetCell(2).ToString().ToUpper()))
                        {
                            cBox_CH3_COMID.Text = sht.GetRow(4).GetCell(2).ToString().ToUpper();
                        }
                        else
                        {
                            MessageBox.Show("串口号不对，请在列表中选择！");
                        }
                        break;
                    }
                case "通道四":
                    {
                        groupBox_channel_4.Enabled = true;
                        checkedListBox_Channel.SetItemChecked(4, true);
                        tBox_CH4_FOGID.Text     = sht.GetRow(5).GetCell(7).ToString();
                        cBox_CH4_BaudRate.Text  = sht.GetRow(5).GetCell(3).ToString();
                        cBox_CH4_DataBit.Text   = sht.GetRow(5).GetCell(4).ToString();
                        cBox_CH4_StopBit.Text   = sht.GetRow(5).GetCell(5).ToString();
                        cBox_CH4_ParityBit.Text = sht.GetRow(5).GetCell(6).ToString();
                        if (ArryPort.Contains(sht.GetRow(5).GetCell(2).ToString().ToUpper()))
                        {
                            cBox_CH4_COMID.Text = sht.GetRow(5).GetCell(2).ToString().ToUpper();
                        }
                        else
                        {
                            MessageBox.Show("串口号不对，请在列表中选择！");
                        }
                        break;
                    }
                case "通道五":
                    {
                        groupBox_channel_5.Enabled = true;
                        checkedListBox_Channel.SetItemChecked(5, true);
                        tBox_CH5_FOGID.Text     = sht.GetRow(6).GetCell(7).ToString();
                        cBox_CH5_BaudRate.Text  = sht.GetRow(6).GetCell(3).ToString();
                        cBox_CH5_DataBit.Text   = sht.GetRow(6).GetCell(4).ToString();
                        cBox_CH5_StopBit.Text   = sht.GetRow(6).GetCell(5).ToString();
                        cBox_CH5_ParityBit.Text = sht.GetRow(6).GetCell(6).ToString();
                        if (ArryPort.Contains(sht.GetRow(6).GetCell(2).ToString().ToUpper()))
                        {
                            cBox_CH5_COMID.Text = sht.GetRow(6).GetCell(2).ToString().ToUpper();
                        }
                        else
                        {
                            MessageBox.Show("串口号不对，请在列表中选择！");
                        }
                        break;
                    }
                case "通道六":
                    {
                        groupBox_channel_6.Enabled = true;
                        checkedListBox_Channel.SetItemChecked(6, true);
                        tBox_CH6_FOGID.Text     = sht.GetRow(7).GetCell(7).ToString();
                        cBox_CH6_BaudRate.Text  = sht.GetRow(7).GetCell(3).ToString();
                        cBox_CH6_DataBit.Text   = sht.GetRow(7).GetCell(4).ToString();
                        cBox_CH6_StopBit.Text   = sht.GetRow(7).GetCell(5).ToString();
                        cBox_CH6_ParityBit.Text = sht.GetRow(7).GetCell(6).ToString();
                        if (ArryPort.Contains(sht.GetRow(7).GetCell(2).ToString().ToUpper()))
                        {
                            cBox_CH6_COMID.Text = sht.GetRow(7).GetCell(2).ToString().ToUpper();
                        }
                        else
                        {
                            MessageBox.Show("串口号不对，请在列表中选择！");
                        }
                        break;
                    }
            }
        }
        /*************************************
        函数名：InitializeConfigFile
        创建日期：2019/10/24
        函数功能：初始化配置文件中串口部分
        函数参数：
        返回值：void
        *************************************/
        public void SetConfigFile()
        {
            int SelectedChannelsNum = 0;
            //读入配置文件
            FileStream rfile = new FileStream(FilePara.ConfigFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            XSSFWorkbook workbook = new XSSFWorkbook(rfile);
            rfile.Close();

            ISheet sht = workbook.GetSheet("通道串口配置");
            //配置转台通道
            IRow row = sht.GetRow(1);
            ICell cell = GetCell(sht, 1, 1);
            if (groupBox_Table.Enabled)
            {
                cell.SetCellValue("True");
                cell = GetCell(sht, 1, 2);
                if (cBox_Table_COMID.SelectedItem == null)
                {
                    cell.SetCellValue(cBox_Table_COMID.SelectedItem.ToString());
                }
                else
                {
                    cell.SetCellValue(cBox_Table_COMID.SelectedItem.ToString());
                }
                
                cell = GetCell(sht, 1, 3);
                cell.SetCellValue(CBox_Table_BaudRate.SelectedItem.ToString());
                cell = GetCell(sht, 1, 4);
                cell.SetCellValue(cBox_Table_DataBit.SelectedItem.ToString());
                cell = GetCell(sht, 1, 5);
                cell.SetCellValue(cBox_Table_StopBit.SelectedItem.ToString());
                cell = GetCell(sht, 1, 6);
                cell.SetCellValue(cBox_Table_ParityBit.SelectedItem.ToString());
                cell = GetCell(sht, 1, 7);
                cell.SetCellValue(tBox_table_ID.Text);
                cell = GetCell(sht, 1, 8);
                cell.SetCellValue("None");
            }
            else
            {
                cell.SetCellValue("False");
            }

            //配置测试通道一
            row = sht.GetRow(2);
            cell = GetCell(sht, 2, 1);
            if (groupBox_channel_1.Enabled)
            {
                cell.SetCellValue("True");
                cell = GetCell(sht, 2, 2);
                if (cBox_CH1_COMID.SelectedItem == null)
                {
                    cell.SetCellValue("null");
                }
                else
                {
                    cell.SetCellValue(cBox_CH1_COMID.SelectedItem.ToString());
                }
                cell = GetCell(sht, 2, 3);
                cell.SetCellValue(cBox_CH1_BaudRate.SelectedItem.ToString());
                cell = GetCell(sht, 2, 4);
                cell.SetCellValue(cBox_CH1_DataBit.SelectedItem.ToString());
                cell = GetCell(sht, 2, 5);
                cell.SetCellValue(cBox_CH1_StopBit.SelectedItem.ToString());
                cell = GetCell(sht, 2, 6);
                cell.SetCellValue(cBox_CH1_ParityBit.SelectedItem.ToString());
                cell = GetCell(sht, 2, 7);
                cell.SetCellValue(tBox_CH1_FOGID.Text);
                cell = GetCell(sht, 2, 8);
                cell.SetCellValue(tBox_CH1_SF.Text);
                SelectedChannelsNum++;
            }
            else
            {
                cell.SetCellValue("False");
            }
            //配置测试通道二
            row = sht.GetRow(3);
            cell = GetCell(sht, 3, 1);
            if (groupBox_channel_2.Enabled)
            {
                cell.SetCellValue("True");
                cell = GetCell(sht, 3, 2);
                if (cBox_CH2_COMID.SelectedItem == null)
                {
                    cell.SetCellValue("null");
                }
                else
                {
                    cell.SetCellValue(cBox_CH2_COMID.SelectedItem.ToString());
                }
                cell = GetCell(sht, 3, 3);
                cell.SetCellValue(cBox_CH2_BaudRate.SelectedItem.ToString());
                cell = GetCell(sht, 3, 4);
                cell.SetCellValue(cBox_CH2_DataBit.SelectedItem.ToString());
                cell = GetCell(sht, 3, 5);
                cell.SetCellValue(cBox_CH2_StopBit.SelectedItem.ToString());
                cell = GetCell(sht, 3, 6);
                cell.SetCellValue(cBox_CH2_ParityBit.SelectedItem.ToString());
                cell = GetCell(sht, 3, 7);
                cell.SetCellValue(tBox_CH2_FOGID.Text);
                cell = GetCell(sht, 3, 8);
                cell.SetCellValue(tBox_CH2_SF.Text);
                SelectedChannelsNum++;
            }
            else
            {
                cell.SetCellValue("False");
            }
            //配置测试通道三
            row = sht.GetRow(4);
            cell = GetCell(sht, 4, 1);
            if (groupBox_channel_3.Enabled)
            {
                cell.SetCellValue("True");
                cell = GetCell(sht, 4, 2);
                if (cBox_CH3_COMID.SelectedItem == null)
                {
                    cell.SetCellValue("null");
                }
                else
                {
                    cell.SetCellValue(cBox_CH3_COMID.SelectedItem.ToString());
                }
                
                cell = GetCell(sht, 4, 3);
                cell.SetCellValue(cBox_CH3_BaudRate.SelectedItem.ToString());
                cell = GetCell(sht, 4, 4);
                cell.SetCellValue(cBox_CH3_DataBit.SelectedItem.ToString());
                cell = GetCell(sht, 4, 5);
                cell.SetCellValue(cBox_CH3_StopBit.SelectedItem.ToString());
                cell = GetCell(sht, 4, 6);
                cell.SetCellValue(cBox_CH3_ParityBit.SelectedItem.ToString());
                cell = GetCell(sht, 4, 7);
                cell.SetCellValue(tBox_CH3_FOGID.Text);
                cell = GetCell(sht, 4, 8);
                cell.SetCellValue(tBox_CH3_SF.Text);
                SelectedChannelsNum++;
            }
            else
            {
                cell.SetCellValue("False");
            }
            //配置测试通道四
            row = sht.GetRow(5);
            cell = GetCell(sht, 5, 1);
            if (groupBox_channel_4.Enabled)
            {
                cell.SetCellValue("True");
                cell = GetCell(sht, 5, 2);
                if (cBox_CH4_COMID.SelectedItem == null)
                {
                    cell.SetCellValue("null");
                }
                else
                {
                    cell.SetCellValue(cBox_CH4_COMID.SelectedItem.ToString());
                }
                cell = GetCell(sht, 5, 3);
                cell.SetCellValue(cBox_CH4_BaudRate.SelectedItem.ToString());
                cell = GetCell(sht, 5, 4);
                cell.SetCellValue(cBox_CH4_DataBit.SelectedItem.ToString());
                cell = GetCell(sht, 5, 5);
                cell.SetCellValue(cBox_CH4_StopBit.SelectedItem.ToString());
                cell = GetCell(sht, 5, 6);
                cell.SetCellValue(cBox_CH4_ParityBit.SelectedItem.ToString());
                cell = GetCell(sht, 5, 7);
                cell.SetCellValue(tBox_CH4_FOGID.Text);
                cell = GetCell(sht, 5, 8);
                cell.SetCellValue(tBox_CH4_SF.Text); 
                SelectedChannelsNum++;
            }
            else
            {
                cell.SetCellValue("False");
            }
            //配置测试通道五
            row = sht.GetRow(6);
            cell = GetCell(sht, 6, 1);
            if (groupBox_channel_5.Enabled)
            {
                cell.SetCellValue("True");
                cell = GetCell(sht, 6, 2);
                if (cBox_CH5_COMID.SelectedItem == null)
                {
                    cell.SetCellValue("null");
                }
                else
                {
                    cell.SetCellValue(cBox_CH5_COMID.SelectedItem.ToString());
                }
                
                cell = GetCell(sht, 6, 3);
                cell.SetCellValue(cBox_CH5_BaudRate.SelectedItem.ToString());
                cell = GetCell(sht, 6, 4);
                cell.SetCellValue(cBox_CH5_DataBit.SelectedItem.ToString());
                cell = GetCell(sht, 6, 5);
                cell.SetCellValue(cBox_CH5_StopBit.SelectedItem.ToString());
                cell = GetCell(sht, 6, 6);
                cell.SetCellValue(cBox_CH5_ParityBit.SelectedItem.ToString());
                cell = GetCell(sht, 6, 7);
                cell.SetCellValue(tBox_CH5_FOGID.Text);
                cell = GetCell(sht, 6, 8);
                cell.SetCellValue(tBox_CH5_SF.Text);
                SelectedChannelsNum++;
            }
            else
            {
                cell.SetCellValue("False");
            }
            //配置测试通道六
            row = sht.GetRow(7);
            cell = GetCell(sht, 7, 1);
            if (groupBox_channel_6.Enabled)
            {
                cell.SetCellValue("True");
                cell = GetCell(sht, 7, 2);
                if (cBox_CH6_COMID.SelectedItem == null)
                {
                    cell.SetCellValue("null");
                }
                else
                {
                    cell.SetCellValue(cBox_CH6_COMID.SelectedItem.ToString());
                }
                
                cell = GetCell(sht, 7, 3);
                cell.SetCellValue(cBox_CH6_BaudRate.SelectedItem.ToString());
                cell = GetCell(sht, 7, 4);
                cell.SetCellValue(cBox_CH6_DataBit.SelectedItem.ToString());
                cell = GetCell(sht, 7, 5);
                cell.SetCellValue(cBox_CH6_StopBit.SelectedItem.ToString());
                cell = GetCell(sht, 7, 6);
                cell.SetCellValue(cBox_CH6_ParityBit.SelectedItem.ToString());
                cell = GetCell(sht, 7, 7);
                cell.SetCellValue(tBox_CH6_FOGID.Text);
                cell = GetCell(sht, 7, 8);
                cell.SetCellValue(tBox_CH6_SF.Text);
                SelectedChannelsNum++;
            }
            else
            {
                cell.SetCellValue("False");
            }
            cell = GetCell(sht, 8, 1);
            cell.SetCellValue(SelectedChannelsNum);
            //写入配置文件
            FileStream wfile = new FileStream(FilePara.ConfigFilePath, FileMode.Open, FileAccess.ReadWrite);
            workbook.Write(wfile);
            wfile.Close();
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
            IRow row = sheet.GetRow(row_num)    == null ? sheet.CreateRow(row_num) : sheet.GetRow(row_num);
            ICell cell  = row.GetCell(cell_num) == null ? row.CreateCell(cell_num) : row.GetCell(cell_num);

            return cell;
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
            SetConfigFile();
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
                groupBox_Table.Enabled            = true;
                CBox_Table_BaudRate.SelectedIndex = 6;
                cBox_Table_DataBit.SelectedIndex  = 1;
                cBox_Table_StopBit.SelectedIndex  = 1;
                cBox_Table_ParityBit.SelectedIndex = 1;
            }
            else
            {
                groupBox_Table.Enabled = false;
            }
            //通道1串口配置使能
            if (checkedListBox_Channel.GetItemChecked(1))
            {
                groupBox_channel_1.Enabled      = true;
                cBox_CH1_BaudRate.SelectedIndex = 8;
                cBox_CH1_DataBit.SelectedIndex  = 1;
                cBox_CH1_StopBit.SelectedIndex  = 1;
                cBox_CH1_ParityBit.SelectedIndex = 2;
            }
            else
            {
                groupBox_channel_1.Enabled = false;
            }
            //通道2串口配置使能
            if (checkedListBox_Channel.GetItemChecked(2))
            {
                groupBox_channel_2.Enabled      = true;
                cBox_CH2_BaudRate.SelectedIndex = 8;
                cBox_CH2_DataBit.SelectedIndex  = 1;
                cBox_CH2_StopBit.SelectedIndex  = 1;
                cBox_CH2_ParityBit.SelectedIndex = 2;
            }
            else
            {
                groupBox_channel_2.Enabled = false;
            }
            //通道3串口配置使能
            if (checkedListBox_Channel.GetItemChecked(3))
            {
                groupBox_channel_3.Enabled      = true;
                cBox_CH3_BaudRate.SelectedIndex = 8;
                cBox_CH3_DataBit.SelectedIndex  = 1;
                cBox_CH3_StopBit.SelectedIndex  = 1;
                cBox_CH3_ParityBit.SelectedIndex = 2;
            }                    
            else                 
            {                    
                groupBox_channel_3.Enabled = false;
            }
            //通道4串口配置使能
            if (checkedListBox_Channel.GetItemChecked(4))
            {
                groupBox_channel_4.Enabled      = true;
                cBox_CH4_BaudRate.SelectedIndex = 8;
                cBox_CH4_DataBit.SelectedIndex  = 1;
                cBox_CH4_StopBit.SelectedIndex  = 1;
                cBox_CH4_ParityBit.SelectedIndex = 2;
            }
            else
            {
                groupBox_channel_4.Enabled = false;
            }
            //通道5串口配置使能
            if (checkedListBox_Channel.GetItemChecked(5))
            {
                groupBox_channel_5.Enabled      = true;
                cBox_CH5_BaudRate.SelectedIndex = 8;
                cBox_CH5_DataBit.SelectedIndex  = 1;
                cBox_CH5_StopBit.SelectedIndex  = 1;
                cBox_CH5_ParityBit.SelectedIndex = 2;
            }
            else
            {
                groupBox_channel_5.Enabled = false;
            }
            //通道6串口配置使能
            if (checkedListBox_Channel.GetItemChecked(6))
            {
                groupBox_channel_6.Enabled      = true;
                cBox_CH6_BaudRate.SelectedIndex = 8;
                cBox_CH6_DataBit.SelectedIndex  = 1;
                cBox_CH6_StopBit.SelectedIndex  = 1;
                cBox_CH6_ParityBit.SelectedIndex = 2;
            }
            else
            {
                groupBox_channel_6.Enabled = false;
            }
        }

      
    }
}
