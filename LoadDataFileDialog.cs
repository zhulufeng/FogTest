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

namespace FOGTestPlatform
{
    public partial class LoadDataFileDialog : Form
    {
        public List<double> scaleFactorList = new List<double>();
        public LoadDataFileDialog()
        {
            InitializeComponent();
        }

        private void Btn_LoadData_CH1_Click(object sender, EventArgs e)
        {
            if (tBox_LoadDataFilePath_CH1.Text != "")
            {
                DialogResult dr;
                dr = MessageBox.Show("该通道已选择数据文件，是否覆盖？", "确认对话框", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2);
                if(dr == DialogResult.Yes)
                {
                    FilePara.dataFileList.Remove(tBox_LoadDataFilePath_CH1.Text);
                }
                else
                {
                    return;
                }
                
            }
            string filePath = OpenDataFile();
            FilePara.dataFileList.Add(filePath);
            scaleFactorList.Add(Convert.ToDouble(tBox_ScaleFactor_CH1.Text));
            tBox_LoadDataFilePath_CH1.Text = filePath;
            HashSet<string> dataFileHashset = new HashSet<string>(FilePara.dataFileList);
            if (FilePara.dataFileList.Count() != dataFileHashset.Count())
            {
                FilePara.dataFileList.RemoveAt(FilePara.dataFileList.Count() - 1);
                scaleFactorList.RemoveAt(FilePara.dataFileList.Count() - 1);
                MessageBox.Show("这个文件已经选过了，请选择其他文件！");
            }
            
        }

        private void Btn_LoadData_CH2_Click(object sender, EventArgs e)
        {
            if (tBox_LoadDataFilePath_CH2.Text != "")
            {
                DialogResult dr;
                dr = MessageBox.Show("该通道已选择数据文件，是否覆盖？", "确认对话框", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2);
                if (dr == DialogResult.Yes)
                {
                    FilePara.dataFileList.Remove(tBox_LoadDataFilePath_CH2.Text);
                }
                else
                {
                    return;
                }
            }
            string filePath = OpenDataFile();
            FilePara.dataFileList.Add(filePath);
            tBox_LoadDataFilePath_CH2.Text = filePath;
            scaleFactorList.Add(Convert.ToDouble(tBox_ScaleFactor_CH2.Text));
            HashSet<string> dataFileHashset = new HashSet<string>(FilePara.dataFileList);
            if (FilePara.dataFileList.Count() != dataFileHashset.Count())
            {
                FilePara.dataFileList.RemoveAt(FilePara.dataFileList.Count() - 1);
                scaleFactorList.RemoveAt(FilePara.dataFileList.Count() - 1);
                MessageBox.Show("这个文件已经选过了，请选择其他文件！");
            }
        }

        private void Btn_LoadData_CH3_Click(object sender, EventArgs e)
        {
            if (tBox_LoadDataFilePath_CH3.Text != "")
            {
                DialogResult dr;
                dr = MessageBox.Show("该通道已选择数据文件，是否覆盖？", "确认对话框", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2);
                if (dr == DialogResult.Yes)
                {
                    FilePara.dataFileList.Remove(tBox_LoadDataFilePath_CH3.Text);
                }
                else
                {
                    return;
                }
            }
            string filePath = OpenDataFile();
            FilePara.dataFileList.Add(filePath);
            tBox_LoadDataFilePath_CH3.Text = filePath;
            scaleFactorList.Add(Convert.ToDouble(tBox_ScaleFactor_CH3.Text));
            HashSet<string> dataFileHashset = new HashSet<string>(FilePara.dataFileList);
            if (FilePara.dataFileList.Count() != dataFileHashset.Count())
            {
                FilePara.dataFileList.RemoveAt(FilePara.dataFileList.Count() - 1);
                scaleFactorList.RemoveAt(FilePara.dataFileList.Count() - 1);
                MessageBox.Show("这个文件已经选过了，请选择其他文件！");
            }
        }
        private void Btn_LoadData_CH4_Click(object sender, EventArgs e)
        {
            if (tBox_LoadDataFilePath_CH4.Text != "")
            {
                DialogResult dr;
                dr = MessageBox.Show("该通道已选择数据文件，是否覆盖？", "确认对话框", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2);
                if (dr == DialogResult.Yes)
                {
                    FilePara.dataFileList.Remove(tBox_LoadDataFilePath_CH4.Text);
                }
                else
                {
                    return;
                }
            }
            string filePath = OpenDataFile();
            FilePara.dataFileList.Add(filePath);
            scaleFactorList.Add(Convert.ToDouble(tBox_ScaleFactor_CH4.Text));
            tBox_LoadDataFilePath_CH4.Text = filePath;
            HashSet<string> dataFileHashset = new HashSet<string>(FilePara.dataFileList);
            if (FilePara.dataFileList.Count() != dataFileHashset.Count())
            {
                FilePara.dataFileList.RemoveAt(FilePara.dataFileList.Count() - 1);
                scaleFactorList.RemoveAt(FilePara.dataFileList.Count() - 1);
                MessageBox.Show("这个文件已经选过了，请选择其他文件！");
            }
        }
        private void Btn_LoadData_CH5_Click(object sender, EventArgs e)
        {
            if (tBox_LoadDataFilePath_CH5.Text != "")
            {
                DialogResult dr;
                dr = MessageBox.Show("该通道已选择数据文件，是否覆盖？", "确认对话框", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2);
                if (dr == DialogResult.Yes)
                {
                    FilePara.dataFileList.Remove(tBox_LoadDataFilePath_CH5.Text);
                }
                else
                {
                    return;
                }
            }
            string filePath = OpenDataFile();
            FilePara.dataFileList.Add(filePath);
            tBox_LoadDataFilePath_CH5.Text = filePath;
            scaleFactorList.Add(Convert.ToDouble(tBox_ScaleFactor_CH5.Text));
            HashSet<string> dataFileHashset = new HashSet<string>(FilePara.dataFileList);
            if (FilePara.dataFileList.Count() != dataFileHashset.Count())
            {
                FilePara.dataFileList.RemoveAt(FilePara.dataFileList.Count() - 1);
                scaleFactorList.RemoveAt(FilePara.dataFileList.Count() - 1);
                MessageBox.Show("这个文件已经选过了，请选择其他文件！");
            }
        }
        private void Btn_LoadData_CH6_Click(object sender, EventArgs e)
        {
            if (tBox_LoadDataFilePath_CH6.Text != "")
            {
                DialogResult dr;
                dr = MessageBox.Show("该通道已选择数据文件，是否覆盖？", "确认对话框", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2);
                if (dr == DialogResult.Yes)
                {
                    FilePara.dataFileList.Remove(tBox_LoadDataFilePath_CH6.Text);
                }
                else
                {
                    return;
                }
            }
            string filePath = OpenDataFile();
            FilePara.dataFileList.Add(filePath);
            tBox_LoadDataFilePath_CH6.Text = filePath;
            scaleFactorList.Add(Convert.ToDouble(tBox_ScaleFactor_CH6.Text));
            HashSet<string> dataFileHashset = new HashSet<string>(FilePara.dataFileList);
            if (FilePara.dataFileList.Count() != dataFileHashset.Count())
            {
                FilePara.dataFileList.RemoveAt(FilePara.dataFileList.Count() - 1);
                scaleFactorList.RemoveAt(FilePara.dataFileList.Count() - 1);
                MessageBox.Show("这个文件已经选过了，请选择其他文件！");
            }
        }

        private string OpenDataFile()
        {
            string filePath = null;
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.InitialDirectory = FilePara.BaseDirectory;
            openFileDialog.DefaultExt = "dat";
            openFileDialog.Filter = "Data File(.dat)|*.dat";
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                filePath = openFileDialog.FileName;
            }
            while (filePath == null)
            {
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    filePath = openFileDialog.FileName;
                }
            }
            return filePath;
        }

        private void checkBox_CH1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox_CH1.Checked)
            {
                Btn_LoadData_CH1.Enabled = true;
            }
            else
            {
                Btn_LoadData_CH1.Enabled = false;
            }
        }

        private void checkBox_CH2_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox_CH2.Checked)
            {
                Btn_LoadData_CH2.Enabled = true;
            }
            else
            {
                Btn_LoadData_CH2.Enabled = false;
            }
        }

        private void checkBox_CH3_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox_CH3.Checked)
            {
                Btn_LoadData_CH3.Enabled = true;
            }
            else
            {
                Btn_LoadData_CH3.Enabled = false;
            }
        }

        private void LoadDataFileDialog_Load(object sender, EventArgs e)
        {
            if (checkBox_CH4.Checked)
            {
                Btn_LoadData_CH4.Enabled = true;
            }
            else
            {
                Btn_LoadData_CH4.Enabled = false;
            }
        }

        private void checkBox_CH5_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox_CH5.Checked)
            {
                Btn_LoadData_CH5.Enabled = true;
            }
            else
            {
                Btn_LoadData_CH5.Enabled = false;
            }
        }

        private void checkBox_CH6_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox_CH6.Checked)
            {
                Btn_LoadData_CH6.Enabled = true;
            }
            else
            {
                Btn_LoadData_CH6.Enabled = false;
            }
        }

        private void Btn_OK_Click(object sender, EventArgs e)
        {

        }
    }
}
