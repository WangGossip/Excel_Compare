using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace Excel_Compare
{
    public partial class Form1 : Form
    {
        /*常量*/
        private static int key_id_row_num_new = 0;//原始文件主码所在列
        private static int key_id_row_num_old = 0;//新文件主码所在列
        /*变量区*/
        private string old_file_name, new_file_name, out_file_folder, out_file1, out_file2;
        private List<string> old_colums, new_colums, old_rows, new_rows;//行、列名称
        private List<int> old_column_stas, new_column_stas, old_row_stas, new_row_stas;//行、列状态数组-1：另一边不存在（多余）;1:另一边存在(直接标记对应下标)
        //List<int> same_old, same_new;

        private void 选择被比较文件ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            /*选择并打开文件，读取数据，关闭文件*/
            OpenFileDialog dialog = new OpenFileDialog
            {
                Multiselect = false,
                Title = "请选择一个符合格式的excel表格",
                Filter = "Excel文件(*.xlsx)|*.xlsx"
            };
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                new_file_name = dialog.FileName;
            }
            dialog.Dispose();
            label6.Text = new_file_name;
        }

        private void 选择原始文件ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            /*选择并打开文件，读取数据，关闭文件*/
            OpenFileDialog dialog = new OpenFileDialog
            {
                Multiselect = false,
                Title = "请选择一个符合格式的excel表格",
                Filter = "Excel文件(*.xlsx)|*.xlsx"
            };
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                old_file_name = dialog.FileName;
            }
            dialog.Dispose();
            label4.Text = new_file_name;
        }

        private void 使用帮助ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("");
        }

        private void 选择输出文件夹ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog dialog = new FolderBrowserDialog();
            dialog.Description = "请选择目标文件夹";
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                if (string.IsNullOrEmpty(dialog.SelectedPath))
                {
                    MessageBox.Show(this, "文件夹路径不能为空", "提示");
                    return;
                }
            }
            out_file_folder = dialog.SelectedPath;
            dialog.Dispose();
            label2.Text = out_file_folder;
        }
        
        protected bool Judge_Input_Full()
        {
            bool input_ok = true;
            if (old_file_name == "")
            {
                input_ok = false;
                MessageBox.Show("清选择原始文件！");
            }
            else if(new_file_name=="")
            {
                input_ok = false;
                MessageBox.Show("清选择待比较文件！");
            }
            else if(out_file_folder=="")
            {
                input_ok = false;
                MessageBox.Show("清选择输出文件夹！");
            }
            return input_ok;
        }

        /// <summary>
        /// 将资源文件复制到目标文件夹，进行修改
        /// </summary>
        protected void Copy_Original_Files()
        {
            out_file1 = out_file_folder + "\\file1.xlsx";
            out_file2 = out_file_folder + "\\file2.xlsx";
            //删掉原有文件
            if (File.Exists(out_file1))
                File.Delete(out_file1);
            if (File.Exists(out_file2))
                File.Delete(out_file2);
            //复制文件
            File.Copy(old_file_name, out_file1);
            File.Copy(new_file_name, out_file2);
            return ;
        }
        private void 开始ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            /*变量区*/
            int old_col_num, new_col_num, old_row_num, new_row_num, i, j;
            FileStream fs1, fs2;
            IWorkbook wk1,wk2;
            ISheet sheet1, sheet2;
            IRow row1, row2;
            if (!Judge_Input_Full())//判断输入
                return ;
            Copy_Original_Files();//先复制旧的文件到新的文件夹中
            //打开文件,获取基础参数
            fs1 = File.OpenRead(out_file1);
            fs2 = File.OpenRead(out_file2);
            wk1 = new XSSFWorkbook(fs1);
            wk2 = new XSSFWorkbook(fs2);
            fs1.Close();
            fs2.Close();
            sheet1 = wk1.GetSheetAt(0);
            sheet2 = wk2.GetSheetAt(0);
            old_row_num = sheet1.LastRowNum + 1;//总行数1
            new_row_num = sheet2.LastRowNum + 1;//总行数2
            row1 = sheet1.GetRow(0);
            row2 = sheet2.GetRow(0);
            old_col_num = row1.LastCellNum;//总列数1
            new_col_num = row2.LastCellNum;//总列数2
            //获取列名称
            for (i = 0; i < old_col_num; i++)
                old_colums.Add(row1.GetCell(i).ToString());
            for (i = 0; i < new_col_num; i++)
                new_colums.Add(row2.GetCell(i).ToString());
            //比较列名
        }

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            old_colums = new List<string>();
            new_colums = new List<string>();
            old_rows = new List<string>();
            new_rows = new List<string>();
        }
    }
}
