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
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;

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
        private List<int> old_column_stas, new_column_stas, old_row_stas, new_row_stas;//行、列状态数组-1：另一边不存在（多余）;其余数字:另一边存在(直接标记对应下标)
        private int file_type1, file_type2;
        //List<int> same_old, same_new;

        private void 选择被比较文件ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            /*选择并打开文件，读取数据，关闭文件*/
            OpenFileDialog dialog = new OpenFileDialog
            {
                Multiselect = false,
                Title = "请选择一个符合格式的excel表格",
                Filter = "Excel文件(*.xlsx)|*.xlsx|旧版本Excel文件(*.xls)|*.xls"
            };
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                new_file_name = dialog.FileName;
                file_type2 = dialog.FilterIndex;
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
                Filter = "Excel文件(*.xlsx)|*.xlsx|旧版本Excel文件(*.xls)|*.xls"
            };
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                old_file_name = dialog.FileName;
                file_type1 = dialog.FilterIndex;
            }
            dialog.Dispose();
            label4.Text += old_file_name;
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
            else if (new_file_name == "")
            {
                input_ok = false;
                MessageBox.Show("清选择待比较文件！");
            }
            else if (out_file_folder == "")
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
            if (file_type1 == 1)
            {
                out_file1 = out_file_folder + "\\file1.xlsx";
            }
            else
            {
                out_file1 = out_file_folder + "\\file1.xls";
            }
            if (file_type2 == 1)
            {
                out_file2 = out_file_folder + "\\file2.xlsx";
            }
            else
            {
                out_file2 = out_file_folder + "\\file2.xls";
            }
            //删掉原有文件
            if (File.Exists(out_file1))
                File.Delete(out_file1);
            if (File.Exists(out_file2))
                File.Delete(out_file2);
            //复制文件
            File.Copy(old_file_name, out_file1);
            File.Copy(new_file_name, out_file2);
            return;
        }
        private void 开始ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            #region 变量
            /*变量区*/
            int old_col_num, new_col_num, old_row_num, new_row_num, i, j, tmp_sta, tmp_col, tmp_row;
            string tmp_column_name, tmp_row_name, tmp_cell, tmp_cell2;
            FileStream fs1, fs2;
            IWorkbook wk1, wk2;
            ISheet sheet1, sheet2;
            IRow row1, row2;
            ICellStyle s_add1, s_add2, s_dif1, s_dif2;//多余格子的格式；不同各自的格式
            #endregion
            #region 输入判断以及变量预处理
            if (!Judge_Input_Full())//判断输入
                return;
            Copy_Original_Files();//先复制旧的文件到新的文件夹中
            //打开文件,获取基础参数
            fs1 = File.OpenRead(out_file1);
            fs2 = File.OpenRead(out_file2);
            if (file_type1 == 1)
            {
                wk1 = new XSSFWorkbook(fs1);
            }
            else
            {
                wk1 = new HSSFWorkbook(fs1);
            }
            if (file_type2 == 1)
            {
                wk2 = new XSSFWorkbook(fs2);
            }
            else
            {
                wk2 = new HSSFWorkbook(fs2);
            }
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
            //设置单元格格式
            s_add1 = wk1.CreateCellStyle();
            s_add1.FillForegroundColor = HSSFColor.LightBlue.Index;
            s_add1.FillPattern = FillPattern.SolidForeground;
            s_add2 = wk2.CreateCellStyle();
            s_add2.FillForegroundColor = HSSFColor.LightBlue.Index;
            s_add2.FillPattern = FillPattern.SolidForeground;
            s_dif1 = wk1.CreateCellStyle();
            s_dif1.FillForegroundColor = HSSFColor.LightOrange.Index;
            s_dif1.FillPattern = FillPattern.SolidForeground;
            s_dif2 = wk2.CreateCellStyle();
            s_dif2.FillForegroundColor = HSSFColor.LightOrange.Index;
            s_dif2.FillPattern = FillPattern.SolidForeground;
            #endregion
            #region 处理不同的列
            //获取列名称
            for (i = 0; i < old_col_num; i++)
                old_colums.Add(row1.GetCell(i).ToString());
            for (i = 0; i < new_col_num; i++)
                new_colums.Add(row2.GetCell(i).ToString());
            //比较列名，要获取列之间的对应关系，以及列的对应状态
            //旧列对应新列状态
            for (i = 0; i < old_col_num; i++)
            {
                tmp_column_name = old_colums[i];
                tmp_sta = new_colums.IndexOf(tmp_column_name);
                old_column_stas.Add(tmp_sta);//不存在则返回-1,否则返回下标
                if (tmp_sta == -1)
                {
                    for (j = 0; j < old_row_num; j++)//第一行也考虑在内
                    {
                        row1 = sheet1.GetRow(j);
                        row1.GetCell(i).CellStyle = s_add1;
                    }
                }
            }
            //新列对应旧列状态
            for (i = 0; i < new_col_num; i++)
            {
                tmp_column_name = new_colums[i];
                tmp_sta = old_colums.IndexOf(tmp_column_name);
                new_column_stas.Add(tmp_sta);
                if (tmp_sta == -1)
                {
                    for (j = 0; j < new_row_num; j++)
                    {
                        row2 = sheet2.GetRow(j);
                        row2.GetCell(i).CellStyle = s_add2;
                    }
                }
            }
            #endregion
            #region 处理不同行
            //获取行名称,第一行不是数据不需要获取
            for (i = 0; i < old_row_num; i++)
            {
                row1 = sheet1.GetRow(i);
                old_rows.Add(row1.GetCell(key_id_row_num_old).ToString());//主码名称
            }
            for (i = 0; i < new_row_num; i++)
            {
                row2 = sheet2.GetRow(i);
                new_rows.Add(row2.GetCell(key_id_row_num_new).ToString());//主码名称
            }
            //旧行对应新行状态
            for (i = 0; i < old_row_num; i++)
            {
                tmp_row_name = old_rows[i];
                tmp_sta = new_rows.IndexOf(tmp_row_name);
                old_row_stas.Add(tmp_sta);
                if (tmp_sta == -1)//不存在对应行
                {
                    row1 = sheet1.GetRow(i);
                    for (j = 0; j < old_col_num; j++)
                    {
                        row1.GetCell(j).CellStyle = s_add1;
                    }
                }
            }
            //新行对应旧行状态
            for (i = 0; i < new_row_num; i++)
            {
                tmp_row_name = new_rows[i];
                tmp_sta = old_rows.IndexOf(tmp_row_name);
                new_row_stas.Add(tmp_sta);
                if (tmp_sta == -1)
                {
                    row2 = sheet2.GetRow(i);
                    for (j = 0; j < old_col_num; j++)
                    {
                        row2.GetCell(j).CellStyle = s_add2;
                    }
                }
            }
            #endregion
            #region 处理不同单元格
            //旧表格
            for (i = 1; i < old_row_num; i++)
            {
                tmp_row = old_row_stas[i];
                if (tmp_row != -1)
                {
                    row1 = sheet1.GetRow(i);
                    row2 = sheet2.GetRow(tmp_row);
                    for (j = 0; j < old_col_num; j++)
                    {
                        if (j != key_id_row_num_old)//非id列
                        {
                            tmp_col = old_column_stas[j];
                            if (tmp_col != -1)
                            {
                                tmp_cell = row1.GetCell(j).ToString();
                                tmp_cell2 = row2.GetCell(tmp_col).ToString();
                                if (string.Compare(tmp_cell, tmp_cell2) != 0)
                                {
                                    row1.GetCell(j).CellStyle = s_dif1;
                                    row2.GetCell(tmp_col).CellStyle = s_dif2;
                                }
                            }
                        }
                    }
                }
            }
            #endregion
            //保存文件
            fs1 = new FileStream(out_file1, FileMode.Open, FileAccess.Write);
            wk1.Write(fs1);
            fs1.Close();
            fs2 = new FileStream(out_file2, FileMode.Open, FileAccess.Write);
            wk2.Write(fs2);
            fs2.Close();
            MessageBox.Show("比较完成!");
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
            old_column_stas = new List<int>();
            new_column_stas = new List<int>();
            old_row_stas = new List<int>();
            new_row_stas = new List<int>();
            file_type1 = 1;
            file_type2 = 1;
        }
    }
}
