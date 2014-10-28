using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using System.Reflection;
using System.IO;
using System.Diagnostics;
using System.Text.RegularExpressions;

namespace WordAndExcel
{
    public partial class Form1 : Form
    {
        static string num = string.Empty;       //学号
        static string name = string.Empty;      //姓名
        static string exp = string.Empty;       //实验
        static string res = string.Empty;       //成绩
        static string fileName = string.Empty;  //文件名
        static bool isFind = false;             //是否找到文件

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            openFileDialog1.Filter = "Excel2003(*.xls)|*.xls|Excel2007(*.xlsx)|*.xlsx";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)  //点击"确定"按钮执行
            {
                if (openFileDialog1.FileName != "")
                {
                    this.textBox1.Text = openFileDialog1.FileName;
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                this.textBox2.Text = folderBrowserDialog1.SelectedPath;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.FileName == "openFileDialog1")
            {
                MessageBox.Show("未选择Excel文档！");
                return;
            }
            if (folderBrowserDialog1.SelectedPath == "")
            {
                MessageBox.Show("未选择Word目录！");
                return;
            }

            //Excel表范围
            int startRow = 2;
            int endRow = 43;
            int startColumn = 8;
            int endColumn = 9;

            //初始化
            string path = folderBrowserDialog1.SelectedPath;
            object excelPath = openFileDialog1.FileName;
            Excel.Application excelApp;	                //Excel应用程序变量
            Excel.Workbook excelDoc;		            //Excel文档变量
            excelApp = new Excel.Application();
            object oMissing = Missing.Value;
            excelDoc = excelApp.Workbooks.Add(excelPath);    //引用模板

            if (File.Exists((string)excelPath))
            {
                File.Delete((string)excelPath);
            }

            //遍历Excel里的数据
            for (int i = startRow; i <= endRow; i++)
            {
                Excel.Worksheet ws = (Excel.Worksheet)excelDoc.Sheets[1];
                Excel.Range r;

                r = (Excel.Range)ws.Cells[i, 1];
                if (r.Value2 == null) continue;
                num = r.Value2.ToString();
                r = (Excel.Range)ws.Cells[i, 2];
                if (r.Value2 == null) continue;
                name = r.Value2.ToString();

                //遍历实验
                for (int j = startColumn; j <= endColumn; j++)
                {
                    r = (Excel.Range)ws.Cells[1, j];
                    exp = r.Value2.ToString();
                    if (r.Value2 == null) continue;
                    r = (Excel.Range)ws.Cells[i, j];
                    res = r.Value2.ToString();
                    if (r.Value2 == null) continue;
                    fileName = num + "_" + exp + "_" + name + ".doc";
                    isFind = false;

                    FindAndOpenDir(path);//从指定的目录中寻找

                    //if (isFind == false)
                    //{
                    //    MessageBox.Show("未找到" + fileName);
                    //}
                }
            }

            //释放资源
            object format = Excel.XlFileFormat.xlWorkbookNormal;
            excelDoc.SaveAs(excelPath, format, oMissing, oMissing, oMissing,
                oMissing, Excel.XlSaveAsAccessMode.xlExclusive, oMissing,
                oMissing, oMissing, oMissing, oMissing);
            excelDoc.Close(oMissing, oMissing, oMissing);
            excelApp.Quit();

            MessageBox.Show("已完成所有操作！");
        }

        public static void FindAndOpenDir(string path)
        {
            if (Directory.Exists(path))
            {
                DirectoryInfo TheFolder = new DirectoryInfo(path);
                foreach (FileInfo NextFile in TheFolder.GetFiles())
                {
                    //string file = path;
                    //file += "\\";
                    //file += NextFile.Name;//file为搜索到的文件的目录+文件名
                    //string suffix = Path.GetExtension(NextFile.Name);
                    //if (suffix == ".doc")
                    //{
                    //    Regex regname = new Regex(name);
                    //    Match matname = regname.Match(NextFile.Name);
                    //    Regex regexp = new Regex(exp);
                    //    Match matexp = regexp.Match(NextFile.Name);
                    //    if (matexp.Success && matname.Success)
                    //    {
                    //        MessageBox.Show(NextFile.Name);
                    //    }
                    //}
                    if (fileName == NextFile.Name)
                    {
                        //初始化
                        isFind = true;
                        string file = path;
                        file += "\\";
                        file += NextFile.Name;//file为搜索到的文件的目录+文件名
                        object filePath = file;
                        Word.Application word;
                        Word.Document doc;
                        object oMissing = Missing.Value;
                        word = new Word.Application();
                        word.Visible = false;
                        doc = word.Documents.Add(ref filePath, ref oMissing, ref oMissing, ref oMissing);

                        //操作
                        string oldString = "成绩评定：			";
                        string newString = "成绩评定：" + res + "			";
                        doc.Content.Find.Text = oldString;
                        object FindText = oldString;
                        object ReplaceWith = newString;
                        object Replace = Word.WdReplace.wdReplaceAll;
                        doc.Content.Find.ClearFormatting();
                        if (doc.Content.Find.Execute(
                            ref FindText, ref oMissing,
                            ref oMissing, ref oMissing,
                            ref oMissing, ref oMissing,
                            ref oMissing, ref oMissing, ref oMissing,
                            ref ReplaceWith, ref Replace,
                            ref oMissing, ref oMissing,
                            ref oMissing, ref oMissing))
                        {
                            
                        }

                        //释放资源
                        object format = Word.WdSaveFormat.wdFormatDocument;
                        doc.SaveAs(filePath, format,
                            ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                            ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                            ref oMissing, ref oMissing, ref oMissing, ref oMissing);
                        doc.Close(ref oMissing, ref oMissing, ref oMissing);
                        word.Quit(ref oMissing, ref oMissing, ref oMissing);
                        return;
                    }
                }
                var subDirs = Directory.EnumerateDirectories(path);
                foreach (var subDir in subDirs)
                {
                    DirectoryInfo dirInfo = new DirectoryInfo(subDir);
                    FindAndOpenDir(subDir);
                }
            }
        }
    }
}
