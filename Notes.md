using MSExcel = Microsoft.Office.Interop.Excel;
using MSWord = Microsoft.Office.Interop.Word;
using System.IO;
using System.Reflection;
using System.Diagnostics;
using Novacode;

打开C盘(目录)：
string fileName = @"C:\";
Process.Start(fileName);(打开某个目录同理)

打开文件：
string fileName = @"C:\Users\Marknoon\Desktop\AAA.txt";
Process.Start(fileName);

打开网站：
string fileName = @"http://www.baidu.com";
Process.Start(fileName);

var doc = DocX.Create(fileName);
doc.InsertParagraph("This is my first paragraph");
doc.Append("AAA");
doc.Save();

string fileName = @"C:\Users\Marknoon\Desktop\AAA.docx";
using (DocX doc = DocX.Create(fileName))
{
	Paragraph p = doc.InsertParagraph("This is my first paragraph!!!");
	p.Append("AAA");
	doc.Save();
	Process.Start(fileName);
}

string fileName = @"C:\Users\Marknoon\Desktop\AAA.docx";
DocX doc = DocX.Create(fileName);
Paragraph p = doc.InsertParagraph("This is my first paragraph");
p.Append("AAA");
doc.Save();
Process.Start(fileName);

//对于Excel的操作
MSExcel.Application excelApp;	//Excel应用程序变量
MSExcel.Workbook excelDoc;		//Excel文档变量
excelApp = new MSExcel.ApplicationClass();	//初始化
//如果已存在，则删除
if (File.Exists((string)path))
{
	File.Delete((string)path);
}
//由于使用的是COM库，因此有许多变量需要用Nothing代替
Object Nothing = Missing.Value;
excelDoc = excelApp.Workbooks.Add(Nothing);
//使用第一个工作表作为插入数据的工作表
MSExcel.Worksheet ws = (MSExcel.Worksheet)excelDoc.Sheets[1];
//声明一个MSExcel.Range 类型的变量r
MSExcel.Range r;
//获得A1到E3处的表格，并赋值
r = ws.get_Range("A1", "E3");
r.Value2 = "数据1";
//读取Excel
object xx = ws.Cells[1, "A"];
ws.Cells[7, "A"] = xx;
//WdSaveFormat为Excel文档的保存格式
object format = MSExcel.XlFileFormat.xlWorkbookNormal;
//将excelDoc文档对象的内容保存为XLSX文档
excelDoc.SaveAs(path, format, Nothing, Nothing, Nothing, Nothing, MSExcel.XlSaveAsAccessMode.xlExclusive, Nothing, Nothing, Nothing, Nothing, Nothing);
//关闭excelDoc文档对象
excelDoc.Close(Nothing, Nothing, Nothing);
//关闭excelApp组件对象
excelApp.Quit();
MessageBox.Show("创建成功！");
Process.Start(path);

操作Word
{
	//命名空间
	using Word = Microsoft.Office.Interop.Word;
	using System.Reflection;
	using System.IO;
	using System.Diagnostics;

	//初始化
	object path = @"E:\CCCXCXXTestDoc.doc";
	Word.Application word;
	Word.Document doc;
	object oMissing = Missing.Value;
	word = new Word.Application();
	word.Visible = false;
	doc = word.Documents.Add(ref path, ref oMissing, ref oMissing, ref oMissing);//引用模板
	
	//操作
	strContent += "XXX";
	doc.Paragraphs.Last.Range.Text = strContent;

	//释放资源
	object format = Word.WdSaveFormat.wdFormatDocument;
	doc.SaveAs(path, format,
		ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
		ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
		ref oMissing, ref oMissing, ref oMissing, ref oMissing);
	doc.Close(ref oMissing, ref oMissing, ref oMissing);
	word.Quit(ref oMissing, ref oMissing, ref oMissing);
	MessageBox.Show("创建成功！");
}

操作Excel
{
	//命名空间
	using Excel = Microsoft.Office.Interop.Excel;
	using System.Reflection;
	
	//初始化
	string str = string.Empty;
	object path = @"C:\Users\Marknoon\Desktop\梦幻西游炼妖.xls";
	object path2 = @"E:\aaa.xls";
	Excel.Application excelApp;	//Excel应用程序变量
	Excel.Workbook excelDoc;		//Excel文档变量
	excelApp = new Excel.Application();	//初始化
	object oMissing = Missing.Value;
	excelDoc = excelApp.Workbooks.Add(path);//引用模板
	
	//操作
	Excel.Worksheet ws = (Excel.Worksheet)excelDoc.Sheets[1];
	Excel.Range r;
	r = ws.get_Range("A1", "E3");
	r.Value2 = "数据1";
	object xx = ws.Cells[1, "A"];
	ws.Cells[7, "A"] = xx;
	
	//取出数据
	Excel.Range r;
	r = (Excel.Range)ws.Cells[1, 1];
	str += r.Value2.ToString();
	MessageBox.Show(str);
	
	//释放资源
	object format = Excel.XlFileFormat.xlWorkbookNormal;
	excelDoc.SaveAs(path2, format2, oMissing, oMissing, oMissing,
		oMissing, Excel.XlSaveAsAccessMode.xlExclusive, oMissing,
		oMissing, oMissing, oMissing, oMissing);
	excelDoc.Close(oMissing, oMissing, oMissing);
	excelApp.Quit();
	MessageBox.Show("创建成功！");
}

//选择文件
openFileDialog1.Filter = "BMP格式图片(*.bmp)|*.bmp|JPG格式图片(*.jpg)|*.jpg";
if (openFileDialog1.ShowDialog() == DialogResult.OK)  //点击"确定"按钮执行
{
	if (openFileDialog1.FileName != "")//图片路径赋值给textBox5
	{
		this.textBox5.Text = openFileDialog1.FileName;
	}
}

//选择目录
if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
{
	textBox1.Text = folderBrowserDialog1.SelectedPath;//显示目录
	//遍历并操作该目录下的对应文件
	
}

//遍历目录下的所有文件夹
if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
{
	//递归遍历该目录下的子目录
	string path = folderBrowserDialog1.SelectedPath;
	FindAndOpenDir(path);
}
public static void FindAndOpenDir(string path)
{
	if (Directory.Exists(path))
	{
		var subDirs = Directory.EnumerateDirectories(path);
		foreach (var subDir in subDirs)
		{
			DirectoryInfo dirInfo = new DirectoryInfo(subDir);
			MessageBox.Show(dirInfo.Name);
			FindAndOpenDir(subDir);
		}
	}
}

//遍历目录下的文件和子目录
public static void FindAndOpenDir(string path)
{
	if (Directory.Exists(path))
	{
		DirectoryInfo TheFolder = new DirectoryInfo(path);
		foreach (DirectoryInfo NextFolder in TheFolder.GetDirectories())
		{
			MessageBox.Show(NextFolder.Name);
		}
		foreach (FileInfo NextFile in TheFolder.GetFiles())
		{
			MessageBox.Show(NextFile.Name);
		}
	}
}

获取文件的后缀
DirectoryInfo TheFolder = new DirectoryInfo(path);
foreach (FileInfo NextFile in TheFolder.GetFiles())
{
	string name = Path.GetExtension(NextFile.Name);
	if (name == ".doc")
		MessageBox.Show(NextFile.Name);
}

doc.Content.Text;//获取文档的文本内容

//替换字符串
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
	MessageBox.Show("执行成功！");
}



