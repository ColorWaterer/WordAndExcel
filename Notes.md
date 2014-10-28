using MSExcel = Microsoft.Office.Interop.Excel;
using MSWord = Microsoft.Office.Interop.Word;
using System.IO;
using System.Reflection;
using System.Diagnostics;
using Novacode;

��C��(Ŀ¼)��
string fileName = @"C:\";
Process.Start(fileName);(��ĳ��Ŀ¼ͬ��)

���ļ���
string fileName = @"C:\Users\Marknoon\Desktop\AAA.txt";
Process.Start(fileName);

����վ��
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

//����Excel�Ĳ���
MSExcel.Application excelApp;	//ExcelӦ�ó������
MSExcel.Workbook excelDoc;		//Excel�ĵ�����
excelApp = new MSExcel.ApplicationClass();	//��ʼ��
//����Ѵ��ڣ���ɾ��
if (File.Exists((string)path))
{
	File.Delete((string)path);
}
//����ʹ�õ���COM�⣬�������������Ҫ��Nothing����
Object Nothing = Missing.Value;
excelDoc = excelApp.Workbooks.Add(Nothing);
//ʹ�õ�һ����������Ϊ�������ݵĹ�����
MSExcel.Worksheet ws = (MSExcel.Worksheet)excelDoc.Sheets[1];
//����һ��MSExcel.Range ���͵ı���r
MSExcel.Range r;
//���A1��E3���ı�񣬲���ֵ
r = ws.get_Range("A1", "E3");
r.Value2 = "����1";
//��ȡExcel
object xx = ws.Cells[1, "A"];
ws.Cells[7, "A"] = xx;
//WdSaveFormatΪExcel�ĵ��ı����ʽ
object format = MSExcel.XlFileFormat.xlWorkbookNormal;
//��excelDoc�ĵ���������ݱ���ΪXLSX�ĵ�
excelDoc.SaveAs(path, format, Nothing, Nothing, Nothing, Nothing, MSExcel.XlSaveAsAccessMode.xlExclusive, Nothing, Nothing, Nothing, Nothing, Nothing);
//�ر�excelDoc�ĵ�����
excelDoc.Close(Nothing, Nothing, Nothing);
//�ر�excelApp�������
excelApp.Quit();
MessageBox.Show("�����ɹ���");
Process.Start(path);

����Word
{
	//�����ռ�
	using Word = Microsoft.Office.Interop.Word;
	using System.Reflection;
	using System.IO;
	using System.Diagnostics;

	//��ʼ��
	object path = @"E:\CCCXCXXTestDoc.doc";
	Word.Application word;
	Word.Document doc;
	object oMissing = Missing.Value;
	word = new Word.Application();
	word.Visible = false;
	doc = word.Documents.Add(ref path, ref oMissing, ref oMissing, ref oMissing);//����ģ��
	
	//����
	strContent += "XXX";
	doc.Paragraphs.Last.Range.Text = strContent;

	//�ͷ���Դ
	object format = Word.WdSaveFormat.wdFormatDocument;
	doc.SaveAs(path, format,
		ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
		ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
		ref oMissing, ref oMissing, ref oMissing, ref oMissing);
	doc.Close(ref oMissing, ref oMissing, ref oMissing);
	word.Quit(ref oMissing, ref oMissing, ref oMissing);
	MessageBox.Show("�����ɹ���");
}

����Excel
{
	//�����ռ�
	using Excel = Microsoft.Office.Interop.Excel;
	using System.Reflection;
	
	//��ʼ��
	string str = string.Empty;
	object path = @"C:\Users\Marknoon\Desktop\�λ���������.xls";
	object path2 = @"E:\aaa.xls";
	Excel.Application excelApp;	//ExcelӦ�ó������
	Excel.Workbook excelDoc;		//Excel�ĵ�����
	excelApp = new Excel.Application();	//��ʼ��
	object oMissing = Missing.Value;
	excelDoc = excelApp.Workbooks.Add(path);//����ģ��
	
	//����
	Excel.Worksheet ws = (Excel.Worksheet)excelDoc.Sheets[1];
	Excel.Range r;
	r = ws.get_Range("A1", "E3");
	r.Value2 = "����1";
	object xx = ws.Cells[1, "A"];
	ws.Cells[7, "A"] = xx;
	
	//ȡ������
	Excel.Range r;
	r = (Excel.Range)ws.Cells[1, 1];
	str += r.Value2.ToString();
	MessageBox.Show(str);
	
	//�ͷ���Դ
	object format = Excel.XlFileFormat.xlWorkbookNormal;
	excelDoc.SaveAs(path2, format2, oMissing, oMissing, oMissing,
		oMissing, Excel.XlSaveAsAccessMode.xlExclusive, oMissing,
		oMissing, oMissing, oMissing, oMissing);
	excelDoc.Close(oMissing, oMissing, oMissing);
	excelApp.Quit();
	MessageBox.Show("�����ɹ���");
}

//ѡ���ļ�
openFileDialog1.Filter = "BMP��ʽͼƬ(*.bmp)|*.bmp|JPG��ʽͼƬ(*.jpg)|*.jpg";
if (openFileDialog1.ShowDialog() == DialogResult.OK)  //���"ȷ��"��ťִ��
{
	if (openFileDialog1.FileName != "")//ͼƬ·����ֵ��textBox5
	{
		this.textBox5.Text = openFileDialog1.FileName;
	}
}

//ѡ��Ŀ¼
if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
{
	textBox1.Text = folderBrowserDialog1.SelectedPath;//��ʾĿ¼
	//������������Ŀ¼�µĶ�Ӧ�ļ�
	
}

//����Ŀ¼�µ������ļ���
if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
{
	//�ݹ������Ŀ¼�µ���Ŀ¼
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

//����Ŀ¼�µ��ļ�����Ŀ¼
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

��ȡ�ļ��ĺ�׺
DirectoryInfo TheFolder = new DirectoryInfo(path);
foreach (FileInfo NextFile in TheFolder.GetFiles())
{
	string name = Path.GetExtension(NextFile.Name);
	if (name == ".doc")
		MessageBox.Show(NextFile.Name);
}

doc.Content.Text;//��ȡ�ĵ����ı�����

//�滻�ַ���
string oldString = "�ɼ�������			";
string newString = "�ɼ�������" + res + "			";
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
	MessageBox.Show("ִ�гɹ���");
}



