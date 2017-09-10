using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Diagnostics;
using DevExpress.Spreadsheet;
using System.Threading;

namespace 汇合表
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        IWorkbook workbook;
        Worksheet workSheet;
         List<string > listPath   = new List<string>();
        private void btnOpenFolder_Click(object sender, EventArgs e)
        {
            DialogResult = openFolder.ShowDialog();
            string folderPath = openFolder.SelectedPath;
            if (folderPath == "") return;
            lblPath.Text = folderPath;
            DirectoryInfo dirInfo = new DirectoryInfo(folderPath);
            FileInfo[] files = dirInfo.GetFiles("*.xls*", SearchOption.AllDirectories);//根目录下
            //FileInfo[] files = dirInfo.GetFiles();
            foreach (FileInfo file in files)
            {
              
                listPath.Add(file.FullName);//将文件夹里的所有文件路径加到listPath

            }
           
        }

        private void btn汇合_Click(object sender, EventArgs e)
        {
            string Path = @"F:\vs系统开发实验\参考\暴雨资料5历时频率分析\" + "总表.xlsx";
            spreadsheet.LoadDocument(Path);
            workbook = spreadsheet.Document;
            workSheet = (Worksheet)workbook.Worksheets[0];
            int j = 2;
            List<string> ListValues = new List<string>();
           
            foreach (string filePath in listPath)
            {
                spreadsheet2.LoadDocument(filePath);
                //Debug.Print(filePath);
                IWorkbook workbook2 = spreadsheet2.Document;
                Worksheet workSheet2 = (Worksheet)workbook2.Worksheets[0];
                string 十分 = workSheet2.Cells[1, 1].Value.ToString();
                //string value = workSheet2.Cells[1, z].Value.ToString();
               
                string fileName = System.IO.Path.GetFileNameWithoutExtension(filePath);//不带扩展名的文件名
                string sql =string.Format( "  select [测站站码],[流域] ,[雨量] ,[河名],[站名],[站别] from [山西省雨量站基本情况一览表$] where [站名]='{0}'",fileName );
                DataTable dt = db.GetTable(sql);
               
                if (dt.Rows.Count == 1)
                {
                    DataRow dr = dt.Rows[0];
                    string code = dr["测站站码"].ToString();
                    string 流域 = dr["流域"].ToString();
                    string 雨量 = dr["雨量"].ToString();
                    string 河名 = dr["河名"].ToString();
                    string 站名 = dr["站名"].ToString();
                    string 站别 = dr["站别"].ToString();
                    workSheet.Cells[j, 1].Value = code;
                    workSheet.Cells[j, 12].Value = 流域;
                    workSheet.Cells[j, 13].Value = 雨量;
                    workSheet.Cells[j, 14].Value = 河名;
                    workSheet.Cells[j, 15].Value = 站名;
                    workSheet.Cells[j, 16].Value = 站别;
                  
                }
                workSheet.Cells[j, 0].Value = fileName;//填入站名
                if (十分 == "60min")
                {


                    string 六十分均值 = workSheet2.Cells[2, 1].DisplayText;
                    string 六十分均值format = workSheet2.Cells[2, 1].NumberFormat;

                    string 六十分Cv值 = workSheet2.Cells[3, 1].DisplayText;
                    string 六十分Cv值format = workSheet2.Cells[3, 1].NumberFormat;


                    string 六小时均值 = workSheet2.Cells[2, 2].DisplayText;
                    string 六小时均值format = workSheet2.Cells[2, 2].NumberFormat;

                    string 六小时Cv值 = workSheet2.Cells[3, 2].DisplayText;
                    string 六小时Cv值format = workSheet2.Cells[3, 2].NumberFormat;


                    string 二十四均值 = workSheet2.Cells[2, 3].DisplayText;
                    string 二十四均值format = workSheet2.Cells[2, 3].NumberFormat;

                    string 二十四Cv值 = workSheet2.Cells[3, 3].DisplayText;
                    string 二十四Cv值format = workSheet2.Cells[3, 3].NumberFormat;


                    string 三天均值 = workSheet2.Cells[2, 4].DisplayText;
                    string 三天均值format = workSheet2.Cells[2, 4].NumberFormat;

                    string 三天Cv值 = workSheet2.Cells[3, 4].DisplayText;
                    string 三天Cv值format = workSheet2.Cells[3, 4].NumberFormat;



                    workSheet.Cells[j, 4].Value = double.Parse(六十分均值);
                    workSheet.Cells[j, 4].NumberFormat = 六十分均值format;
                    workSheet.Cells[j, 5].Value = double.Parse(六十分Cv值);
                    workSheet.Cells[j, 5].NumberFormat = 六十分Cv值format;

                    workSheet.Cells[j, 6].Value = double.Parse(六小时均值);
                    workSheet.Cells[j, 6].NumberFormat = 六小时均值format;
                    workSheet.Cells[j, 7].Value = double.Parse(六小时Cv值);
                    workSheet.Cells[j, 7].NumberFormat = 六小时Cv值format;

                    workSheet.Cells[j, 8].Value = double.Parse(二十四均值);
                    workSheet.Cells[j, 8].NumberFormat = 二十四均值format;

                    workSheet.Cells[j, 9].Value = double.Parse(二十四Cv值);
                    workSheet.Cells[j, 9].NumberFormat = 二十四Cv值format;

                    workSheet.Cells[j, 10].Value = double.Parse(三天均值);
                    workSheet.Cells[j, 10].NumberFormat = 三天均值format;

                    workSheet.Cells[j, 11].Value = double.Parse(三天Cv值);
                    workSheet.Cells[j, 11].NumberFormat = 三天Cv值format;
                }
                else if(十分=="10min")
                {
                    string 十分均值 = workSheet2.Cells[2, 1].DisplayText;
                    string 十分均值format = workSheet2.Cells[2, 1].NumberFormat;

                    string 十分Cv值 = workSheet2.Cells[3, 1].DisplayText;
                    string 十分Cv值format = workSheet2.Cells[3, 1].NumberFormat;


                    string 六十分均值 = workSheet2.Cells[2, 2].DisplayText;
                    string 六十分均值format = workSheet2.Cells[2, 2].NumberFormat;

                    string 六十分Cv值 = workSheet2.Cells[3, 2].DisplayText;
                    string 六十分Cv值format = workSheet2.Cells[3, 2].NumberFormat;


                    string 六小时均值 = workSheet2.Cells[2, 3].DisplayText;
                    string 六小时均值format = workSheet2.Cells[2, 3].NumberFormat;

                    string 六小时Cv值 = workSheet2.Cells[3, 3].DisplayText;
                    string 六小时Cv值format = workSheet2.Cells[3, 3].NumberFormat;


                    string 二十四均值 = workSheet2.Cells[2, 4].DisplayText;
                    string 二十四均值format = workSheet2.Cells[2, 4].NumberFormat;

                    string 二十四Cv值 = workSheet2.Cells[3, 4].DisplayText;
                    string 二十四Cv值format = workSheet2.Cells[3,4].NumberFormat;


                    string 三天均值 = workSheet2.Cells[2, 5].DisplayText;
                    string 三天均值format = workSheet2.Cells[2, 5].NumberFormat;

                    string 三天Cv值 = workSheet2.Cells[3, 5].DisplayText;
                    string 三天Cv值format = workSheet2.Cells[3,5].NumberFormat;


                    workSheet.Cells[j, 2].Value =  double.Parse(十分均值);
                    workSheet.Cells[j, 2].NumberFormat = 十分均值format;
                    workSheet.Cells[j, 3].Value =  double.Parse(十分Cv值);
                    workSheet.Cells[j, 3].NumberFormat = 十分Cv值format;

                    workSheet.Cells[j, 4].Value =  double.Parse(六十分均值);
                    workSheet.Cells[j, 4].NumberFormat = 六十分均值format;

                    workSheet.Cells[j, 5].Value =  double.Parse(六十分Cv值);
                    workSheet.Cells[j, 5].NumberFormat = 六十分均值format;


                    workSheet.Cells[j, 6].Value =  double.Parse(六小时均值);
                    workSheet.Cells[j, 6].NumberFormat = 六小时均值format;

                    workSheet.Cells[j, 7].Value =  double.Parse(六小时Cv值);
                    workSheet.Cells[j, 7].NumberFormat = 六小时Cv值format;


                    workSheet.Cells[j, 8].Value =  double.Parse(二十四均值);
                    workSheet.Cells[j, 8].NumberFormat = 二十四均值format;

                    workSheet.Cells[j, 9].Value =  double.Parse(二十四Cv值);
                    workSheet.Cells[j, 9].NumberFormat = 二十四Cv值format;

                    double zhi;
                    double.TryParse(三天均值, out zhi);
                    workSheet.Cells[j, 10].Value = zhi;
                    workSheet.Cells[j, 10].NumberFormat = 三天均值format;

                    double zhi1;
                    double.TryParse(三天Cv值, out zhi1);
                    workSheet.Cells[j, 11].Value = zhi1;
                    workSheet.Cells[j, 11].NumberFormat = 三天Cv值format;

                  
                }
                else if (十分=="6h")
                {
                    string 六小时均值 = workSheet2.Cells[2, 1].DisplayText;
                    string 六小时均值format = workSheet2.Cells[2, 1].NumberFormat;

                    string 六小时Cv值 = workSheet2.Cells[3, 1].DisplayText;
                    string 六小时Cv值format = workSheet2.Cells[3, 1].NumberFormat;


                    string 二十四均值 = workSheet2.Cells[2, 2].DisplayText;
                    string 二十四均值format = workSheet2.Cells[2, 2].NumberFormat;

                    string 二十四Cv值 = workSheet2.Cells[3, 2].DisplayText;
                    string 二十四Cv值format = workSheet2.Cells[3, 2].NumberFormat;


                    string 三天均值 = workSheet2.Cells[2, 3].DisplayText;
                    string 三天均值format = workSheet2.Cells[2, 3].NumberFormat;

                    string 三天Cv值 = workSheet2.Cells[3, 3].DisplayText;
                    string 三天Cv值format = workSheet2.Cells[3, 3].NumberFormat;


                    workSheet.Cells[j, 6].Value =  double.Parse(六小时均值);
                    workSheet.Cells[j, 2].NumberFormat = 六小时均值format;

                    workSheet.Cells[j, 7].Value =  double.Parse(六小时Cv值);
                    workSheet.Cells[j, 2].NumberFormat = 六小时Cv值format;


                    workSheet.Cells[j, 8].Value =  double.Parse(二十四均值);
                    workSheet.Cells[j, 2].NumberFormat = 二十四均值format;

                    workSheet.Cells[j, 9].Value =  double.Parse(二十四Cv值);
                    workSheet.Cells[j, 2].NumberFormat = 二十四Cv值format;


                    workSheet.Cells[j, 10].Value =  double.Parse(三天均值);
                    workSheet.Cells[j, 2].NumberFormat = 三天均值format;

                    workSheet.Cells[j, 11].Value =  double.Parse(三天Cv值);
                    workSheet.Cells[j, 2].NumberFormat = 三天Cv值format;


                }
                
                j++;
            }
            workbook.SaveDocument(@"C:\Users\henrik\Desktop\暴雨资料5历时频率分析\新总表.xlsx");
        }

        private void btnSave_Click(object sender, EventArgs e)
        {

        }

        

        
    }
}
