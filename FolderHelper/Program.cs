using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;

namespace FolderHelper
{
    class Program
    {
        static void Main(string[] args)
        {
            string model = ConfigurationSettings.AppSettings["Model"];
            string sourcePath = ConfigurationSettings.AppSettings["SourcePath"];
            string targetPath = AppDomain.CurrentDomain.BaseDirectory + "/FileList.xlsx";

            var helper = new FolderHelper();

            Console.WriteLine("羅伊德：阿尼亞，該寫作業了!");
            Console.WriteLine("阿尼亞：@@");

            if (model == "1")
            {
                Console.WriteLine("羅伊德：題目是：讀取指定路徑下所有檔名，並匯出FileList.xlsx");
                Console.WriteLine("阿尼亞：T_T");
                Console.WriteLine("...");
                Console.WriteLine("...");
                Console.WriteLine("...");
                helper.GetAllFilesNameExportExcel(sourcePath, targetPath);
            }
            else if (model == "2")
            {
                Console.WriteLine("羅伊德：題目是：讀取FileList.xlsx並依內容將檔案重新命名");
                Console.WriteLine("阿尼亞：Q_Q");
                Console.WriteLine("...");
                Console.WriteLine("...");
                Console.WriteLine("...");
                helper.RenameFileFromExcel(targetPath);
                
            }
            else
            {
                Console.WriteLine("羅伊德：恩!?今天沒有作業");
            }

            Console.WriteLine("約  兒：這樣的話，我來做晚餐吧!!");
            Console.WriteLine("羅伊德：!!");
            Console.WriteLine("阿尼亞：!!");
            Console.WriteLine("龐  德：!!");
            Console.ReadKey();
        }
    }

    class FolderHelper
    {
        DataTable baseDT;

        public void GetAllFilesNameExportExcel(string folderPath, string excelFile)
        {
            try
            {
                if (File.Exists(excelFile))
                    File.Delete(excelFile);

                setDataTableSchema();
                dirSearch(folderPath);
                Excel.ExportExcelFromDataTable(baseDT, excelFile);

                Console.WriteLine("羅伊德：好厲害，完成了!!");
                Console.WriteLine("阿尼亞：V_V");
            }
            catch (Exception ex)
            {
                Console.WriteLine("");
                Console.WriteLine("阿尼亞不會寫程式，把這個問題交給父親");
                Console.WriteLine(ex.Message);
                Console.WriteLine("");
            }
        }

        public void RenameFileFromExcel(string excelPath)
        {
            try
            {
                baseDT = Excel.LoadExcelAsDataTable(excelPath);
                renameFile(baseDT);

                Console.WriteLine("羅伊德：好厲害，完成了!!");
                Console.WriteLine("阿尼亞：V_V");
            }
            catch (Exception ex)
            {
                Console.WriteLine("");
                Console.WriteLine("阿尼亞不會寫程式，把這個問題交給父親");
                Console.WriteLine(ex.Message);
                Console.WriteLine("");
            }
        }

        private void dirSearch(string dir)
        {
            foreach (var d in Directory.GetDirectories(dir))
                dirSearch(d);

            var dd = new DirectoryInfo(dir);
            var fileArray = dd.GetFiles();

            foreach (var f in fileArray)
            {
                var row = baseDT.NewRow();
                row["Path"] = f.FullName;
                row["FileName"] = f.Name;
                baseDT.Rows.Add(row);
            }
        }

        private void setDataTableSchema()
        {
            baseDT = new DataTable();
            baseDT.Columns.Add(new DataColumn("Path", typeof(string)));
            baseDT.Columns.Add(new DataColumn("FileName", typeof(string)));
        }

        private void renameFile(DataTable dt)
        {
            foreach (DataRow row in dt.Rows)
            {
                if (File.Exists(row[0].ToString()))
                {
                    var oldName = Path.GetFileName(row[0].ToString());
                    var newName = row[1].ToString();

                    File.Move(row[0].ToString(), row[0].ToString().Replace(oldName, newName));
                }
            }
        }
    }
}
