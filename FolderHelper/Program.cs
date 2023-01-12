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

            if (model == "1")
            {
                Console.WriteLine("執行模式一：將指定路徑下所有檔名匯出成Excel");
                helper.GetAllFilesNameExportExcel(sourcePath, targetPath);
            }
            if (model == "2")
            {
                Console.WriteLine("執行模式二：依Excel設定變更檔案名稱");
                helper.RenameFileFromExcel(targetPath);
            }

            Console.WriteLine("執行完畢");
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
            }
            catch (Exception ex)
            {
                Console.WriteLine("");
                Console.WriteLine("阿尼亞不會寫程式，把這個交給父親");
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
            }
            catch (Exception ex)
            {
                Console.WriteLine("");
                Console.WriteLine("阿尼亞不會寫程式，把這個交給父親");
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
