using System;
using System.Data;
using System.IO;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;

namespace FolderHelper
{
    public class Excel
    {
        public static void ExportExcelFromDataTable(DataTable dt, string targetPath)
        {
            using (FileStream fs = File.Create(targetPath))
            {
                int i = 0, j = 0;

                IWorkbook wb = new XSSFWorkbook();
                ISheet sheet = wb.CreateSheet("Simple");
                sheet.CreateRow(0);

                for (i = 0; i < dt.Columns.Count; i++)
                    sheet.GetRow(0).CreateCell(i).SetCellValue(dt.Columns[i].ColumnName);

                for (i = 1; i <= dt.Rows.Count; i++)
                {
                    sheet.CreateRow(i);
                    for (j = 0; j < dt.Columns.Count; j++)
                        sheet.GetRow(i).CreateCell(j).SetCellValue(dt.Rows[i - 1][j].ToString());
                }

                wb.Write(fs);
                wb.Close();
            }
        }

        public static DataTable LoadExcelAsDataTable(String xlsFilename)
        {
            FileInfo fi = new FileInfo(xlsFilename);
            using (FileStream fstream = new FileStream(fi.FullName, FileMode.Open))
            {
                IWorkbook wb;
                if (fi.Extension == ".xlsx")
                    wb = new XSSFWorkbook(fstream); // excel2007
                else
                    wb = new HSSFWorkbook(fstream); // excel97

                // 只取第一個sheet。
                ISheet sheet = wb.GetSheetAt(0);

                // target
                DataTable table = new DataTable();

                // 由第一列取標題做為欄位名稱
                IRow headerRow = sheet.GetRow(0);
                int cellCount = headerRow.LastCellNum; // 取欄位數
                for (int i = headerRow.FirstCellNum; i < cellCount; i++)
                {
                    //table.Columns.Add(new DataColumn(headerRow.GetCell(i).StringCellValue, typeof(double)));
                    table.Columns.Add(new DataColumn(headerRow.GetCell(i).StringCellValue));
                }

                // 略過第零列(標題列)，一直處理至最後一列
                for (int i = (sheet.FirstRowNum + 1); i <= sheet.LastRowNum; i++)
                {
                    IRow row = sheet.GetRow(i);
                    if (row == null) continue;

                    DataRow dataRow = table.NewRow();

                    //依先前取得的欄位數逐一設定欄位內容
                    for (int j = row.FirstCellNum; j < cellCount; j++)
                    {
                        ICell cell = row.GetCell(j);
                        if (cell != null)
                        {
                            //如要針對不同型別做個別處理，可善用.CellType判斷型別
                            //再用.StringCellValue, .DateCellValue, .NumericCellValue...取值

                            switch (cell.CellType)
                            {
                                case CellType.Numeric:
                                    dataRow[j] = cell.NumericCellValue;
                                    break;
                                default: // String
                                         //此處只簡單轉成字串
                                    dataRow[j] = cell.StringCellValue;
                                    break;
                            }
                        }
                    }

                    table.Rows.Add(dataRow);
                }

                // success
                return table;
            }
        }
    }
}
