using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.SS.Util; // 用于设置单元格样式 
using System.IO;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MyWorkDiary
{
    public class ExcelHelper
    {
        public void AppendRowToMonthlySheet(string filePath, string[] rowData)
        {
            try
            {
                // 检查文件是否存在
                if (!File.Exists(filePath))
                {
                    // 如果文件不存在，则创建一个新的Excel文件
                    CreateNewExcelFile(filePath);
                }

                using (FileStream fs = new FileStream(filePath, FileMode.OpenOrCreate))
                {
                    IWorkbook workbook = null;
                    // 根据文件扩展名选择不同的Workbook类型
                    if (filePath.EndsWith(".xlsx"))
                    {
                        workbook = new XSSFWorkbook(fs);
                    }
                    else if (filePath.EndsWith(".xls"))
                    {
                        workbook = new HSSFWorkbook(fs);
                    }
                    else
                    {
                        throw new ArgumentException("不支持的文件格式");
                    }

                    // 获取当前月份
                    string currentMonth = DateTime.Now.ToString("yyyy-MM");

                    // 获取或创建指定月份的Sheet
                    ISheet sheet = GetOrCreateSheet(workbook, currentMonth);

                    // 添加标题行（如果不存在）
                    AddTitleRow(sheet);

                    // 检查是否已经有相同日期的数据
                    bool isDuplicate = CheckForDuplicateDate(sheet, rowData[2]);
                    if (isDuplicate)
                    {
                        Console.WriteLine("已有相同日期的数据，无法追加。");
                        throw new Exception("已有相同日期的数据，无法追加。");
                    }

                    // 获取最后一行的索引
                    int lastRowIndex = sheet.LastRowNum;

                    // 创建新行
                    IRow row = sheet.CreateRow(lastRowIndex + 1);

                    // 向新行中添加数据
                    for (int i = 0; i < rowData.Length; i++)
                    {
                        row.CreateCell(i).SetCellValue(rowData[i]);
                    }

                    // 将更改保存回文件
                    using (FileStream fileOut = new FileStream(filePath, FileMode.Create))
                    {
                        workbook.Write(fileOut);
                    }
                }
            }
            catch (IOException ex)
            {
                Console.WriteLine("文件被占用，请关闭正在使用的Excel文件后再试。");
                Console.WriteLine("详细错误信息: " + ex.Message);
                throw new Exception("文件被占用，请关闭正在使用的Excel文件后再试。");
            }
            catch (Exception ex)
            {
                Console.WriteLine("发生了一个错误: " + ex.Message);
                throw new Exception(ex.Message);
            }
        }

        private void CreateNewExcelFile(string filePath)
        {
            // 创建一个新的Excel文件
            IWorkbook workbook = new XSSFWorkbook(); // 使用XSSFWorkbook创建.xlsx文件

            // 创建第一个Sheet
            ISheet sheet = workbook.CreateSheet(DateTime.Now.ToString("yyyy-MM"));

            // 添加标题行
            AddTitleRow(sheet);

            // 保存新文件
            using (FileStream fileOut = new FileStream(filePath, FileMode.Create))
            {
                workbook.Write(fileOut);
            }
        }

        private ISheet GetOrCreateSheet(IWorkbook workbook, string sheetName)
        {
            // 检查是否已有该Sheet
            ISheet sheet = workbook.GetSheet(sheetName);
            if (sheet == null)
            {
                // 如果不存在，则创建新的Sheet
                sheet = workbook.CreateSheet(sheetName);
            }
            return sheet;
        }

        private void AddTitleRow(ISheet sheet)
        {
            // 检查是否已经有标题行
            if (sheet.LastRowNum >= 0 && sheet.GetRow(0) != null)
            {
                return; // 已经有标题行，无需重复添加
            }

            // 创建标题行
            IRow titleRow = sheet.CreateRow(0);

            // 设置标题行的数据
            titleRow.CreateCell(0).SetCellValue("标题");
            titleRow.CreateCell(1).SetCellValue("内容");
            titleRow.CreateCell(2).SetCellValue("时间");

            // 设置标题行的样式
            ICellStyle titleStyle = sheet.Workbook.CreateCellStyle();
            titleStyle.SetFont(CreateBoldFont(sheet.Workbook));
            titleStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;

            // 应用样式
            foreach (ICell cell in titleRow)
            {
                cell.CellStyle = titleStyle;
            }
        }

        private IFont CreateBoldFont(IWorkbook workbook)
        {
            IFont font = workbook.CreateFont();
            font.FontName = "Arial";
            font.FontHeightInPoints = 12;
            font.IsBold = true;
            return font;
        }

        private bool CheckForDuplicateDate(ISheet sheet, string dateValue)
        {
            // 检查是否有相同日期的数据
            for (int i = 1; i <= sheet.LastRowNum; i++)
            {
                IRow row = sheet.GetRow(i);
                if (row != null && row.Cells.Count > 2)
                {
                    ICell dateCell = row.GetCell(2);
                    if (dateCell != null && dateCell.StringCellValue == dateValue)
                    {
                        return true;
                    }
                }
            }
            return false;
        }
    }
}
