using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace 串口助手
{
    class ExcelHelper
    {
        /// <summary>
        /// 从Excel读取数据
        /// 只支持单表
        /// </summary>
        /// <param name="FilePath">文件路径</param>
        public static DataTable ReadFromExcel(string FilePath)
        {
            DataTable result = null;
            IWorkbook wk = null;
            string extension = System.IO.Path.GetExtension(FilePath); //获取扩展名
            try
            {
                using (FileStream fs = File.OpenRead(FilePath))
                {
                    if (extension.Equals(".xls")) //2003
                    {
                        wk = new HSSFWorkbook(fs);
                    }
                    else                         //2007以上
                    {
                        wk = new XSSFWorkbook(fs);
                    }
                }

                //读取当前表数据
                ISheet sheet = wk.GetSheetAt(0);

                //构建DataTable
                IRow row = sheet.GetRow(2);
                result = BuildDataTable(row);
                if (result != null)
                {
                    if (sheet.LastRowNum >= 1)
                    {
                        for (int i = 3; i < sheet.LastRowNum + 1; i++)
                        {
                            IRow temp_row = sheet.GetRow(i);
                            if (temp_row == null) { continue; }//2019-01-14 修复 行为空时会出错
                            List<object> itemArray = new List<object>();
                            for (int j = 0; j < result.Columns.Count; j++)//解决Excel超出DataTable列问题    lqwvje20181027
                            {
                                //itemArray.Add(temp_row.GetCell(j) == null ? string.Empty : temp_row.GetCell(j).ToString());
                                itemArray.Add(GetValueType(temp_row.GetCell(j)));//解决 导入Excel  时间格式问题  lqwvje 20180904
                            }

                            result.Rows.Add(itemArray.ToArray());
                        }
                    }
                }
                return result;

            }
            catch (Exception ex)
            {
                return null;
            }
        }

     

        private static DataTable BuildDataTable(IRow Row)
        {
            DataTable result = null;
            if (Row.Cells.Count > 0)
            {
                result = new DataTable();
                for (int i = 0; i < Row.LastCellNum; i++)
                {
                    if (Row.GetCell(i) != null)
                    {
                        result.Columns.Add(Row.GetCell(i).ToString());
                    }
                }
            }
            return result;
        }

        /// <summary>
        /// 获取单元格类型
        /// </summary>
        /// <param name="cell"></param>
        /// <returns></returns>
        private static object GetValueType(ICell cell)
        {
            if (cell == null)
                return null;
            switch (cell.CellType)
            {
                case CellType.Blank: //BLANK:  
                    return null;
                case CellType.Boolean: //BOOLEAN:  
                    return cell.BooleanCellValue;
                case CellType.Numeric: //NUMERIC:  
                    if (DateUtil.IsCellDateFormatted(cell))
                    {
                        return cell.DateCellValue;
                    }
                    return cell.NumericCellValue;
                case CellType.String: //STRING:  
                    return cell.StringCellValue;
                case CellType.Error: //ERROR:  
                    return cell.ErrorCellValue;
                case CellType.Formula: //FORMULA:  
                    cell.SetCellType(CellType.String);
                    return cell.StringCellValue;
                default:
                    return "=" + cell.CellFormula;
            }
        }
    }
}
