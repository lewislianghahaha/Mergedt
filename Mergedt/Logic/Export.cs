using System;
using System.Data;
using System.IO;
using NPOI.XSSF.UserModel;

namespace Mergedt.Logic
{
    public class Export
    {
        /// <summary>
        /// 导出数据至EXCEL
        /// </summary>
        /// <param name="fileAddress"></param>
        /// <param name="temp">MEASUREMENT_COLOR 表头导出记录</param>
        /// <param name="tempdtl">MEASUREMENT 表体导出记录</param>
        /// <returns></returns>
        public bool ExportDtToExcel(string fileAddress, DataTable temp, DataTable tempdtl)
        {
            var result = true;
            var sheetcount = 0;  //记录所需的sheet页总数
            var rownum = 1;
            
            try
            {
                //声明一个WorkBook
                var xssfWorkbook=new XSSFWorkbook();

                //先执行MEASUREMENT_COLOR sheet页(注:1)先列表temp行数判断需拆分多少个sheet表进行填充; 以一个sheet表有9W行记录填充为基准)
                sheetcount = temp.Rows.Count % 100000 == 0 ? temp.Rows.Count / 100000 : temp.Rows.Count / 100000 + 1;
                //i为EXCEL的Sheet页数ID
                for (var i = 1; i <= sheetcount; i++)
                {
                    //创建sheet页
                    var sheet = xssfWorkbook.CreateSheet("MEASUREMENT_COLOR" + i);
                    //创建"标题行"
                    var row = sheet.CreateRow(0);
                    //创建"MEASUREMENT_COLOR"sheet页各列标题
                    for (var j = 0; j < temp.Columns.Count; j++)
                    {
                        //设置列宽度
                        sheet.SetColumnWidth(j, (int)((20 + 0.72) * 256));
                        //创建标题
                        switch (j)
                        {
                            #region SetCellValue
                            case 0:
                                row.CreateCell(j).SetCellValue("BMMEASUREMENTID");
                                break;
                            case 1:
                                row.CreateCell(j).SetCellValue("COLORCODE");
                                break;
                            case 2:
                                row.CreateCell(j).SetCellValue("FORMULAVERSIONDATE");
                                break;
                            case 3:
                                row.CreateCell(j).SetCellValue("DIFFUSECOARSENESS");
                                break;
                            case 4:
                                row.CreateCell(j).SetCellValue("CREATEDDATE");
                                break;
                            case 5:
                                row.CreateCell(j).SetCellValue("MEASUREMENTTIME");
                                break;
                                #endregion
                        }
                    }

                    //计算进行循环的起始行
                    var startrow = (i - 1) * 100000;
                    //计算进行循环的结束行
                    var endrow = i == sheetcount ? temp.Rows.Count : i * 100000;

                    //每一个sheet表显示90000行  
                    for (var j = startrow; j < endrow; j++)
                    {
                        //创建行
                        row = sheet.CreateRow(rownum);
                        //循环获取DT内的列值记录
                        for (var k = 0; k < temp.Columns.Count; k++)
                        {
                            row.CreateCell(k).SetCellValue(Convert.ToString(temp.Rows[j][k]));
                        }
                        rownum++;
                    }
                    //当一个SHEET页填充完毕后,需将变量初始化
                    rownum = 1;
                }


                //创建"MEASUREMENT"
                //先执行MEASUREMENT sheet页(注:1)先列表temp行数判断需拆分多少个sheet表进行填充;以一个sheet表有9W行记录填充为基准)
                sheetcount = tempdtl.Rows.Count % 100000 == 0 ? tempdtl.Rows.Count / 100000 : tempdtl.Rows.Count / 100000 + 1;
                //i为EXCEL的Sheet页数ID
                for (var i = 1; i <= sheetcount; i++)
                {
                    //创建sheet页
                    var sheet = xssfWorkbook.CreateSheet("MEASUREMENT" + i);
                    //创建"标题行"
                    var row = sheet.CreateRow(0);
                    //创建"MEASUREMENT"sheet页各列标题
                    for (var j = 0; j < tempdtl.Columns.Count; j++)
                    {
                        //设置列宽度
                        sheet.SetColumnWidth(j, (int)((20 + 0.72) * 256));
                        //创建标题
                        switch (j)
                        {
                            #region CellValue
                            case 0:
                                row.CreateCell(j).SetCellValue("BMAID");
                                break;
                            case 1:
                                row.CreateCell(j).SetCellValue("ANGLE");
                                break;
                            case 2:
                                row.CreateCell(j).SetCellValue("BMMEASUREMENTID");
                                break;
                            case 3:
                                row.CreateCell(j).SetCellValue("L");
                                break;
                            case 4:
                                row.CreateCell(j).SetCellValue("A");
                                break;
                            case 5:
                                row.CreateCell(j).SetCellValue("B");
                                break;
                            case 6:
                                row.CreateCell(j).SetCellValue("C");
                                break;
                            case 7:
                                row.CreateCell(j).SetCellValue("H");
                                break;
                            case 8:
                                row.CreateCell(j).SetCellValue("BMA_R400");
                                break;
                            case 9:
                                row.CreateCell(j).SetCellValue("BMA_R410");
                                break;
                            case 10:
                                row.CreateCell(j).SetCellValue("BMA_R420");
                                break;
                            case 11:
                                row.CreateCell(j).SetCellValue("BMA_R430");
                                break;
                            case 12:
                                row.CreateCell(j).SetCellValue("BMA_R440");
                                break;
                            case 13:
                                row.CreateCell(j).SetCellValue("BMA_R450");
                                break;
                            case 14:
                                row.CreateCell(j).SetCellValue("BMA_R460");
                                break;
                            case 15:
                                row.CreateCell(j).SetCellValue("BMA_R470");
                                break;
                            case 16:
                                row.CreateCell(j).SetCellValue("BMA_R480");
                                break;
                            case 17:
                                row.CreateCell(j).SetCellValue("BMA_R490");
                                break;
                            case 18:
                                row.CreateCell(j).SetCellValue("BMA_R500");
                                break;
                            case 19:
                                row.CreateCell(j).SetCellValue("BMA_R510");
                                break;
                            case 20:
                                row.CreateCell(j).SetCellValue("BMA_R520");
                                break;
                            case 21:
                                row.CreateCell(j).SetCellValue("BMA_R530");
                                break;
                            case 22:
                                row.CreateCell(j).SetCellValue("BMA_R540");
                                break;
                            case 23:
                                row.CreateCell(j).SetCellValue("BMA_R550");
                                break;
                            case 24:
                                row.CreateCell(j).SetCellValue("BMA_R560");
                                break;
                            case 25:
                                row.CreateCell(j).SetCellValue("BMA_R570");
                                break;
                            case 26:
                                row.CreateCell(j).SetCellValue("BMA_R580");
                                break;
                            case 27:
                                row.CreateCell(j).SetCellValue("BMA_R590");
                                break;
                            case 28:
                                row.CreateCell(j).SetCellValue("BMA_R600");
                                break;
                            case 29:
                                row.CreateCell(j).SetCellValue("BMA_R610");
                                break;
                            case 30:
                                row.CreateCell(j).SetCellValue("BMA_R620");
                                break;
                            case 31:
                                row.CreateCell(j).SetCellValue("BMA_R630");
                                break;
                            case 32:
                                row.CreateCell(j).SetCellValue("BMA_R640");
                                break;
                            case 33:
                                row.CreateCell(j).SetCellValue("BMA_R650");
                                break;
                            case 34:
                                row.CreateCell(j).SetCellValue("BMA_R660");
                                break;
                            case 35:
                                row.CreateCell(j).SetCellValue("BMA_R670");
                                break;
                            case 36:
                                row.CreateCell(j).SetCellValue("BMA_R680");
                                break;
                            case 37:
                                row.CreateCell(j).SetCellValue("BMA_R690");
                                break;
                            case 38:
                                row.CreateCell(j).SetCellValue("BMA_R700");
                                break;
                                #endregion
                        }
                    }

                    //计算进行循环的起始行
                    var startrow = (i - 1) * 100000;
                    //计算进行循环的结束行
                    var endrow = i == sheetcount ? tempdtl.Rows.Count : i * 100000;

                    //获取DT内的行记录步骤(重)
                    //每一个sheet表显示100000行
                    for (var j = startrow; j < endrow; j++)
                    {
                        //创建行
                        row = sheet.CreateRow(rownum);
                        //循环获取DT内的列值记录
                        for (var k = 0; k < tempdtl.Columns.Count; k++)
                        {
                            row.CreateCell(k).SetCellValue(Convert.ToDouble(tempdtl.Rows[j][k]));
                        }
                        rownum++;
                    }
                    //当一个SHEET页填充完毕后,需将变量初始化
                    rownum = 1;
                }

                ////写入数据
                var file = new FileStream(fileAddress, FileMode.Create);
                xssfWorkbook.Write(file);
                file.Close();
            }
            catch (Exception)
            {
                result = false;
            }
            return result;
        }
    }
}
