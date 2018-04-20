using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms.VisualStyles;
using Excel = Microsoft.Office.Interop.Excel;
using TopssLogger;

namespace CommonTool
{
    public static class WorksheetExtension
    {
        /// <summary>
        /// Chn 设置列宽 单位为字符个数
        /// </summary>
        /// <param name="workSheet"></param>
        /// <param name="columnWidth"></param>
        public static void SetColumnWidth(this Excel.Worksheet workSheet, double columnWidth)
        {
            workSheet.Columns.ColumnWidth = columnWidth;
        }

        /// <summary>
        /// Chn 设置列宽 单位像素
        /// </summary>
        /// <param name="workSheet"></param>
        /// <param name="columnWidth"></param>
        public static void SetColumnWidth(this Excel.Worksheet workSheet, int columnWidth)
        {
            //CHN 转成像素,列宽是默认字体的宽度为单位的,表示显示多少个字符
            //此处没有用字符宽度，默认以1/10英寸为单位转换
            workSheet.Columns.ColumnWidth = columnWidth / 9.6;
        }

        /// <summary>
        /// Chn 设置行高 单位磅
        /// </summary>
        /// <param name="workSheet"></param>
        /// <param name="rowHeight"></param>
        public static void SetRowHeight(this Excel.Worksheet workSheet, double rowHeight)
        {
            //CHN 转成像素,行高是以磅为单位,DPI=96
            workSheet.Rows.RowHeight = rowHeight;
        }

        /// <summary>
        /// Chn 设置行高 单位像素
        /// </summary>
        /// <param name="workSheet"></param>
        /// <param name="rowHeight"></param>
        public static void SetRowHeight(this Excel.Worksheet workSheet, int rowHeight)
        {
            //CHN 转成像素,行高是以磅为单位,DPI=96
            workSheet.Rows.RowHeight = rowHeight * 72 / 96;
        }

        /// <summary>
        /// Chn 行宽自适应
        /// </summary>
        /// <param name="workSheet"></param>
        public static void SetColumnWidthAutoFit(this Excel.Worksheet workSheet)
        {
            workSheet.Columns.EntireColumn.AutoFit();
            //workSheet.Columns.AutoFit();
        }

        /// <summary>
        /// Chn 插入图片
        /// </summary>
        /// <param name="workSheet"></param>
        /// <param name="picturePath"></param>
        /// <param name="picLeft"></param>
        /// <param name="picTop"></param>
        /// <returns></returns>
        public static Excel.Shape InsertPictureByShape(this Excel.Worksheet workSheet, string picturePath, float picLeft, float picTop)
        {
            Size picSize = ImageHelper.GetDimensions(picturePath);
            int picWidth = picSize.Width;
            int picHeight = picSize.Height;

            Excel.Shape pic = workSheet.Shapes.AddPicture(picturePath, Microsoft.Office.Core.MsoTriState.msoFalse,
               Microsoft.Office.Core.MsoTriState.msoTrue,
               picLeft, picTop, picWidth, picHeight);

            return pic;
        }

        /// <summary>
        /// Chn 插入图片
        /// </summary>
        /// <param name="workSheet"></param>
        /// <param name="picturePath"></param>
        /// <param name="picLeft"></param>
        /// <param name="picTop"></param>
        /// <returns></returns>
        public static Excel.Picture InsertPictureByPicture(this Excel.Worksheet workSheet, string picturePath, float picLeft, float picTop)
        {
            Size picSize = ImageHelper.GetDimensions(picturePath);
            int picWidth = picSize.Width;
            int picHeight = picSize.Height;

            Excel.Pictures pics = (Excel.Pictures)workSheet.Pictures(Type.Missing);
            Excel.Picture pic = pics.Insert(picturePath, Type.Missing);

            pic.Left = (double)picLeft;
            pic.Top = (double)picTop;
            pic.Width = (double)picWidth;
            pic.Height = (double)picHeight;

            return pic;
        }

        public static Excel.Range GetRange(this Excel.Worksheet workSheet, int startRow, int startColumn,int endRow,int endColumn)
        {
            Excel.Range startCell = workSheet.Cells[startRow, startColumn];
            Excel.Range endCell = workSheet.Cells[endRow, endColumn];

            Excel.Range range = workSheet.Range[startCell, endCell];
            return range;
        }

        public static Excel.Range GetCellRange(this Excel.Worksheet workSheet, int row, int column)
        {
            //Excel.Range range = workSheet.Range[workSheet.Cells[row, column], workSheet.Cells[row, column]];

            Excel.Range range = (Excel.Range)workSheet.Cells[row, column];
            return range;
        }

        public static Excel.Range GetRowRange(this Excel.Worksheet workSheet, int row)
        {
            Excel.Range range = (Excel.Range)workSheet.Rows[row, Type.Missing];
            return range;
        }

        public static Excel.Range GetColumnRange(this Excel.Worksheet workSheet,int column)
        {
            Excel.Range range = (Excel.Range)workSheet.Columns[Type.Missing, column];
            return range;
        }

        /// <summary>
        /// Chn 合并单元格
        /// </summary>
        /// <param name="workSheet"></param>
        /// <param name="startRow"></param>
        /// <param name="startColumn"></param>
        /// <param name="endRow"></param>
        /// <param name="endColumn"></param>
        public static void MergeCells(this Excel.Worksheet workSheet, int startRow, int startColumn, int endRow, int endColumn)
        {
            Excel.Range range = GetRange(workSheet, startRow, startColumn, endRow, endColumn);

            range.Merge(Type.Missing);
        }

        /// <summary>
        /// Chn 合并单元格
        /// </summary>
        /// <param name="workSheet"></param>
        /// <param name="startRow"></param>
        /// <param name="startColumn"></param>
        /// <param name="endRow"></param>
        /// <param name="endColumn"></param>
        public static void MergeCells(this Excel.Worksheet workSheet, Excel.Range range)
        {
            range.Merge(Type.Missing);
        }

        /// <summary>
        /// Chn 设置公式
        /// </summary>
        /// <param name="ws"></param>
        /// <param name="formula"></param>
        /// <param name="Startx"></param>
        /// <param name="Starty"></param>
        /// <param name="Endx"></param>
        /// <param name="Endy"></param>
        public static void SetFormula(this Excel.Worksheet workSheet, string formula, int startRow, int startColumn, int endRow, int endColumn)
        {
            Excel.Range range = GetRange(workSheet, startRow, startColumn, endRow, endColumn);

            range.Formula = formula;
        }

        /// <summary>
        /// Chn 设置字体
        /// </summary>
        /// <param name="workSheet"></param>
        /// <param name="range"></param>
        /// <param name="fontName"></param>
        /// <param name="fontSize"></param>
        /// <param name="fontColor"></param>
        public static void SetFont(this Excel.Worksheet workSheet, Excel.Range range, string fontName, int fontSize, ExcelColorIndex? fontColor)
        {
            Excel.Font font = range.Font;

            font.Name = fontName;
            font.Size = fontSize;
            if(fontColor!=null)
               font.ColorIndex = fontColor;
        }

         /// <summary>
        /// Chn 设置字体
        /// </summary>
        /// <param name="workSheet"></param>
        /// <param name="range"></param>
        /// <param name="fontName"></param>
        /// <param name="fontSize"></param>
        /// <param name="fontColor"></param>
        public static void SetFont(this Excel.Worksheet workSheet, Excel.Range range,ExcelFont excelFont)
        {
             if (excelFont != null)
             {
                 Excel.Font font = range.Font;

                 font.Name = excelFont.FontName;
                 font.Size = excelFont.FontSize;
                 font.ColorIndex = excelFont.FontColor;
             }
        }

        /// <summary>
        /// Chn 设置水平对齐方式
        /// </summary>
        /// <param name="workSheet"></param>
        /// <param name="range"></param>
        /// <param name="HorizontalAlignment"></param>
        public static  void SetHorizontalAlignment(this Excel.Worksheet workSheet, Excel.Range range, Excel.XlHAlign? horizontalAlignment)
        {
            if (horizontalAlignment != null)
                range.HorizontalAlignment = horizontalAlignment;
        }

        /// <summary>
        /// Chn 设置垂直对齐方式
        /// </summary>
        /// <param name="workSheet"></param>
        /// <param name="range"></param>
        /// <param name="verticalAlignment"></param>
        public static void SetVerticalAlignment(this Excel.Worksheet workSheet, Excel.Range range, Excel.XlVAlign? verticalAlignment)
        {
            if (verticalAlignment != null)
                range.VerticalAlignment = verticalAlignment;
        }

        /// <summary>
        /// Chn 设置边框样式
        /// </summary>
        /// <param name="workSheet"></param>
        /// <param name="range"></param>
        /// <param name="lineStyle"></param>
        /// <param name="weight"></param>
        public static void SetBorder(this Excel.Worksheet workSheet, Excel.Range range, Excel.XlLineStyle lineStyle, Excel.XlBorderWeight weight, ExcelColorIndex? borderColor)
        {
            range.Borders.LineStyle = lineStyle;
            range.Borders.Weight = weight;
            if (borderColor != null)
               range.Borders.ColorIndex = borderColor;
        }

        /// <summary>
        /// Chn 设置边框样式
        /// </summary>
        /// <param name="workSheet"></param>
        /// <param name="range"></param>
        /// <param name="lineStyle"></param>
        /// <param name="weight"></param>
        public static void SetBorder(this Excel.Worksheet workSheet, Excel.Range range, ExcelBorder excelBorder)
        {
            if (excelBorder != null)
            {
                range.Borders.LineStyle = excelBorder.LineStyle;
                range.Borders.Weight = excelBorder.Weight;
                range.Borders.ColorIndex = excelBorder.BorderColor;
            }
        }

        public static void SetRowHeight(this Excel.Worksheet workSheet, Excel.Range range,double rowHeight)
        {
            range.RowHeight = rowHeight;
        }

        public static void SetColumnWidth(this Excel.Worksheet workSheet, Excel.Range range, double columnWidth)
        {
            range.ColumnWidth = columnWidth;
        }

        /// <summary>
        /// Chn 设置背景色
        /// </summary>
        /// <param name="workSheet"></param>
        /// <param name="range"></param>
        /// <param name="lineStyle"></param>
        /// <param name="weight"></param>
        public static void SetBackgroundColor(this Excel.Worksheet workSheet, Excel.Range range, ExcelColorIndex? backgroundColor)
        {
            if (backgroundColor != null)
                range.Interior.ColorIndex = backgroundColor;

            //range.Cells.Interior.ColorIndex = backgroundColor.
        }

        public static void SetCellValue(this Excel.Worksheet workSheet, int row, int column, object value)
        {
            workSheet.Cells[row, column] = value;
        }

        /// <summary>
        /// 设置区域内的内容(批量设置能提高速度)
        /// </summary>
        /// <param name="workSheet"></param>
        /// <param name="range"></param>
        /// <param name="stringArray"></param>
        public static void SetRangeValue(this Excel.Worksheet workSheet, Excel.Range range, string[,] stringArray)
        {
            range.Value2 = stringArray;
        }

        public static bool InsertExcelText(this Excel.Worksheet workSheet, int row, int column, ExcelText excelText)
        {
            try
            {
                workSheet.SetCellValue(row, column, excelText.Text);

                Excel.Range range = workSheet.GetCellRange(row, column);

                workSheet.SetFont(range, excelText.Font);

                workSheet.SetHorizontalAlignment(range, excelText.GetHorizontalAlignment());

                return true;
            }
            catch (Exception ex)
            {
                LogManager.GetLogger().Write(LogLevel.Error, "InsertExcelText Err", ex);
                return false;
            }
        }

        #region InsertExcelTable

        private static void SetExcelTableColumnWidth(this Excel.Worksheet workSheet, int row, int column, List<ExcelTableColumnHeader> columnHeaderList)
        {
            int rangeStartRow = row;
            int rangeEndRow = row;
            int rangeStartColumn = column;
            int rangeEndColumn = column;
            for (int j = 0; j < columnHeaderList.Count; j++)
            {
                ExcelTableColumnHeader columnHeader = columnHeaderList[j];
                rangeEndColumn += columnHeader.UseExcelColumnCount - 1;

                Excel.Range columnRange = workSheet.GetRange(rangeStartRow, rangeStartColumn, rangeEndRow,
                        rangeEndColumn);

                //设置宽度
                if (columnHeader.Width != null)
                {
                    double columnWidth = (double)columnHeader.Width / columnHeader.UseExcelColumnCount;
                    workSheet.SetColumnWidth(columnRange, columnWidth);
                }

                rangeStartColumn = rangeEndColumn + 1;
                rangeEndColumn = rangeStartColumn;
            }
        }

        private static void SetExcelTableColumnStyle(this Excel.Worksheet workSheet, int startRow,int endRow, int column,List<ExcelTableColumnHeader> columnHeaderList)
        {
            int rangeStartRow = startRow;
            int rangeEndRow = endRow;
            int rangeStartColumn = column;
            int rangeEndColumn = column;
            for (int j = 0; j < columnHeaderList.Count; j++)
            {
                ExcelTableColumnHeader columnHeader = columnHeaderList[j];
                rangeEndColumn += columnHeader.UseExcelColumnCount - 1;

                Excel.Range columnRange = workSheet.GetRange(rangeStartRow, rangeStartColumn, rangeEndRow,
                        rangeEndColumn);

                //设置对齐方式
                workSheet.SetHorizontalAlignment(columnRange, columnHeader.GetHorizontalAlignment());
                workSheet.SetVerticalAlignment(columnRange, columnHeader.GetVerticalAlignment());

                rangeStartColumn = rangeEndColumn + 1;
                rangeEndColumn = rangeStartColumn;
            }
        }

        private static void MergeExcelTableCells(this Excel.Worksheet workSheet,
            int startRow,int endRow,int startColumn, List<ExcelTableColumnHeader> columnHeaderList)
        {
            for (int i = startRow; i <= endRow; i++)
            {
                int rangeStartColumn = startColumn;
                int rangeEndColumn = startColumn;

                for (int j = 0; j < columnHeaderList.Count; j++)
                {
                    ExcelTableColumnHeader columnHeader = columnHeaderList[j];
                    rangeEndColumn += columnHeader.UseExcelColumnCount - 1;

                    if (columnHeader.UseExcelColumnCount > 1)
                    {
                        Excel.Range columnRange = workSheet.GetRange(i, rangeStartColumn, i, rangeEndColumn);
                        workSheet.MergeCells(columnRange);
                    }

                    rangeEndColumn += 1;
                    rangeStartColumn = rangeEndColumn;
                }
            }
        }

        public static void InsertExcelTable(this Excel.Worksheet workSheet, int row, int column, ExcelTable excelTable,Action<int,int> callback,out Action stopFunc)
        {
            bool stopFlag = false;

            #region Chn 停止方法
            stopFunc = delegate()
            {
                stopFlag = true;
            };
            #endregion

            #region Chn 判断停止
            if (stopFlag == true)
                return;
            #endregion

            ExcelTableHeader tableHeader = excelTable.TableHeader;
            ExcelTableContent tableContent = excelTable.TableContent;
            Excel.Range tableHeaderRange=null;
            Excel.Range tableContentRange = null;

            List <ExcelTableColumnHeader> columnHeaderList = tableHeader.Columns.GetSortedColumnList();
            int excelColumnCount = tableHeader.Columns.ExcelColumnCount;

            //Chn 设置列样式
            workSheet.SetExcelTableColumnWidth(row, column, columnHeaderList);

            //Chn 合并表头单元格
            workSheet.MergeExcelTableCells(row,row,column,columnHeaderList);

            int headerStartRow=row;
            int headerEndRow=row;
            int headerStartColumn = column;
            int headerEndColumn = column + excelColumnCount - 1;

            //Chn 设置表头(内容，字体...)
            if (tableHeader.Visible)
            {
                tableHeaderRange = workSheet.GetRange(headerStartRow, headerStartColumn, headerEndRow, headerEndColumn);

                string[,] headerStringArray = excelTable.GetHeaderStringArray();
                workSheet.SetBackgroundColor(tableHeaderRange, tableHeader.BackgroundColor);
                workSheet.SetFont(tableHeaderRange, tableHeader.Font);
                workSheet.SetBorder(tableHeaderRange, tableHeader.Border);
                if (tableHeader.RowHeight!=null)
                    workSheet.SetRowHeight(tableHeaderRange, (double)tableHeader.RowHeight);
                workSheet.SetRangeValue(tableHeaderRange, headerStringArray);
            }

            #region Chn 判断停止
            if (stopFlag == true)
                return;
            #endregion
           
            if (tableContent.TableData == null)
                return;

            int contentRowCount = excelTable.TableContent.TableData.Rows.Count;

            if (tableHeader.Visible)
            {
                //Chn 设置列样式
                workSheet.SetExcelTableColumnStyle(row + 1, row + contentRowCount, column, columnHeaderList);
                //Chn 合并表体单元格
                workSheet.MergeExcelTableCells(row + 1, row + contentRowCount, column, columnHeaderList);
            }
            else
            {
                //Chn 设置列样式
                workSheet.SetExcelTableColumnStyle(row, row + contentRowCount - 1, column, columnHeaderList);
                //Chn 合并表体单元格
                workSheet.MergeExcelTableCells(row, row + contentRowCount - 1, column, columnHeaderList);
            }

            int contentStartRow;
            int contentEndRow;
            int contentStartColumn;
            int contentEndColumn;

            if (tableHeader.Visible)
            {
                contentStartRow = row + 1;
                contentEndRow = row + contentRowCount;
                contentStartColumn = column;
                contentEndColumn = column + excelColumnCount - 1;
            }
            else
            {
                contentStartRow = row;
                contentEndRow = row + contentRowCount-1;
                contentStartColumn = column;
                contentEndColumn = column + excelColumnCount - 1;
            }

            tableContentRange = workSheet.GetRange(contentStartRow, contentStartColumn, contentEndRow, contentEndColumn);

            //Chn 设置表体(内容，字体...)
            string[,] contentStringArray = excelTable.GetContentStringArray();
            workSheet.SetBackgroundColor(tableContentRange, tableContent.BackgroundColor);
            workSheet.SetFont(tableContentRange, tableContent.Font);
            workSheet.SetBorder(tableContentRange, tableContent.Border);
            if (tableContent.RowHeight!=null)
                workSheet.SetRowHeight(tableContentRange,(double)tableContent.RowHeight);
            workSheet.SetRangeValue(tableContentRange, contentStringArray);

            #region Chn 获取图片插入位置方法

            CellRectangle columnFirstCellRactangle = null;
            CellRectangle ret = new CellRectangle();
            Func<int,int,CellRectangle> getColumnCellRectangle = delegate(int tableRow, int tableColumn)
            {
                if (columnFirstCellRactangle == null)
                {
                    int cellRow = row + tableRow - 1;
                    if (tableHeader.Visible)
                        cellRow = cellRow + 1;

                    int cellStartColumn = column;
                    int cellEndColumn = column;
                    for (int i = 0; i < tableColumn; i++)
                    {
                        ExcelTableColumnHeader columnHeader = columnHeaderList[i];

                        cellStartColumn = cellEndColumn;
                        cellEndColumn += columnHeader.UseExcelColumnCount;
                    }
                    cellEndColumn--;

                    Excel.Range rng = workSheet.GetRange(cellRow,cellStartColumn,cellRow,cellEndColumn);

                    double top = Convert.ToSingle(rng.Top);
                    double left = Convert.ToSingle(rng.Left);
                    double width = Convert.ToSingle(rng.Width);
                    double height = Convert.ToSingle(rng.Height);

                    columnFirstCellRactangle = new CellRectangle(left, top, width, height);
                }

                double retTop = columnFirstCellRactangle.Top + (tableRow - 1) * columnFirstCellRactangle.Height;
                double retLeft = columnFirstCellRactangle.Left;
                double retWidth = columnFirstCellRactangle.Width;
                double retHeight = columnFirstCellRactangle.Height;

                ret.Top = retTop;
                ret.Left = retLeft;
                ret.Width = retWidth;
                ret.Height = retHeight;

                return ret;
            };

            #endregion

            #region Chn 判断停止
            if (stopFlag == true)
                return;
            #endregion

            #region Chn 插入图片
            for (int j = 0; j < columnHeaderList.Count; j++)
            {
                ExcelTableColumnHeader columnHeader = columnHeaderList[j];
                if (columnHeader.ColumnType == ExcelColumnType.picture)
                {
                    DataTable dt = excelTable.TableContent.TableData;
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        DataRow dataRow = dt.Rows[i];
                        string picPath = dataRow[columnHeader.ColumnName].ToString();

                        CellRectangle rect = getColumnCellRectangle(i + 1, j + 1);

                        if (File.Exists(picPath))
                        {
                            Excel.Shape pic = workSheet.InsertPictureByShape(picPath, (float) rect.Left,
                                (float) rect.Top);

                            //Chn 缩放
                            if (pic.Width > rect.Width || pic.Height > rect.Height)
                            {
                                double widthK = pic.Width/rect.Width;
                                double heightK = pic.Height/rect.Height;

                                if (widthK > heightK)
                                {
                                    pic.Width = (float)(pic.Width / widthK);
                                    pic.Height = (float)(pic.Height / widthK);
                                }
                                else
                                {
                                    pic.Width = (float)(pic.Width / heightK);
                                    pic.Height = (float)(pic.Height / heightK);
                                }
                            }
                           
                            //Chn 居中
                            pic.Left += (float)((rect.Width - pic.Width) / 2);
                            pic.Top += (float)((rect.Height - pic.Height) / 2);
                        }

                        if(callback!=null)
                           callback(i+1, dt.Rows.Count);

                        #region Chn 判断停止
                        if (stopFlag == true)
                            return;
                        #endregion
                    }

                    columnFirstCellRactangle = null;
                }
            }
            #endregion
        }

        #endregion
    }
}
