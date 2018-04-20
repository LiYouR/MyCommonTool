using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace CommonTool
{
    public class ExcelTable
    {
        public ExcelTableHeader TableHeader { get; set; }

        public ExcelTableContent TableContent { get; set; }

        public ExcelTable()
        {
            TableHeader=new ExcelTableHeader();
            TableContent=new ExcelTableContent();
        }

        public string[,] GetHeaderStringArray()
        {
            List<ExcelTableColumnHeader> columnHeaderList = TableHeader.Columns.GetSortedColumnList();
            int excelColumnCount = TableHeader.Columns.ExcelColumnCount;

            string[,] array = new string[1, excelColumnCount];

            int column = 0;
            for (int j = 0; j < columnHeaderList.Count; j++)
            {
                ExcelTableColumnHeader columnHeader = columnHeaderList[j];

                array[0, column] = columnHeader.DisplayName;
                column += columnHeader.UseExcelColumnCount;
            }

            return array;
        }

        public string[,] GetContentStringArray()
        {
            List<ExcelTableColumnHeader> columnHeaderList = TableHeader.Columns.GetSortedColumnList();
            int excelColumnCount = TableHeader.Columns.ExcelColumnCount;

            string[,] array = new string[TableContent.TableData.Rows.Count, excelColumnCount];

            for (int i = 0; i < TableContent.TableData.Rows.Count; i++)
            {
                DataRow dataRow = TableContent.TableData.Rows[i];
                int column = 0;
                for (int j = 0; j < columnHeaderList.Count; j++)
                {
                    ExcelTableColumnHeader columnHeader = columnHeaderList[j];

                    if (columnHeader.ColumnType == ExcelColumnType.text)
                        array[i, column] = dataRow[columnHeader.ColumnName].ToString();
                    else if (columnHeader.ColumnType == ExcelColumnType.picture)
                        array[i, column] = "";
                    else if (columnHeader.ColumnType==ExcelColumnType.rowNo)
                        array[i, column] = (i + 1).ToString();
                    else if (columnHeader.ColumnType == ExcelColumnType.empty)
                        array[i, column] = "";

                    column += columnHeader.UseExcelColumnCount;
                }
            }

            return array;
        }
    }

    public class ExcelTableHeader
    {
        public ExcelFont Font { get; set; }

        public ExcelBorder Border { get; set; }

        public ExcelColorIndex? BackgroundColor { get; set; }

        public double? RowHeight { get; set; }

        private ExcelTableColumnHeaderCollection _Columns;
        public ExcelTableColumnHeaderCollection Columns
        {
            get { return _Columns; }
        }

        public bool Visible { get; set; }

        public ExcelTableHeader()
        {
            Font = null;
            Border = null;
            BackgroundColor = null;
            RowHeight = 20;
            _Columns = new ExcelTableColumnHeaderCollection();
            Visible = true;
            RowHeight = null;
        }
    }

    public class ExcelTableContent
    {
        public ExcelFont Font { get; set; }

        public ExcelBorder Border { get; set; }

        public ExcelColorIndex? BackgroundColor { get; set; }

        public double? RowHeight { get; set; }

        public DataTable TableData { get; set; }

        public ExcelTableContent()
        {
            Font = null;
            Border = null;
            BackgroundColor = null;
            RowHeight = 20;
            TableData = null;
            RowHeight = null;
        }
    }

    #region Column

    public class ExcelTableColumnHeaderCollection
    {
        private List<ExcelTableColumnHeader> _ColumnList = new List<ExcelTableColumnHeader>();

        public int Count
        {
            get
            {
                int count = 0;
                for (int i = 0; i < _ColumnList.Count; i++)
                {
                    ExcelTableColumnHeader header = _ColumnList[i];
                    if (header.Visible)
                        count++;
                }
                return count;
            }
        }

        public int ExcelColumnCount
        {
            get
            {
                int excelColumnCount = 0;
                for (int i = 0; i < _ColumnList.Count; i++)
                {
                    ExcelTableColumnHeader header = _ColumnList[i];
                    if (header.Visible)
                        excelColumnCount += header.UseExcelColumnCount;
                }
                return excelColumnCount;
            }
        }

        public ExcelTableColumnHeader this[int index]
        {
            get { return _ColumnList[index]; }
        }

        public ExcelTableColumnHeader this[string key]
        {
            get
            {
                for (int i = 0; i < _ColumnList.Count; i++)
                {
                    ExcelTableColumnHeader columnHeader = _ColumnList[i];
                    if (key == columnHeader.ColumnName)
                        return columnHeader;
                }
                return null;
            }
        }

        public void Add(ExcelTableColumnHeader value)
        {
            _ColumnList.Add(value);
        }

        public List<ExcelTableColumnHeader> GetSortedColumnList()
        {
            List<ExcelTableColumnHeader> columnList = new List<ExcelTableColumnHeader>();
            for (int i = 0; i < _ColumnList.Count; i++)
            {
                ExcelTableColumnHeader header = _ColumnList[i];
                if (header.Visible)
                    columnList.Add(header);
            }

            columnList.Sort((p1, p2) =>
            {
                return p1.ColumnNo - p2.ColumnNo;
            });

            return columnList;
        }
    }

    public class ExcelTableColumnHeader
    {
        public string ColumnName { get; set; }

        public string DisplayName { get; set; }

        public int ColumnNo { get; set; }

        public double? Width { get; set; }

        public ExcelColumnType ColumnType { get; set; }

        public bool Visible { get; set; }

        public int UseExcelColumnCount { get; set; }

        private Excel.XlHAlign? _HorizontalAlignment;

        private Excel.XlVAlign? _VerticalAlignment;

        public ExcelTableColumnHeader()
        {
            ColumnName = "";
            DisplayName = "";
            ColumnNo = 1;
            Width = null;
            ColumnType = ExcelColumnType.text;
            Visible = true;
            UseExcelColumnCount = 1;
            _HorizontalAlignment = null;
            _VerticalAlignment = null;
        }

        public void SetHorizontalAlignment(Excel.XlHAlign horizontalAlignment)
        {
            _HorizontalAlignment = horizontalAlignment;
        }

        public Excel.XlHAlign? GetHorizontalAlignment()
        {
            return _HorizontalAlignment;
        }

        public void SetVerticalAlignment(Excel.XlVAlign verticalAlignment)
        {
            _VerticalAlignment = verticalAlignment;
        }

        public Excel.XlVAlign? GetVerticalAlignment()
        {
            return _VerticalAlignment;
        }
    }

    public enum ExcelColumnType
    {
        rowNo = 1,
        text=2,
        picture=3,
        empty=4,
    }

    #endregion
}
