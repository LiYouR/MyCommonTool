using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using TopssLogger;

namespace CommonTool
{
    public class ExcelEditor
    {
        private string _FilePath;

        private Excel.Application _ExcelApplication;
        private Excel.Workbooks _ExcelWorkbooks;
        private Excel.Workbook _ExcelWorkbook;

        private Excel.Sheets _ExcelWorkSheets;
        public Excel.Sheets ExcelWorkSheets
        {
            get { return _ExcelWorkSheets; }
        }

        public OfficeVersion Version
        {
            get
            {
                return GetVersion();
            }
        }

        //Chn 是否显示Excel程序
        public bool Visible
        {
            get
            {
                return _ExcelApplication.Visible;
            }
            set
            {
                _ExcelApplication.Visible = value;
            }
        }

        //Chn 是否刷新屏幕，设为false可以提升速度
        public bool ScreenUpdating
        {
            get
            {
                return _ExcelApplication.ScreenUpdating;
            }
            set
            {
                _ExcelApplication.ScreenUpdating = value;
            }
        }

        public ExcelEditor()
        {
            _ExcelApplication = new Excel.Application();

            //2018/02/27 ADD BY daijun #725 SPC機能の検査結果比較画面で結果出力が正常に出力できない（#8147）START
            if(_ExcelApplication != null)
                LogManager.GetLogger().Write(LogLevel.Info, "Create Excel.Application Success");
            else
                LogManager.GetLogger().Write(LogLevel.Info, "Create Excel.Application Failed");

            GetVersion();

            if (_ExcelApplication != null)
			//2018/02/27 ADD BY daijun #725 SPC機能の検査結果比較画面で結果出力が正常に出力できない（#8147）End
                _ExcelWorkbooks = _ExcelApplication.Workbooks;

            //2018/02/27 ADD BY daijun #725 SPC機能の検査結果比較画面で結果出力が正常に出力できない（#8147）START
            if (_ExcelWorkbooks != null)
                LogManager.GetLogger().Write(LogLevel.Info, "Get Excel.Workbooks Success");
            else
                LogManager.GetLogger().Write(LogLevel.Info, "Get Excel.Workbooks Failed");
            //2018/02/27 ADD BY daijun #725 SPC機能の検査結果比較画面で結果出力が正常に出力できない（#8147）End
        }

        private void Init()
        {
            //Chn 不显示任何提示，直接保存
            _ExcelApplication.DisplayAlerts = false;
            _ExcelApplication.AlertBeforeOverwriting = false;

            _ExcelApplication.Visible = false;
            _ExcelApplication.ScreenUpdating = false;
        }

        
        private OfficeVersion GetVersion()
        {
            if (_ExcelApplication == null)
                return OfficeVersion.NotInstall;

            string version = _ExcelApplication.Version;

            //2018/02/27 ADD BY daijun #725 SPC機能の検査結果比較画面で結果出力が正常に出力できない（#8147）START
            string logMsg = string.Format("Office Version : {0}", version);
            LogManager.GetLogger().Write(LogLevel.Info, logMsg);
            //2018/02/27 ADD BY daijun #725 SPC機能の検査結果比較画面で結果出力が正常に出力できない（#8147）End

            version = version.Trim();

            OfficeVersion officeVersion;
            switch (version)
            {
                case "8.0":
                    officeVersion = OfficeVersion.Office97;
                    break;
                case "9.0":
                    officeVersion = OfficeVersion.Office2000;
                    break;
                case "10.0":
                    officeVersion = OfficeVersion.Office2002;
                    break;
                case "11.0":
                    officeVersion = OfficeVersion.Office2003;
                    break;
                case "12.0":
                    officeVersion = OfficeVersion.Office2007;
                    break;
                case "14.0":
                    officeVersion = OfficeVersion.Office2010;
                    break;
                case "15.0":
                    officeVersion = OfficeVersion.Office2013;
                    break;
                default:
                    officeVersion = OfficeVersion.Unknown;
                    break;
            }

            return officeVersion;
        }

        public void Create(string filePath, string defaultSheetName)
        {
            _ExcelWorkbook = _ExcelWorkbooks.Add(true);

            //2018/02/27 ADD BY daijun #725 SPC機能の検査結果比較画面で結果出力が正常に出力できない（#8147）START
            if (_ExcelWorkbook != null)
                LogManager.GetLogger().Write(LogLevel.Info, "Add Excel.Workbook Success");
            else
                LogManager.GetLogger().Write(LogLevel.Info, "Add Excel.Workbook Failed");
            //2018/02/27 ADD BY daijun #725 SPC機能の検査結果比較画面で結果出力が正常に出力できない（#8147）End

            _ExcelWorkSheets = _ExcelWorkbook.Worksheets;

            //2018/02/27 ADD BY daijun #725 SPC機能の検査結果比較画面で結果出力が正常に出力できない（#8147）START
            if (_ExcelWorkSheets != null)
                LogManager.GetLogger().Write(LogLevel.Info, "Get Excel.Sheets Success");
            else
                LogManager.GetLogger().Write(LogLevel.Info, "Get Excel.Sheets Failed");
            //2018/02/27 ADD BY daijun #725 SPC機能の検査結果比較画面で結果出力が正常に出力できない（#8147）End

            Excel.Worksheet sheet = (Excel.Worksheet)_ExcelWorkSheets[1];
            sheet.Name = defaultSheetName;

            Init();

            _FilePath = filePath;
        }

        public void Open(string filePath)
        {
            _ExcelWorkbook = _ExcelWorkbooks.Add(filePath);

            //2018/02/27 ADD BY daijun #725 SPC機能の検査結果比較画面で結果出力が正常に出力できない（#8147）START
            if (_ExcelWorkbook != null)
                LogManager.GetLogger().Write(LogLevel.Info, "Add Excel.Workbook Success");
            else
                LogManager.GetLogger().Write(LogLevel.Info, "Add Excel.Workbook Failed");
            //2018/02/27 ADD BY daijun #725 SPC機能の検査結果比較画面で結果出力が正常に出力できない（#8147）End

            _ExcelWorkSheets = _ExcelWorkbook.Worksheets;

            //2018/02/27 ADD BY daijun #725 SPC機能の検査結果比較画面で結果出力が正常に出力できない（#8147）START
            if (_ExcelWorkSheets != null)
                LogManager.GetLogger().Write(LogLevel.Info, "Get Excel.Sheets Success");
            else
                LogManager.GetLogger().Write(LogLevel.Info, "Get Excel.Sheets Failed");
            //2018/02/27 ADD BY daijun #725 SPC機能の検査結果比較画面で結果出力が正常に出力できない（#8147）End

            Init();

            _FilePath = filePath;
        }

        /// <summary>
        /// Chn 添加一个工作表
        /// </summary>
        /// <param name="SheetName"></param>
        /// <returns></returns>
        public Excel.Worksheet AddSheet(string sheetName)
        {
            Excel.Worksheet sheet = (Excel.Worksheet)_ExcelWorkSheets.Add(Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            sheet.Name = sheetName;
            return sheet;
        }

        /// <summary>
        /// Chn 获取当前显示的工作表
        /// </summary>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        public Excel.Worksheet GetActiveSheet(string sheetName)
        {
            Excel.Worksheet sheet = (Excel.Worksheet)_ExcelWorkbook.ActiveSheet;
            return sheet;
        }

        public bool Save()
        {
            return SaveAs(_FilePath);
        }

        public bool SaveAs(string filePath)
        {
            try
            {
                string ext = Path.GetExtension(filePath);
                ext = ext.Trim();
                ext = ext.ToLower();

                int excelVersion;

                switch (ext)
                {
                    case ".xls":
                        excelVersion = 56;
                        break;
                    case ".xlsx":
                        excelVersion = 51;
                        break;
                    case ".csv":
                        excelVersion = 6;
                        break;
                    default:
                        excelVersion = 51;
                        break;
                }

                _ExcelWorkbook.SaveAs((object)filePath, excelVersion, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                return true;
            }
            catch (Exception ex)
            {
                LogManager.GetLogger().Write(LogLevel.Error, "ExcelEditor SaveAs Err", ex);
                return false;
            }
        }

       
        public bool Dispose()
        {
            IntPtr intPtr = new IntPtr(_ExcelApplication.Hwnd);
            int excelApplicationProcessId;
            GetWindowThreadProcessId(intPtr, out excelApplicationProcessId);

            try
            {
                _ExcelWorkbook.Close(Type.Missing, Type.Missing, Type.Missing);
                _ExcelWorkbooks.Close();
                _ExcelApplication.Quit();

                return true;
            }
            catch (Exception ex)
            {
                LogManager.GetLogger().Write(LogLevel.Error,"ExcelEditor Dispose Err",ex);
                return false;
            }
            finally
            {
                Process p = Process.GetProcessById(excelApplicationProcessId);
                if (p != null)
                    p.Kill();

                _ExcelWorkSheets = null;
                _ExcelWorkbook = null;
                _ExcelWorkbooks = null;
                _ExcelApplication = null;

                GC.Collect();
            }
        }


        [DllImport("User32.dll")]
        public static extern int GetWindowThreadProcessId(IntPtr hWnd, out int ProcessId);

        public void KillExcelApp(Excel.Application app)
        {
            app.Quit();
            IntPtr intptr = new IntPtr(app.Hwnd);
            int id;
            GetWindowThreadProcessId(intptr, out id);
            
        }
    }

    /// <summary>
    /// Chn Office 板本
    /// </summary>
    public enum OfficeVersion
    {
        NotInstall=0,
        Unknown = 1,
        Office97 = 2,
        Office2000 = 3,
        Office2002 = 4,
        Office2003 = 5,
        Office2007 = 6,
        Office2010 = 7,
        Office2013 = 8,
    }

    /// <summary>
    /// Chn 常用颜色定义,对应Excel中颜色名
    /// </summary>
    public enum ExcelColorIndex
    {
        无色 = -4142, 自动 = -4105, 黑色 = 1, 褐色 = 53, 橄榄 = 52, 深绿 = 51, 深青 = 49,
        深蓝 = 11, 靛蓝 = 55, 灰色80 = 56, 深红 = 9, 橙色 = 46, 深黄 = 12, 绿色 = 10,
        青色 = 14, 蓝色 = 5, 蓝灰 = 47, 灰色50 = 16, 红色 = 3, 浅橙色 = 45, 酸橙色 = 43,
        海绿 = 50, 水绿色 = 42, 浅蓝 = 41, 紫罗兰 = 13, 灰色40 = 48, 粉红 = 7,
        金色 = 44, 黄色 = 6, 鲜绿 = 4, 青绿 = 8, 天蓝 = 33, 梅红 = 54, 灰色25 = 15,
        玫瑰红 = 38, 茶色 = 40, 浅黄 = 36, 浅绿 = 35, 浅青绿 = 34, 淡蓝 = 37, 淡紫 = 39,
        白色 = 2,
    }

    /// <summary>
    /// Chn Excel字体
    /// </summary>
    public class ExcelFont
    {
        public string FontName;
        public int FontSize;
        public ExcelColorIndex? FontColor;

        public ExcelFont()
        {

        }

        public ExcelFont(string fontName,int fontSize,ExcelColorIndex? fontcolor)
        {
            FontName = fontName;
            FontSize = fontSize;
            FontColor = fontcolor;
        }
    }

    /// <summary>
    /// Chn Excel边框
    /// </summary>
    public class ExcelBorder
    {
        public Excel.XlLineStyle LineStyle;

        public Excel.XlBorderWeight Weight;

        public ExcelColorIndex? BorderColor;

        public ExcelBorder()
        {

        }

        public ExcelBorder(Excel.XlLineStyle lineStyle, Excel.XlBorderWeight weight, ExcelColorIndex? borderColor)
        {
            LineStyle = lineStyle;
            Weight = weight;
            BorderColor = borderColor;
        }
    }

    public class CellRectangle
    {
        public double Top { get; set; }
        public double Left { get; set; }
        public double Width { get; set; }
        public double Height { get; set; }

        public CellRectangle()
        {
        }

        public CellRectangle(double left, double top, double width, double height)
        {
            Left=left;
            Top = top;
            Width = width;
            Height = height;
        }
    }
}
