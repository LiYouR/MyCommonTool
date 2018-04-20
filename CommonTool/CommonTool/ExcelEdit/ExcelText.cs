using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;

namespace CommonTool
{
    public class ExcelText
    {
        public ExcelFont Font { get; set; }

        private Excel.XlHAlign? _HorizontalAlignment;

        public string Text { get; set; }

        public ExcelText()
        {
            Font = null;
            _HorizontalAlignment = null;
            Text = "";
        }

        public void SetHorizontalAlignment(Excel.XlHAlign horizontalAlignment)
        {
            _HorizontalAlignment = horizontalAlignment;
        }

        public Excel.XlHAlign? GetHorizontalAlignment()
        {
            return _HorizontalAlignment;
        }
    }
}
