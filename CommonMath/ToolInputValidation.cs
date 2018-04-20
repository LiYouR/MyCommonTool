using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace CommonMath
{
    public static class ToolInputValidation
    {
        /// <summary>
        /// only input decimal and alternative translations
        /// </summary>
        /// <param name="e"></param>
        /// <param name="control"></param>
        public static void OnlyInputDecimal(KeyPressEventArgs e, TextBox control)
        {
            //Detection has been input "-"  
            bool IsContainsDot = control.Text.Contains(".");

            //Detection can only enter numbers and Backspace
            if (((int)e.KeyChar < 48 || (int)e.KeyChar > 57) && (int)e.KeyChar != 8)
                e.Handled = true;
            else if (IsContainsDot && (e.KeyChar == 46)) //if input "-" , and input "-" again
            {
                e.Handled = true;
            }

            if ((int)e.KeyChar == 46)                       //is "."    
            {
                if (control.Text.Length == 1)             // "-" back is not allowed "."
                {
                    if (control.Text[0] == 45)
                    {
                        e.Handled = true;
                        return;
                    }
                }

                if (control.Text.Length <= 0)
                    e.Handled = true;           //"." can not int the first index.   
                else
                {
                    float f;
                    float oldf;
                    bool b1 = false, b2 = false;
                    b1 = float.TryParse(control.Text, out oldf);
                    b2 = float.TryParse(control.Text + e.KeyChar.ToString(), out f);
                    if (b2 == false)
                    {
                        if (b1 == true)
                            e.Handled = true;
                        else
                            e.Handled = false;
                    }
                    else
                        e.Handled = false;
                }
            }

            // input code is "-"
            if (e.KeyChar == 45)
            {
                if (control.Text.Length > 0)
                {
                    if (control.Text[0] != e.KeyChar && control.SelectionStart == 0) //when index == 0, "-" can be input
                    {
                        e.Handled = false;           // "-" Can only be entered once                              
                    }
                    else
                    {
                        e.Handled = true;
                    }
                }
                else
                {
                    e.Handled = false;              // "-" Can only int the first index.
                }
            }
        }
    }
}
