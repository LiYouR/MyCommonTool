using System;

namespace CommonMath
{
    public static class ToolMath
    {
        /// <summary>
        /// DOUBLE型比較する際に、精度の桁数
        /// </summary>
        private const int ACCURACY = 6;

        /// <summary>
        /// DOUBLE型比較の精度
        /// </summary>
        private static readonly double DOUBLE_ACCURACY = Math.Pow(0.1, ACCURACY);

        /// <summary>
        /// double→int 四捨五入
        /// </summary>
        /// <param name="d"></param>
        /// <returns></returns>
        public static int Round(double d)
        {
            return (int)Math.Round(d);
        }

        /// <summary>
        /// double型の比較
        /// </summary>
        /// <param name="d1"></param>
        /// <param name="d2"></param>
        /// <returns></returns>
        public static bool EqualsDouble(double d1, double d2)
        {
            return Math.Abs(d1 - d2) < DOUBLE_ACCURACY;
        }
    }
}
