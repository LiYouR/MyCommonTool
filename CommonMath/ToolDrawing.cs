using System;
using CommonMath.DataClass;

namespace CommonMath
{
    public static class ToolDrawing
    {
        /// <summary>
        /// 時計回り
        /// </summary>
        /// <param name="point"></param>
        /// <param name="dAngle"></param>
        /// <returns></returns>
        public static PointD Rotate_Clockwise(PointD point, double dAngle)
        {
            return PointDMulMatrix(point, MakeRotateMatrix_Clockwise(dAngle));
        }

        /// <summary>
        /// 反時計回り
        /// </summary>
        /// <param name="point"></param>
        /// <param name="dAngle"></param>
        /// <returns></returns>
        public static PointD Rotate_Counterclockwise(PointD point, double dAngle)
        {
            return PointDMulMatrix(point, MakeRotateMatrix_Counterclockwise(dAngle));
        }

        /// <summary>
        /// 時計回りMatrix
        /// </summary>
        /// <param name="dAngle"></param>
        /// <returns></returns>
        public static double[][] MakeRotateMatrix_Clockwise(double dAngle)
        {
            double dAngleRadians = ConvertDegreesToRadians(dAngle);
            return new double[][] { new double[] { Math.Cos(dAngleRadians), -Math.Sin(dAngleRadians) },
                                    new double[] { Math.Sin(dAngleRadians),  Math.Cos(dAngleRadians) }};
        }

        /// <summary>
        /// 反時計回りMatrix
        /// </summary>
        /// <param name="dAngle"></param>
        /// <returns></returns>
        public static double[][] MakeRotateMatrix_Counterclockwise(double dAngle)
        {
            double dAngleRadians = ConvertDegreesToRadians(dAngle);
            return new double[][] { new double[] {  Math.Cos(dAngleRadians), Math.Sin(dAngleRadians) },
                                    new double[] { -Math.Sin(dAngleRadians), Math.Cos(dAngleRadians) }};
        }

        /// <summary>
        /// PointDとMatrixを掛け合わせる
        /// </summary>
        /// <param name="point"></param>
        /// <param name="dMatrix"></param>
        /// <returns></returns>
        public static PointD PointDMulMatrix(PointD point, double[][] dMatrix)
        {
            return new PointD(point.X * dMatrix[0][0] + point.Y * dMatrix[1][0], point.X * dMatrix[0][1] + point.Y * dMatrix[1][1]);
        }

        /// <summary>
        /// 外接枠算出
        /// </summary>
        /// <param name="pointArray"></param>
        /// <returns></returns>
        public static RectangleD MakeOutlineRectangle(PointD[] pointArray)
        {
            //////////////////////
            // 座標系
            //    ↑Y
            //    ｜
            //    ｜
            //    ｜
            //  (0,0)－－－－－→X
            ///////////////////
            if (pointArray.Length == 0)
            {
                return RectangleD.Empty;
            }

            double left = pointArray[0].X;
            double right = pointArray[0].X;
            double top = pointArray[0].Y;
            double bottom = pointArray[0].Y;


            for (int i = 1; i < pointArray.Length; i++)
            {
                if (left > pointArray[i].X)
                {
                    left = pointArray[i].X;
                }
                if (right < pointArray[i].X)
                {
                    right = pointArray[i].X;
                }
                if (top < pointArray[i].Y)
                {
                    top = pointArray[i].Y;
                }
                if (bottom > pointArray[i].Y)
                {
                    bottom = pointArray[i].Y;
                }
            }
            return new RectangleD(left, top, right - left, top - bottom);
        }

        /// <summary>
        /// 角度→弧度転換
        /// </summary>
        /// <param name="degrees"></param>
        /// <returns></returns>
        public static double ConvertDegreesToRadians(double degrees)
        {
            return (Math.PI / 180) * degrees;
        }

        /// <summary>
        /// 点が矩形枠の範囲にあるかどうか判断
        /// </summary>
        /// <param name="pt"></param>
        /// <param name="frame"></param>
        /// <returns></returns>
        public static bool IsPointInFrame_Rectangle(PointD pt, FrameD frame)
        {
            //////////////////////
            // 座標系
            //    ↑Y
            //    ｜
            //    ｜
            //    ｜
            //  (0,0)－－－－－→X
            ///////////////////
            pt.X = pt.X - frame.X;
            pt.Y = pt.Y - frame.Y;

            pt.Rotate_Clockwise(frame.Angle);

            if (pt.X > -frame.Width / 2 && pt.X < frame.Width / 2 && pt.Y > -frame.Height / 2 && pt.Y < frame.Height / 2)
            {
                return true;
            }
            return false;
        }

        /// <summary>
        /// 点が円形の範囲にあるかどうか判断
        /// </summary>
        /// <param name="pt"></param>
        /// <param name="frame"></param>
        /// <returns></returns>
        public static bool IsPointInFrame_Ellipse(PointD pt, FrameD frame)
        {
            //////////////////////
            // 座標系
            //    ↑Y
            //    ｜
            //    ｜
            //    ｜
            //  (0,0)－－－－－→X
            ///////////////////
            pt.X = pt.X - frame.X;
            pt.Y = pt.Y - frame.Y;

            pt.Rotate_Clockwise(frame.Angle);

            double radius;
            if (!ToolMath.EqualsDouble(frame.Width, frame.Height))
            {

                double dScale;
                if (frame.Width > frame.Height)
                {
                    dScale = frame.Width / frame.Height;
                    pt.Y = dScale * pt.Y;
                    radius = frame.Width / 2;
                }
                else
                {
                    dScale = frame.Height / frame.Width;
                    pt.X = dScale * pt.X;
                    radius = frame.Height / 2;
                }
            }
            else
            {
                radius = frame.Width / 2;
            }

            if (Math.Sqrt(pt.X * pt.X + pt.Y * pt.Y) < radius)
            {
                return true;
            }
            return false;
        }

        /// <summary>
        /// 点がポリゴンの範囲にあるかどうか判断
        /// </summary>
        /// <param name="pt"></param>
        /// <param name="frame"></param>
        /// <returns></returns>
        public static bool IsPointInFrame_Polygon(PointD pt, FrameD frame)
        {
            //////////////////////
            // 座標系
            //    ↑Y
            //    ｜
            //    ｜
            //    ｜
            //  (0,0)－－－－－→X
            ///////////////////
            pt.X = pt.X - frame.X;
            pt.Y = pt.Y - frame.Y;

            pt.Rotate_Clockwise(frame.Angle);

            int i, j;
            bool bRet = false;

            int nvert = frame.PolygonArray.Length;
            double testx = pt.X;
            double testy = pt.Y;
            PointD[] arrayP = frame.PolygonArray;

            for (i = 0, j = nvert - 1; i < nvert; j = i++)
            {
                if (((arrayP[i].Y > testy) != (arrayP[j].Y > testy)) &&
                      (testx < (arrayP[j].X - arrayP[i].X) * (testy - arrayP[i].Y) / (arrayP[j].Y - arrayP[i].Y) + arrayP[i].X))
                {
                    bRet = !bRet;
                }

            }
            return bRet;
        }
    }
}
