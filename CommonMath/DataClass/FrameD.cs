using System;
using System.Runtime.InteropServices;
using System.ComponentModel;
using System.Diagnostics.Contracts;
using System.Globalization;

namespace CommonMath.DataClass
{
    /// <summary>
    /// 枠の種類
    /// </summary>
    public enum FrameType
    {
        RECTANGLE = 0, //矩形
        ELLIPSE   = 1, //円形
        POLYGON   = 2  //ポリゴン
    }

    /// <summary>
    /// 
    /// </summary>
    [Serializable]
    [ComVisible(true)]
    public class FrameD
    {
        private double _x;
        private double _y;
        private double _width;
        private double _height;
        private double _angle;
        private FrameType _frameType;
        private PointD[] _polygonArray;

        public static readonly FrameD Empty = new FrameD(0, 0, 0, 0, 0, 0, null);

        public FrameD(double x, double y, double width, double height, double angle, FrameType frameType, PointD[] polygonArray)
        {
            this._x = x;
            this._y = y;
            this._width = width;
            this._height = height;
            this._angle = angle;
            this._frameType = frameType;
            if (null == polygonArray)
            {
                this._polygonArray = new PointD[0];
            }
            else
            {
                this._polygonArray = polygonArray;
            }
        }

        /// <summary>
        ///  sets the x-coordinate of the Center
        /// </summary>
        public double X
        {
            get
            {
                return _x;
            }
            set
            {
                _x = value;
            }
        }

        /// <summary>
        ///  sets the y-coordinate of the Center
        /// </summary>
        public double Y
        {
            get
            {
                return _y;
            }
            set
            {
                _y = value;
            }
        }

        /// <summary>
        ///  Gets or sets the width of the frame
        /// </summary>
        public double Width
        {
            get
            {
                return _width;
            }
            set
            {
                _width = value;
            }
        }

        /// <summary>
        ///  Gets or sets the height of the frame
        /// </summary>
        public double Height
        {
            get
            {
                return _height;
            }
            set
            {
                _height = value;
            }
        }

        /// <summary>
        ///  Gets or sets the angle of the frame
        /// </summary>
        public double Angle
        {
            get
            {
                return _angle;
            }
            set
            {
                _angle = value;
            }
        }

        /// <summary>
        ///  Gets or sets the type of the frame
        /// </summary>
        public FrameType FrameType
        {
            get
            {
                return _frameType;
            }
            set
            {
                _frameType = value;
            }
        }

        /// <summary>
        ///  Gets or sets the polygon point of the frame
        /// </summary>
        public PointD[] PolygonArray
        {
            get
            {
                return _polygonArray;
            }
            set
            {
                _polygonArray = value;
            }
        }


        [Browsable(false)]
        public PointD Location
        {
            get
            {
                return new PointD(X, Y);
            }
            set
            {
                X = value.X;
                Y = value.Y;
            }
        }

        /// <summary>
        ///  Determines if the specfied point is contained within the
        ///       rectangular region defined by this
        /// </summary>
        /// <param name="dx"></param>
        /// <param name="dy"></param>
        /// <returns></returns>
        [Pure]
        public bool Contains(double dx, double dy)
        {
            return Contains(new PointD(dx,dy));
        }

        /// <summary>
        ///  Determines if the specfied point is contained within the
        ///       rectangular region defined by this
        /// </summary>
        /// <param name="pt"></param>
        /// <returns></returns>
        [Pure]
        public bool Contains(PointD pt)
        {
            bool bRet = false;

            switch (_frameType)
            {
                case FrameType.RECTANGLE:
                    bRet = ToolDrawing.IsPointInFrame_Rectangle(pt, this);
                    break;
                case FrameType.ELLIPSE:
                    bRet = ToolDrawing.IsPointInFrame_Ellipse(pt, this);
                    break;
                case FrameType.POLYGON:
                    bRet = ToolDrawing.IsPointInFrame_Polygon(pt, this);
                    break;
                default:
                    break;

            }
            return bRet;
        }

        /// <summary>
        ///  Determines if the specfied point is contained within the
        ///       rectangular region defined by this
        /// </summary>
        /// <returns></returns>
        [Pure]
        public RectangleD GetRectangle_CenterOrigin()
        {
            return new RectangleD(-_width / 2, _height / 2, _width , _height);
        }

        /// <summary>
        ///  Determines if the specfied point is contained within the
        ///       rectangular region defined by this
        /// </summary>
        /// <returns></returns>
        [Pure]
        public RectangleD GetOutlineRectangle()
        {
            PointD[] ptArray = null;

            switch (_frameType)
            {
                case FrameType.RECTANGLE:
                case FrameType.ELLIPSE:
                    ptArray = new PointD[] { new PointD(-_width / 2, _height / 2),
                                        new PointD(_width / 2, _height / 2), 
                                        new PointD(_width / 2, -_height / 2),
                                        new PointD(-_width / 2, -_height / 2)};
                    break;
                case FrameType.POLYGON:
                    ptArray = new PointD[_polygonArray.Length];

                    for (int i = 0; i < _polygonArray.Length; i++)
                    {
                        ptArray[i] = new PointD(_polygonArray[i].X, _polygonArray[i].Y);
                    }
                    break;
                default:
                    break;

            }

            if (ptArray != null)
            {
                foreach (PointD pt in ptArray)
                {
                    pt.Rotate_Counterclockwise(Angle);
                    pt.X += _x;
                    pt.Y += _y;
                }
                return ToolDrawing.MakeOutlineRectangle(ptArray);
            }
            return RectangleD.Empty;
        }

        /// <summary>
        ///  GetOutlineRectangle Without Angle
        /// </summary>
        /// <returns></returns>
        [Pure]
        public RectangleD GetOutlineRectangleNoAngle()
        {
            PointD[] ptArray = null;

            switch (_frameType)
            {
                case FrameType.RECTANGLE:
                case FrameType.ELLIPSE:
                    ptArray = new PointD[] { new PointD(-_width / 2, _height / 2),
                                        new PointD(_width / 2, _height / 2), 
                                        new PointD(_width / 2, -_height / 2),
                                        new PointD(-_width / 2, -_height / 2)};
                    break;
                case FrameType.POLYGON:
                    ptArray = new PointD[_polygonArray.Length];

                    for (int i = 0; i < _polygonArray.Length; i++)
                    {
                        ptArray[i] = new PointD(_polygonArray[i].X, _polygonArray[i].Y);
                    }
                    break;
                default:
                    break;

            }

            if (ptArray != null)
            {
                return ToolDrawing.MakeOutlineRectangle(ptArray);
            }
            return RectangleD.Empty;
        }

        // !! Not in C++ version

        /// <devdoc>
        ///    <para>
        ///       Converts the attributes of this <see cref='System.Drawing.Rectangle'/> to a
        ///       human readable string.
        ///    </para>
        /// </devdoc>
        public override string ToString()
        {
            return "{X=" + X.ToString(CultureInfo.CurrentCulture) + ",Y=" + Y.ToString(CultureInfo.CurrentCulture) +
            ",Width=" + Width.ToString(CultureInfo.CurrentCulture) +
            ",Height=" + Height.ToString(CultureInfo.CurrentCulture) +
            ",Angle =" + Angle.ToString(CultureInfo.CurrentCulture) +
            ",Type=" + FrameType.ToString() + "}";
        }
    }
}
