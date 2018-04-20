using System;
using System.Runtime.InteropServices;
using System.ComponentModel;
using System.Diagnostics.Contracts;
using System.Globalization;

namespace CommonMath.DataClass
{
    /// <devdoc>
    ///    <para>
    ///       Stores the location and size of a rectangular region. For
    ///       more advanced region functions use a <see cref='System.Drawing.Region'/>
    ///       object.
    ///    </para>
    /// </devdoc>
    [Serializable]
    [ComVisible(true)]
    public class RectangleD
    {

        /// <devdoc>
        ///    <para>
        ///       Stores the location and size of a rectangular region. For
        ///       more advanced region functions use a <see cref='System.Drawing.Region'/>
        ///       object.
        ///    </para>
        /// </devdoc>
        public static readonly RectangleD Empty = new RectangleD(0,0,0,0);

        private double _x;
        private double _y;
        private double _width;
        private double _height;

        /// <devdoc>
        ///    <para>
        ///       Initializes a new instance of the <see cref='System.Drawing.Rectangle'/>
        ///       class with the specified location and size.
        ///    </para>
        /// </devdoc>
        public RectangleD(double x, double y, double width, double height)
        {
            this._x = x;
            this._y = y;
            this._width = width;
            this._height = height;
        }

        /// <devdoc>
        ///    <para>
        ///       Gets or sets the coordinates of the
        ///       upper-left corner of the rectangular region represented by this <see cref='System.Drawing.Rectangle'/>.
        ///    </para>
        /// </devdoc>
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

        /// <devdoc>
        ///    Gets or sets the x-coordinate of the
        ///    upper-left corner of the rectangular region defined by this <see cref='System.Drawing.Rectangle'/>.
        /// </devdoc>
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

        /// <devdoc>
        ///    Gets or sets the y-coordinate of the
        ///    upper-left corner of the rectangular region defined by this <see cref='System.Drawing.Rectangle'/>.
        /// </devdoc>
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

        /// <devdoc>
        ///    Gets or sets the width of the rectangular
        ///    region defined by this <see cref='System.Drawing.Rectangle'/>.
        /// </devdoc>
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

        /// <devdoc>
        ///    Gets or sets the width of the rectangular
        ///    region defined by this <see cref='System.Drawing.Rectangle'/>.
        /// </devdoc>
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

        /// <devdoc>
        ///    <para>
        ///       Gets the x-coordinate of the upper-left corner of the
        ///       rectangular region defined by this <see cref='System.Drawing.Rectangle'/> .
        ///    </para>
        /// </devdoc>
        [Browsable(false)]
        public double Left
        {
            get
            {
                return X;
            }
        }

        /// <devdoc>
        ///    <para>
        ///       Gets the y-coordinate of the upper-left corner of the
        ///       rectangular region defined by this <see cref='System.Drawing.Rectangle'/>.
        ///    </para>
        /// </devdoc>
        [Browsable(false)]
        public double Top
        {
            get
            {
                return Y;
            }
        }

        /// <devdoc>
        ///    <para>
        ///       Gets the x-coordinate of the lower-right corner of the
        ///       rectangular region defined by this <see cref='System.Drawing.Rectangle'/>.
        ///    </para>
        /// </devdoc>
        [Browsable(false)]
        public double Right
        {
            get
            {
                return X + Width;
            }
        }

        /// <devdoc>
        ///    <para>
        ///       Gets the y-coordinate of the lower-right corner of the
        ///       rectangular region defined by this <see cref='System.Drawing.Rectangle'/>.
        ///    </para>
        /// </devdoc>
        [Browsable(false)]
        public double Bottom
        {
            get
            {
                return Y + Height;
            }
        }

        /// <devdoc>
        ///    <para>
        ///       Tests whether this <see cref='System.Drawing.Rectangle'/> has a <see cref='System.Drawing.Rectangle.Width'/>
        ///       or a <see cref='System.Drawing.Rectangle.Height'/> of 0.
        ///    </para>
        /// </devdoc>
        [Browsable(false)]
        public bool IsEmpty
        {
            get
            {
                return ToolMath.EqualsDouble(_height, 0) && ToolMath.EqualsDouble(_width, 0) && ToolMath.EqualsDouble(_x, 0) && ToolMath.EqualsDouble(_y, 0);
                // C++ uses this definition:
                // return(Width <= 0 )|| (Height <= 0);
            }
        }

        /// <devdoc>
        ///    <para>
        ///       Tests whether <paramref name="obj"/> is a <see cref='System.Drawing.Rectangle'/> with
        ///       the same location and size of this Rectangle.
        ///    </para>
        /// </devdoc>
        public override bool Equals(object obj)
        {
            if (!(obj is RectangleD))
                return false;

            RectangleD comp = (RectangleD)obj;

            return (ToolMath.EqualsDouble(comp.X, X)) &&
            (ToolMath.EqualsDouble(comp.Y, Y)) &&
            (ToolMath.EqualsDouble(comp.Width, Width)) &&
            (ToolMath.EqualsDouble(comp.Height, Height));
        }

        /// <devdoc>
        ///    <para>
        ///       Determines if the specfied point is contained within the
        ///       rectangular region defined by this <see cref='System.Drawing.Rectangle'/> .
        ///    </para>
        /// </devdoc>
        [Pure]
        public bool Contains(double dx, double dy)
        {
            return X <= dx &&
            dx < X + Width &&
            Y <= dy &&
            dy < Y + Height;
        }

        /// <devdoc>
        ///    <para>
        ///       Determines if the specfied point is contained within the
        ///       rectangular region defined by this <see cref='System.Drawing.Rectangle'/> .
        ///    </para>
        /// </devdoc>
        [Pure]
        public bool Contains(PointD pt)
        {
            return Contains(pt.X, pt.Y);
        }

        /// <devdoc>
        ///    <para>
        ///       Determines if the rectangular region represented by
        ///    <paramref name="rect"/> is entirely contained within the rectangular region represented by 
        ///       this <see cref='System.Drawing.Rectangle'/> .
        ///    </para>
        /// </devdoc>
        [Pure]
        public bool Contains(RectangleD rect)
        {
            return (X <= rect.X) &&
            ((rect.X + rect.Width) <= (X + Width)) &&
            (Y <= rect.Y) &&
            ((rect.Y + rect.Height) <= (Y + Height));
        }

        // !! Not in C++ version
        /// <devdoc>
        ///    <para>[To be supplied.]</para>
        /// </devdoc>
        public override int GetHashCode()
        {
            return unchecked((int)((UInt32)X ^
                        (((UInt32)Y << 13) | ((UInt32)Y >> 19)) ^
                        (((UInt32)Width << 26) | ((UInt32)Width >> 6)) ^
                        (((UInt32)Height << 7) | ((UInt32)Height >> 25))));
        }

        /// <devdoc>
        ///    <para>
        ///       Inflates this <see cref='System.Drawing.Rectangle'/>
        ///       by the specified amount.
        ///    </para>
        /// </devdoc>
        public void Inflate(double dWidth, double dHeight)
        {
            X -= dWidth;
            Y -= dHeight;
            Width += 2 * dWidth;
            Height += 2 * dHeight;
        }


        /// <devdoc>
        ///    <para>
        ///       Creates a <see cref='System.Drawing.Rectangle'/>
        ///       that is inflated by the specified amount.
        ///    </para>
        /// </devdoc>
        // !! Not in C++
        public static RectangleD Inflate(RectangleD rect, double x, double y)
        {
            RectangleD r = rect;
            r.Inflate(x, y);
            return r;
        }

        /// <devdoc>
        ///     Creates a Rectangle that represents the intersection between this Rectangle and rect.
        /// </devdoc>
        public void Intersect(RectangleD rect)
        {
            RectangleD result = Intersect(rect, this);

            X = result.X;
            Y = result.Y;
            Width = result.Width;
            Height = result.Height;
        }

        /// <devdoc>
        ///    Creates a rectangle that represents the intersetion between a and
        ///    b. If there is no intersection, null is returned.
        /// </devdoc>
        public static RectangleD Intersect(RectangleD a, RectangleD b)
        {
            double x1 = Math.Max(a.X, b.X);
            double x2 = Math.Min(a.X + a.Width, b.X + b.Width);
            double y1 = Math.Max(a.Y, b.Y);
            double y2 = Math.Min(a.Y + a.Height, b.Y + b.Height);

            if (x2 >= x1
                && y2 >= y1)
            {

                return new RectangleD(x1, y1, x2 - x1, y2 - y1);
            }
            return Empty;
        }

        /// <devdoc>
        ///     Determines if this rectangle intersets with rect.
        /// </devdoc>
        [Pure]
        public bool IntersectsWith(RectangleD rect)
        {
            return (rect.X < X + Width) &&
            (X < (rect.X + rect.Width)) &&
            (rect.Y < Y + Height) &&
            (Y < rect.Y + rect.Height);
        }

        /// <devdoc>
        ///    <para>
        ///       Creates a rectangle that represents the union between a and
        ///       b.
        ///    </para>
        /// </devdoc>
        [Pure]
        public static RectangleD Union(RectangleD a, RectangleD b)
        {
            double x1 = Math.Min(a.X, b.X);
            double x2 = Math.Max(a.X + a.Width, b.X + b.Width);
            double y1 = Math.Min(a.Y, b.Y);
            double y2 = Math.Max(a.Y + a.Height, b.Y + b.Height);

            return new RectangleD(x1, y1, x2 - x1, y2 - y1);
        }

        /// <devdoc>
        ///    <para>
        ///       Adjusts the location of this rectangle by the specified amount.
        ///    </para>
        /// </devdoc>
        public void Offset(PointD pos)
        {
            Offset(pos.X, pos.Y);
        }

        /// <devdoc>
        ///    Adjusts the location of this rectangle by the specified amount.
        /// </devdoc>
        public void Offset(double dx, double dy)
        {
            X += dx;
            Y += dy;
        }

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
            ",Height=" + Height.ToString(CultureInfo.CurrentCulture) + "}";
        }
    }
}
