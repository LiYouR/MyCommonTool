using System;
using System.ComponentModel;
using System.Globalization;
using System.Runtime.InteropServices;

namespace CommonMath.DataClass
{
    [Serializable]
    [ComVisible(true)]
    public class PointD
    {
        /// <devdoc>
        ///    <para>
        ///       Creates a new instance of the <see cref='PointD'/> class
        ///       with member data left uninitialized.
        ///    </para>
        /// </devdoc>
        public static readonly PointD Empty = new PointD(0, 0);
        private double _x;
        private double _y;
        /**
         * Create a new Point object at the given location
         */
        /// <devdoc>
        ///    <para>
        ///       Initializes a new instance of the <see cref='PointD'/> class
        ///       with the specified coordinates.
        ///    </para>
        /// </devdoc>
        public PointD(double x, double y)
        {
            this._x = x;
            this._y = y;
        }


        /// <devdoc>
        ///    <para>
        ///       Gets a value indicating whether this <see cref='PointD'/> is empty.
        ///    </para>
        /// </devdoc>
        [Browsable(false)]
        public bool IsEmpty
        {
            get
            {
                return ToolMath.EqualsDouble(_x, 0d) && ToolMath.EqualsDouble(_y, 0d);
            }
        }

        /// <devdoc>
        ///    <para>
        ///       Gets the x-coordinate of this <see cref='PointD'/>.
        ///    </para>
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
        ///    <para>
        ///       Gets the y-coordinate of this <see cref='PointD'/>.
        ///    </para>
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
        ///    <para>[To be supplied.]</para>
        /// </devdoc>
        public override bool Equals(object obj)
        {
            if (!(obj is PointD)) return false;
            PointD comp = (PointD)obj;
            return
            ToolMath.EqualsDouble(comp.X, X) &&
            ToolMath.EqualsDouble(comp.Y, Y) &&
            comp.GetType() == GetType();
        }

        /// <devdoc>
        ///    <para>[To be supplied.]</para>
        /// </devdoc>
        public override int GetHashCode()
        {
// ReSharper disable once BaseObjectGetHashCodeCallInGetHashCode
            return base.GetHashCode();
        }

        public override string ToString()
        {
            return string.Format(CultureInfo.CurrentCulture, "{{X={0}, Y={1}}}", _x, _y);
        }

        /// <devdoc>
        ///    <para>Counterclockwise Rotate</para>
        /// </devdoc>
        public void Rotate_Counterclockwise(double dAngle)
        {
            PointD pointTemp = ToolDrawing.Rotate_Counterclockwise(this, dAngle);
            _x = pointTemp.X;
            _y = pointTemp.Y;
        }

        /// <devdoc>
        ///    <para>Clockwise Rotate</para>
        /// </devdoc>
        public void Rotate_Clockwise(double dAngle)
        {
            PointD pointTemp = ToolDrawing.Rotate_Clockwise(this, dAngle);
            _x = pointTemp.X;
            _y = pointTemp.Y;
        }
    }
}
