using System.Windows;

namespace shell_style_drag_n_drop
{
    public static class PointExtensions
    {
        public static Win32Point ToWin32Point(this Point source)
        {
            Win32Point point;
            point.x = (int)source.X;
            point.y = (int)source.Y;

            return point;
        }
    }
}
