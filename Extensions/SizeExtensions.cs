using System.Drawing;

namespace shell_style_drag_n_drop
{
    public static class SizeExtensions
    {
        public static Win32Size ToWin32Size(this Size source)
        {
            Win32Size size;
            size.cx = source.Width;
            size.cy = source.Height;

            return size;
        }
    }
}
