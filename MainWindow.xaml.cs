using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.Windows;
using System.Windows.Input;
using System.Windows.Interop;
using ComIDataObject = System.Runtime.InteropServices.ComTypes.IDataObject;

namespace shell_style_drag_n_drop
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private const string FileName = "Test.txt";
        private const string FileContent = "Lorem ipsum dolor sit amet, consetetur sadipscing elitr, sed diam nonumy eirmod tempor invidunt ut labore et dolore magna aliquyam erat, sed diam voluptua. At vero eos et accusam et justo duo dolores et ea rebum. Stet clita kasd gubergren, no sea takimata sanctus est Lorem ipsum dolor sit amet. Lorem ipsum dolor sit amet, consetetur sadipscing elitr, sed diam nonumy eirmod tempor invidunt ut labore et dolore magna aliquyam erat, sed diam voluptua. At vero eos et accusam et justo duo dolores et ea rebum. Stet clita kasd gubergren, no sea takimata sanctus est Lorem ipsum dolor sit amet.";

        private System.Windows.Point _mouseDownPosition;

        public MainWindow()
        {
            InitializeComponent();

            AllowDrop = true;
        }

        protected override void OnDragEnter(DragEventArgs e)
        {
            e.Effects = DragDropEffects.Copy;

            var wp = e.GetPosition(this).ToWin32Point();

            var hwndTarget = new WindowInteropHelper(this).Handle;

            var dropHelper = (IDropTargetHelper)new DragDropHelper();

            dropHelper.DragEnter(hwndTarget, (ComIDataObject)e.Data, wp, (int)e.Effects);
        }

        protected override void OnDragOver(DragEventArgs e)
        {
            e.Effects = DragDropEffects.Copy;

            var wp = e.GetPosition(this).ToWin32Point();

            var dropHelper = (IDropTargetHelper)new DragDropHelper();
            dropHelper.DragOver(ref wp, (int)e.Effects);
        }

        protected override void OnDragLeave(DragEventArgs e)
        {
            IDropTargetHelper dropHelper = (IDropTargetHelper)new DragDropHelper();
            dropHelper.DragLeave();
        }

        protected override void OnDrop(DragEventArgs e)
        {
            e.Effects = DragDropEffects.Copy;

            System.Windows.Point p = e.GetPosition(this);

            Win32Point pt;
            pt.x = (int)p.X;
            pt.y = (int)p.Y;

            IDropTargetHelper dropHelper = (IDropTargetHelper)new DragDropHelper();

            dropHelper.Drop((ComIDataObject)e.Data, ref pt, (int)e.Effects);

        }

        private void Grid_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            _mouseDownPosition = e.GetPosition(this);
        }

        private void Grid_PreviewMouseMove(object sender, MouseEventArgs e)
        {
            var currentPosition = e.GetPosition(null);
            if (e.LeftButton == MouseButtonState.Pressed && CheckPosition(_mouseDownPosition - currentPosition))
            {
                // Create DataObject
                var dataObject = new VirtualFileDataObject();

                // Set Thumbnail
                SetDragThumbnail(dataObject);

                // DoDragDrop
                DragDrop.DoDragDrop(sender as DependencyObject, dataObject, DragDropEffects.Copy);
            }
        }

        private bool CheckPosition(Vector position)
        {
            return Math.Abs(position.X) > SystemParameters.MinimumHorizontalDragDistance
                || Math.Abs(position.Y) > SystemParameters.MinimumVerticalDragDistance;
        }

        private void SetDragThumbnail(ComIDataObject dataObject)
        {
            var bitmap = new Bitmap(100, 100, PixelFormat.Format32bppArgb);

            using (var graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.Magenta);
                graphics.DrawEllipse(Pens.Blue, 20, 20, 60, 60);
            }

            var dragImage = new ShDragImage
            {
                sizeDragImage = bitmap.Size.ToWin32Size(),
                ptOffset = new System.Windows.Point(50, 95).ToWin32Point(),
                hbmpDragImage = bitmap.GetHbitmap(),
                crColorKey = Color.Magenta.ToArgb()
            };

            var dragHelper = (IDragSourceHelper)new DragDropHelper();
            dragHelper.InitializeFromBitmap(ref dragImage, dataObject);
        }
    }
}
