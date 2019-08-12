using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace shell_style_drag_n_drop
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private const string FileName = "Test.txt";
        private const string FileContent = "Lorem ipsum dolor sit amet, consetetur sadipscing elitr, sed diam nonumy eirmod tempor invidunt ut labore et dolore magna aliquyam erat, sed diam voluptua. At vero eos et accusam et justo duo dolores et ea rebum. Stet clita kasd gubergren, no sea takimata sanctus est Lorem ipsum dolor sit amet. Lorem ipsum dolor sit amet, consetetur sadipscing elitr, sed diam nonumy eirmod tempor invidunt ut labore et dolore magna aliquyam erat, sed diam voluptua. At vero eos et accusam et justo duo dolores et ea rebum. Stet clita kasd gubergren, no sea takimata sanctus est Lorem ipsum dolor sit amet.";

        private Point _mouseDownPosition;

        public MainWindow()
        {
            InitializeComponent();
        }

        private void Grid_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            _mouseDownPosition = e.GetPosition(null);
        }

        private void Grid_PreviewMouseMove(object sender, MouseEventArgs e)
        {
            var currentPosition = e.GetPosition(null);
            if (e.LeftButton == MouseButtonState.Pressed && CheckPosition(_mouseDownPosition - currentPosition))
            {
                // Create DataObject

                // Set Thumbnail

                // DoDragDrop
            }
        }

        private bool CheckPosition(Vector position)
        {
            return Math.Abs(position.X) > SystemParameters.MinimumHorizontalDragDistance
                || Math.Abs(position.Y) > SystemParameters.MinimumVerticalDragDistance;
        }
    }
}
