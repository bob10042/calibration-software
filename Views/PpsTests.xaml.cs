using System.Windows.Controls;
using AGXCalibrationUI.ViewModels;

namespace AGXCalibrationUI.Views
{
    public partial class PpsTests : UserControl
    {
        public PpsTests()
        {
            InitializeComponent();
            DataContext = new PpsTestsViewModel();
        }
    }
}
