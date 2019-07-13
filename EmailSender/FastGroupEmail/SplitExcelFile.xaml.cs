using System.Windows;
using EmailNotesSend;

namespace FastGroupEmail
{
    /// <summary>
    /// Interaction logic for SplitExcelFile.xaml
    /// </summary>
    public partial class SplitExcelFile : Window
    {
        public SplitExcelFile()
        {
            InitializeComponent();
        }
        private void CloseWindowSplitExcel(object sender, RoutedEventArgs e)
        {
            MessageBoxResult dr = MessageBox.Show(Constants.exitAsking, Constants.confirmation, MessageBoxButton.YesNo);

            if (dr == MessageBoxResult.Yes)
            {               
                this.Close();             
            }
        }
        private void ClearWindowSplitExcel(object sender, RoutedEventArgs e)
        {
            pathExcel.Clear();
            pathSave.Clear();
        }

        private void StartWindowSplitExcel(object sender, RoutedEventArgs e)
        {
            string _excelFile = pathExcel.Text;
            string _saveFiles = pathSave.Text;

            Split _split = new Split();

            if(_excelFile.Equals(string.Empty))
            {
                MessageBox.Show(Constants.emptyPath);
                return;
            }
            else if(_saveFiles.Equals(string.Empty))
            {
                MessageBox.Show(Constants.emptyPath);
                return;
            }
            else
            {
                _split.SplitExcel(_excelFile, _saveFiles);
            }

                   
        }
    }
}
