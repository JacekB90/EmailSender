using System;
using System.Windows;
using EmailNotesSend;

namespace FastGroupEmail
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private static readonly FileReader _fileReader = new FileReader();
        private static readonly EmailSender _emailSender = new EmailSender();

        public MainWindow()
        {
            InitializeComponent();
        }
        public void SendEmails(object sender, RoutedEventArgs e)
        {
            try
            {
                string _subject = subject.Text;
                string _message = message.Text;
                string _attachments = attachments.Text;
                string _adress = adress.Text;
                int counter = 0;

                if (_attachments.Equals(string.Empty))
                {
                    MessageBox.Show(Constants.emptyPath);
                    return;
                }
                else if (_adress.Equals(string.Empty))
                {
                    MessageBox.Show(Constants.emptyPath);
                    return;
                }
                else if (_subject.Equals(string.Empty))
                {
                    MessageBox.Show(Constants.emptySubject);
                    return;
                }
                else if (_message.Equals(string.Empty))
                {
                    MessageBox.Show(Constants.emptyMessage);
                    return;
                }

                string[] mailList = _fileReader.ReadMailList(_adress);
                Array.Sort(mailList);
                string[] fileList = _fileReader.ReadAttachments(_attachments);
                Array.Sort(fileList);              
               
                foreach (string attachment in fileList)
                {
                    foreach (string address in mailList)
                    {
                        if (System.IO.Path.GetFileName(attachment).Split('.')[0] == address.Split(',')[0])
                        {
                            //MessageBox.Show("There is a match " + attachment + " = " + address.Split(',')[1] + "temat: " + _subject + "tresc " + _message);
                            _emailSender.NotesMailSend(address.Split(',')[1], _subject, _message, attachment);
                            counter += 1;                        
                        }
                       
                    }
                                      
                }
                if (counter == 0)
                {
                    MessageBox.Show(Constants.noMatchesStatement + counter + Constants.countMessage);
                    return;
                }
                if (counter == 1)
                {
                    MessageBox.Show(counter + Constants.countMessage);
                }
                else
                {
                    MessageBox.Show(counter + Constants.countMessagesPlural);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
        public void ClearInscribedText(object sender, RoutedEventArgs e)
        {
            subject.Clear();
            message.Clear();
            attachments.Clear();
            adress.Clear();
        }
        public void AskAboutExitProgram(object sender, RoutedEventArgs e)
        {         
            MessageBoxResult dr = MessageBox.Show(Constants.exitAsking, Constants.confirmation, MessageBoxButton.YesNo);

                if (dr == MessageBoxResult.Yes)
               { 
                    System.Windows.Application.Current.Shutdown();
               }
        }
        private void OpenNewWindowToExcelSplit(object sender, RoutedEventArgs e)
        {
            SplitExcelFile _split = new SplitExcelFile();
            _split.ShowDialog();
        }
    }
}
        
        
