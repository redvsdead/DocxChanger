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
using System.IO;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace WPF_Hello_World
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            processButton.IsEnabled = false;
        }

        

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            var result = MessageBox.Show(oldDate.Text + " заменяется на " + newDate.Text + ".\nВы уверены?", "Подтверждение", 
                            MessageBoxButton.YesNo, MessageBoxImage.Question);
            if (result == MessageBoxResult.Yes)
            {
                string[] fileList = Directory.GetFiles(path.Text);
                string docText = null;
                WordprocessingDocument wordDoc;
                StreamReader sr;
                Regex regexText;
                StreamWriter sw;
                for (int i = 0; i < fileList.Length; ++i)
                {
                    using (wordDoc =
                           WordprocessingDocument.Open(fileList[i], true))
                    {
                        using (sr = new StreamReader(wordDoc.MainDocumentPart.GetStream()))
                        {
                            docText = sr.ReadToEnd();
                        }
                        regexText = new Regex(oldDate.Text);
                        docText = regexText.Replace(docText, newDate.Text);
                        using (sw = new StreamWriter(wordDoc.MainDocumentPart.GetStream(FileMode.Create)))
                        {
                            sw.Write(docText);
                        }
                        //C:\Users\Asus\Desktop\docsdocs
                    }
                }
            }
        }

        private void OldDate_TextChanged(object sender, TextChangedEventArgs e)
        {
            processButton.IsEnabled = (path.Text.Length > 0) && (oldDate.Text.Length > 1) && (newDate.Text.Length > 1);
        }

        private void NewDate_TextChanged(object sender, TextChangedEventArgs e)
        {
            processButton.IsEnabled = (path.Text.Length > 0) && (oldDate.Text.Length > 1) && (newDate.Text.Length > 1);
        }
    }
}
