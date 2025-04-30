using System.Windows;
using System.IO;
using Microsoft.Win32;
using Excel_Сonverter.BL;

namespace Excel_Сonverter.WPF_Client
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "xlsx файлы|*.xlsx";

            var resault = openFileDialog.ShowDialog();

            try
            {
                if (resault != null)
                {
                    this.Title = openFileDialog.FileName;
                    var path = openFileDialog.FileName;

                    ExelService service = new ExelService();

                    using (var fileStream = File.OpenRead(path))
                    {
                        using (var memoryStream = new MemoryStream())
                        {
                            fileStream.CopyTo(memoryStream);
                            memoryStream.Position = 0;
                            var r = service.GetModelExels(memoryStream);
                            ReportService reportService = new ReportService();
                            var otch = reportService.GetModelReports(r);
                            Stream st = service.SaveReports(otch);

                            SaveFileDialog saveFileDialog = new SaveFileDialog();
                            saveFileDialog.Filter = "xlsx файлы|*.xlsx";

                            if (saveFileDialog.ShowDialog() == true)
                            {
                                using (FileStream saveStream = new FileStream(saveFileDialog.FileName, FileMode.Create, FileAccess.Write))
                                {
                                    st.CopyTo(saveStream);
                                }
                                st.Close();

                                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo()
                                {
                                    FileName = saveFileDialog.FileName,
                                    UseShellExecute = true
                                });

                            }

                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }
    }
}