using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
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
using ExcelDataReader;
using System.Data;

namespace Diplom
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private DataTableCollection tableCollection = null;
        IExcelDataReader edr;
        public MainWindow()
        {
            InitializeComponent();
        }

        private void btnFile_Click(object sender, RoutedEventArgs e)
        {
            // Открытие файла в проводнике с расширением EXCEL
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel file (*.xlxs)|*.xlsx|All Files(*.*)|*.*";
            openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            if (openFileDialog.ShowDialog() == true)
            {
                txbFile.Text = File.ReadAllText(openFileDialog.FileName);
            }
            dtgView.ItemsSource = readFile(openFileDialog.FileName); 
            try
            {
                dtgView.ItemsSource = readFile(openFileDialog.FileName);
              
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
            }
        }
        private DataView readFile(string fileNames)
        {
            //Вывод EXCEL файла в datagrid
            var extension = fileNames.Substring(fileNames.LastIndexOf('.'));
            FileStream stream = File.Open(fileNames, FileMode.Open, FileAccess.Read);
            if (extension == ".xlsx")
                edr = ExcelReaderFactory.CreateOpenXmlReader(stream);
            else if (extension == ".xls")
                edr = ExcelReaderFactory.CreateBinaryReader(stream);

            var conf = new ExcelDataSetConfiguration
            {
                ConfigureDataTable = x => new ExcelDataTableConfiguration
                {
                    UseHeaderRow = true
                }
            };
            DataSet dataSet = edr.AsDataSet(conf);
            DataView dtView = dataSet.Tables[0].AsDataView();

            edr.Close();
            return dtView;
        }

        private void btnSql_Click(object sender, RoutedEventArgs e)
        {
            if(dtgView.ItemsSource == null)
            {
                MessageBox.Show("Нету данных для создания базы");
            }
            else
            {

            }
        }

        private void btnjSON_Click(object sender, RoutedEventArgs e)
        {
            //Выгрузка jSON файла
            this.dtgView.SelectAllCells();
            this.dtgView.ClipboardCopyMode = DataGridClipboardCopyMode.IncludeHeader;
            ApplicationCommands.Copy.Execute(null, this.dtgView);
            this.dtgView.UnselectAllCells();
            String result = (string)Clipboard.GetData(DataFormats.CommaSeparatedValue);

            Microsoft.Win32.SaveFileDialog dlg = new Microsoft.Win32.SaveFileDialog();
            dlg.FileName = "jSON file";
            dlg.DefaultExt = ".json";
            dlg.Filter = "jSON files (.json)|*.json";

            Nullable<bool> _result = dlg.ShowDialog();

            string filePath = "";
            if (_result == true) filePath = dlg.FileName;

            try
            {
                StreamWriter sw = new StreamWriter(filePath);
                sw.WriteLine(result);
                sw.Close();
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show(ex.Message.ToString());
            }
        }
    }
}
