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
using System.Data.SqlClient;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using MessageBox = System.Windows.Forms.MessageBox;
using OpenFileDialog = Microsoft.Win32.OpenFileDialog;
using Clipboard = System.Windows.Clipboard;
using DataFormats = System.Windows.DataFormats;
using System.Reflection;
using DataGrid = System.Windows.Controls.DataGrid;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Xml.Linq;
using System.Windows.Markup;
using System.Runtime.Remoting.Contexts;

namespace Diplom.Pages
{
    /// <summary>
    /// Логика взаимодействия для UserPG.xaml
    /// </summary>
    public partial class UserPG : Page
    {
        private DataTableCollection tableCollection = null;
        IExcelDataReader edr;
        public UserPG()
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
        public DataView readFile(string fileNames)
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
            if (dtgView.ItemsSource != null)
            {
                SqlConnection con = new SqlConnection("Data Source=DESK_HP_MINI\\SQLEXPRESS01;Integrated Security=true;");
                con.Open();
                if (con != null && con.State == ConnectionState.Open)
                {
                    MessageBox.Show("Успешно подключено");
                    //CheckTables(sender,e);
                    string readString = "Create Database TestOL";
                    string Read = "Use [TestOL] CREATE TABLE Oleg(\r\n  ID INT  PRIMARY KEY,\r\n  Name VARCHAR(20) NOT NULL,\r\n  Surname INT DEFAULT 0\r\n);";
                    SqlCommand sqlCommand = new SqlCommand(Read, con);
                    SqlCommand readCommand = new SqlCommand(readString, con);
                    using (SqlDataReader dataRead = readCommand.ExecuteReader())
                    {

                    }
                    using (SqlDataReader Datard = sqlCommand.ExecuteReader())
                    {
                        if (Datard != null)
                        {
                            while (Datard.Read())
                            {
                                this.dtgView.SelectAllCells();
                                this.dtgView.ClipboardCopyMode = DataGridClipboardCopyMode.IncludeHeader;
                                ApplicationCommands.Copy.Execute(null, this.dtgView);
                                this.dtgView.UnselectAllCells();
                                var result = (TextWriter)Clipboard.GetData(DataFormats.CommaSeparatedValue);
                            }
                        }
                    }
                    con.Close();
                }
                else
                {
                    MessageBox.Show("Неудалось подключиться к серверу");
                    return;
                }
            }
            else
            {
                MessageBox.Show("Нету данных для создания базы");
            }
        }

        private static void insertToDataBase(TextWriter result)
        {
            throw new NotImplementedException();
        }

        private void btnjSON_Click(object sender, RoutedEventArgs e)
        {
            if (dtgView.ItemsSource != null)
            {
                this.dtgView.SelectAllCells();
                this.dtgView.ClipboardCopyMode = DataGridClipboardCopyMode.IncludeHeader;
                ApplicationCommands.Copy.Execute(null, this.dtgView);
                this.dtgView.UnselectAllCells();
                var result = (TextWriter)Clipboard.GetData(DataFormats.CommaSeparatedValue);

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
            else
            {
                MessageBox.Show("Нету данных для создания файла");
            }
        }
        private void CheckTables(object sender, RoutedEventArgs e)
        {
            DataRowView dataRow = (DataRowView)dtgView.SelectedItem;
            List<DataGrid> tab = new List<DataGrid>();
            int index = dtgView.CurrentCell.Column.DisplayIndex;
            List<DataGrid> cellValue = new List<DataGrid>((int)dataRow.Row.ItemArray[index]);
            for (int i = 0; i < dtgView.Columns.Count; i++)
            {
                for (int j = 0; j < cellValue.Count; j++)
                {
                    if (cellValue[j] == cellValue[j + 1])
                    {
                        tab = cellValue;
                    }
                }
            }
        }
    }
}
