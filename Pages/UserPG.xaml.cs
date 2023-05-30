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
using System.Data.OleDb;
using System.Windows.Threading;
using System.Web.Script.Serialization;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.TaskbarClock;

namespace Diplom.Pages
{
    /// <summary>
    /// Логика взаимодействия для UserPG.xaml
    /// </summary>
    public partial class UserPG : Page
    {
        // SqlConnection cn = new SqlConnection(@"Data Source=имя сервера;Initial Catalog=имя БД;Integrated Security=False;Persist Security Info=True;User ID=логин;Password=пароль");
        IExcelDataReader edr;
        string ds;
        string temp;
        SqlConnection con = new SqlConnection("Data Source=DESK_HP_MINI\\SQLEXPRESS;Integrated Security=true;");
        void timer_Tick(object sender, EventArgs e)
        {
            LiveTimeLabel.Content = DateTime.Now.ToString("D");
            LBTime.Content = DateTime.Now.ToString("F");
        }
        public UserPG()
        {
            InitializeComponent();
            DispatcherTimer LiveTime = new DispatcherTimer();
            LiveTime.Interval = TimeSpan.FromSeconds(1);
            LiveTime.Tick += timer_Tick;
            LiveTime.Start();
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
                edr = ExcelReaderFactory.CreateReader(stream);
            else if (extension == ".xls")
                edr = ExcelReaderFactory.CreateReader(stream);

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
                this.dtgView.SelectAllCells();
                this.dtgView.ClipboardCopyMode = DataGridClipboardCopyMode.IncludeHeader;
                ApplicationCommands.Copy.Execute(null, this.dtgView);
                this.dtgView.UnselectAllCells();
                var result = (string)Clipboard.GetData(DataFormats.CommaSeparatedValue);
                Microsoft.Win32.SaveFileDialog dlg = new Microsoft.Win32.SaveFileDialog();
                dlg.FileName = "csv file";
                dlg.DefaultExt = ".csv";
                dlg.Filter = "CSV files (.csv)|*.csv";
                Nullable<bool> _result = dlg.ShowDialog();

                string filePath = "";
                if (_result == true) filePath = dlg.FileName;

                try
                {
                    StreamWriter sw = new StreamWriter(filePath);
                    sw.Write(result);
                    sw.Close();
                }
                catch (Exception ex)
                {
                    System.Windows.MessageBox.Show(ex.Message.ToString());
                }
                temp = LiveTimeLabel.Content.ToString();
                ds = LBTime.Content.ToString();
            // if (dtgView.ItemsSource != null)
            // {
            con.Open();
                if (con != null && con.State == ConnectionState.Open)
                {
                    string readString = "Create Database [" + temp + "]";
                    SqlCommand readCommand = new SqlCommand(readString, con);
                    using (SqlDataReader dataRead = readCommand.ExecuteReader())
                    {
                        MessageBox.Show("База успешно создана");
                    }
                    string insert = @"USE ["+ temp + "] BEGIN CREATE TABLE [dbo].["+ ds + "] (        [ID] [INT] NOT NULL ,        [Name] [NVARCHAR](max) NULL,        [SecondName] [NVARCHAR](max) NULL,        [email] [NVARCHAR](max) NULL,        [Company] [NVARCHAR](max) NULL,    )    BULK INSERT ["+ ds + "]     FROM 'C:\\Users\\olezh\\OneDrive\\Рабочий стол\\csv file.csv'    WITH    (        CODEPAGE = '1253',        FIELDTERMINATOR = ',',        CHECK_CONSTRAINTS    ) END";
                    SqlCommand insCommand = new SqlCommand(insert, con);
                    using (SqlDataReader insdata = insCommand.ExecuteReader())
                    {
                        MessageBox.Show("Данные успешно добавлены");
                    }
                    con.Close();
                }
           // else
           // {
           //     MessageBox.Show("Неудалось подключиться к серверу");
           //     return;
           // }
           // }
           // else
           // {
           //     MessageBox.Show("Нету данных");
           //     return;
           // }
        }
        private void btnjSON_Click(object sender, RoutedEventArgs e)
        {
            if (dtgView.ItemsSource != null)
            {
                this.dtgView.SelectAllCells();
                this.dtgView.ClipboardCopyMode = DataGridClipboardCopyMode.IncludeHeader;
                ApplicationCommands.Copy.Execute(null, this.dtgView);
                this.dtgView.UnselectAllCells();
                var result = (string)Clipboard.GetData(DataFormats.CommaSeparatedValue);
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
                    sw.Write(result);
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
    }
}