using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
/////////////////////////////////////////
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using ExcelDataReader;
using System.Data;
using System.IO;
using System.Data.SqlClient;
using System.Text.RegularExpressions;
using System.Reflection;
using System.Xml.Linq;
using System.Windows.Markup;
using System.Runtime.Remoting.Contexts;
using System.Data.OleDb;

namespace ConsoleApp1
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hello World");
            System.MSSQLMAster.BD_IS();
            //System.IO.TextReader _TextReader = System.IO;

            Console.ReadLine();
        }
    }
}
namespace System
{
    public static class MSSQLMAster
    {
        
        /// <summary>
        /// Diplom.Pages.MSSQLMAster.BD_IS();
        /// </summary>
        /// <returns></returns>
        public static bool BD_IS(
            string _connectionString= "Persist Security Info = False; User ID = sa; Initial Catalog = master; Data Source = BD - KIP\\SQLEXPRESS",
            string _cmdText=";"
        )
        {
            SqlConnection p_SqlConnection = new SqlConnection(_connectionString);
            p_SqlConnection.Open();
            using (SqlDataReader insdata = new SqlCommand(_cmdText, p_SqlConnection).ExecuteReader())
            {
                Console.WriteLine("Данные успешно добавлены");
                return true;
            }
            p_SqlConnection.Close();
            return !true;
        }
    }
}
