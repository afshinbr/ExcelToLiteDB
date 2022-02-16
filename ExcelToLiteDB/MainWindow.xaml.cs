using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Forms;
using System.Windows.Media;
using LiteDB;
using ExcelApp = Microsoft.Office.Interop.Excel;

namespace ExcelToLiteDB
{

    public partial class MainWindow : Window
    {
        private string _excelPath;
        public MainWindow()
        {
            InitializeComponent();
        }

        private void BtnRun_OnClick(object sender, RoutedEventArgs e)
        {
            try
            {
                // Check if user entered database name
                string dataBaseName = TextDbName.Text;
                if (string.IsNullOrWhiteSpace(dataBaseName))
                    throw new Exception("Please enter database name.");

                // Check if user entered password
                string password = TextPassword.Text;
                if (string.IsNullOrWhiteSpace(password))
                    throw new Exception("Please enter password.");

                // Check if user entered Table
                string tableName = TextTableName.Text;
                if (string.IsNullOrWhiteSpace(tableName))
                    throw new Exception("Please enter table name.");

                // Check if user select an excel file.
                ExcelApp.Application excelApp = new ExcelApp.Application();
                if (string.IsNullOrWhiteSpace(_excelPath))
                    throw new Exception("Please select an excel file.");

                // Open dialog and ask about database folder location
                FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog();
                folderBrowserDialog.Description = "Please select save location";
                var result = folderBrowserDialog.ShowDialog();
                if (result == System.Windows.Forms.DialogResult.OK)
                {
                    // Show a message that confirms converting process is started
                    ShowMessage(Brushes.Black, "Please wait...");

                    //Create a connection string based on database location and password
                    string saveLocation = Path.Combine(folderBrowserDialog.SelectedPath, dataBaseName);
                    string connectionString = $"Filename={saveLocation};Password={password}";

                    // Use connectionString to Create a database
                    var db = new LiteDatabase(connectionString);

                    // Create a table in the database
                    var table = db.GetCollection(tableName);

                    //Open the first sheet in the excel file and count row and column number
                    ExcelApp.Workbook excelBook = excelApp.Workbooks.Open(_excelPath);
                    ExcelApp._Worksheet excelSheet = excelBook.Sheets[1];
                    ExcelApp.Range excelRange = excelSheet.UsedRange;
                    int rows = excelRange.Rows.Count;
                    int cols = excelRange.Columns.Count;

                    // In this section we try to determine type of the cells
                    // We created an empty list "types" and add type of each column in it.
                    // For example types[0] determines type of the first column.
                    List<string> types = new List<string>();
                    for (int i = 1; i <= cols; i++)
                    {
                        // We save all the cells in a column into a list "cells"
                        List<string> cells = new List<string>();
                        for (int j = 2; j <= rows; j++)
                        {
                            cells.Add(excelRange.Cells[j, i].Value2.ToString());
                        }

                        //  determine type of a column based on the "cells" and add it to the "types"
                        string type = DetectType(cells);
                        types.Add(type);
                    }

                    // In this loop we add each row into the table
                    // First we determine first Bson title from the header of the column. Then we get type of each cell based on
                    // "Types" list. Then parse the value to the desire type. At the end, we add Id for each row and insert the
                    // Bson file in the table

                    for (int i = 2; i <= rows; i++)
                    {
                        BsonDocument bson = new BsonDocument();
                        for (int j = 1; j <= cols; j++)
                        {
                            string type = types[j - 1];
                            string columnName = excelRange.Cells[1, j].Value2.ToString();
                            bson.Add(columnName, TypeConvertor(type, excelRange.Cells[i, j].Value2.ToString()));

                        }

                        bson["_id"] = i - 1;
                        table.Insert(bson);

                    }

                    excelApp.Quit();
                    db.Dispose();
                    ShowMessage(Brushes.DarkGreen, "Database successfully created in " + folderBrowserDialog.SelectedPath);
                }


            }
            catch (Exception exception)
            {
                ShowMessage(Brushes.DarkRed, "Error: " + exception.Message);
            }

            

        }

        // Get path and name of the excel file
        private void BtnOpenFile_OnClick(object sender, RoutedEventArgs e)
        {
            OpenFileDialog getExcelFile = new OpenFileDialog();
            getExcelFile.Filter = "xls files (*.xls)|*.xls|xlsx files (*.xlsx)|*.xlsx";
            getExcelFile.FilterIndex = 2;
            getExcelFile.Multiselect = false;
            var result = getExcelFile.ShowDialog();
            if (result == System.Windows.Forms.DialogResult.OK)
            {
                _excelPath = getExcelFile.FileName;
                TextFileName.Text = getExcelFile.SafeFileName;
            }


        }

        //Convert a value based on the given type
        public object TypeConvertor(string type, string input)
        {
            if (type == "int")
                return Int32.Parse(input);
            if (type == "double")
                return double.Parse(input);
            if (type == "bool")
                return bool.Parse(input);
            return input;

        }

        // We use Regex to determine type of list.
        // If all members only contain numbers, it will return int.
        // If it contains numbers but one of them also contains "." , it will return double
        // If it only contains false or true, it will return bool.
        // Else, it will return string.
        public string DetectType(List<string> input)
        {
            string type = "string";
            if (input.All(x => Regex.IsMatch(x, @"^[0-9\.]*$")))
            {
                type = input.Any(x => Regex.IsMatch(x, @"\.")) ? "double" : "int";
            }

            if (input.All(x => (x.ToLower() == "false") || x.ToLower() == "true"))
                type = "bool";

            return type;

        }

        public void ShowMessage(SolidColorBrush color, string message)
        {
            TextLog.Foreground = color;
            TextLog.Text = message;
        }


    }
}
