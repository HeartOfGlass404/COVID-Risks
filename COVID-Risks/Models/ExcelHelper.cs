using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WinFormsApp1.Models;
using Excel = Microsoft.Office.Interop.Excel;

namespace COVID_Risks.Models
{
    class ExcelHelper : IDisposable
    {
        private Excel.Application _excel;
        private Workbook? _workbook;
        private Worksheet? _worksheet;
        private string? _filePath;

        public ExcelHelper()
        {
            _excel = new Excel.Application();
        }

        public void Dispose()
        {
            try
            {
                _workbook.Close();
                _excel.Quit();
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }

        internal bool Open(string _filePath)
        {
            try
            {
                if(File.Exists(_filePath))
                {
                    _workbook = _excel.Workbooks.Open(_filePath);
                }
                else
                {
                    _workbook = _excel.Workbooks.Add();
                    _filePath= Path.GetFullPath(_filePath);
                }
                return true;
            }
            catch (Exception e)
            {

                Console.WriteLine(e.Message);
                return false;
            }
        }

        internal void Save()
        {
            if (!string.IsNullOrEmpty(_filePath))
            {
                _workbook.SaveAs(_filePath);
                _filePath = null; ;
            }
            else
            {
                _workbook.Save();
            }
        }

        internal void Create()
        {
            _excel.Workbooks.Add();
            _excel.Visible = true;
            _excel.Range["D1"].Value = "Акушерский анамнез";
            _excel.Range["D1", "F1"].Merge();
            _excel.Range["G1"].Value = "Соматический анамнез";
            _excel.Range["G1", "I1"].Merge();
            _excel.Range["J1"].Value = "Течение беременности";
            _excel.Range["J1", "AG1"].Merge();
            _excel.Range["AH1"].Value = "Течение SARS-CoV-2";
            _excel.Range["AH1", "BB1"].Merge();

            _workbook.SaveAs("D:\\Desktop\\test.output.xlsx");
        }

        internal bool Set(string collumn, int row, string data)
        {
            try
            {
                ((Excel.Worksheet)_excel.ActiveSheet).Cells[row, collumn].Value = data; 
                return true;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return false;
            }
        }

        public static List<Person> GetPersonData(String _filePath)
        {
            Excel.Application application = new Excel.Application();
            application.DisplayAlerts= false;
            Excel.Workbook workbook = application.Workbooks.Open(_filePath);
            Excel.Worksheet worksheet = workbook.Worksheets[1];


            List<Person> personlist = new List<Person>();


            int rows = worksheet.UsedRange.Rows.Count;
            int cols = worksheet.UsedRange.Columns.Count;






            for (int i = 4; i <= rows; i++)
            {


                if (worksheet.Cells[i, 1].Value2 != null)
                {
                    personlist.Add(new Person(
                        Convert.ToString(worksheet.Cells[i, 1].Value2),
                        Convert.ToInt16(worksheet.Cells[i, 2].Value2),
                        Convert.ToString(worksheet.Cells[i, 3].Value2),
                        Convert.ToString(worksheet.Cells[i, 4].Value2),
                        Convert.ToString(worksheet.Cells[i, 5].Value2),
                        Convert.ToString(worksheet.Cells[i, 6].Value2),
                        Convert.ToString(worksheet.Cells[i, 7].Value2),
                        Convert.ToString(worksheet.Cells[i, 8].Value2),
                        Convert.ToString(worksheet.Cells[i, 9].Value2),
                        Convert.ToString(worksheet.Cells[i, 10].Value2),
                        Convert.ToString(worksheet.Cells[i, 11].Value2),
                        Convert.ToString(worksheet.Cells[i, 12].Value2),
                        Convert.ToString(worksheet.Cells[i, 13].Value2),
                        Convert.ToString(worksheet.Cells[i, 14].Value2),
                        Convert.ToString(worksheet.Cells[i, 15].Value2),
                        Convert.ToString(worksheet.Cells[i, 16].Value2),
                        Convert.ToString(worksheet.Cells[i, 17].Value2),
                        Convert.ToString(worksheet.Cells[i, 18].Value2),
                        Convert.ToString(worksheet.Cells[i, 19].Value2),
                        Convert.ToString(worksheet.Cells[i, 20].Value2)));
                }


            }

            return personlist;
        }
   
    }
}
