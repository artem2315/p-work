using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace ConsoleAppStudentMarks
{
    class Program
    {
        static void Main(string[] args)
        {
            ExcelReader excelReader = new ExcelReader();
            Console.WriteLine("Process start!");
            try
            {
                Console.WriteLine("Result: " + excelReader.ReadWriteResult().ToString("#.###") +"%");
                Console.WriteLine("Process finish! Check result file.");
            }
            catch(Exception e)
            {
                Console.WriteLine("Error!");
            }

            Console.ReadLine();
        }
    }

    class Mark
    {
        public int IdStudent { get; set; }
        public int MarkVal { get; set; } = 0;
        public string ParseMark
        {
            set
            {
                int resParse = 0;
                Int32.TryParse(value, out resParse); MarkVal = resParse;
            }
        }
    }

    class ExcelReader
    {
        public Workbook DucumentFile = null;
        public Application ExcelApplic = null;
        public Worksheet SheetPaper = null;
        public List<Mark> marksArray = new List<Mark>();
        List<IGrouping<int, Mark>> studentsArray;

        public float ReadWriteResult()
        { // работа с ексель файлами
            ExcelApplic = new Application();
            ExcelApplic.Visible = false;
            DucumentFile = ExcelApplic.Workbooks.Open(System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "STUDENT_MARKS.xls"));
            SheetPaper = (Worksheet)DucumentFile.Sheets[1];
            var rowsCount = SheetPaper.Cells.SpecialCells(XlCellType.xlCellTypeLastCell).Row;

            for (int i = 2; i <= rowsCount; i++)
            {
                Array MyValues = SheetPaper.get_Range("A" +i, "M" + i).Cells.Value;
                marksArray.Add(new Mark
                {
                    IdStudent = Convert.ToInt32(MyValues.GetValue(1, 1)),
                    ParseMark = MyValues.GetValue(1, 6).ToString()
                });
            }

            float result = GetResult();
            ExcelApplic = new Application();
            DucumentFile = ExcelApplic.Workbooks.Open(System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ResultFile.xlsx"));
            SheetPaper = (Worksheet)DucumentFile.Sheets[1];
            SheetPaper.Cells[1, 2] = studentsArray.Count();
            SheetPaper.Cells[2, 2] = result;
            DucumentFile.Save();
            return result;
        }

        public float GetResult()
        { // подсчет отличников
            int resultCount = 0;
            studentsArray = marksArray.GroupBy(x => x.IdStudent).ToList();
            foreach(var x in studentsArray)
            {
                bool best = true;
                foreach (var y in x) { if (y.MarkVal < 90){ best = false; } }
                if (best) { resultCount++; }
            }
            return ((float)100 / studentsArray.Count) * resultCount;
        }
    }
}
