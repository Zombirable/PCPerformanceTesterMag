using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using System.Data;

namespace PCPerformanceTester
{
    class Program
    {
        static string dir = Environment.CurrentDirectory.Replace("\\","\\\\");
        static string machineName = Environment.MachineName;
        static List<TestResult> Results = new List<TestResult>();
        static void Main(string[] args)
        {
            Console.WriteLine("init");
            string manyFileZipPath = $"{Environment.CurrentDirectory}\\testing_7-z\\7-zip-execute.bat";
            string videoZipPath = $"{Environment.CurrentDirectory}\\testing_7-z\\7-zip-vid_execute.bat";

            for(int i=0;i<5;i++)
            {
                Execute7zipBatch(manyFileZipPath, "7ZIP_CODE_EXTRACTING");

                ExcelTest();

                Execute7zipBatch(videoZipPath, "7ZIP_VIDEO_EXTRACTING");

                WordTest();
            }
        }

        private static void Execute7zipBatch(string path, string type)
        {
            Console.WriteLine($"Executing batch: {path}");
            int ExitCode;
            ProcessStartInfo ProcessInfo;
            Process process;

            ProcessInfo = new ProcessStartInfo(path);
            ProcessInfo.CreateNoWindow = true;
            ProcessInfo.UseShellExecute = false;
            ProcessInfo.RedirectStandardError = true;
            ProcessInfo.RedirectStandardOutput = true;

            Console.WriteLine($"Starting task");
            Stopwatch watch = new Stopwatch();
            watch.Start();
            process = Process.Start(ProcessInfo);
            process.WaitForExit();
            watch.Stop();
            Console.WriteLine($"Ending task");

            string output = process.StandardOutput.ReadToEnd();
            string error = process.StandardError.ReadToEnd();

            ExitCode = process.ExitCode;
            Console.WriteLine($"Elapsed time: {watch.ElapsedMilliseconds}");
            Results.Add(new TestResult(machineName, type, dir, watch.ElapsedMilliseconds));
            Console.WriteLine("ExitCode: " + ExitCode.ToString(), "Execute7zipBat");
            process.Close();
        }


        private static DataTable createDataTable()
        {
            Console.WriteLine("Создание наполнения таблицы");

            DataTable outDT = new DataTable();
            outDT.Columns.Add("COL_1", typeof(string));
            outDT.Columns.Add("COL_2", typeof(string));
            outDT.Columns.Add("COL_3", typeof(string));
            outDT.Columns.Add("COL_4", typeof(string));
            outDT.Columns.Add("COL_5", typeof(string));

            for(int i=0;i<1000000;i++)
            {
                string GUID_1 = Guid.NewGuid().ToString();
                string GUID_2 = Guid.NewGuid().ToString();
                string GUID_3 = Guid.NewGuid().ToString();
                string GUID_4 = Guid.NewGuid().ToString();
                string GUID_5 = Guid.NewGuid().ToString();

                outDT.Rows.Add(GUID_1, GUID_2, GUID_3, GUID_4, GUID_5);
            }

            Console.WriteLine("Наполнение таблицы создано");
            return outDT;
        }


        private static void OutputCSV(List<TestResult> TestResults)
        {
            List<string> Text = new List<string>();
            foreach(TestResult result in TestResults)
            {
                Text.Add($"{result.MachineName};{result.TestType};{result.StartDirectory};{result.ElapsedTime};{result.TestTimestamp.ToString("dd:MM:yyyy HH:mm:ss")}");
            }
            File.WriteAllLines($"{Environment.CurrentDirectory}\\PCOutput\\RESULT_{machineName}_{DateTime.Now.ToString("dd:MM:yyyy HH:mm:ss:ffffff")}.CSV", Text.ToArray());
        }

        private static DataTable prepareDataTable(DataTable table)
        {

            Console.WriteLine("Подготовка таблицы");
            DataTable outDT = table;

            foreach (DataRow row in outDT.Rows)
            {
                for (int i = 0; i < row.ItemArray.Count(); i++)
                {
                    try
                    {
                        if (row[i] == DBNull.Value)
                            row[i] = "NULL";
                    }
                    catch
                    {
                    }
                }
                string comment = row.ItemArray[row.ItemArray.Count() - 1].ToString().TrimStart('\n', '\r').TrimEnd('\n', '\r');
                row[row.ItemArray.Count() - 1] = (object)comment;
            }
            Console.WriteLine("Подготовка таблицы завершена");
            return outDT;
        }

        public static void ExcelTest()
        {
            

            DataTable outDT = prepareDataTable(createDataTable());
            int rowcount = outDT.Rows.Count;
            Console.WriteLine("Создание Excel файла отчета по чек-листам");

            Console.WriteLine("Формирование строк");
            object[] Header = new object[outDT.Columns.Count];

            for (int i = 0; i < outDT.Columns.Count; i++)
            {
                Header[i] = outDT.Columns[i].ColumnName;
            }

            object[,] Cells = new object[outDT.Rows.Count, outDT.Columns.Count];

            for (int j = 0; j < outDT.Rows.Count; j++)
            {
                for (int i = 0; i < outDT.Columns.Count; i++)
                {
                    Cells[j, i] = outDT.Rows[j][i];

                }
            }

            Stopwatch watch = new Stopwatch();
            watch.Start();

            var excelApp = new Excel.Application();
            var workbook = excelApp.Workbooks.Add();
            try
            {
                excelApp.Range["A1:E1"].Value = Header;
                excelApp.Range["A2:E" + (outDT.Rows.Count + 1).ToString()].Value = Cells;
                outDT = null;
                Cells = null;
                Header = null;
                Console.WriteLine("Создание форматирования");

                excelApp.Range["A" + (rowcount + 1).ToString()].Columns.ColumnWidth = 20;
                excelApp.Range["B" + (rowcount + 1).ToString()].Columns.ColumnWidth = 20;
                excelApp.Range["C" + (rowcount + 1).ToString()].Columns.ColumnWidth = 20;
                excelApp.Range["D" + (rowcount + 1).ToString()].Columns.ColumnWidth = 20;
                excelApp.Range["E" + (rowcount + 1).ToString()].Columns.ColumnWidth = 20;

                excelApp.Range["A1:G1"].Cells.Font.Bold = true;

                excelApp.Range["A1:E" + (rowcount + 1).ToString()].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                excelApp.Range["A1:E" + (rowcount + 1).ToString()].Style.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                excelApp.Range["A1:E" + (rowcount + 1).ToString()].Style.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                excelApp.Range["S1:E" + (rowcount + 1).ToString()].Style.WrapText = true;
                excelApp.Range["A1:E" + (rowcount + 1).ToString()].Rows.AutoFit();

                Console.WriteLine("Вывод документа");
                string filepath = $"{Environment.CurrentDirectory}\\testing_excel\\text_excel_{DateTime.Now.ToString("dd_MM_yyyy_HH_mm_ss")}.xlsx";
                workbook.SaveAs(filepath);

                workbook.Close();
                excelApp.Quit();
                
                watch.Stop();
                Console.WriteLine($"Time elapsed: {watch.ElapsedMilliseconds}");
                Results.Add(new TestResult(machineName, "EXCEL_TEST_WRITE", dir, watch.ElapsedMilliseconds));

                Console.WriteLine("Открытие сохранённого документа");
                watch = new Stopwatch();
                watch.Start();
                excelApp = new Excel.Application();
                workbook=excelApp.Workbooks.Open(filepath);

                workbook.Close();
                excelApp.Quit();
                watch.Stop();
                Console.WriteLine($"Time elapsed: {watch.ElapsedMilliseconds}");
                Results.Add(new TestResult(machineName, "EXCEL_TEST_READ", dir, watch.ElapsedMilliseconds));
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка: {ex.Message}");
                excelApp.Quit();
                watch.Stop();
            }
        }

        public static void WordTest()
        {
            Console.WriteLine("Инициализация Word");
            string doc_path = $"{Environment.CurrentDirectory}\\testing_word\\text_word_{DateTime.Now.ToString("dd_MM_yyyy_HH_mm_ss")}.docx";

            Console.WriteLine("Создание документа");
            Stopwatch watch = new Stopwatch();
            watch.Start();
            Word.Application oWordApp = new Microsoft.Office.Interop.Word.Application();
            Word.Document oWordDoc = new Word.Document();

            try
            {
               for(int i=0;i<200;i++)
                {
                    Word.Paragraph para = oWordDoc.Paragraphs.Add();

                    string text = File.ReadAllText($"{Environment.CurrentDirectory}\\testing_word\\text_for_word.txt");

                    para.Range.Text = text;
                    para.Range.InsertParagraphAfter();                  
                }

                
                oWordDoc.SaveAs2(doc_path);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка: {ex.Message}");
            }

            oWordDoc.Close();
            oWordApp.Application.Quit(0);
            watch.Stop();
            Console.WriteLine($"Time elapsed: {watch.ElapsedMilliseconds}");
            Results.Add(new TestResult(machineName, "WORD_TEST_WRITE", dir, watch.ElapsedMilliseconds));
            Console.WriteLine("Чтение документа");

            watch = new Stopwatch();
            watch.Start();
            oWordApp = new Word.Application();

            try
            {
                
                oWordDoc = oWordApp.Documents.Open(doc_path);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка: {ex.Message}");
            }

            oWordDoc.Close();
            oWordApp.Application.Quit(0);
            watch.Stop();
            Results.Add(new TestResult(machineName, "WORD_TEST_READ", dir, watch.ElapsedMilliseconds));
            Console.WriteLine($"Time elapsed: {watch.ElapsedMilliseconds}");
            Console.WriteLine("Чтение документа завершено");
        }
    }
}
