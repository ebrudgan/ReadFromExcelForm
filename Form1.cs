using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ReadFromExcelForm
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            {
                var input = textBox1.Text.ToLowerAndTurkishCharacter();
                var excelDataList = new List<ExcelSheet>();
                //Create COM Objects. Create a COM object for everything that is referenced
                Excel.Application xlApp = new Excel.Application();
                Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
                string desktop = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
                Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(Path.Combine(desktop, "Inventiv_Giris_Cikis_Raporu.xlsx"));
                Excel._Worksheet xlWorksheet = xlWorkbook.Sheets["Giriş Çıkış"];
                Excel.Range xlRange = xlWorksheet.UsedRange;
                int rowCount = xlRange.Rows.Count;
                int colCount = xlRange.Columns.Count;

                //iterate over the rows and columns and print to the console as it appears in the file
                //excel is not zero based!!
                string[] formats = { "dd.MM.yyyy HH:mm:ss", "d.MM.yyyy HH:mm:ss" };
                for (int i = 2; i <= rowCount; i++)
                {
                    try
                    {
                        excelDataList.Add(
                            new ExcelSheet
                            {
                                RowNum = i - 1,
                                FullName = xlRange.Cells[i, 1].Value2.ToString(),
                                Action = xlRange.Cells[i, 2].Text == "Çıkış" ? ActionType.Out : ActionType.In,
                                Date = DateTime.ParseExact(xlRange.Cells[i, 3].Value2.ToString(), formats, CultureInfo.InvariantCulture, DateTimeStyles.None),
                                Terminal = xlRange.Cells[i, 4].Value2.ToString(),
                            }
                        );
                    }
                    catch (Exception ex)
                    {
                        throw ex;
                    }
                }

               // string pazartesi = CultureInfo.GetCultureInfo("tr-TR").DateTimeFormat.DayNames[(int)(DateTime.Date.DayOfWeek) + 1];

                var KisininSiraliGirisSaatleri = excelDataList.Where(c => c.FullName.ToLowerAndTurkishCharacter() == input && c.Action == ActionType.In).Select(c => c.Date).OrderBy(c => c.Date).ToList();
                var KisininSiraliCikisSaatleri = excelDataList.Where(c => c.FullName.ToLowerAndTurkishCharacter() == input && c.Action == ActionType.Out).Select(c => c.Date).OrderBy(c => c.Date).ToList();
                TimeSpan toplamsaat = default(TimeSpan);
                var totalHours = default(double);
                for (int i = 0; i < KisininSiraliGirisSaatleri.Count; i++)
                {
                    toplamsaat = KisininSiraliCikisSaatleri[i] - KisininSiraliGirisSaatleri[i];
                    totalHours += toplamsaat.TotalHours;

                    
                }
                MessageBox.Show("Toplam Çalışma Saati:" + totalHours);
                #region Handle Garbage Collector and Clean Up
                //cleanup
                GC.Collect();
                GC.WaitForPendingFinalizers();
                //rule of thumb for releasing com objects:
                //  never use two dots, all COM objects must be referenced and released individually
                //  ex: [somthing].[something].[something] is bad
                //release com objects to fully kill excel process from running in the background
                Marshal.ReleaseComObject(xlRange);
                Marshal.ReleaseComObject(xlWorksheet);
                //close and release
                xlWorkbook.Close();
                Marshal.ReleaseComObject(xlWorkbook);
                //quit and release
                xlApp.Quit();
                Marshal.ReleaseComObject(xlApp);
                #endregion

            }

        }
    }
}

        
