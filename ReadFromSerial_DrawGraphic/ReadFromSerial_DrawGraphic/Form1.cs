using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO.Ports;
using System.Threading;
using Excel = Microsoft.Office.Interop.Excel; //excel referans kütüphaneleri ekleniyor!


namespace ReadFromSerial_DrawGraphic
{
    public partial class Form1 : Form
    {
        public SerialPort myport;
        Thread thread1;

        private string excelFilePath = string.Empty;
        Excel.Workbook xlWorkBook;
        Excel.Worksheet xlWorkSheet;
        object misValue;

        DateTime theTime; // zamanı excel dosyasına yazdırmak icin olusturuldu

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            myport = new SerialPort(); // SerialPort olusturuluyor
            myport.BaudRate= 9600; // SerialPort hızı belirleniyor
            myport.PortName = "COM11"; // Kullanılacak Port belirleniyor
            myport.Open(); // port aciliyor
            
            chart1.Series["Series1"].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.FastLine; // cizdirilecek grafik türü
            chart1.ChartAreas[0].AxisX.MajorGrid.Enabled = false; // ekseninde chart taki arka plani temizliyoruz
            chart1.ChartAreas[0].AxisY.MajorGrid.Enabled = false;
            chart1.ChartAreas[0].AxisY.Maximum = 40;
            chart1.ChartAreas[0].AxisY.Minimum = 15;

            misValue = System.Reflection.Missing.Value;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            thread1 = new Thread(new ThreadStart(drawGraph));
            thread1.Start();      
        }
        private void drawGraph()
        {

            Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

            try
            {
                xlWorkBook = xlApp.Workbooks.Open("C:\\Users\\DELL\\Desktop\\ReadFromSerial_DrawGraphic\\csharp-Excel.xls", 0, false, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1); //Get all the sheets in the workbook

                while (thread1.IsAlive)
                {
                    //son satır bulunuyor excel dosyasındaki
                    Excel.Range last = xlWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
                    Excel.Range range = xlWorkSheet.get_Range("A1", last);
                    int lastUsedRow = last.Row;
                    int lastUsedColumn = last.Column;

                    string ReceiveData = myport.ReadLine(); // comdan degeri okuyuruz

                    // alınan degerdeki stringleri temizleyerek sadece double değeri yakalıyor
                    string[] HeatingData = ReceiveData.Split(':');
                    string[] HeatingData2 = HeatingData[1].Split('D');
                    var result = HeatingData2[0];
                    double heating = Convert.ToDouble(result);

                    theTime = DateTime.Now; // anlik olarak zamani ogreniyoruz!
                    string zaman = theTime.ToString("yyyy/MM/dd HH:mm:ss");

                    Thread.Sleep(1000); // ilk threadi anlik olarak durduruyor ve Invoke ile GUI threadini ulasip cizdiriyor! 
                    this.Invoke((MethodInvoker)delegate
                       {
                           chart1.Series["Series1"].Points.AddY(result);
                           // excel dosyasındaki son yazılan satırdan bir sonraki satıra sıcaklığı yazdırıyor
                           xlWorkSheet.Cells[lastUsedRow+1, 2] = (heating / 100);
                           xlWorkSheet.Cells[lastUsedRow + 1, 1] = zaman;
                       });
                }
            }
            catch
            {
                // MessageBox.Show("Dosya bulunamadı");
                xlWorkBook = xlApp.Workbooks.Add(misValue);
                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

                xlWorkSheet.Cells[1, 1] = "Zaman";
                xlWorkSheet.Cells[1, 2] = "Sıcaklık Celcius";

                xlWorkBook.SaveAs("C:\\Users\\DELL\\Desktop\\ReadFromSerial_DrawGraphic\\csharp-Excel.xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                xlWorkBook.Close(true, misValue, misValue);
                xlApp.Quit();

                releaseObject(xlWorkSheet);
                releaseObject(xlWorkBook);
                releaseObject(xlApp);

                MessageBox.Show("Dosya oluşturuldu , proje klasörünüzde bulunmaktadır");
            }
        }
        private void releaseObject(object obj) //gecmis datayı , objeleri temizlemek icin
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Exception Occured while releasing object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }
        private void button2_Click(object sender, EventArgs e)
        {
            // cizdirmeyi durdurmak icin ilk threadi sonlandiriyor!
            thread1.Abort();
            xlWorkBook.Close(misValue, misValue, misValue);
            //xlApp.Quit();

            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            //releaseObject(xlApp);
        }
    }
}
