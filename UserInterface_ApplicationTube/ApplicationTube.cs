using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using System.Windows.Forms;
using System.Data.OleDb;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace WindowsFormsApplication2
{
    public partial class ApplicationTube : Form
    {
        string excelFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "DataExcel.xls");
        SqlConn data = new SqlConn();

        public ApplicationTube()
        {
            InitializeComponent();
        }

        private void menuCalibration_Click(object sender, EventArgs e)
        {
            Calibration calibration = new Calibration();
            calibration.Show();
        }

        private void buttonExit_Click(object sender, EventArgs e)
        {
            DialogResult rs = new DialogResult();
            rs = MessageBox.Show("Do You Want to Exit?", "Exit", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (rs == DialogResult.Yes)
            {
                Application.Exit();
            }
            else
            {
                this.Show();
            }
        }

        private void Form2_Load(object sender, EventArgs e)
        {

            DeviceType();
            Machine();
            DisplayExcel("DATA");
            this.Menu = new MainMenu();

            //MENU
            MenuItem item1 = new MenuItem("MENU");
            this.Menu.MenuItems.Add(item1);
                item1.MenuItems.Add("CALIBRATION", new EventHandler(menuCalibration_Click));
                item1.MenuItems.Add("DATA_SQL", new EventHandler(menuSave_Click));
        }

        private void SaveExcelFile(string savePath)
        {
            // Buka aplikasi Excel
            Excel.Application excelApp = new Excel.Application();
            excelApp.Visible = false;

            // Buka file Excel
            Excel.Workbook excelWorkbook = excelApp.Workbooks.Open(excelFilePath);

            try
            {
                // Simpan file Excel ke lokasi yang dipilih
                excelWorkbook.SaveAs(savePath, Excel.XlFileFormat.xlWorkbookDefault);
                MessageBox.Show("File successfully saved in: " + savePath);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error occurred: " + ex.Message);
            }
            finally
            {
                // Tutup aplikasi Excel
                excelWorkbook.Close(false);
                excelApp.Quit();

                // Bersihkan objek Excel dari memori
                releaseObject(excelWorkbook);
                releaseObject(excelApp);
            }
        }

        public void DisplayExcel(string sheetName)
        {
            // Buka aplikasi Excel
            Excel.Application excelApp = new Excel.Application();
            excelApp.Visible = false;

            // Buka file Excel
            Excel.Workbook excelWorkbook = excelApp.Workbooks.Open(excelFilePath);

            // Cari sheet dengan nama yang sesuai
            Excel.Worksheet excelWorksheet = null;
            foreach (Excel.Worksheet sheet in excelWorkbook.Sheets)
            {
                if (sheet.Name == sheetName)
                {
                    excelWorksheet = sheet;
                    break;
                }
            }

            // Jika sheet tidak ditemukan, keluarkan pesan kesalahan
            if (excelWorksheet == null)
            {
                MessageBox.Show("Sheet dengan nama '" + sheetName + "' tidak ditemukan!");
                excelWorkbook.Close(false);
                excelApp.Quit();
                releaseObject(excelWorkbook);
                releaseObject(excelApp);
                return;
            }

            // Simpan file Excel ke tempat sementara (bisa juga disimpan di temp path)
            string tempFilePath = System.IO.Path.GetTempFileName() + ".html";
            excelWorksheet.SaveAs(tempFilePath, Excel.XlFileFormat.xlHtml);

            // Tampilkan file HTML di WebBrowser
            webBrowser1.Navigate(tempFilePath);

            // Tutup aplikasi Excel
            excelWorkbook.Close(false);
            excelApp.Quit();

            // Bersihkan objek Excel dari memori
            releaseObject(excelWorksheet);
            releaseObject(excelWorkbook);
            releaseObject(excelApp);
        }

        private void releaseObject(object obj)
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

        private void ButtonPrint_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Excel Files|*.xls;*.xlsx";
            saveFileDialog.Title = "Save an Excel File";
            saveFileDialog.DefaultExt = "xlsx";
            saveFileDialog.FileName = "MyExcelFile.xlsx"; // Default file name

            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                string savePath = saveFileDialog.FileName;
                SaveExcelFile(savePath);
            }
        }

        private void menuSave_Click(object sender, EventArgs e)
        {
            DataSql Dsql = new DataSql();
            Dsql.Show();
        }

        private void buttonExecute_Click(object sender, EventArgs e)
        {   
            // Koneksi ke SQL Server
            string connectionString = "Data Source=etcbtmsql01\\digitmont;Initial Catalog=dbDigitmontDetection;User ID=Metis;Password=Metis123";
            string query = "SELECT * FROM TESTER_APPLICATION_TUBE WHERE DATE >= @StartDate AND DATE <= @EndDate AND MACHINE_NUMBER = @Machine AND DEVICE_TYPE = @DeviceType";

            // Tanggal rentang
            DateTime startDate = dateTimePicker1.Value.Date;
            DateTime endDate = dateTimePicker2.Value.Date;
            string machine = CB_Machine.Text;
            string deviceType = CB_DeviceType.Text;
            DataTable dataTable = new DataTable();

            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                using (SqlCommand command = new SqlCommand(query, connection))
                {
                    // Parameter
                    command.Parameters.AddWithValue("@StartDate", startDate);
                    command.Parameters.AddWithValue("@EndDate", endDate);
                    command.Parameters.AddWithValue("@Machine", machine);
                    command.Parameters.AddWithValue("@DeviceType", deviceType);
                    connection.Open();
                    using (SqlDataAdapter adapter = new SqlDataAdapter(command))
                    {
                        adapter.Fill(dataTable);
                    }
                }
            }
            clearExcel();

            // Jika tidak ada data yang sesuai, tampilkan pesan
            if (dataTable.Rows.Count == 0)
            {
                MessageBox.Show("This Data Not Found!!!, Please Find Other Data");
            }
            else
            {
                // Menyisipkan data ke Excel
                Excel.Application excelApp = new Excel.Application();
                excelApp.Visible = false;

                // Buka workbook yang sudah ada
                Excel.Workbook workbook = excelApp.Workbooks.Open(excelFilePath);
                Excel.Worksheet worksheet = workbook.Sheets["DATA"];

                int currentColumn1 = 3;
                int currentColumn2 = 4;

                worksheet.Cells[3, 19] = GlobalVariable.tempLow;
                worksheet.Cells[3, 20] = GlobalVariable.tempHigh;

                worksheet.Cells[3, 23] = GlobalVariable.Dv;
                worksheet.Cells[3, 24] = GlobalVariable.Dv;

                worksheet.Cells[4, 23] = "-" + GlobalVariable.Dv;
                worksheet.Cells[4, 24] = "-" + GlobalVariable.Dv;

                worksheet.Cells[5, 23] = GlobalVariable.Range;
                worksheet.Cells[5, 24] = GlobalVariable.Range;

                /*
                if (CB_Machine.SelectedItem.ToString() == "COMP1")
                {
                    worksheet.Cells[3, 10] = "Comparative : # 01";
                }
                else if (CB_Machine.SelectedItem.ToString() == "COMP2")
                {
                    worksheet.Cells[3, 10] = "Comparative : # 02";
                }
                else if (CB_Machine.SelectedItem.ToString() == "COMP3")
                {
                    worksheet.Cells[3, 10] = "Comparative : # 03";
                }
                else if (CB_Machine.SelectedItem.ToString() == "COMP4")
                {
                    worksheet.Cells[3, 10] = "Comparative : # 04";
                }
                else if (CB_Machine.SelectedItem.ToString() == "COMP5")
                {
                    worksheet.Cells[3, 10] = "Comparative : # 05";
                }
                else if (CB_Machine.SelectedItem.ToString() == "COMP6")
                {
                    worksheet.Cells[3, 10] = "Comparative : # 06";
                }
                else if (CB_Machine.SelectedItem.ToString() == "COMP7")
                {
                    worksheet.Cells[3, 10] = "Comparative : # 07";
                }
                else if (CB_Machine.SelectedItem.ToString() == "COMP8")
                {
                    worksheet.Cells[3, 10] = "Comparative : # 08";
                }
                else if (CB_Machine.SelectedItem.ToString() == "COMP9")
                {
                    worksheet.Cells[3, 10] = "Comparative : # 09";
                }
                else if (CB_Machine.SelectedItem.ToString() == "COMP10")
                {
                    worksheet.Cells[3, 10] = "Comparative : # 10";
                }
                else if (CB_Machine.SelectedItem.ToString() == "COMP11")
                {
                    worksheet.Cells[3, 10] = "Comparative : # 11";
                }
                else if (CB_Machine.SelectedItem.ToString() == "COMP12")
                {
                    worksheet.Cells[3, 10] = "Comparative : # 12";
                }
                else if (CB_Machine.SelectedItem.ToString() == "COMP13")
                {
                    worksheet.Cells[3, 10] = "Comparative : # 13";
                }
                else if (CB_Machine.SelectedItem.ToString() == "COMP14")
                {
                    worksheet.Cells[3, 10] = "Comparative : # 14";
                }
                else if (CB_Machine.SelectedItem.ToString() == "COMP15")
                {
                    worksheet.Cells[3, 10] = "Comparative : # 15";
                }
                else if (CB_Machine.SelectedItem.ToString() == "COMP16")
                {
                    worksheet.Cells[3, 10] = "Comparative : # 16";
                }
                else if (CB_Machine.SelectedItem.ToString() == "COMP17")
                {
                    worksheet.Cells[3, 10] = "Comparative : # 17";
                }
                else if (CB_Machine.SelectedItem.ToString() == "COMP18")
                {
                    worksheet.Cells[3, 10] = "Comparative : # 18";
                }
                else if (CB_Machine.SelectedItem.ToString() == "COMP19")
                {
                    worksheet.Cells[3, 10] = "Comparative : # 19";
                }*/
                string comparativeText = "Comparative : # ";
                if (CB_Machine.SelectedItem != null)
                {
                    worksheet.Cells[3, 10] = comparativeText + CB_Machine.SelectedItem.ToString().Replace("COMP", "");
                }

                // Mengisi data dari DataTable ke Excel
                for (int i = 0; i < dataTable.Rows.Count; i++)
                {
                    string filePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Datainput.txt");
                    string fileContent = File.ReadAllText(filePath);
                    string tempSetting = dataTable.Rows[i]["BB_TEMP_SETTING"].ToString();
                    if (fileContent.Contains(CB_DeviceType.Text))
                    {
                        if (tempSetting == "HIGH      ")
                        {
                            // Menulis data ke sel Excel
                            worksheet.Cells[3, 2] = dataTable.Rows[i]["DEVICE_TYPE"];
                            worksheet.Cells[7, currentColumn1] = dataTable.Rows[i]["DATE"];
                            worksheet.Cells[8, currentColumn1] = dataTable.Rows[i]["SHIFT"] + "/" + dataTable.Rows[i]["START"] + "-" + dataTable.Rows[i]["FINISH"];
                            worksheet.Cells[10, currentColumn1] = dataTable.Rows[i]["TEST_JIG"];
                            worksheet.Cells[22, currentColumn1] = dataTable.Rows[i]["DONE_BY"];

                            //HIGH
                            worksheet.Cells[12, currentColumn2] = dataTable.Rows[i]["GS_READING"];
                            worksheet.Cells[13, currentColumn2] = dataTable.Rows[i]["DUT1_READING"];
                            worksheet.Cells[14, currentColumn2] = dataTable.Rows[i]["DUT2_READING"];
                            worksheet.Cells[15, currentColumn2] = dataTable.Rows[i]["DUT3_READING"];
                            worksheet.Cells[19, currentColumn2] = dataTable.Rows[i]["BUYOFF_RESULT"];
                            worksheet.Cells[20, currentColumn2] = dataTable.Rows[i]["MC_OFFSET_OLD"];
                            worksheet.Cells[21, currentColumn2] = dataTable.Rows[i]["MC_OFFSET_NEW"];
                            currentColumn2 += 2;
                            currentColumn1 += 2;
                        }
                        else if (tempSetting == "LOW       ")
                        {
                            // Menulis data ke sel Excel
                            worksheet.Cells[3, 2] = dataTable.Rows[i]["DEVICE_TYPE"];
                            worksheet.Cells[7, currentColumn1] = dataTable.Rows[i]["DATE"];
                            worksheet.Cells[8, currentColumn1] = dataTable.Rows[i]["SHIFT"] + "/" + dataTable.Rows[i]["START"] + "-" + dataTable.Rows[i]["FINISH"];
                            worksheet.Cells[10, currentColumn1] = dataTable.Rows[i]["TEST_JIG"];
                            worksheet.Cells[22, currentColumn1] = dataTable.Rows[i]["DONE_BY"];

                            //LOW
                            worksheet.Cells[12, currentColumn1] = dataTable.Rows[i]["GS_READING"];
                            worksheet.Cells[13, currentColumn1] = dataTable.Rows[i]["DUT1_READING"];
                            worksheet.Cells[14, currentColumn1] = dataTable.Rows[i]["DUT2_READING"];
                            worksheet.Cells[15, currentColumn1] = dataTable.Rows[i]["DUT3_READING"];
                            worksheet.Cells[19, currentColumn1] = dataTable.Rows[i]["BUYOFF_RESULT"];
                            worksheet.Cells[20, currentColumn1] = dataTable.Rows[i]["MC_OFFSET_OLD"];
                            worksheet.Cells[21, currentColumn1] = dataTable.Rows[i]["MC_OFFSET_NEW"];
                            currentColumn1 += 2;
                            currentColumn2 += 2;
                        }
                    }
                    else
                    {
                        if (tempSetting == "HIGH      ")
                        {
                            // Menulis data ke sel Excel
                            worksheet.Cells[3, 2] = dataTable.Rows[i]["DEVICE_TYPE"];
                            worksheet.Cells[7, currentColumn1] = dataTable.Rows[i]["DATE"];
                            worksheet.Cells[8, currentColumn1] = dataTable.Rows[i]["SHIFT"] + "/" + dataTable.Rows[i]["START"] + "-" + dataTable.Rows[i]["FINISH"];
                            worksheet.Cells[10, currentColumn1] = dataTable.Rows[i]["TEST_JIG"];
                            worksheet.Cells[22, currentColumn1] = dataTable.Rows[i]["DONE_BY"];

                            //HIGH
                            worksheet.Cells[12, currentColumn2] = dataTable.Rows[i]["GS_READING"];
                            worksheet.Cells[13, currentColumn2] = dataTable.Rows[i]["DUT1_READING"];
                            worksheet.Cells[14, currentColumn2] = dataTable.Rows[i]["DUT2_READING"];
                            worksheet.Cells[15, currentColumn2] = dataTable.Rows[i]["DUT3_READING"];
                            worksheet.Cells[19, currentColumn2] = dataTable.Rows[i]["BUYOFF_RESULT"];
                            worksheet.Cells[20, currentColumn2] = dataTable.Rows[i]["MC_OFFSET_OLD"];
                            worksheet.Cells[21, currentColumn2] = dataTable.Rows[i]["MC_OFFSET_NEW"];
                            currentColumn2 += 2;

                        }
                        else if (tempSetting == "LOW       ")
                        {
                            // Menulis data ke sel Excel
                            worksheet.Cells[3, 2] = dataTable.Rows[i]["DEVICE_TYPE"];
                            worksheet.Cells[7, currentColumn1] = dataTable.Rows[i]["DATE"];
                            worksheet.Cells[8, currentColumn1] = dataTable.Rows[i]["SHIFT"] + "/" + dataTable.Rows[i]["START"] + "-" + dataTable.Rows[i]["FINISH"];
                            worksheet.Cells[10, currentColumn1] = dataTable.Rows[i]["TEST_JIG"];
                            worksheet.Cells[22, currentColumn1] = dataTable.Rows[i]["DONE_BY"];

                            //LOW
                            worksheet.Cells[12, currentColumn1] = dataTable.Rows[i]["GS_READING"];
                            worksheet.Cells[13, currentColumn1] = dataTable.Rows[i]["DUT1_READING"];
                            worksheet.Cells[14, currentColumn1] = dataTable.Rows[i]["DUT2_READING"];
                            worksheet.Cells[15, currentColumn1] = dataTable.Rows[i]["DUT3_READING"];
                            worksheet.Cells[19, currentColumn1] = dataTable.Rows[i]["BUYOFF_RESULT"];
                            worksheet.Cells[20, currentColumn1] = dataTable.Rows[i]["MC_OFFSET_OLD"];
                            worksheet.Cells[21, currentColumn1] = dataTable.Rows[i]["MC_OFFSET_NEW"];
                            currentColumn1 += 2;
                        }
                    }
                }
                // Simpan workbook
                workbook.Save();
      
                // Tutup workbook
                workbook.Close();

                MessageBox.Show("Data On Machine " + CB_Machine.Text + " and Device Type " + CB_DeviceType.Text + " From The " + dateTimePicker1.Value.Date.ToString("yyyy-MM-dd") + " To The " + dateTimePicker2.Value.Date.ToString("yyyy-MM-dd") + " Was Found Successfully");
                DisplayExcel("DATA");
            }
        }

        private void clearExcel()
        {
            // Inisialisasi aplikasi Excel
            Excel.Application excelApp = new Excel.Application();
            excelApp.Visible = false;

            // Buka workbook
            Excel.Workbook workbook = excelApp.Workbooks.Open(excelFilePath);
            Excel.Worksheet worksheet = workbook.Sheets[1];// Worksheet pertama
            
            // Range untuk sel-sel yang ingin dihapus datanya (A12:T12)
            Excel.Range rangeToDelete1 = worksheet.Range["C8:BB8"];
            Excel.Range rangeToDelete11 = worksheet.Range["C7:BB7"];
            Excel.Range rangeToDelete2 = worksheet.Range["C10:BB10"];
            Excel.Range rangeToDelete3 = worksheet.Range["C12:BB12"];
            Excel.Range rangeToDelete4 = worksheet.Range["C13:BB13"];
            Excel.Range rangeToDelete5 = worksheet.Range["C14:BB14"];
            Excel.Range rangeToDelete6 = worksheet.Range["C15:BB15"];
            Excel.Range rangeToDelete7 = worksheet.Range["C19:BB19"];
            Excel.Range rangeToDelete8 = worksheet.Range["C20:BB20"];
            Excel.Range rangeToDelete9 = worksheet.Range["C21:BB21"];
            Excel.Range rangeToDelete10 = worksheet.Range["C22:BB22"];


            // Hapus konten dari sel-sel tertentu
            rangeToDelete1.ClearContents();
            rangeToDelete2.ClearContents();
            rangeToDelete3.ClearContents();
            rangeToDelete4.ClearContents();
            rangeToDelete5.ClearContents();
            rangeToDelete6.ClearContents();
            rangeToDelete7.ClearContents();
            rangeToDelete8.ClearContents();
            rangeToDelete9.ClearContents();
            rangeToDelete10.ClearContents();
            rangeToDelete11.ClearContents();

            // Simpan perubahan dan tutup workbook
            workbook.Save();
            workbook.Close();

            // Tutup aplikasi Excel
            excelApp.Quit();

            // Release objek Excel yang digunakan
            System.Runtime.InteropServices.Marshal.ReleaseComObject(rangeToDelete1);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(rangeToDelete2);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(rangeToDelete3);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(rangeToDelete4);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(rangeToDelete5);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(rangeToDelete6);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(rangeToDelete7);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(rangeToDelete8);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(rangeToDelete9);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(rangeToDelete10);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(rangeToDelete11);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {

        }

        private void CB_Machine_SelectedIndexChanged(object sender, EventArgs e)
        {
          
        }

        void ConfigData()
        {
            string deviceNumber = CB_DeviceType.Text.Trim();
            string filePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Config.txt"); // Ganti dengan path ke file Anda

            if (File.Exists(filePath))
            {
                var lines = File.ReadAllLines(filePath);
                var deviceData = GetDeviceData(lines, deviceNumber);
                if (deviceData != null)
                {
                    labelResult.Text = @"DV : " + GlobalVariable.Dv + "," + "  Range : " + GlobalVariable.Range;
                }
                else
                {
                    labelResult.Text = @"DV : 0, Range : 0 ";
                    MessageBox.Show("Data not found, Please Input data " + CB_DeviceType.Text + " on txt File");
                }
            }
            else
            {
                MessageBox.Show("File not found");
            }
        }

        private string GetDeviceData(string[] lines, string deviceNumber)
        {
            bool isDeviceFound = false;
            string deviceData = "";

            foreach (var line in lines)
            {
                if (line.Contains(@"Device = " + deviceNumber))
                {
                    isDeviceFound = true;
                }
                else if (isDeviceFound && (line.StartsWith("Device") || line.Trim() == ""))
                {
                    break;
                }
                else if (isDeviceFound)
                {
                    if (line.StartsWith("DV ="))
                    {
                        GlobalVariable.Dv = int.Parse(line.Split('=')[1].Trim());
                    }
                    else if (line.StartsWith("Range ="))
                    {
                        GlobalVariable.Range = int.Parse(line.Split('=')[1].Trim());
                    }
                    else if (line.StartsWith("TempHigh ="))
                    {
                        GlobalVariable.tempHigh = int.Parse(line.Split('=')[1].Trim());
                    }
                    else if (line.StartsWith("TempLow ="))
                    {
                        GlobalVariable.tempLow = int.Parse(line.Split('=')[1].Trim());
                    }
                }
            }

            return isDeviceFound ? deviceData : null;
        }

        private void CB_DeviceType_SelectedIndexChanged(object sender, EventArgs e)
        {
            ConfigData();
        }

        void DeviceType()
        {
            string filePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "DeviceType.txt"); // Path to your text file

            if (File.Exists(filePath))
            {
                try
                {
                    CB_DeviceType.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
                    CB_DeviceType.AutoCompleteSource = AutoCompleteSource.ListItems;
                    string[] lines = File.ReadAllLines(filePath);
                    CB_DeviceType.Items.AddRange(lines);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error reading file: " + ex.Message);
                }
            }
            else
            {
                MessageBox.Show("File not found: " + filePath);
            }
        }

        void Machine()
        {
            string filePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Machine.txt"); // Path to your text file
            if (File.Exists(filePath))
            {
                try
                {
                    CB_Machine.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
                    CB_Machine.AutoCompleteSource = AutoCompleteSource.ListItems;
                    string[] lines = File.ReadAllLines(filePath);
                    CB_Machine.Items.AddRange(lines);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error reading file: " + ex.Message);
                }
            }
            else
            {
                MessageBox.Show("File not found: " + filePath);
            }
        }
    }
}
