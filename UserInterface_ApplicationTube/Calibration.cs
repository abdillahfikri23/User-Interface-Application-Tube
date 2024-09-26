using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Diagnostics;
using System.Windows.Forms.DataVisualization.Charting;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace WindowsFormsApplication2
{
    public partial class Calibration : Form
    {
        public DataGridView DataGridViewReference { get; set; }
        private BackgroundWorker backgroundWorker;
        Stopwatch stopwatch = new Stopwatch();
        Stopwatch timer = new Stopwatch();
        double[] dataarrayH = new double[3];
        double[] dataarrayL = new double[3];
        int dut_Number = 0;
        int temp_Number = 0;
        enum dataType { Vtobj = 1, Vtobj_LPfiltered, Vtamb, Vtamb_LPfiltered, Tamb, PT100_Temperature, Supply_Current_IDD, ATPM, ATEMP };
        bool isStop = false;
        SqlConn data = new SqlConn();

        public Calibration()
        {
            InitializeComponent();

            // Inisialisasi Chart
            chart1.Titles.Add("VTObj");
            chart2.Titles.Add("VTAmb");

            // Inisialisasi SelectedIndexTemp
            cb_Temp.SelectedIndexChanged += comboBox2_SelectedIndexChanged;

            // Inisialisasi BackgroundWorker
            backgroundWorker = new BackgroundWorker();
            backgroundWorker.DoWork += backgroundWorker1_DoWork;
            backgroundWorker.ProgressChanged += backgroundWorker1_ProgressChanged;
            backgroundWorker.WorkerReportsProgress = true;
            backgroundWorker.RunWorkerCompleted += backgroundWorker1_RunWorkerCompleted;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            DeviceType();
            Machine();
            ApplicationSet up_set = new ApplicationSet();
            bool error = up_set.PortInit1("COM4");
            if (error != true)
            {
                //MessageBox.Show("Initialization is Success!!!", "Success");
            }
            else
            {
                MessageBox.Show("Initialization is Error!!!", "Error");
            }
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            DialogResult rs = new DialogResult();
            rs = MessageBox.Show("Do You Want to Close?", "Close", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (rs == DialogResult.Yes)
            {
                ApplicationSet._SerialPort1.Close();
                this.Close();
            }
            else
            {
                this.Show();
            }
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            temp_Number = cb_Temp.SelectedIndex;
            if (cb_Temp.SelectedItem.ToString() == "HIGH")
            {
                TB_TempMachine.Text = GlobalVariable.tempHigh.ToString();
                //HIGH GS
                textBoxVTob1H.BackColor = Color.White;
                textBoxVTob1H.Enabled = true;

                textBoxVTob2H.BackColor = Color.White;
                textBoxVTob2H.Enabled = true;

                textBoxVTob3H.BackColor = Color.White;
                textBoxVTob3H.Enabled = true;

                //BUTTON CALCULATE AVERAGE GS
                button2.Enabled = false;
                button2.BackColor = Color.Yellow;

                //GS READING HIGH
                textBoxGSReadingH.Enabled = true;
                textBoxGSReadingH.BackColor = Color.White;

                //GS RANGE HIGH
                textBoxRangeH.Enabled = true;
                textBoxRangeH.BackColor = Color.White;

                //Average
                textBoxAvgVToH.BackColor = Color.White;

                //DV 
                textBoxDVH.BackColor = Color.White;

                //Offset
                textBoxOffsetNEWH.BackColor = Color.White;
                textBoxOffsetOLDH.BackColor = Color.White;
            }
            else if (cb_Temp.SelectedItem.ToString() == "LOW")
            {
                TB_TempMachine.Text = GlobalVariable.tempLow.ToString();
                //LOW GS
                textBoxVTob1L.BackColor = Color.White;
                textBoxVTob1L.Enabled = true;

                textBoxVTob2L.BackColor = Color.White;
                textBoxVTob2L.Enabled = true;

                textBoxVTob3L.BackColor = Color.White;
                textBoxVTob3L.Enabled = true;

                //BUTTON CALCULATE AVERAGE GS
                button2.Enabled = false;
                button2.BackColor = Color.Yellow;

                //GS READING LOW
                textBoxGSReadingL.Enabled = true;
                textBoxGSReadingL.BackColor = Color.White;

                //GS RANGE LOW
                textBoxRangeL.Enabled = true;
                textBoxRangeL.BackColor = Color.White;

                //Average
                textBoxAvgVToL.BackColor = Color.White;

                //DV
                textBoxDVL.BackColor = Color.White;

                //Offset
                textBoxOffsetNEWL.BackColor = Color.White;
                textBoxOffsetOLDL.BackColor = Color.White;
            }
        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(textBoxGSReadingH.Text) && !string.IsNullOrEmpty(textBoxGSReadingL.Text))
            {
                CalculateHighAndLow();
            }
            else
            {
                CalculateHighOrLow();
            }
        }

        void CalculateHighAndLow()
        {
            double dataAverageH = dataarrayH.Average();
            double dataRangeH = dataarrayH.Max() - dataarrayH.Min();

            double dataAverageL = dataarrayL.Average();
            double dataRangeL = dataarrayL.Max() - dataarrayL.Min();

            if (string.IsNullOrEmpty(textBoxGSReadingH.Text))
            {
                MessageBox.Show("Please Input The GS_Reading High Value!!!");
            }
            else if (string.IsNullOrEmpty(textBoxOffsetOLDH.Text))
            {
                MessageBox.Show("Please Input The Offset Old High Value!!!");
            }
            else if (string.IsNullOrEmpty(textBoxGSReadingL.Text))
            {
                MessageBox.Show("Please Input The GS_Reading Low Value!!!");
            }
            else if (string.IsNullOrEmpty(textBoxOffsetOLDL.Text))
            {
                MessageBox.Show("Please Input The Offset Old Low Value!!!");
            }
            else if (string.IsNullOrEmpty(tb_Valueoff.Text))
            {
                MessageBox.Show("Please Input The Offset Value!!!");
            }
            else
            {
                button2.Enabled = false;
                button2.BackColor = Color.Yellow;
                button2.Text = "CALCULATE";

                //HIGH
                textBoxAvgVToH.Text = dataAverageH.ToString("0");
                textBoxRangeH.Text = dataRangeH.ToString("0");

                Double offset_OldH = Convert.ToDouble(textBoxOffsetOLDH.Text);
                GlobalVariable.average_gsH = Convert.ToDouble(textBoxAvgVToH.Text);
                GlobalVariable.gsreadingH = Convert.ToDouble(textBoxGSReadingH.Text);
                GlobalVariable.rangeH = Convert.ToDouble(textBoxRangeH.Text);
                GlobalVariable.ValueOff = Convert.ToDouble(tb_Valueoff.Text);

                Double DV_HIGH = GlobalVariable.gsreadingH - dataAverageH;
                textBoxDVH.Text = DV_HIGH.ToString("N1");
                Double DV_HighAbs = Math.Abs(DV_HIGH);

                if (DV_HighAbs <= GlobalVariable.Dv && GlobalVariable.rangeH <= GlobalVariable.Range)
                {
                    textBoxJudgeH.Text = "PASS";
                    textBoxJudgeH.BackColor = Color.Green;
                    textBoxOffsetNEWH.BackColor = Color.Green;
                }
                else
                {
                    textBoxJudgeH.Text = "FAIL";
                    textBoxJudgeH.BackColor = Color.Red;
                    textBoxOffsetNEWH.BackColor = Color.Red;
                }

                Double Offset_NewH = offset_OldH + (DV_HIGH * GlobalVariable.ValueOff);
                textBoxOffsetNEWH.Text = Offset_NewH.ToString("N1");

                //LOW
                textBoxAvgVToL.Text = dataAverageL.ToString("0");
                textBoxRangeL.Text = dataRangeL.ToString("0");

                Double offset_OldL = Convert.ToDouble(textBoxOffsetOLDL.Text);
                GlobalVariable.average_gsL = Convert.ToDouble(textBoxAvgVToL.Text);
                GlobalVariable.gsreadingL = Convert.ToDouble(textBoxGSReadingL.Text);
                GlobalVariable.rangeL = Convert.ToDouble(textBoxRangeL.Text);
                GlobalVariable.ValueOff = Convert.ToDouble(tb_Valueoff.Text);

                Double DV_LOW = GlobalVariable.gsreadingL - dataAverageL;
                textBoxDVL.Text = DV_LOW.ToString("N1");
                Double DV_LowAbs = Math.Abs(DV_LOW);

                if (DV_LowAbs <= GlobalVariable.Dv && GlobalVariable.rangeL <= GlobalVariable.Range)
                {
                    textBoxJudgeL.Text = "PASS";
                    textBoxJudgeL.BackColor = Color.Green;
                    textBoxOffsetNEWL.BackColor = Color.Green;
                }
                else
                {
                    textBoxJudgeL.Text = "FAIL";
                    textBoxJudgeL.BackColor = Color.Red;
                    textBoxOffsetNEWL.BackColor = Color.Red;
                }
                Double Offset_NewL = offset_OldL + (DV_LOW * GlobalVariable.ValueOff);
                textBoxOffsetNEWL.Text = Offset_NewL.ToString("N1");
            }
        }

        void CalculateHighOrLow()
        {
            string selectedTemp = cb_Temp.SelectedItem.ToString();
            if (selectedTemp == "HIGH")
            {
                double dataAverageH = dataarrayH.Average();
                double dataRangeH = dataarrayH.Max() - dataarrayH.Min();

                if (string.IsNullOrEmpty(textBoxGSReadingH.Text))
                {
                    MessageBox.Show("Please Input The GS_Reading High Value!!!");
                }
                else if (string.IsNullOrEmpty(textBoxOffsetOLDH.Text))
                {
                    MessageBox.Show("Please Input The Offset Old High Value!!!");
                }
                else if (string.IsNullOrEmpty(tb_Valueoff.Text))
                {
                    MessageBox.Show("Please Input The Offset Value!!!");
                }
                else
                {
                    button2.Enabled = false;
                    button2.BackColor = Color.Yellow;
                    button2.Text = "CALCULATE";

                    //HIGH
                    textBoxAvgVToH.Text = dataAverageH.ToString("0");
                    textBoxRangeH.Text = dataRangeH.ToString("0");

                    Double offset_OldH = Convert.ToDouble(textBoxOffsetOLDH.Text);
                    GlobalVariable.average_gsH = Convert.ToDouble(textBoxAvgVToH.Text);
                    GlobalVariable.gsreadingH = Convert.ToDouble(textBoxGSReadingH.Text);
                    GlobalVariable.rangeH = Convert.ToDouble(textBoxRangeH.Text);
                    GlobalVariable.ValueOff = Convert.ToDouble(tb_Valueoff.Text);

                    Double DV_HIGH = GlobalVariable.gsreadingH - dataAverageH;
                    textBoxDVH.Text = DV_HIGH.ToString("N1");
                    Double DV_HighAbs = Math.Abs(DV_HIGH);

                    if (DV_HighAbs <= GlobalVariable.Dv && GlobalVariable.rangeH <= GlobalVariable.Range)
                    {
                        textBoxJudgeH.Text = "PASS";
                        textBoxJudgeH.BackColor = Color.Green;
                        textBoxOffsetNEWH.BackColor = Color.Green;
                    }
                    else
                    {
                        textBoxJudgeH.Text = "FAIL";
                        textBoxJudgeH.BackColor = Color.Red;
                        textBoxOffsetNEWH.BackColor = Color.Red;
                    }

                    Double Offset_NewH = offset_OldH + (DV_HIGH * GlobalVariable.ValueOff);
                    textBoxOffsetNEWH.Text = Offset_NewH.ToString("N1");
                }
            }
            else if (selectedTemp == "LOW")
            {
                double dataAverageL = dataarrayL.Average();
                double dataRangeL = dataarrayL.Max() - dataarrayL.Min();

                if (string.IsNullOrEmpty(textBoxGSReadingL.Text))
                {
                    MessageBox.Show("Please Input The GS_Reading Low Value!!!");
                }
                else if (string.IsNullOrEmpty(textBoxOffsetOLDL.Text))
                {
                    MessageBox.Show("Please Input The Offset Old Low Value!!!");
                }
                else if (string.IsNullOrEmpty(tb_Valueoff.Text))
                {
                    MessageBox.Show("Please Input The Offset Value!!!");
                }
                else
                {
                    button2.Enabled = false;
                    button2.BackColor = Color.Yellow;
                    button2.Text = "CALCULATE";

                    //LOW
                    textBoxAvgVToL.Text = dataAverageL.ToString("0");
                    textBoxRangeL.Text = dataRangeL.ToString("0");

                    Double offset_OldL = Convert.ToDouble(textBoxOffsetOLDL.Text);
                    GlobalVariable.average_gsL = Convert.ToDouble(textBoxAvgVToL.Text);
                    GlobalVariable.gsreadingL = Convert.ToDouble(textBoxGSReadingL.Text);
                    GlobalVariable.rangeL = Convert.ToDouble(textBoxRangeL.Text);
                    GlobalVariable.ValueOff = Convert.ToDouble(tb_Valueoff.Text);

                    Double DV_LOW = GlobalVariable.gsreadingL - dataAverageL;
                    textBoxDVL.Text = DV_LOW.ToString("N1");
                    Double DV_LowAbs = Math.Abs(DV_LOW);

                    if (DV_LowAbs <= GlobalVariable.Dv && GlobalVariable.rangeL <= GlobalVariable.Range)
                    {
                        textBoxJudgeL.Text = "PASS";
                        textBoxJudgeL.BackColor = Color.Green;
                        textBoxOffsetNEWL.BackColor = Color.Green;
                    }
                    else
                    {
                        textBoxJudgeL.Text = "FAIL";
                        textBoxJudgeL.BackColor = Color.Red;
                        textBoxOffsetNEWL.BackColor = Color.Red;
                    }
                    Double Offset_NewL = offset_OldL + (DV_LOW * GlobalVariable.ValueOff);
                    textBoxOffsetNEWL.Text = Offset_NewL.ToString("N1");
                }
            }
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            stopwatch.Start();
            timer.Start();
            int time = 0;
            switch (dut_Number)
            {
                case 0:
                    time = 4;
                    break;
                case 1:
                    time = 4;
                    break;
                case 2:
                    time = 4;
                    break;
                case 3:
                    time = 9;
                    break;
                default:
                    break;
            }

            while (stopwatch.Elapsed.TotalMinutes < time)
            {
                //Initialization Aplication Set
                ApplicationSet up_set = new ApplicationSet();
                int read_dataVtobj = up_set.read_data_Median3((int)dataType.Vtobj_LPfiltered);
                int read_dataVtamb = up_set.read_data_Median3((int)dataType.Vtamb_LPfiltered);

                GlobalVariable.newValue = read_dataVtobj;
                GlobalVariable.Vtamb_Value = read_dataVtamb;

                // Melaporkan kemajuan ke MainForm
                backgroundWorker.ReportProgress((int)GlobalVariable.newValue);

                //Stop Running
                if (isStop)
                {
                    isStop = false;
                    break;
                }

                if (time == 9)
                {

                }
                else
                {
                    switch (temp_Number)
                    {
                        case 0:
                            double vtObjHAbs = Math.Abs(Convert.ToDouble(textBoxGSReadingH.Text) - GlobalVariable.newValue);
                            if (GlobalVariable.rangeValue <= 5 && vtObjHAbs <= GlobalVariable.Dv)
                            {
                                timer.Start();
                                if (timer.Elapsed.TotalMinutes >= 2)
                                {
                                    MessageBox.Show("The value of VTobj has stabilized!!!", "DUT");
                                    timer.Reset();
                                    e.Cancel = true;
                                    return;
                                }
                            }
                            else
                            {
                                timer.Reset();
                            }
                            break;
                        case 1:
                            double vtObjLAbs = Math.Abs(Convert.ToDouble(textBoxGSReadingL.Text) - GlobalVariable.newValue);
                            if (GlobalVariable.rangeValue <= 5 && vtObjLAbs <= GlobalVariable.Dv)
                            {
                                timer.Start();
                                if (timer.Elapsed.TotalMinutes >= 2)
                                {
                                    MessageBox.Show("The value of VTobj has stabilized!!!", "DUT");
                                    timer.Reset();
                                    e.Cancel = true;
                                    return;
                                }
                            }
                            else
                            {
                                timer.Reset();
                            }
                            break;
                        default:
                            break;
                    }
                }
            }
        }

        private void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            DateTime startTime;
            startTime = DateTime.Now;

            //Update Process Time
            UpdateTextBoxWaktu(String.Format("{0:mm\\:ss}", stopwatch.Elapsed));
            textBoxRangeTime.Text = String.Format("{0:mm\\:ss}", timer.Elapsed);

            // Menampilkan Nilai VTobj di Chart
            chart1.Series["VTobj_Value"].Points.AddY(GlobalVariable.newValue);
            chart1.ChartAreas[0].AxisY.Minimum = 1100;
            chart1.ChartAreas[0].AxisY.Maximum = 4000;

            // Menampilkan Nilai VTamb di Chart
            chart2.Series["VTamb_Value"].Points.AddY(GlobalVariable.Vtamb_Value);
            chart2.ChartAreas[0].AxisY.Minimum = 1100;
            chart2.ChartAreas[0].AxisY.Maximum = 4000;

            //Disabled combobox temp dan dut
            cb_Temp.Enabled = false;
            cb_DUT.Enabled = false;

            //Mendapatkan nilai range real time
            Series series = chart1.Series["VTobj_Value"];

            // Ambil jumlah titik data dalam series
            int count = series.Points.Count;

            // Tentukan indeks awal untuk memulai iterasi
            int startIndex1 = Math.Max(count - 10, 0); //Average Vtobj value
            int startIndex2 = Math.Max(count - 100, 0); //Range stabil value

            double max = double.MinValue;
            double min = double.MaxValue;

            double total = 0;

            // Lakukan iterasi melalui titik data yang dipilih
            for (int i = startIndex1; i < count; i++)
            {
                DataPoint point = series.Points[i];
                double yValue = point.YValues[0];
                total += yValue; // Tambahkan nilai ke total
            }

            for (int k = startIndex2; k < count; k++)
            {
                DataPoint point = series.Points[k];
                double yValue = point.YValues[0];
                if (yValue > max)
                {
                    max = yValue;
                }
                if (yValue < min)
                {
                    min = yValue;
                }
            }

            Double average = total / Math.Min(10, count - startIndex1);
            GlobalVariable.movingAverage = Convert.ToInt32(average);
            GlobalVariable.rangeValue = max - min;
            textBoxRangeRT.Text = GlobalVariable.rangeValue.ToString();

            if (cb_Temp.SelectedItem.ToString() == "HIGH")
            {
                switch (dut_Number)
                {
                    case 0:
                        textBoxVTob1H.Text = GlobalVariable.movingAverage.ToString();
                        dataarrayH[0] = GlobalVariable.movingAverage;
                        string DUT1_Value = "VTobj = " + GlobalVariable.newValue.ToString() + ", VTamb = " + GlobalVariable.Vtamb_Value.ToString();
                        textBoxValue.AppendText(DUT1_Value + Environment.NewLine);
                        textBoxVTob1H.BackColor = Color.Green;
                        break;
                    case 1:
                        textBoxVTob2H.Text = GlobalVariable.movingAverage.ToString();
                        dataarrayH[1] = GlobalVariable.movingAverage;
                        string DUT2_Value = "VTobj = " + GlobalVariable.newValue.ToString() + ", VTamb = " + GlobalVariable.Vtamb_Value.ToString();
                        textBoxValue.AppendText(DUT2_Value + Environment.NewLine);
                        textBoxVTob2H.BackColor = Color.Green;
                        break;
                    case 2:
                        textBoxVTob3H.Text = GlobalVariable.movingAverage.ToString();
                        dataarrayH[2] = GlobalVariable.movingAverage;
                        string DUT3_Value = "VTobj = " + GlobalVariable.newValue.ToString() + ", VTamb = " + GlobalVariable.Vtamb_Value.ToString();
                        textBoxValue.AppendText(DUT3_Value + Environment.NewLine);
                        textBoxVTob3H.BackColor = Color.Green;
                        break;
                    case 3:
                        //textBoxGSReadingH.Text = GlobalVariable.newValue.ToString();
                        textBoxGSReadingH.Text = GlobalVariable.movingAverage.ToString();
                        string GS_Value = "VTobj = " + GlobalVariable.newValue.ToString() + ", VTamb = " + GlobalVariable.Vtamb_Value.ToString();
                        textBoxValue.AppendText(GS_Value + Environment.NewLine);
                        textBoxGSReadingH.BackColor = Color.Green;
                        break;
                    default:
                        break;
                }
            }
            else if (cb_Temp.SelectedItem.ToString() == "LOW")
            {
                switch (dut_Number)
                {
                    case 0:
                        textBoxVTob1L.Text = GlobalVariable.movingAverage.ToString();
                        dataarrayL[0] = GlobalVariable.movingAverage;
                        string DUT1_Value = "VTobj = " + GlobalVariable.newValue.ToString() + ", VTamb = " + GlobalVariable.Vtamb_Value.ToString();
                        textBoxValue.AppendText(DUT1_Value + Environment.NewLine);
                        textBoxVTob1L.BackColor = Color.Green;
                        break;
                    case 1:
                        textBoxVTob2L.Text = GlobalVariable.movingAverage.ToString();
                        dataarrayL[1] = GlobalVariable.movingAverage;
                        string DUT2_Value = "VTobj = " + GlobalVariable.newValue.ToString() + ", VTamb = " + GlobalVariable.Vtamb_Value.ToString();
                        textBoxValue.AppendText(DUT2_Value + Environment.NewLine);
                        textBoxVTob2L.BackColor = Color.Green;
                        break;
                    case 2:
                        textBoxVTob3L.Text = GlobalVariable.movingAverage.ToString();
                        dataarrayL[2] = GlobalVariable.movingAverage;
                        string DUT3_Value = "VTobj = " + GlobalVariable.newValue.ToString() + ", VTamb = " + GlobalVariable.Vtamb_Value.ToString();
                        textBoxValue.AppendText(DUT3_Value + Environment.NewLine);
                        textBoxVTob3L.BackColor = Color.Green;
                        break;
                    case 3:
                        //textBoxGSReadingL.Text = GlobalVariable.newValue.ToString();
                        textBoxGSReadingL.Text = GlobalVariable.movingAverage.ToString();
                        string GS_Value = "VTobj = " + GlobalVariable.newValue.ToString() + ", VTamb = " + GlobalVariable.Vtamb_Value.ToString();
                        textBoxValue.AppendText(GS_Value + Environment.NewLine);
                        textBoxGSReadingL.BackColor = Color.Green;
                        break;
                    default:
                        break;
                }
            }
        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            cb_Temp.Enabled = true;
            cb_DUT.Enabled = true;
            string selectedTemp = cb_Temp.SelectedItem.ToString();
            if (selectedTemp == "HIGH")
            {
                textBoxVTob1H.BackColor = Color.White;
                textBoxVTob2H.BackColor = Color.White;
                textBoxVTob3H.BackColor = Color.White;
                textBoxGSReadingH.BackColor = Color.White;
                /*
                if (!string.IsNullOrEmpty(textBoxVTob1H.Text) && !string.IsNullOrEmpty(textBoxVTob2H.Text) && !string.IsNullOrEmpty(textBoxVTob3H.Text) && !string.IsNullOrEmpty(textBoxVTob1L.Text) && !string.IsNullOrEmpty(textBoxVTob2L.Text) && !string.IsNullOrEmpty(textBoxVTob3L.Text))
                {
                    button2.Enabled = true;
                    button2.BackColor = Color.ForestGreen;
                    button2.Text = "CALCULATE IS READY";
                }*/
                if (!string.IsNullOrEmpty(textBoxVTob1H.Text) && !string.IsNullOrEmpty(textBoxVTob2H.Text) && !string.IsNullOrEmpty(textBoxVTob3H.Text))
                {
                    button2.Enabled = true;
                    button2.BackColor = Color.ForestGreen;
                    button2.Text = "CALCULATE IS READY";
                }

            }
            else if (selectedTemp == "LOW")
            {
                textBoxVTob1L.BackColor = Color.White;
                textBoxVTob2L.BackColor = Color.White;
                textBoxVTob3L.BackColor = Color.White;
                textBoxGSReadingL.BackColor = Color.White;
                /*
                if (!string.IsNullOrEmpty(textBoxVTob1H.Text) && !string.IsNullOrEmpty(textBoxVTob2H.Text) && !string.IsNullOrEmpty(textBoxVTob3H.Text) && !string.IsNullOrEmpty(textBoxVTob1L.Text) && !string.IsNullOrEmpty(textBoxVTob2L.Text) && !string.IsNullOrEmpty(textBoxVTob3L.Text))
                {
                    button2.Enabled = true;
                    button2.BackColor = Color.ForestGreen;
                    button2.Text = "CALCULATE IS READY";
                }*/
                if (!string.IsNullOrEmpty(textBoxVTob1L.Text) && !string.IsNullOrEmpty(textBoxVTob2L.Text) && !string.IsNullOrEmpty(textBoxVTob3L.Text))
                {
                    button2.Enabled = true;
                    button2.BackColor = Color.ForestGreen;
                    button2.Text = "CALCULATE IS READY";
                }
            }

            buttonEnter.BackColor = Color.Yellow;
            buttonEnter.Text = "START";
            isRunning = false;

            buttonStop.BackColor = Color.Yellow;
            buttonStop.Enabled = false;
            timer.Reset();

            //textBoxValue.Clear();

            //Chart1 and chart2
            chart1.Series["VTobj_Value"].Points.Clear();
            chart2.Series["VTamb_Value"].Points.Clear();
        }

        private void UpdateTextBoxWaktu(string text)
        {
            // Jika perlu Invoke (memanggil dari thread yang berbeda), maka kita menggunakan Invoke
            if (textBoxIterasi.InvokeRequired)
            {
                textBoxIterasi.Invoke(new Action(() => textBoxIterasi.Text = text));
            }
            else
            {
                // Jika tidak perlu Invoke (sudah di thread UI yang benar), kita langsung atur nilai Text
                textBoxIterasi.Text = text;
            }
        }

        private void buttonClear_Click(object sender, EventArgs e)
        {
            DialogResult rs = new DialogResult();
            rs = MessageBox.Show("Do You Want to Clear?", "Clear", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (rs == DialogResult.Yes)
            {
                cb_Machine.Text = "";
                cb_DUT.Text = "";
                cb_Temp.Text = "";
                cb_DeviceType.Text = "";
                textBoxJig.Clear();
                textBoxSerial.Clear();
                textBoxGSReadingH.Clear();
                textBoxGSReadingL.Clear();
                textBoxRangeH.Clear();
                textBoxRangeL.Clear();
                textBoxVTob1H.Clear();
                textBoxVTob1L.Clear();
                textBoxVTob2H.Clear();
                textBoxVTob2L.Clear();
                textBoxVTob3H.Clear();
                textBoxVTob3L.Clear();
                textBoxAvgVToH.Clear();
                textBoxAvgVToL.Clear();
                textBoxOffsetNEWH.Clear();
                textBoxOffsetNEWL.Clear();
                textBoxOffsetOLDH.Clear();
                textBoxOffsetOLDL.Clear();
                textBoxJudgeH.Clear();
                textBoxJudgeL.Clear();
                textBoxIterasi.Clear();

                //BUTTON CALCULATE AVERAGE GS
                button2.Enabled = false;
                button2.BackColor = Color.Yellow;

                //JUDGE HIGH dan LOW
                textBoxJudgeH.BackColor = Color.Silver;
                textBoxJudgeL.BackColor = Color.Silver;

                //HIGH
                textBoxVTob1H.BackColor = Color.Silver;
                textBoxVTob1H.Enabled = false;

                textBoxVTob2H.BackColor = Color.Silver;
                textBoxVTob1H.Enabled = false;

                textBoxVTob3H.BackColor = Color.Silver;
                textBoxVTob3H.Enabled = false;

                //LOW
                textBoxVTob1L.BackColor = Color.Silver;
                textBoxVTob1L.Enabled = false;

                textBoxVTob2L.BackColor = Color.Silver;
                textBoxVTob2L.Enabled = false;

                textBoxVTob3L.BackColor = Color.Silver;
                textBoxVTob3L.Enabled = false;

                //GS READING HIGH
                textBoxGSReadingH.Enabled = false;
                textBoxGSReadingH.BackColor = Color.Silver;

                //GS RANGE HIGH
                textBoxRangeH.Enabled = false;
                textBoxRangeH.BackColor = Color.Silver;

                //GS READING LOW
                textBoxGSReadingL.Enabled = false;
                textBoxGSReadingL.BackColor = Color.Silver;

                //GS RANGE LOW
                textBoxGSReadingL.Enabled = false;
                textBoxRangeL.BackColor = Color.Silver;

                //OffsetOLD HIGH
                textBoxOffsetOLDH.BackColor = Color.White;

                //OffsetNEW HIGH
                textBoxOffsetNEWH.BackColor = Color.White;

                //OffsetOLD LOW
                textBoxOffsetOLDL.BackColor = Color.White;

                //OffsetNEW LOW
                textBoxOffsetNEWL.BackColor = Color.White;

                //DVH and DVL
                textBoxDVH.Clear();
                textBoxDVL.Clear();

                //Chart1 and chart2
                chart1.Series["VTobj_Value"].Points.Clear();
                chart2.Series["VTamb_Value"].Points.Clear();

                textBoxRangeTime.Clear();
                textBoxRangeRT.Clear();

                //Value
                textBoxValue.Clear();

                //Shift
                cb_Shift.Text = "";

                //DoneBy
                textBoxDoneBy.Clear();

                //Start
                textBoxStart.Clear();

            }
            else
            {
                this.Show();
            }
        }

        bool isRunning = false;
        bool timeStart = false;
        private void buttonEnter_Click(object sender, EventArgs e)
        {
            DateTime startTime;
            startTime = DateTime.Now;
            if (timeStart == true)
            {

            }
            else
            {
                timeStart = true;
                textBoxStart.Text = startTime.ToString("HH:mm:ss");
            }

            if (cb_DUT.Text != "" && cb_Temp.Text != "" && cb_DeviceType.Text != "" && cb_Machine.Text != "" && cb_Shift.Text != "" && textBoxJig.Text != "" && textBoxDoneBy.Text != "" && textBoxSerial.Text != "" && isRunning == false)
            {
                if (cb_Temp.SelectedItem.ToString() == "HIGH")
                {
                    if (textBoxGSReadingH.Text != "" || cb_DUT.SelectedItem.ToString() == "GS")
                    {
                        switch (dut_Number)
                        {
                            case 0:
                                if (!string.IsNullOrEmpty(textBoxVTob1H.Text))
                                {
                                    DialogResult rsH0 = new DialogResult();
                                    rsH0 = MessageBox.Show("Do You Want to Countinue?", "DUT1_HIGH", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                                    if (rsH0 == DialogResult.Yes)
                                    {
                                        isRunning = !isRunning;
                                        if (isRunning)
                                        {
                                            buttonEnter.Text = "RUNNING";
                                            buttonEnter.BackColor = Color.Green;
                                            buttonStop.BackColor = Color.Red;
                                            buttonStop.Enabled = true;
                                            isRunning = true;
                                            stopwatch.Reset();
                                            backgroundWorker.RunWorkerAsync();
                                        }
                                    }
                                    else
                                    {
                                        this.Show();
                                    }
                                }
                                else
                                {
                                    isRunning = !isRunning;
                                    if (isRunning)
                                    {
                                        buttonEnter.Text = "RUNNING";
                                        buttonEnter.BackColor = Color.Green;
                                        buttonStop.BackColor = Color.Red;
                                        buttonStop.Enabled = true;
                                        //textBoxStart.Text = startTime.ToString("HH:mm:ss");
                                        isRunning = true;
                                        stopwatch.Reset();
                                        backgroundWorker.RunWorkerAsync();
                                    }
                                }
                                break;
                            case 1:
                                if (!string.IsNullOrEmpty(textBoxVTob2H.Text))
                                {
                                    DialogResult rsH1 = new DialogResult();
                                    rsH1 = MessageBox.Show("Do You Want to Countinue?", "DUT2_HIGH", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                                    if (rsH1 == DialogResult.Yes)
                                    {
                                        isRunning = !isRunning;
                                        if (isRunning)
                                        {
                                            buttonEnter.Text = "RUNNING";
                                            buttonEnter.BackColor = Color.Green;
                                            buttonStop.BackColor = Color.Red;
                                            buttonStop.Enabled = true;
                                            isRunning = true;
                                            stopwatch.Reset();
                                            backgroundWorker.RunWorkerAsync();
                                        }
                                    }
                                    else
                                    {
                                        this.Show();
                                    }
                                }
                                else
                                {
                                    isRunning = !isRunning;
                                    if (isRunning)
                                    {
                                        buttonEnter.Text = "RUNNING";
                                        buttonEnter.BackColor = Color.Green;
                                        buttonStop.BackColor = Color.Red;
                                        buttonStop.Enabled = true;
                                        isRunning = true;
                                        stopwatch.Reset();
                                        backgroundWorker.RunWorkerAsync();
                                    }
                                }
                                break;
                            case 2:
                                if (!string.IsNullOrEmpty(textBoxVTob3H.Text))
                                {
                                    DialogResult rsH2 = new DialogResult();
                                    rsH2 = MessageBox.Show("Do You Want to Countinue?", "DUT3_HIGH", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                                    if (rsH2 == DialogResult.Yes)
                                    {
                                        isRunning = !isRunning;
                                        if (isRunning)
                                        {
                                            buttonEnter.Text = "RUNNING";
                                            buttonEnter.BackColor = Color.Green;
                                            buttonStop.BackColor = Color.Red;
                                            buttonStop.Enabled = true;
                                            isRunning = true;
                                            stopwatch.Reset();
                                            backgroundWorker.RunWorkerAsync();
                                        }
                                    }
                                    else
                                    {
                                        this.Show();
                                    }
                                }
                                else
                                {
                                    isRunning = !isRunning;
                                    if (isRunning)
                                    {
                                        buttonEnter.Text = "RUNNING";
                                        buttonEnter.BackColor = Color.Green;
                                        buttonStop.BackColor = Color.Red;
                                        buttonStop.Enabled = true;
                                        isRunning = true;
                                        stopwatch.Reset();
                                        backgroundWorker.RunWorkerAsync();
                                    }
                                }
                                break;
                            case 3:
                                isRunning = !isRunning;
                                if (isRunning)
                                {
                                    buttonEnter.Text = "RUNNING";
                                    buttonEnter.BackColor = Color.Green;
                                    buttonStop.BackColor = Color.Red;
                                    buttonStop.Enabled = true;
                                    isRunning = true;
                                    stopwatch.Reset();
                                    backgroundWorker.RunWorkerAsync();
                                }
                                break;
                            default:
                                break;
                        }
                    }
                    else
                    {
                        MessageBox.Show("Please Input The GSReading High Value");
                    }
                }
                else if (cb_Temp.SelectedItem.ToString() == "LOW")
                {
                    if (textBoxGSReadingL.Text != "" || cb_DUT.SelectedItem.ToString() == "GS")
                    {
                        switch (dut_Number)
                        {
                            case 0:
                                if (!string.IsNullOrEmpty(textBoxVTob1L.Text))
                                {
                                    DialogResult rsL0 = new DialogResult();
                                    rsL0 = MessageBox.Show("Do You Want to Countinue?", "DUT1_LOW", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                                    if (rsL0 == DialogResult.Yes)
                                    {
                                        isRunning = !isRunning;
                                        if (isRunning)
                                        {
                                            buttonEnter.Text = "RUNNING";
                                            buttonEnter.BackColor = Color.Green;
                                            buttonStop.BackColor = Color.Red;
                                            buttonStop.Enabled = true;
                                            isRunning = true;
                                            stopwatch.Reset();
                                            backgroundWorker.RunWorkerAsync();
                                        }
                                    }
                                    else
                                    {
                                        this.Show();
                                    }
                                }
                                else
                                {
                                    isRunning = !isRunning;
                                    if (isRunning)
                                    {
                                        buttonEnter.Text = "RUNNING";
                                        buttonEnter.BackColor = Color.Green;
                                        buttonStop.BackColor = Color.Red;
                                        buttonStop.Enabled = true;
                                        isRunning = true;
                                        stopwatch.Reset();
                                        backgroundWorker.RunWorkerAsync();
                                    }
                                }
                                break;
                            case 1:
                                if (!string.IsNullOrEmpty(textBoxVTob2L.Text))
                                {
                                    DialogResult rsL1 = new DialogResult();
                                    rsL1 = MessageBox.Show("Do You Want to Countinue?", "DUT2_LOW", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                                    if (rsL1 == DialogResult.Yes)
                                    {
                                        isRunning = !isRunning;
                                        if (isRunning)
                                        {
                                            buttonEnter.Text = "RUNNING";
                                            buttonEnter.BackColor = Color.Green;
                                            buttonStop.BackColor = Color.Red;
                                            buttonStop.Enabled = true;
                                            isRunning = true;
                                            stopwatch.Reset();
                                            backgroundWorker.RunWorkerAsync();
                                        }
                                    }
                                    else
                                    {
                                        this.Show();
                                    }
                                }
                                else
                                {
                                    isRunning = !isRunning;
                                    if (isRunning)
                                    {
                                        buttonEnter.Text = "RUNNING";
                                        buttonEnter.BackColor = Color.Green;
                                        buttonStop.BackColor = Color.Red;
                                        buttonStop.Enabled = true;
                                        isRunning = true;
                                        stopwatch.Reset();
                                        backgroundWorker.RunWorkerAsync();
                                    }
                                }
                                break;
                            case 2:
                                if (!string.IsNullOrEmpty(textBoxVTob3L.Text))
                                {
                                    DialogResult rsL2 = new DialogResult();
                                    rsL2 = MessageBox.Show("Do You Want to Countinue?", "DUT3_LOW", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                                    if (rsL2 == DialogResult.Yes)
                                    {
                                        isRunning = !isRunning;
                                        if (isRunning)
                                        {
                                            buttonEnter.Text = "RUNNING";
                                            buttonEnter.BackColor = Color.Green;
                                            buttonStop.BackColor = Color.Red;
                                            buttonStop.Enabled = true;
                                            isRunning = true;
                                            stopwatch.Reset();
                                            backgroundWorker.RunWorkerAsync();
                                        }
                                    }
                                    else
                                    {
                                        this.Show();
                                    }
                                }
                                else
                                {
                                    isRunning = !isRunning;
                                    if (isRunning)
                                    {
                                        buttonEnter.Text = "RUNNING";
                                        buttonEnter.BackColor = Color.Green;
                                        buttonStop.BackColor = Color.Red;
                                        buttonStop.Enabled = true;
                                        isRunning = true;
                                        stopwatch.Reset();
                                        backgroundWorker.RunWorkerAsync();
                                    }
                                }
                                break;
                            case 3:
                                isRunning = !isRunning;
                                if (isRunning)
                                {
                                    buttonEnter.Text = "RUNNING";
                                    buttonEnter.BackColor = Color.Green;
                                    buttonStop.BackColor = Color.Red;
                                    buttonStop.Enabled = true;
                                    isRunning = true;
                                    stopwatch.Reset();
                                    backgroundWorker.RunWorkerAsync();
                                }
                                break;
                            default:
                                break;
                        }
                    }
                    else
                    {
                        MessageBox.Show("Please Input The GSReading Low Value");
                    }
                }
            }
            else if (string.IsNullOrEmpty(cb_Temp.Text))
            {
                MessageBox.Show("Please Input The Item Of Temp!!!");
            }
            else if (string.IsNullOrEmpty(cb_DUT.Text))
            {
                MessageBox.Show("Please Input The Item Of DUT!!!");
            }
            else if (string.IsNullOrEmpty(cb_Machine.Text))
            {
                MessageBox.Show("Please Input The Item Of Machine!!!");
            }
            else if (string.IsNullOrEmpty(cb_DeviceType.Text))
            {
                MessageBox.Show("Please Input The Device Type!!!");
            }
            else if (string.IsNullOrEmpty(cb_Shift.Text))
            {
                MessageBox.Show("Please Input The Item Of Shift!!!");
            }
            else if (string.IsNullOrEmpty(textBoxJig.Text))
            {
                MessageBox.Show("Please Input The Jig!!!");
            }
            else if (string.IsNullOrEmpty(textBoxDoneBy.Text))
            {
                MessageBox.Show("Please Input The Done By!!!");
            }
            else if (string.IsNullOrEmpty(textBoxSerial.Text))
            {
                MessageBox.Show("Please Input The Serial Number!!!");
            }
        }

        private void cb_DUT_SelectedIndexChanged(object sender, EventArgs e)
        {
            dut_Number = cb_DUT.SelectedIndex;
        }

        private void buttonStop_Click(object sender, EventArgs e)
        {
            DialogResult stop = new DialogResult();
            stop = MessageBox.Show("Do You Want to Stop Running?", "Stop", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (stop == DialogResult.Yes)
            {
                isStop = true;
            }
            else
            {
                this.Show();
            }
        }

        void ConfigData()
        {
            string deviceNumber = cb_DeviceType.Text.Trim();
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
                    TB_TempMachine.Text = "0";
                    labelResult.Text = @"DV : 0, Range : 0 ";
                    MessageBox.Show("Data not found, Please Input data " + cb_DeviceType.Text + " on txt File");
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

        void DeviceType()
        {
            string filePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "DeviceType.txt"); // Path to your text file

            if (File.Exists(filePath))
            {
                try
                {
                    cb_DeviceType.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
                    cb_DeviceType.AutoCompleteSource = AutoCompleteSource.ListItems;
                    string[] lines = File.ReadAllLines(filePath);
                    cb_DeviceType.Items.AddRange(lines);
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
                    cb_Machine.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
                    cb_Machine.AutoCompleteSource = AutoCompleteSource.ListItems;
                    string[] lines = File.ReadAllLines(filePath);
                    cb_Machine.Items.AddRange(lines);
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

        private void buttonSave_Click(object sender, EventArgs e)
        {
            /*
            DateTime finishTime;
            finishTime = DateTime.Now;
            textBoxFinish.Text = finishTime.ToString("HH:mm:ss");

            // SAVE TO SQL SERVER
            data.year = textBoxDate.Value.Year;
            //data.date = textBoxDate.Text;
            data.date = textBoxDate.Value.ToString("yyyy-MM-dd");
            data.machineNumber = cb_Machine.Text;
            data.deviceType = cb_DeviceType.Text;
            data.serialNumber = textBoxSerial.Text;
            data.shift = cb_Shift.Text;
            data.start = textBoxStart.Text;
            data.finish = textBoxFinish.Text;
            data.testJig = textBoxJig.Text;
            data.doneBy = textBoxDoneBy.Text;

            //HIGH
            data.tempSettingH = labelHigh.Text;
            data.gs_ReadingH = Convert.ToDouble(textBoxGSReadingH.Text);
            data.dut1_ReadingH = Convert.ToDouble(textBoxVTob1H.Text);
            data.dut2_ReadingH = Convert.ToDouble(textBoxVTob2H.Text);
            data.dut3_ReadingH = Convert.ToDouble(textBoxVTob3H.Text);
            data.avg_ReadingH = Convert.ToDouble(textBoxAvgVToH.Text);
            data.dv_ReadingH = Convert.ToDouble(textBoxDVH.Text);
            data.range_ReadingH = Convert.ToDouble(textBoxRangeH.Text);
            data.mc_OffsetOldH = Convert.ToDouble(textBoxOffsetOLDH.Text);
            data.mc_OffsetNewH = Convert.ToDouble(textBoxOffsetNEWH.Text);
            data.buy0ffResultH = textBoxJudgeH.Text;
            //LOW
            data.tempSettingL = labelLow.Text;
            data.gs_ReadingL = Convert.ToDouble(textBoxGSReadingL.Text);
            data.dut1_ReadingL = Convert.ToDouble(textBoxVTob1L.Text);
            data.dut2_ReadingL = Convert.ToDouble(textBoxVTob2L.Text);
            data.dut3_ReadingL = Convert.ToDouble(textBoxVTob3L.Text);
            data.avg_ReadingL = Convert.ToDouble(textBoxAvgVToL.Text);
            data.dv_ReadingL = Convert.ToDouble(textBoxDVL.Text);
            data.range_ReadingL = Convert.ToDouble(textBoxRangeL.Text);
            data.mc_OffsetOldL = Convert.ToDouble(textBoxOffsetOLDL.Text);
            data.mc_OffsetNewL = Convert.ToDouble(textBoxOffsetNEWL.Text);
            data.buy0ffResultL = textBoxJudgeL.Text;

            //data.Insertdata();
            bool success = data.Insert(data);
            if (success == true)
            {
                MessageBox.Show("Data has been saved!!!");
                ApplicationSet._SerialPort1.Close();
                this.Close();
            }
            else
            {
                MessageBox.Show("Failed to Saved Data!!!", "Try Again");
            }*/
            if (!string.IsNullOrEmpty(textBoxGSReadingH.Text) && !string.IsNullOrEmpty(textBoxGSReadingL.Text))
            {
                SaveHighAndLow();
            }
            else
            {
                SaveHighOrLow();
            }
        }

        void SaveHighAndLow()
        {
            DateTime finishTime;
            finishTime = DateTime.Now;
            textBoxFinish.Text = finishTime.ToString("HH:mm:ss");

            // SAVE TO SQL SERVER
            data.year = textBoxDate.Value.Year;
            data.date = textBoxDate.Value.ToString("yyyy-MM-dd");
            data.machineNumber = cb_Machine.Text;
            data.deviceType = cb_DeviceType.Text;
            data.serialNumber = textBoxSerial.Text;
            data.shift = cb_Shift.Text;
            data.start = textBoxStart.Text;
            data.finish = textBoxFinish.Text;
            data.testJig = textBoxJig.Text;
            data.doneBy = textBoxDoneBy.Text;

            //HIGH
            data.tempSettingH = labelHigh.Text;
            data.gs_ReadingH = Convert.ToDouble(textBoxGSReadingH.Text);
            data.dut1_ReadingH = Convert.ToDouble(textBoxVTob1H.Text);
            data.dut2_ReadingH = Convert.ToDouble(textBoxVTob2H.Text);
            data.dut3_ReadingH = Convert.ToDouble(textBoxVTob3H.Text);
            data.avg_ReadingH = Convert.ToDouble(textBoxAvgVToH.Text);
            data.dv_ReadingH = Convert.ToDouble(textBoxDVH.Text);
            data.range_ReadingH = Convert.ToDouble(textBoxRangeH.Text);
            data.mc_OffsetOldH = Convert.ToDouble(textBoxOffsetOLDH.Text);
            data.mc_OffsetNewH = Convert.ToDouble(textBoxOffsetNEWH.Text);
            data.buy0ffResultH = textBoxJudgeH.Text;

            //LOW
            data.tempSettingL = labelLow.Text;
            data.gs_ReadingL = Convert.ToDouble(textBoxGSReadingL.Text);
            data.dut1_ReadingL = Convert.ToDouble(textBoxVTob1L.Text);
            data.dut2_ReadingL = Convert.ToDouble(textBoxVTob2L.Text);
            data.dut3_ReadingL = Convert.ToDouble(textBoxVTob3L.Text);
            data.avg_ReadingL = Convert.ToDouble(textBoxAvgVToL.Text);
            data.dv_ReadingL = Convert.ToDouble(textBoxDVL.Text);
            data.range_ReadingL = Convert.ToDouble(textBoxRangeL.Text);
            data.mc_OffsetOldL = Convert.ToDouble(textBoxOffsetOLDL.Text);
            data.mc_OffsetNewL = Convert.ToDouble(textBoxOffsetNEWL.Text);
            data.buy0ffResultL = textBoxJudgeL.Text;

            bool success = data.Insert(data);
            if (success == true)
            {
                MessageBox.Show("Data has been saved!!!");
                ApplicationSet._SerialPort1.Close();
                this.Close();
            }
            else
            {
                MessageBox.Show("Failed to Saved Data!!!", "Try Again");
            }
        }
        void SaveHighOrLow()
        {
            DateTime finishTime;
            finishTime = DateTime.Now;
            textBoxFinish.Text = finishTime.ToString("HH:mm:ss");

            // SAVE TO SQL SERVER
            data.year = textBoxDate.Value.Year;
            data.date = textBoxDate.Value.ToString("yyyy-MM-dd");
            data.machineNumber = cb_Machine.Text;
            data.deviceType = cb_DeviceType.Text;
            data.serialNumber = textBoxSerial.Text;
            data.shift = cb_Shift.Text;
            data.start = textBoxStart.Text;
            data.finish = textBoxFinish.Text;
            data.testJig = textBoxJig.Text;
            data.doneBy = textBoxDoneBy.Text;

            string selectedTemp = cb_Temp.SelectedItem.ToString();
            if (selectedTemp == "HIGH")
            {
                //HIGH
                data.tempSettingH = labelHigh.Text;
                data.gs_ReadingH = Convert.ToDouble(textBoxGSReadingH.Text);
                data.dut1_ReadingH = Convert.ToDouble(textBoxVTob1H.Text);
                data.dut2_ReadingH = Convert.ToDouble(textBoxVTob2H.Text);
                data.dut3_ReadingH = Convert.ToDouble(textBoxVTob3H.Text);
                data.avg_ReadingH = Convert.ToDouble(textBoxAvgVToH.Text);
                data.dv_ReadingH = Convert.ToDouble(textBoxDVH.Text);
                data.range_ReadingH = Convert.ToDouble(textBoxRangeH.Text);
                data.mc_OffsetOldH = Convert.ToDouble(textBoxOffsetOLDH.Text);
                data.mc_OffsetNewH = Convert.ToDouble(textBoxOffsetNEWH.Text);
                data.buy0ffResultH = textBoxJudgeH.Text;

                bool success = data.InsertHigh(data);
                if (success == true)
                {
                    MessageBox.Show("Data has been saved!!!");
                    ApplicationSet._SerialPort1.Close();
                    this.Close();
                }
                else
                {
                    MessageBox.Show("Failed to Saved Data!!!", "Try Again");
                }
            }
            else if (selectedTemp == "LOW")
            {
                //LOW
                data.tempSettingL = labelLow.Text;
                data.gs_ReadingL = Convert.ToDouble(textBoxGSReadingL.Text);
                data.dut1_ReadingL = Convert.ToDouble(textBoxVTob1L.Text);
                data.dut2_ReadingL = Convert.ToDouble(textBoxVTob2L.Text);
                data.dut3_ReadingL = Convert.ToDouble(textBoxVTob3L.Text);
                data.avg_ReadingL = Convert.ToDouble(textBoxAvgVToL.Text);
                data.dv_ReadingL = Convert.ToDouble(textBoxDVL.Text);
                data.range_ReadingL = Convert.ToDouble(textBoxRangeL.Text);
                data.mc_OffsetOldL = Convert.ToDouble(textBoxOffsetOLDL.Text);
                data.mc_OffsetNewL = Convert.ToDouble(textBoxOffsetNEWL.Text);
                data.buy0ffResultL = textBoxJudgeL.Text;

                bool success = data.InsertLow(data);
                if (success == true)
                {
                    MessageBox.Show("Data has been saved!!!");
                    ApplicationSet._SerialPort1.Close();
                    this.Close();
                }
                else
                {
                    MessageBox.Show("Failed to Saved Data!!!", "Try Again");
                }
            }
        }

        private void cb_DeviceType_SelectedIndexChanged(object sender, EventArgs e)
        {
            ConfigData();
            if (cb_Temp.SelectedItem == null)
            {
                TB_TempMachine.Text = "0";
            }
            else
            {
                if (cb_Temp.SelectedItem.ToString() == "HIGH")
                {
                    TB_TempMachine.Text = GlobalVariable.tempHigh.ToString();
                }
                else if (cb_Temp.SelectedItem.ToString() == "LOW")
                {
                    TB_TempMachine.Text = GlobalVariable.tempLow.ToString();
                }
            }
        }
    }
}
