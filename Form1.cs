using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using NPOI;
using System.Security.Cryptography;
using NPOI.SS.UserModel;
using System.IO;
using NPOI.XSSF.UserModel;
using System.Runtime.Remoting;


namespace ProgramP3
{
    public partial class Form1 : Form
    {
        double signal_noise;

        double[] sinyal_hasil = new double[15000];

        double[] sinyal_hasil_geser = new double[15000];
        double[] sinyal_hasil_filter = new double[15000];

        double[,] sinyal_sensor = new double[15000, 15000];
        double[] sinyal_noise = new double[15000];

        int amplitudoNoise = 0;
        int frekuensiNoise = 0;

        double[] Realx = new double[15000];
        double[] Imajiner = new double[15000];
        double[] absolute = new double[15000];

        string fileDataLocation = "";
        int urutan = 0;

        int jumlahdataSensor = 0;

        string cek_filter2;
        string cek_modifikasi_sensor;
        int Msensor = 0;
        int frekuensiCuttOffsensor = 0;

        public Form1()
        {
            InitializeComponent();

        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void chart1_Click(object sender, EventArgs e)
        {

        }

        private void chart2_Click(object sender, EventArgs e)
        {

        }

        private void chart3_Click(object sender, EventArgs e)
        {

        }

        private void button5_Click(object sender, EventArgs e)
        {
            button6.Visible = true;

            try
            {
                OpenFileDialog dialog = new OpenFileDialog();
                dialog.Filter = "xlsx files(*.xlsx)|*.xlsx| All Files(*.*)|*.*";

                if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    fileDataLocation = dialog.FileName;
                    label12.Visible = true;
                    label12.Text = fileDataLocation;

                }
            }

            catch (Exception)
            {
                MessageBox.Show("tidak berhasil membuka file");
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            groupBox6.Visible = true;
            groupBox7.Visible = true;
            groupBox9.Visible = true;
            groupBox1.Visible = true;

            chart1.Series.Clear();

            Random random = new Random();
            Color[] colors = new Color[] { Color.Blue, Color.Green, Color.Black, Color.Orange, Color.Violet, Color.Chocolate ,Color.AliceBlue , Color.DarkBlue , Color.DarkGreen};

            int randomColorIndex = random.Next(colors.Length);

            DataTable dtTable = new DataTable();
            ISheet sheet;

            using (var fStream = new FileStream(fileDataLocation, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                fStream.Position = 0;
                XSSFWorkbook objWorkbook = new XSSFWorkbook(fStream);

                sheet = objWorkbook.GetSheetAt(0);
                IRow row;
                ICell cell;

                int countCells = 2;
                int rowCount = sheet.LastRowNum;

                for (int i = 1; i < countCells; i++)
                {
                    urutan++;

                    string garis;
                    garis = "Sensor " + urutan;
                    chart1.Series.Add(garis);
                    chart1.Series[garis].Color = colors[randomColorIndex];
                    chart1.Series[garis].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.FastLine;

                    randomColorIndex++;

                    double isiCell = 0;

                    for (int j = 3; j < rowCount + 1; j++)
                    {
                        row = sheet.GetRow(j);

                        isiCell = double.Parse(row.GetCell(i).NumericCellValue.ToString());

                        int t = j - 3;

                        sinyal_sensor[urutan, t] = isiCell;
                    }

                    for (int k = 1; k < jumlahdataSensor; k++)
                    {
                        double a = Convert.ToDouble(k);
                        double b = a / 1000;

                        chart1.Series[garis].Points.AddXY(b, sinyal_sensor[urutan, k]);
                    }
                }          
            }
        }

        private void button9_Click(object sender, EventArgs e)
        {
            chart4.Series.Clear();

            chart4.Visible = true;
            panel3.Visible = true;

            button10.Visible = true;

            SettingNoiseSignal();
            GenerateNoiseSignal();

            button7.Visible = true;
            button8.Visible = true;
        }

        private void GenerateNoiseSignal()
        {
            string garis;
            garis = "Sinyal noise";
            chart4.Series.Add(garis);
            chart4.Series[garis].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.FastLine;

            for (int i = 1; i <= 1000 + 1; i++)
            {
                double a = Convert.ToDouble(i);
                double b = a / 1000;

                signal_noise = amplitudoNoise * Math.Sin((2 * 3.14 * frekuensiNoise) * b + 0);

                sinyal_noise[i] = signal_noise;

                chart4.Series[garis].Points.AddXY(b, signal_noise);
                chart4.Series[garis].Color = Color.Red;
            }
        }
        private void SettingNoiseSignal()
        {
            trackBar10.Value = Convert.ToInt32(textBox10.Text);
            trackBar9.Value = Convert.ToInt32(textBox11.Text);

            frekuensiNoise = Convert.ToInt32(textBox10.Text);
            amplitudoNoise = Convert.ToInt32(textBox11.Text);
        }

        private void trackBar10_ValueChanged(object sender, EventArgs e)
        {
            chart4.Series.Clear();
            frekuensiNoise = trackBar10.Value;

            textBox10.Text = frekuensiNoise.ToString();

            GenerateNoiseSignal();
        }

        private void trackBar9_ValueChanged(object sender, EventArgs e)
        {
            chart4.Series.Clear();
            amplitudoNoise = trackBar9.Value;

            textBox11.Text = amplitudoNoise.ToString();

            GenerateNoiseSignal();
        }

        private void button12_Click(object sender, EventArgs e)
        {
            groupBox6.Visible = true;
            groupBox7.Visible = true;
            groupBox9.Visible = true;

            chart2.Series.Clear();

            string garis;
            garis = "Sinyal Sensor";
            chart2.Series.Add(garis);
            chart2.Series[garis].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.FastLine;

            for (int i = 1; i < jumlahdataSensor; i++)
            {
                double a = Convert.ToDouble(i);
                double b = a / 1000;

                sinyal_hasil[i] = sinyal_sensor[1, i] ;

                chart2.Series[garis].Points.AddXY(b, sinyal_hasil[i]);
                chart2.Series[garis].Color = Color.Blue;
            }
        }

        private void button13_Click(object sender, EventArgs e)
        {
            groupBox6.Visible = true;
            groupBox7.Visible = true;
            groupBox9.Visible = true;

            chart2.Series.Clear();

            string garis;
            garis = "Sinyal hasil x";
            chart2.Series.Add(garis);
            chart2.Series[garis].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.FastLine;

            for (int i = 1; i < jumlahdataSensor; i++)
            {
                double a = Convert.ToDouble(i);
                double b = a / 1000;

                sinyal_hasil[i] = sinyal_sensor[1, i] * sinyal_sensor[2, i] * sinyal_sensor[3, i];

                chart2.Series[garis].Points.AddXY(b, sinyal_hasil[i]);
                chart2.Series[garis].Color = Color.Blue;
            }
        }

        private void button10_Click(object sender, EventArgs e)
        {
            button10.Visible = false;

            button7.Visible = false;
            button8.Visible = false;

            panel3.Visible = false;

            chart2.Series.Clear();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            penjumlahansensor();
            cek_modifikasi_sensor = "penjumlahan";
        }

        private void button8_Click(object sender, EventArgs e)
        {
            perkaliansensor();
            cek_modifikasi_sensor = "perkalian";
        }

        private void penjumlahansensor()
        {
            chart2.Series.Clear();

            string garis;
            garis = "Sinyal hasil +";
            chart2.Series.Add(garis);
            chart2.Series[garis].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.FastLine;

            for (int i = 1; i < jumlahdataSensor; i++)
            {
                double a = Convert.ToDouble(i);
                double b = a / 1000;

                sinyal_hasil[i] = sinyal_sensor[1, i] + sinyal_noise[i];

                chart2.Series[garis].Points.AddXY(b, sinyal_hasil[i]);
                chart2.Series[garis].Color = Color.Blue;
            }
        }

        private void perkaliansensor()
        {
            chart2.Series.Clear();

            string garis;
            garis = "Sinyal hasil x";
            chart2.Series.Add(garis);
            chart2.Series[garis].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.FastLine;

            for (int i = 1; i < jumlahdataSensor; i++)
            {
                double a = Convert.ToDouble(i);
                double b = a / 1000;

                sinyal_hasil[i] = sinyal_sensor[1, i] * sinyal_noise[i];

                chart2.Series[garis].Points.AddXY(b, sinyal_hasil[i]);
                chart2.Series[garis].Color = Color.Blue;
            }
        }

        // dft datasheet //
        private void button11_Click(object sender, EventArgs e)
        {
            chart3.Series.Clear();

            string garis;
            garis = "DFT Filter";

            chart3.Series.Add(garis);
            chart3.Series[garis].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.FastLine;

            for (int i = 0; i <=jumlahdataSensor; i++)
            {
                Realx[i] = 0;
                Imajiner[i] = 0;

                for (int k = 0; k <= jumlahdataSensor; k++)
                {
                    Realx[i] = Realx[i] + (sinyal_hasil_filter[k] * Math.Cos(2 * Math.PI * k * i / jumlahdataSensor));
                    Imajiner[i] = Imajiner[i] - (sinyal_hasil_filter[k] * Math.Sin(2 * Math.PI * k * i / jumlahdataSensor));
                }

                //Realx[i] = (1 / jumlahdataSensor) * Realx[i];
                //Imajiner[i] = (1 / jumlahdataSensor) * Imajiner[i];

                absolute[i] = Math.Sqrt(Math.Pow(Realx[i], 2) + Math.Pow(Imajiner[i], 2));

                progressBar2.Value = i;
            }

            for (int i = 1; i <= jumlahdataSensor; i++)
            {
                chart3.Series[garis].Points.AddXY(i*10, absolute[i]/ jumlahdataSensor);;
                chart3.Series[garis].Color = Color.Red;
            }

        }

        private void button16_Click(object sender, EventArgs e)
        {
            jumlahdataSensor = Convert.ToInt32(textBox14.Text);
        }

        private void button15_Click(object sender, EventArgs e)
        {
            if (comboBox2.Text != null)
            {
                cek_filter2 = comboBox2.Text;

                trackBar12.Value = 1;
                textBox13.Text = null;
                label20.Text = "Nilai :";
            }

            button10.Visible = false;
        }

        private void trackBar12_ValueChanged(object sender, EventArgs e)
        {


            switch (cek_modifikasi_sensor)
            {
                case "penjumlahan":
                    {
                        switch (cek_filter2)
                        {
                            case "Low pass":
                                {
                                    chart2.Series.Clear();

                                    penjumlahansensor();

                                    label20.Text = "Nilai low pass :";

                                    frekuensiCuttOffsensor = trackBar12.Value;
                                    textBox13.Text = frekuensiCuttOffsensor.ToString();

                                    lowPassFilter2();
                                }
                                break;

                            case "High pass":
                                {
                                    chart2.Series.Clear();

                                    penjumlahansensor();

                                    label20.Text = "Nilai high pass :";

                                    frekuensiCuttOffsensor = trackBar12.Value;
                                    textBox13.Text = frekuensiCuttOffsensor.ToString();

                                    highPassFilter2();
                                }
                                break;

                            case "Moving average":
                                {
                                    chart2.Series.Clear();

                                    penjumlahansensor();

                                    label20.Text = "Nilai moving average :";

                                    Msensor = trackBar12.Value;
                                    textBox13.Text = Msensor.ToString();

                                    movingAverageFilter2();
                                }
                                break;
                        }
                    }
                    break;

                case "perkalian":
                    {
                        switch (cek_filter2)
                        {
                            case "Low pass":
                                {
                                    chart2.Series.Clear();

                                    perkaliansensor();

                                    label20.Text = "Nilai low pass :";

                                    frekuensiCuttOffsensor = trackBar12.Value;
                                    textBox13.Text = frekuensiCuttOffsensor.ToString();

                                    lowPassFilter2();
                                }
                                break;

                            case "High pass":
                                {
                                    chart2.Series.Clear();

                                    perkaliansensor();

                                    label20.Text = "Nilai high pass :";

                                    frekuensiCuttOffsensor = trackBar12.Value;
                                    textBox13.Text = frekuensiCuttOffsensor.ToString();

                                    highPassFilter2();
                                }
                                break;

                            case "Moving average":
                                {
                                    chart2.Series.Clear();

                                    perkaliansensor();

                                    label20.Text = "Nilai moving average :";

                                    Msensor = trackBar12.Value;
                                    textBox13.Text = Msensor.ToString();

                                    movingAverageFilter2();
                                }
                                break;
                        }

                    }
                    break;
            }

            switch (cek_filter2)
            {
                case "Low pass":
                    {
                        chart2.Series.Clear();

                        penjumlahansensor();

                        label20.Text = "Nilai low pass :";

                        frekuensiCuttOffsensor = trackBar12.Value;
                        textBox13.Text = frekuensiCuttOffsensor.ToString();

                        lowPassFilter2();
                    }
                    break;

                case "High pass":
                    {
                        chart2.Series.Clear();

                        penjumlahansensor();

                        label20.Text = "Nilai high pass :";

                        frekuensiCuttOffsensor = trackBar12.Value;
                        textBox13.Text = frekuensiCuttOffsensor.ToString();

                        highPassFilter2();
                    }
                    break;

                case "Moving average":
                    {
                        chart2.Series.Clear();

                        penjumlahansensor();

                        label20.Text = "Nilai moving average :";

                        Msensor = trackBar12.Value;
                        textBox13.Text = Msensor.ToString();

                        movingAverageFilter2();
                    }
                    break;

            }
        }

        private void movingAverageFilter2()
        {
            string garis;
            garis = "Sinyal hasil filter";
            chart2.Series.Add(garis);
            chart2.Series[garis].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.FastLine;

            for (int k = 0; k <= jumlahdataSensor + Msensor; k++)
            {
                sinyal_hasil_geser[k] = 0;
            }

            for (int k = Msensor; k <= jumlahdataSensor + Msensor; k++)
            {
                sinyal_hasil_geser[k - 1] = sinyal_hasil[k - Msensor];
            }

            for (int i = 0; i <= jumlahdataSensor; i++)
            {
                double a = Convert.ToDouble(i);
                double b = a / 1000;

                sinyal_hasil_filter[i] = 0;

                for (int j = i; j <= (Msensor - 1) + i; j++)
                {
                    int k = i - j;
                    sinyal_hasil_filter[i] = sinyal_hasil_filter[i] + sinyal_hasil_geser[i - k];
                }

                sinyal_hasil_filter[i] = sinyal_hasil_filter[i] / Msensor;

                chart2.Series[garis].Points.AddXY(b, sinyal_hasil_filter[i]);
                chart2.Series[garis].Color = Color.Red;
            }
        }

        private void lowPassFilter2()
        {
            double rc = 1.0 / (2.0 * Math.PI * frekuensiCuttOffsensor);
            double dt = 1.0 / jumlahdataSensor;
            double alpha = dt / (rc + dt);

            string garis;
            garis = "Sinyal hasil filter";
            chart2.Series.Add(garis);
            chart2.Series[garis].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.FastLine;

            for (int i = 1; i <= jumlahdataSensor; i++)
            {
                double a = Convert.ToDouble(i);
                double b = a / 1000;

                sinyal_hasil_filter[i] = alpha * sinyal_hasil[i] + (1 - alpha) * sinyal_hasil_filter[i - 1];

                chart2.Series[garis].Points.AddXY(b, sinyal_hasil_filter[i]);
                chart2.Series[garis].Color = Color.Red;
            }
        }

        private void highPassFilter2()
        {
            double rc = 1.0 / (2.0 * Math.PI * (100 - frekuensiCuttOffsensor));
            double dt = 1.0 / jumlahdataSensor;
            double alpha = dt / (rc + dt);

            string garis;
            garis = "Sinyal hasil filter";
            chart2.Series.Add(garis);
            chart2.Series[garis].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.FastLine;

            for (int i = 1; i <= jumlahdataSensor; i++)
            {
                double a = Convert.ToDouble(i);
                double b = a / 1000;

                sinyal_hasil_filter[i] = alpha * sinyal_hasil[i] + (1 - alpha) * sinyal_hasil_filter[i - 1];

                chart2.Series[garis].Points.AddXY(b, sinyal_hasil_filter[i]);
                chart2.Series[garis].Color = Color.Red;
            }
        }

        private void button17_Click(object sender, EventArgs e)
        {
            chart1.Series.Clear();
            chart2.Series.Clear();
            chart3.Series.Clear();
            chart5.Series.Clear();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            chart5.Series.Clear();

            string garis;
            garis = "DFT Modifikasi";

            chart5.Series.Add(garis);
            chart5.Series[garis].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.FastLine;

            for (int i = 0; i <= jumlahdataSensor; i++)
            {
                Realx[i] = 0;
                Imajiner[i] = 0;

                for (int k = 0; k <= jumlahdataSensor; k++)
                {
                    Realx[i] = Realx[i] + (sinyal_hasil[k] * Math.Cos(2 * Math.PI * k * i / jumlahdataSensor));
                    Imajiner[i] = Imajiner[i] - (sinyal_hasil[k] * Math.Sin(2 * Math.PI * k * i / jumlahdataSensor));
                }

                //Realx[i] = (1 / jumlahdataSensor) * Realx[i];
                //Imajiner[i] = (1 / jumlahdataSensor) * Imajiner[i];

                absolute[i] = Math.Sqrt(Math.Pow(Realx[i], 2) + Math.Pow(Imajiner[i], 2));

                progressBar1.Value = i;
            }

            for (int i = 1; i <= jumlahdataSensor; i++)
            {
                chart5.Series[garis].Points.AddXY(i * 10, absolute[i] / jumlahdataSensor); ;
                chart5.Series[garis].Color = Color.Red;
            }
        }

        /// ////////// end of tab page 2 ///////// ///

    }
}
