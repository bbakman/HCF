using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using System.IO;
using ExcelDataReader;
using Extreme.Mathematics;
using Extreme.Statistics;
using Extreme.Mathematics.Curves;
using System.Windows.Forms.DataVisualization.Charting;

namespace HCF_Calculation
{
    public partial class Form1 : Form
    {
        public Form1()
        {

            InitializeComponent();
            
        }
        
        DataTableCollection dtc;
        List<double> xl = new List<double>();
        List<double> yl = new List<double>();
        List<double> x1 = new List<double>();//Top x and y
        List<double> y1 = new List<double>();
        List<double> x2 = new List<double>();//Bottom x and y
        List<double> y2 = new List<double>(); 
        List<double> xf = new List<double>();//Fittet x and y
        List<double> yf = new List<double>();
        List<double> xa = new List<double>();//Average x and y
        List<double> ya = new List<double>();
        List<double> xd = new List<double>();//Derivative x and y
        List<double> yd = new List<double>();
        string asa = "";
        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {

            

        }
         
        private void Button1_Click(object sender, EventArgs e)
        {
            using(OpenFileDialog openfile = new OpenFileDialog())
            {
                if(openfile.ShowDialog()==DialogResult.OK)
                {
                    textBox1.Text = openfile.FileName;

                    using (var stream = File.Open(openfile.FileName, FileMode.Open, FileAccess.Read))
                    {
                        using (IExcelDataReader excelreader = ExcelDataReader.ExcelReaderFactory.CreateReader(stream))
                        {
                            DataSet result = excelreader.AsDataSet(new ExcelDataSetConfiguration()
                            {
                                ConfigureDataTable = (x) => new ExcelDataTableConfiguration() { UseHeaderRow = false }
                            }) ;
                            dtc = result.Tables;
                            comboBox1.Items.Clear();
                            foreach (DataTable table in dtc) comboBox1.Items.Add(table.TableName);
                        }

                    }

                }

            }

        }

        private void button2_Click(object sender, EventArgs e)
        {
            if(comboBox1.SelectedIndex!= null && comboBox1.SelectedIndex>=0)
            {
                DataTable dt = dtc[comboBox1.SelectedIndex];
                string[] xs = dt.Rows.OfType<DataRow>().Select(k => k[0].ToString()).ToArray();
                string[] ys = dt.Rows.OfType<DataRow>().Select(k => k[1].ToString()).ToArray();
                int a = xs.Length;




                for (int i = 0; i < a; i++)
                {


                    xl.Add(Convert.ToDouble(xs[i]));
                    yl.Add(Convert.ToDouble(ys[i]));


                }


                var x = xl.ToArray().ToVector();
                var y = yl.ToArray().ToVector();



                int maxindex = x.MaxIndex();
                int minindex = x.MinIndex();

                int sp = maxindex;
                int ep = minindex;

                if (minindex < maxindex)
                {
                    sp = minindex;
                    ep = maxindex;
                }

                //var xc = x.Concat(x).ToArray().ToVector();
                //var yc = y.Concat(y).ToArray().ToVector();

                for (int i = sp; i < a; i++)
                {

                    x1.Add(x[i]);
                    y1.Add(y[i]);

                }

                for (int i = sp; i >= 0; i--)
                {

                    x2.Add(x[i]);
                    y2.Add(y[i]);

                }

                //x2.Add(x[ep]);
                //y2.Add(y[ep]);

                //Refit the with polynomial aproximation

                xf = x1;
                yf = y1;

                //Define fitting parameters

                double dp;          //Degree of polinom
                double percentile = 1;  //Percentage of previous and next data

                //List<double> xvalue = new List<double>(7);
                //List<double> yvalue = new List<double>(7);

                var x1v = x1.ToArray().ToVector();
                var x2v = x2.ToArray().ToVector();
                var xav = xa.ToArray().ToVector();
                var xfv = xf.ToArray().ToVector();

                var y1v = y1.ToArray().ToVector();
                var y2v = y2.ToArray().ToVector();
                var yav = ya.ToArray().ToVector();
                var yfv = yf.ToArray().ToVector();
                var residual = Vector.Create<double>(8);
                var finder = x2.ToArray().ToVector();
                var arac = x2.ToArray().ToVector();
                double aranan;
                finder.SetValue(1);
                int range = ((int)(a * percentile / 100));
                //arac = (Extreme.Mathematics.LinearAlgebra.DenseVector<double>)(x2v - finder * aranan);
                //int indis = arac.AbsoluteMinIndex(); 
                var xvalue = Vector.Create<double>(range * 2 + 1);
                var yvalue = Vector.Create<double>(range * 2 + 1);

                LinearCurveFitter fitter1 = new LinearCurveFitter();
                LinearCurveFitter fitter2 = new LinearCurveFitter();
                LinearCurveFitter fitter3 = new LinearCurveFitter();
                LinearCurveFitter fitter4 = new LinearCurveFitter();
                LinearCurveFitter fitter5 = new LinearCurveFitter();
                LinearCurveFitter fitter6 = new LinearCurveFitter();
                LinearCurveFitter fitter7 = new LinearCurveFitter();
                LinearCurveFitter fitter8 = new LinearCurveFitter();

                fitter1.Curve = new Polynomial(1);
                fitter2.Curve = new Polynomial(2);
                fitter3.Curve = new Polynomial(3);
                fitter4.Curve = new Polynomial(4);
                fitter5.Curve = new Polynomial(5);
                fitter6.Curve = new Polynomial(6);
                fitter7.Curve = new Polynomial(7);
                fitter8.Curve = new Polynomial(8);

                for (int i = range + 1; i < (x1v.Length - range - 2); i++)
                {

                    aranan = x1v[i];
                    arac = (Extreme.Mathematics.LinearAlgebra.DenseVector<double>)(x2v - finder * aranan);
                    int indis = arac.AbsoluteMinIndex();
                    if (indis < range || (indis + range) > (x2v.Length - 1))
                    {

                        range--;

                        for (int j = 0; j <= (2 * range); j++)
                        {

                            xvalue[j] = x2v[indis + j - range];
                            yvalue[j] = y2v[indis + j - range];

                        }

                        range++;

                    }
                    else
                    {
                        for (int j = 0; j <= (2 * range); j++)
                        {

                            xvalue[j] = x2v[indis + j - range];
                            yvalue[j] = y2v[indis + j - range];

                        }

                    }

                    //LinearCurveFitter fitter = new LinearCurveFitter();

                    //fitter.Curve = new Polynomial(6);

                    //xvalue = Vector.Create<double>(8,5,6,4,7,5);

                    for (int j = 0; j <= (2 * range); j++)
                    {

                        xvalue[j] = x2v[indis + j - range];
                        yvalue[j] = y2v[indis + j - range];

                    }


                    //xvalue = x2.ToArray().Skip(indis - 3).Take(indis + 3).ToArray();
                    //yvalue = y2.ToArray().Skip(indis - 3).Take(indis + 3).ToArray();


                    fitter1.XValues = xvalue;
                    fitter1.YValues = yvalue;
                    fitter2.XValues = xvalue;
                    fitter2.YValues = yvalue;

                    fitter3.XValues = xvalue;
                    fitter3.YValues = yvalue;
                    fitter4.XValues = xvalue;
                    fitter4.YValues = yvalue;

                    fitter5.XValues = xvalue;
                    fitter5.YValues = yvalue;
                    fitter6.XValues = xvalue;
                    fitter6.YValues = yvalue;

                    fitter7.XValues = xvalue;
                    fitter7.YValues = yvalue;
                    fitter8.XValues = xvalue;
                    fitter8.YValues = yvalue;


                    fitter1.Fit();
                    fitter2.Fit();
                    fitter3.Fit();
                    fitter4.Fit();


                    fitter5.Fit();
                    fitter6.Fit();
                    fitter7.Fit();
                    fitter8.Fit();


                    residual[0] = fitter1.Residuals.Norm();
                    residual[1] = fitter2.Residuals.Norm();
                    residual[2] = fitter3.Residuals.Norm();
                    residual[3] = fitter4.Residuals.Norm();


                    residual[4] = fitter5.Residuals.Norm();
                    residual[5] = fitter6.Residuals.Norm();
                    residual[6] = fitter7.Residuals.Norm();
                    residual[7] = fitter8.Residuals.Norm();


                    int bestfit = residual.ToArray().MinIndex();
                    yfv[i] = fitter8.Curve.GetDerivative().ValueAt(xfv[i]);


                    /*
                    if (i==0)
                    { 
                        yfv[i] = fitter1.Curve.GetDerivative().ValueAt(xfv[i]);
                    }
                    else if (i==1)
                    {
                        yfv[i] = fitter2.Curve.GetDerivative().ValueAt(xfv[i]);
                    }
                    else if (i==2)
                    {
                        yfv[i] = fitter3.Curve.GetDerivative().ValueAt(xfv[i]);
                    }
                    else if (i==3)
                    {
                        yfv[i] = fitter4.Curve.GetDerivative().ValueAt(xfv[i]);
                    }
                    else if (i==4)
                    {
                        yfv[i] = fitter5.Curve.GetDerivative().ValueAt(xfv[i]);
                    }
                    else if (i==5)
                    {
                        yfv[i] = fitter6.Curve.GetDerivative().ValueAt(xfv[i]);
                    }
                    else if (i==6)
                    {
                        yfv[i] = fitter7.Curve.GetDerivative().ValueAt(xfv[i]);
                    }
                    else if (i==7)
                    {
                        yfv[i] = fitter8.Curve.GetDerivative().ValueAt(xfv[i]);
                    }        
                    */
                }


                yav = (Extreme.Mathematics.LinearAlgebra.DenseVector<double>)((yfv + y1v) / 2);


                //Top curve and bottom curve data is ready here

                //After that calculate average and dy/dx curve


                var displacement1 = chart1.Series[0];
                displacement1.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Line;
                displacement1.Points.DataBindXY(x1v.ToArray(), y1v.ToArray());


                var displacement2 = chart1.Series[1];
                displacement2.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Line;
                displacement2.Points.DataBindXY(x1v.ToArray(), yfv.ToArray());


                var displacement3 = chart1.Series[2];
                displacement3.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Line;
                displacement3.Points.DataBindXY(x1v.ToArray(), yav.ToArray());
            }
            else
            {
                MessageBox.Show("Kolon seç", "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Error);
                
            }

           

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            DataTable dt = dtc[comboBox1.SelectedIndex];
            dataGridView1.DataSource = dt;
            dataGridView1.Columns[0].HeaderText = "Displacement";
            dataGridView1.Columns[1].HeaderText = "Force";
            
        }

        private void button3_Click(object sender, EventArgs e)
        {
            LinearCurveFitter fitter = new LinearCurveFitter();
            var ones = Vector.Create<double>(1, 1, 1, 1, 1, 1, 1, 1);
            var deflectionData = Vector.Create<double>(1,5,2,0,3,6, 1, 5, 2, 0, 3, 6, 1, 5, 2, 0, 3, 6);
            var loadData = Vector.Create<double>(0,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17);
            fitter.Curve = new Polynomial(10);
            fitter.XValues = loadData; 
            fitter.YValues = deflectionData;
            //fitter.WeightFunction = WeightFunctions.OneOverXSquared;
            fitter.Fit();
            

            var displacement1 = chart1.Series[0];

            displacement1.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Line;

            displacement1.Points.DataBindXY(loadData.ToArray(), deflectionData.ToArray());
            
            var solution = fitter.BestFitParameters;
            Polynomial polynom1 = new Polynomial(3);


            polynom1[2] = 1;
            polynom1[1] = -5;
            polynom1[0] = 0;
            var x = new linspace();
            x.series(0, 20, 5);
            textBox1.Text = polynom1.ValueAt(2).ToString();
            Curve derivative = polynom1.GetDerivative();
            double c1 = solution[0];
            double c2 = solution[1];
            double c3 = solution[2];
            double c4 = solution[3];
            var s = fitter.GetStandardDeviations();
            s = x.series(0, 50, 1000);
            var y = s.Sin()+(2*s).Cos();
            Console.WriteLine("Solution (weighted observations):");
            Console.WriteLine("c1: {0,20:E10} {1,20:E10}", solution[0], s[0]);
            Console.WriteLine("c2: {0,20:E10} {1,20:E10}", solution[1], s[1]);
            Console.WriteLine("c3: {0,20:E10} {1,20:E10}", solution[2], s[2]);
            Console.WriteLine("c4: {0,20:E10} {1,20:E10}", solution[3], s[3]);
            Console.WriteLine("c5: {0,20:E10} {1,20:E10}", solution[4], s[4]);
            
            deflectionData = Vector.Create<double>(0, 0.5, 1, 1.5, 2);
            loadData = deflectionData; 

            textBox2.Text = Convert.ToString(c1)+"  "+ Convert.ToString(c2) +"  " +Convert.ToString(c3);
            var deflectionData1 = deflectionData; 
            var displacement2 = chart1.Series[1];

             
            displacement2.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Line;
            displacement2.Points.DataBindXY(s.ToArray(), y.ToArray());
        }

        private void Form1_Load(object sender, EventArgs e)
        {

            // Define Charts
            var displacement1 = chart1.Series.Add("Anvil Table");
            var displacement2 = chart1.Series.Add("Equipment");
            var displacement3 = chart1.Series.Add("Average");

        }

        private void button4_Click(object sender, EventArgs e)
        {
               string clipboardText = Clipboard.GetText(TextDataFormat.Text);
               dataGridView1.DataSource = clipboardText;
               // dataGridView1.Columns[0].HeaderText = "Displacement";
               // dataGridView1.Columns[1].HeaderText = "Force"; 
               // Do whatever you need to do with clipboardText
        }
        System.Drawing.Point? prevPosition = null;
        ToolTip tooltip = new ToolTip();

        private void chart1_MouseMove(object sender, MouseEventArgs e)
        {
            var pos = e.Location;
            if (prevPosition.HasValue && pos == prevPosition.Value)
                return;
            tooltip.RemoveAll();
            prevPosition = pos;
            var results = chart1.HitTest(pos.X, pos.Y, false, ChartElementType.DataPoint); // set ChartElementType.PlottingArea for full area, not only DataPoints
            foreach (var result in results)
            {
                if (result.ChartElementType == ChartElementType.DataPoint) // set ChartElementType.PlottingArea for full area, not only DataPoints
                {
                    var yVal = result.ChartArea.AxisY.PixelPositionToValue(pos.Y);
                    var xVal = result.ChartArea.AxisX.PixelPositionToValue(pos.X);
                    tooltip.Show(((double)yVal).ToString()+"-"+ ((double)xVal).ToString(), chart1, pos.X, pos.Y - 15);
                }
            }
        }
    }
}
