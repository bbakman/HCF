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
            //
            if (textBox1.Text!=null)
            {
                button2.Enabled = true;
            }
            else
            {
                button2.Enabled = false;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {          
                DataTable dt = dtc[0];
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

                xf = x1;
                yf = y1;
                           
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

                var ysv = yf.ToArray().ToVector();
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
                fitter8.Curve = new Polynomial(6);


                for (int i = 0; i<range; i++ )
                {
                    aranan = x1v[i];
                    arac = (Extreme.Mathematics.LinearAlgebra.DenseVector<double>)(x2v - finder * aranan);
                    int indis = arac.AbsoluteMinIndex();
                   

                }


                for (int i = 0; i < x1v.Length; i++)
                {

                    aranan = x1v[i];
                    arac = (Extreme.Mathematics.LinearAlgebra.DenseVector<double>)(x2v - finder * aranan);
                    int indis = arac.AbsoluteMinIndex();
                    if (indis < range )
                    {

                        

                        for (int j = 0; j <= (2 * range); j++)
                        {

                            xvalue[j] = x2v[ j ];
                            yvalue[j] = y2v[ j ];

                        }

                        

                    }
                    else if ((indis + range) > (x2v.Length - 1))

                    {
                        for (int j = 0; j <= (2 * range); j++)
                        {
                            
                            xvalue[j] = x2v[x2v.Length - 2 * range - 1+j];
                            yvalue[j] = y2v[x2v.Length - 2 * range - 1+j];

                        }

                    }

                    else
                    {
                        for (int j = 0; j <= (2 * range); j++)
                        {

                            xvalue[j] = x2v[indis + j - range];
                            yvalue[j] = y2v[indis + j - range];

                        }

                    }      
                    
                    fitter8.XValues = xvalue;
                    fitter8.YValues = yvalue;                    
                    fitter8.Fit();                  

                    yfv[i] = fitter8.Curve.ValueAt(xfv[i]);
                    ysv[i] = fitter8.Curve.GetDerivative().ValueAt(xfv[i]);
                                    
                }

                yav = (Extreme.Mathematics.LinearAlgebra.DenseVector<double>)((yfv + y1v) / 2);

                var curve1 = chart1.Series[0];
                curve1.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Line;
                curve1.Points.DataBindXY(x1v.ToArray(), y1v.ToArray());
                var curve2 = chart1.Series[1];
                curve2.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Line;
                curve2.Points.DataBindXY(x2v.ToArray(), y2v.ToArray());
                var curve3 = chart1.Series[2];
                curve3.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Line;
                curve3.Points.DataBindXY(x1v.ToArray(), yav.ToArray());
                var curve4 = chart1.Series[3];
                curve4.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Line;
                curve4.Points.DataBindXY(xfv.ToArray(), ysv.ToArray());       

        }             
        private void Form1_Load(object sender, EventArgs e)
        {
            var curve1 = chart1.Series.Add("Load");
            var curve2 = chart1.Series.Add("Rebound");
            var curve3 = chart1.Series.Add("Average");
            var curve4 = chart1.Series.Add("Stiffness");
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
                    tooltip.Show(((double)xVal).ToString("0.###")+" , "+ ((double)yVal).ToString("0.###"), chart1, pos.X, pos.Y - 15);
                }
            }
        }
    }
}
