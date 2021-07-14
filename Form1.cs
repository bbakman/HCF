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
        //Defining Lists
        DataTableCollection dtc;
        List<double> xl = new List<double>();
        List<double> yl = new List<double>();
        List<double> x1 = new List<double>();//Loading x and y
        List<double> y1 = new List<double>();
        List<double> x2 = new List<double>();//Rebound x and y
        List<double> y2 = new List<double>(); 
        List<double> xf = new List<double>();//Fitted Rebound x and y
        List<double> yf = new List<double>();
        List<double> xa = new List<double>();//Average x and y
        List<double> ya = new List<double>();
        List<double> xs = new List<double>();//Stiffness x and y
        List<double> ys = new List<double>();
       
        
        private void Button1_Click(object sender, EventArgs e)
        {
            //Select File Button
            //Standart Excel Read Code
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
                        }
                    }
                }
            }
            //Boolean Operation for button 
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
        {       //Create Data Table   
                DataTable dt = dtc[0];
                //Separate Stiffness and 
                string[] xs = dt.Rows.OfType<DataRow>().Select(k => k[0].ToString()).ToArray();
                string[] ys = dt.Rows.OfType<DataRow>().Select(k => k[1].ToString()).ToArray();
                int datalenght = xs.Length;

                for (int i = 0; i < datalenght; i++)
                {
                    xl.Add(Convert.ToDouble(xs[i]));
                    yl.Add(Convert.ToDouble(ys[i]));
                }
                //List to vector for mathemtical operations
                var x = xl.ToArray().ToVector();
                var y = yl.ToArray().ToVector();
                
                int maxindex = x.MaxIndex();//İndex of max values measurement
                int minindex = x.MinIndex();//İndex of min values measurement
                int sp = maxindex;
                //Detecting turn point of data it would be min value or max value so that we need boolean operations
                if (minindex < maxindex)
                {
                    sp = minindex;                    
                }
                //Creating new vectors from starting points
                for (int i = sp; i < datalenght; i++)
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

                int dp = 6;          //Degree of polinom
                double percentile = 1;  //Percentage of previous and next data

                //List to vector
                var x1v = x1.ToArray().ToVector();
                var x2v = x2.ToArray().ToVector();
                var xav = xa.ToArray().ToVector();
                var xfv = xf.ToArray().ToVector();

                var y1v = y1.ToArray().ToVector();
                var y2v = y2.ToArray().ToVector();
                var yav = ya.ToArray().ToVector();
                var yfv = yf.ToArray().ToVector();

                var ysv = yf.ToArray().ToVector();               
                var finder = x2.ToArray().ToVector();
                var arac = x2.ToArray().ToVector();
                double wanted;//
                finder.SetValue(1);//Change to all value to 1 such as ones() in matlab
                int range = ((int)(datalenght * percentile / 100));
                //Define data segment for polynomial fitting from percentile of data parameter
                var xvalue = Vector.Create<double>(range * 2 + 1);
                var yvalue = Vector.Create<double>(range * 2 + 1);

                //Create curve fitter
                LinearCurveFitter fitter = new LinearCurveFitter();
                //Define ploynamial curve fitting from given value
                fitter.Curve = new Polynomial(dp);


                for (int i = 0; i<range; i++ )
                {
                    wanted = x1v[i];
                    //arac is using for lookup table method
                    //x2v=[5,4,,2,9,6] 
                    //wanted = 8;
                    //abs([5,4,2,9,6]-[1,1,1,1,1]*8)=[3,4,6,1,2]
                    //min index 4 we use 4 for looking to x2v
                    arac = (Extreme.Mathematics.LinearAlgebra.DenseVector<double>)(x2v - finder * wanted);
                    int indis = arac.AbsoluteMinIndex();               
                }
                for (int i = 0; i < x1v.Length; i++)
                {
                    wanted = x1v[i];
                    arac = (Extreme.Mathematics.LinearAlgebra.DenseVector<double>)(x2v - finder * wanted);
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
                    fitter.XValues = xvalue;
                    fitter.YValues = yvalue;                    
                    fitter.Fit();                  
                    yfv[i] = fitter.Curve.ValueAt(xfv[i]);
                    ysv[i] = fitter.Curve.GetDerivative().ValueAt(xfv[i]);                                    
                }
                //Average of Loading and Rebound Data
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
                var curve4 = chart2.Series[0];
                curve4.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Line;
                curve4.Points.DataBindXY(xfv.ToArray(), ysv.ToArray());     
        }             

        private void Form1_Load(object sender, EventArgs e)
        {
            var curve1 = chart1.Series.Add("Load");
            var curve2 = chart1.Series.Add("Rebound");
            var curve3 = chart1.Series.Add("Average");
            var curve4 = chart2.Series.Add("Stiffness");
        }              
        //Usşing for mouse data read
        System.Drawing.Point? prevPosition = null;
        ToolTip tooltip = new ToolTip();

        private void chart2_MouseMove(object sender, MouseEventArgs e)
        {
            var pos = e.Location;
            if (prevPosition.HasValue && pos == prevPosition.Value)
                return;
            tooltip.RemoveAll();
            prevPosition = pos;
            var results = chart2.HitTest(pos.X, pos.Y, false, ChartElementType.DataPoint); // set ChartElementType.PlottingArea for full area, not only DataPoints
            foreach (var result in results)
            {
                if (result.ChartElementType == ChartElementType.DataPoint) // set ChartElementType.PlottingArea for full area, not only DataPoints
                {
                    var yVal = result.ChartArea.AxisY.PixelPositionToValue(pos.Y);
                    var xVal = result.ChartArea.AxisX.PixelPositionToValue(pos.X);
                    tooltip.Show(((double)xVal).ToString("0.###") + " , " + ((double)yVal).ToString("0.###"), chart2, pos.X, pos.Y - 15);
                }
            }
        }
    }
}
