using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Extreme.Mathematics;

namespace HCF_Calculation
{
    class linspace
    {
        
        
        public Vector<double> series(double startingpoint, double endpoint, int numbers)
        {
            //var seri = Vector.Create<double>(5);
            Vector<double> seri = Vector.Create<double>(numbers);
            double sp = startingpoint;
            double ep = endpoint;
            double n = numbers;
            int sayac = 0;
            double dt = (ep - sp) / (n - 1);
            for (int i = 0; i <= n-1; i++)
                {
                seri[i] = sp;
                sp = sp + dt;
                }


            return seri;
        }

    }
}
