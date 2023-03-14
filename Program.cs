using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportePTGPS
{
    class Program
    {
        static void Main(string[] args)
        {
            Metodos _met = new Metodos();
            
            string dayNow = DateTime.UtcNow.AddDays(0).ToString("dddd");
            if (dayNow == "lunes" || dayNow == "Monday")
            {
                _met.SendDataEmailWeekEnd();
            }
            else
            {
                _met.SendDataEmaiYesterday();
            }
        }
    }
}
