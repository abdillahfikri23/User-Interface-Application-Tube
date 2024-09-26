using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using System.Configuration;
using System.Data;

namespace WindowsFormsApplication2
{
    class GlobalVariable
    {
        public static Double average_gsH { get; set; }
        public static Double average_gsL { get; set; }
        public static Double gsreadingH { get; set; }
        public static Double gsreadingL { get; set; }
        public static Double rangeH { get; set; }
        public static Double rangeL { get; set; }
        public static Double average_maxH { get; set; }
        public static Double average_minH { get; set; }
        public static Double average_maxL { get; set; }
        public static Double average_minL { get; set; }
        public static Double Dut1H_Value { get; set; }
        public static Double Dut2H_Value { get; set; }
        public static Double Dut3H_Value { get; set; }
        public static Double Dut1L_Value { get; set; }
        public static Double Dut2L_Value { get; set; }
        public static Double Dut3L_Value { get; set; }
        public static Double DV_ValueH { get; set; }
        public static Double DV_ValueL { get; set; }
        public static Double currentValue { get; set; }
        public static Double newValue { get; set; }
        public static Double Vtamb_Value { get; set; }
        public static Double rangeValue { get; set; }
        public static Double movingAverage { get; set; }
        public static Double tempHigh { get; set; }
        public static Double tempLow { get; set; }
        public static int Dv { get; set; }
        public static int Range { get; set; }
        public static Double ValueOff { get; set; }
    }
}
