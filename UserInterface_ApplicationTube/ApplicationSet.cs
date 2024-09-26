using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO.Ports;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApplication2
{
    class ApplicationSet
    {
        public static SerialPort _SerialPort1 { get; set; }

        public bool PortInit1(string PortName)
        {
            try
            {
                //PortInit 1
                _SerialPort1 = new SerialPort(PortName, 19200, Parity.None, 8, StopBits.One);
                _SerialPort1.Handshake = Handshake.None;
                _SerialPort1.ReadTimeout = 250;
                _SerialPort1.WriteTimeout = 250;
                _SerialPort1.Close();
                _SerialPort1.Open();
                if (_SerialPort1.IsOpen == true)
                {
                    return false;
                }
                else
                {
                    MessageBox.Show("Error : Port inilitialization error", "error Message");
                    return true;
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("Error : Port Inilitialization Error" + ex.ToString(), "error Message");
                return true;
            }
        }

        public int read_data_Median3(int Ring)
        {
            byte Num = Convert.ToByte(Ring);
            byte[] queryEeprom = { 127, 2, 83, Num };

            _SerialPort1.Write(queryEeprom, 0, 4);

            //Thread.Sleep(250);
            Thread.Sleep(195);
            int bytes = _SerialPort1.BytesToRead;
            byte[] buffer = new byte[2];

            if (bytes > 1)
            {
                _SerialPort1.Read(buffer, 0, 2);
                var hexstring = BitConverter.ToString(buffer);

                if (hexstring != "")
                {
                    hexstring = hexstring.Replace("-", "");
                    int decval = int.Parse(hexstring, System.Globalization.NumberStyles.HexNumber);
                    return decval;
                }
                else
                {
                    return 0;
                }
            }
            else
            {
                return 0;
            }
        }
    }
}
