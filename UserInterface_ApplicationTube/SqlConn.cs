using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using System.Data;
using System.Windows.Forms;

namespace WindowsFormsApplication2
{
    class SqlConn
    {
        // Variable Database
        public int year { get; set; }
        public string date { get; set; }
        public string machineNumber { get; set; }
        public string deviceType { get; set; }
        public string serialNumber { get; set; }
        public string shift { get; set; }
        public string start { get; set; }
        public string finish { get; set; }
        public string testJig { get; set; }
        public string tempSettingH { get; set; }
        public double gs_ReadingH { get; set; }
        public double dut1_ReadingH { get; set; }
        public double dut2_ReadingH { get; set; }
        public double dut3_ReadingH { get; set; }
        public double avg_ReadingH { get; set; }
        public double dv_ReadingH { get; set; }
        public double range_ReadingH { get; set; }
        public double mc_OffsetOldH { get; set; }
        public double mc_OffsetNewH { get; set; }
        public string buy0ffResultH { get; set; }
        public string tempSettingL { get; set; }
        public double gs_ReadingL { get; set; }
        public double dut1_ReadingL { get; set; }
        public double dut2_ReadingL { get; set; }
        public double dut3_ReadingL { get; set; }
        public double avg_ReadingL { get; set; }
        public double dv_ReadingL { get; set; }
        public double range_ReadingL { get; set; }
        public double mc_OffsetOldL { get; set; }
        public double mc_OffsetNewL { get; set; }
        public string buy0ffResultL { get; set; }
        public string doneBy { get; set; }

        string connectionString = "Data Source=etcbtmsql01\\digitmont;Initial Catalog=dbDigitmontDetection;User ID=Metis;Password=Metis123";

        // SQL SERVER DATABASE
        public DataTable Select()
        {
            SqlConnection conn = new SqlConnection();
            conn.ConnectionString = "Data Source=etcbtmsql01\\digitmont;Initial Catalog=dbDigitmontDetection;User ID=Metis;Password=Metis123";
            DataTable dt = new DataTable();
            try
            {
                string sql = "SELECT * FROM TESTER_APPLICATION_TUBE";
                SqlCommand cmd = new SqlCommand(sql, conn);
                SqlDataAdapter adapter = new SqlDataAdapter(cmd);
                conn.Open();
                adapter.Fill(dt);
                
            }
            catch (Exception ex)
            {
                MessageBox.Show("An error occurred: " + ex.Message);
            }
            finally
            {
                conn.Close();
            }
            return dt;
        }

        //Insert Data Into Database
        public bool Insert(SqlConn data)
        {
            bool isSuccess = false;
            SqlConnection conn = new SqlConnection(connectionString);
            try
            {
                //string sql = "INSERT INTO TESTER_APPLICATION_TUBE (Device, Machine, Date, Shift, Temp, Jig, SerialName) VALUES(@Device, @Machine, @Date, @Shift, @Temp, @Jig, @SerialName)";
                string sql1 = "INSERT INTO TESTER_APPLICATION_TUBE (YEAR, DATE, MACHINE_NUMBER, DEVICE_TYPE, SERIAL_NUMBER, SHIFT, START, FINISH, TEST_JIG, BB_TEMP_SETTING, GS_READING, DUT1_READING, DUT2_READING, DUT3_READING, AVG_READING, DV_READING, RANGE_READING, MC_OFFSET_OLD, MC_OFFSET_NEW, BUYOFF_RESULT, DONE_BY) VALUES (@Year, @Date, @MachineNumber, @DeviceType, @SerialNumber, @Shift, @Start, @Finish, @TestJig, @TempSetting, @Gs_Reading, @Dut1_Reading, @Dut2_Reading, @Dut3_Reading, @Avg_Reading, @Dv_Reading, @Range_Reading, @Mc_OffsetOld, @Mc_OffsetNew, @BuyoffResult, @DoneBy)";
                SqlCommand cmd1 = new SqlCommand(sql1, conn);

                // Add parameters
                cmd1.Parameters.AddWithValue("@Year", year);
                cmd1.Parameters.AddWithValue("@Date", date);
                cmd1.Parameters.AddWithValue("@MachineNumber", machineNumber);
                cmd1.Parameters.AddWithValue("@DeviceType", deviceType);
                cmd1.Parameters.AddWithValue("@SerialNumber", serialNumber);
                cmd1.Parameters.AddWithValue("@Shift", shift);
                cmd1.Parameters.AddWithValue("@Start", start);
                cmd1.Parameters.AddWithValue("@Finish", finish);
                cmd1.Parameters.AddWithValue("@TestJig", testJig);
                cmd1.Parameters.AddWithValue("@TempSetting", tempSettingH);
                cmd1.Parameters.AddWithValue("@Gs_Reading", gs_ReadingH);
                cmd1.Parameters.AddWithValue("@Dut1_Reading", dut1_ReadingH);
                cmd1.Parameters.AddWithValue("@Dut2_Reading", dut2_ReadingH);
                cmd1.Parameters.AddWithValue("@Dut3_Reading", dut3_ReadingH);
                cmd1.Parameters.AddWithValue("@Avg_Reading", avg_ReadingH);
                cmd1.Parameters.AddWithValue("@Dv_Reading", dv_ReadingH);
                cmd1.Parameters.AddWithValue("@Range_Reading", range_ReadingH);
                cmd1.Parameters.AddWithValue("@Mc_OffsetOld", mc_OffsetOldH);
                cmd1.Parameters.AddWithValue("@Mc_OffsetNew", mc_OffsetNewH);
                cmd1.Parameters.AddWithValue("@BuyoffResult", buy0ffResultH);
                cmd1.Parameters.AddWithValue("@DoneBy", doneBy);

                string sql2 = "INSERT INTO TESTER_APPLICATION_TUBE (YEAR, DATE, MACHINE_NUMBER, DEVICE_TYPE, SERIAL_NUMBER, SHIFT, START, FINISH, TEST_JIG, BB_TEMP_SETTING, GS_READING, DUT1_READING, DUT2_READING, DUT3_READING, AVG_READING, DV_READING, RANGE_READING, MC_OFFSET_OLD, MC_OFFSET_NEW, BUYOFF_RESULT, DONE_BY) VALUES (@Year, @Date, @MachineNumber, @DeviceType, @SerialNumber, @Shift, @Start, @Finish, @TestJig, @TempSetting, @Gs_Reading, @Dut1_Reading, @Dut2_Reading, @Dut3_Reading, @Avg_Reading, @Dv_Reading, @Range_Reading, @Mc_OffsetOld, @Mc_OffsetNew, @BuyoffResult, @DoneBy)";
                SqlCommand cmd2 = new SqlCommand(sql2, conn);

                // Add parameters
                cmd2.Parameters.AddWithValue("@Year", year);
                cmd2.Parameters.AddWithValue("@Date", date);
                cmd2.Parameters.AddWithValue("@MachineNumber", machineNumber);
                cmd2.Parameters.AddWithValue("@DeviceType", deviceType);
                cmd2.Parameters.AddWithValue("@SerialNumber", serialNumber);
                cmd2.Parameters.AddWithValue("@Shift", shift);
                cmd2.Parameters.AddWithValue("@Start", start);
                cmd2.Parameters.AddWithValue("@Finish", finish);
                cmd2.Parameters.AddWithValue("@TestJig", testJig);
                cmd2.Parameters.AddWithValue("@TempSetting", tempSettingL);
                cmd2.Parameters.AddWithValue("@Gs_Reading", gs_ReadingL);
                cmd2.Parameters.AddWithValue("@Dut1_Reading", dut1_ReadingL);
                cmd2.Parameters.AddWithValue("@Dut2_Reading", dut2_ReadingL);
                cmd2.Parameters.AddWithValue("@Dut3_Reading", dut3_ReadingL);
                cmd2.Parameters.AddWithValue("@Avg_Reading", avg_ReadingL);
                cmd2.Parameters.AddWithValue("@Dv_Reading", dv_ReadingL);
                cmd2.Parameters.AddWithValue("@Range_Reading", range_ReadingL);
                cmd2.Parameters.AddWithValue("@Mc_OffsetOld", mc_OffsetOldL);
                cmd2.Parameters.AddWithValue("@Mc_OffsetNew", mc_OffsetNewL);
                cmd2.Parameters.AddWithValue("@BuyoffResult", buy0ffResultL);
                cmd2.Parameters.AddWithValue("@DoneBy", doneBy);

                conn.Open();
                int rows1 = cmd1.ExecuteNonQuery();
                if (rows1 > 0)
                {
                    isSuccess = true;
                }
                else
                {
                    isSuccess = false;
                }

                int rows2 = cmd2.ExecuteNonQuery();
                if (rows2 > 0)
                {
                    isSuccess = true;
                }
                else
                {
                    isSuccess = false;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("An error occurred: " + ex.Message);
            }
            finally
            {
                conn.Close();
            }
            return isSuccess;
        }

        public bool InsertHigh(SqlConn data)
        {
            bool isSuccess = false;
            SqlConnection conn = new SqlConnection(connectionString);
            try
            {
                //string sql = "INSERT INTO TESTER_APPLICATION_TUBE (Device, Machine, Date, Shift, Temp, Jig, SerialName) VALUES(@Device, @Machine, @Date, @Shift, @Temp, @Jig, @SerialName)";
                string sql1 = "INSERT INTO TESTER_APPLICATION_TUBE (YEAR, DATE, MACHINE_NUMBER, DEVICE_TYPE, SERIAL_NUMBER, SHIFT, START, FINISH, TEST_JIG, BB_TEMP_SETTING, GS_READING, DUT1_READING, DUT2_READING, DUT3_READING, AVG_READING, DV_READING, RANGE_READING, MC_OFFSET_OLD, MC_OFFSET_NEW, BUYOFF_RESULT, DONE_BY) VALUES (@Year, @Date, @MachineNumber, @DeviceType, @SerialNumber, @Shift, @Start, @Finish, @TestJig, @TempSetting, @Gs_Reading, @Dut1_Reading, @Dut2_Reading, @Dut3_Reading, @Avg_Reading, @Dv_Reading, @Range_Reading, @Mc_OffsetOld, @Mc_OffsetNew, @BuyoffResult, @DoneBy)";
                SqlCommand cmd1 = new SqlCommand(sql1, conn);

                // Add parameters
                cmd1.Parameters.AddWithValue("@Year", year);
                cmd1.Parameters.AddWithValue("@Date", date);
                cmd1.Parameters.AddWithValue("@MachineNumber", machineNumber);
                cmd1.Parameters.AddWithValue("@DeviceType", deviceType);
                cmd1.Parameters.AddWithValue("@SerialNumber", serialNumber);
                cmd1.Parameters.AddWithValue("@Shift", shift);
                cmd1.Parameters.AddWithValue("@Start", start);
                cmd1.Parameters.AddWithValue("@Finish", finish);
                cmd1.Parameters.AddWithValue("@TestJig", testJig);
                cmd1.Parameters.AddWithValue("@TempSetting", tempSettingH);
                cmd1.Parameters.AddWithValue("@Gs_Reading", gs_ReadingH);
                cmd1.Parameters.AddWithValue("@Dut1_Reading", dut1_ReadingH);
                cmd1.Parameters.AddWithValue("@Dut2_Reading", dut2_ReadingH);
                cmd1.Parameters.AddWithValue("@Dut3_Reading", dut3_ReadingH);
                cmd1.Parameters.AddWithValue("@Avg_Reading", avg_ReadingH);
                cmd1.Parameters.AddWithValue("@Dv_Reading", dv_ReadingH);
                cmd1.Parameters.AddWithValue("@Range_Reading", range_ReadingH);
                cmd1.Parameters.AddWithValue("@Mc_OffsetOld", mc_OffsetOldH);
                cmd1.Parameters.AddWithValue("@Mc_OffsetNew", mc_OffsetNewH);
                cmd1.Parameters.AddWithValue("@BuyoffResult", buy0ffResultH);
                cmd1.Parameters.AddWithValue("@DoneBy", doneBy);

                conn.Open();
                int rows1 = cmd1.ExecuteNonQuery();
                if (rows1 > 0)
                {
                    isSuccess = true;
                }
                else
                {
                    isSuccess = false;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("An error occurred: " + ex.Message);
            }
            finally
            {
                conn.Close();
            }
            return isSuccess;
        }

        public bool InsertLow(SqlConn data)
        {
            bool isSuccess = false;
            SqlConnection conn = new SqlConnection(connectionString);
            try
            {
                string sql2 = "INSERT INTO TESTER_APPLICATION_TUBE (YEAR, DATE, MACHINE_NUMBER, DEVICE_TYPE, SERIAL_NUMBER, SHIFT, START, FINISH, TEST_JIG, BB_TEMP_SETTING, GS_READING, DUT1_READING, DUT2_READING, DUT3_READING, AVG_READING, DV_READING, RANGE_READING, MC_OFFSET_OLD, MC_OFFSET_NEW, BUYOFF_RESULT, DONE_BY) VALUES (@Year, @Date, @MachineNumber, @DeviceType, @SerialNumber, @Shift, @Start, @Finish, @TestJig, @TempSetting, @Gs_Reading, @Dut1_Reading, @Dut2_Reading, @Dut3_Reading, @Avg_Reading, @Dv_Reading, @Range_Reading, @Mc_OffsetOld, @Mc_OffsetNew, @BuyoffResult, @DoneBy)";
                SqlCommand cmd2 = new SqlCommand(sql2, conn);

                // Add parameters
                cmd2.Parameters.AddWithValue("@Year", year);
                cmd2.Parameters.AddWithValue("@Date", date);
                cmd2.Parameters.AddWithValue("@MachineNumber", machineNumber);
                cmd2.Parameters.AddWithValue("@DeviceType", deviceType);
                cmd2.Parameters.AddWithValue("@SerialNumber", serialNumber);
                cmd2.Parameters.AddWithValue("@Shift", shift);
                cmd2.Parameters.AddWithValue("@Start", start);
                cmd2.Parameters.AddWithValue("@Finish", finish);
                cmd2.Parameters.AddWithValue("@TestJig", testJig);
                cmd2.Parameters.AddWithValue("@TempSetting", tempSettingL);
                cmd2.Parameters.AddWithValue("@Gs_Reading", gs_ReadingL);
                cmd2.Parameters.AddWithValue("@Dut1_Reading", dut1_ReadingL);
                cmd2.Parameters.AddWithValue("@Dut2_Reading", dut2_ReadingL);
                cmd2.Parameters.AddWithValue("@Dut3_Reading", dut3_ReadingL);
                cmd2.Parameters.AddWithValue("@Avg_Reading", avg_ReadingL);
                cmd2.Parameters.AddWithValue("@Dv_Reading", dv_ReadingL);
                cmd2.Parameters.AddWithValue("@Range_Reading", range_ReadingL);
                cmd2.Parameters.AddWithValue("@Mc_OffsetOld", mc_OffsetOldL);
                cmd2.Parameters.AddWithValue("@Mc_OffsetNew", mc_OffsetNewL);
                cmd2.Parameters.AddWithValue("@BuyoffResult", buy0ffResultL);
                cmd2.Parameters.AddWithValue("@DoneBy", doneBy);

                conn.Open();
                int rows2 = cmd2.ExecuteNonQuery();
                if (rows2 > 0)
                {
                    isSuccess = true;
                }
                else
                {
                    isSuccess = false;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("An error occurred: " + ex.Message);
            }
            finally
            {
                conn.Close();
            }
            return isSuccess;
        }

        public void Insertdata()
        {
            try
            {
                // Connection String
                string connectionString = "Data Source=etcbtmsql01\\digitmont;Initial Catalog=dbDigitmontDetection;User ID=Metis;Password=Metis123";

                // Create a new DataTable
                DataTable dataTable = new DataTable();
               
                // Add columns to the DataTable
                dataTable.Columns.Add("YEAR", typeof(string));
                dataTable.Columns.Add("DATE", typeof(string));
                dataTable.Columns.Add("MACHINE_NUMBER", typeof(string));
                dataTable.Columns.Add("DEVICE_TYPE", typeof(string));
                dataTable.Columns.Add("SERIAL_NUMBER", typeof(string));
                dataTable.Columns.Add("SHIFT", typeof(string));
                dataTable.Columns.Add("START", typeof(string));
                dataTable.Columns.Add("FINISH", typeof(string));
                dataTable.Columns.Add("TEST_JIG", typeof(string));
                dataTable.Columns.Add("BB_TEMP_SETTING", typeof(string));
                dataTable.Columns.Add("GS_READING", typeof(int));
                dataTable.Columns.Add("DUT1_READING", typeof(int));
                dataTable.Columns.Add("DUT2_READING", typeof(int));
                dataTable.Columns.Add("DUT3_READING", typeof(int));
                dataTable.Columns.Add("AVG_READING", typeof(int));
                dataTable.Columns.Add("DV_READING", typeof(int));
                dataTable.Columns.Add("RANGE_READING", typeof(int));
                dataTable.Columns.Add("MC_OFFSET_OLD", typeof(int));
                dataTable.Columns.Add("MC_OFFSET_NEW", typeof(int));
                dataTable.Columns.Add("BUYOFF_RESULT", typeof(string));
                dataTable.Columns.Add("DONE_BY", typeof(string));
                
                // Add rows to DataTable
                DataRow row = dataTable.NewRow();
                row["YEAR"] = year;
                row["DATE"] = date;
                row["MACHINE_NUMBER"] = machineNumber;
                row["DEVICE_TYPE"] = deviceType;
                row["SERIAL_NUMBER"] = serialNumber;
                row["SHIFT"] = shift;
                row["START"] = start;
                row["FINISH"] = finish;
                row["TEST_JIG"] = testJig;
                row["BB_TEMP_SETTING"] = tempSettingH;
                row["GS_READING"] = gs_ReadingH;
                row["DUT1_READING"] = dut1_ReadingH;
                row["DUT2_READING"] = dut2_ReadingH;
                row["DUT3_READING"] = dut3_ReadingH;
                row["AVG_READING"] = avg_ReadingH;
                row["DV_READING"] = dv_ReadingH;
                row["RANGE_READING"] = range_ReadingH;
                row["MC_OFFSET_OLD"] = mc_OffsetOldH;
                row["MC_OFFSET_NEW"] = mc_OffsetNewH;
                row["BUYOFF_RESULT"] = buy0ffResultH;
                row["DONE_BY"] = doneBy;

                dataTable.Rows.Add(row);

                // Create a new SqlBulkCopy object
                using (SqlBulkCopy bulkCopy = new SqlBulkCopy(connectionString))
                {
                    // Set the destination table name
                    bulkCopy.DestinationTableName = "TESTER_APPLICATION_TUBE";

                    // Write the data to the database
                    bulkCopy.WriteToServer(dataTable);
                }

                //Console.WriteLine("Data inserted successfully.");
            }
            catch (Exception ex)
            {
                MessageBox.Show("An error occurred: " + ex.Message);
            }
        }

        //Method to Update data
        public bool Update(SqlConn data)
        {
            bool isSuccess = false;
            SqlConnection conn = new SqlConnection();
            try
            {
                string sql = "UPDATE TESTER_APPLICATION_TUBE SET Device=@Device Machine=@Machine, Date=@Date, Shift=@Shift, Temp=@Temp, Jig=@Jig, SerialName=@SerialName";
                SqlCommand cmd = new SqlCommand(sql, conn);

                // Add parameters
                cmd.Parameters.AddWithValue("@Year", year);
                cmd.Parameters.AddWithValue("@Date", date);
                cmd.Parameters.AddWithValue("@MachineNumber", machineNumber);

                conn.Open();
                int rows = cmd.ExecuteNonQuery();
                if (rows > 0)
                {
                    isSuccess = true;
                }
                else
                {
                    isSuccess = false;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("An error occurred: " + ex.Message);
            }
            finally
            {
                conn.Close();
            }
            return isSuccess;
        }

        public bool Delete(SqlConn data)
        {
            bool isSuccess = false;
            SqlConnection conn = new SqlConnection(connectionString);
            try
            {
                string sql = "DELETE FROM TESTER_APPLICATION_TUBE WHERE MACHINE_NUMBER=@MachineNumber OR DEVICE_TYPE=@DeviceType";
                SqlCommand cmd = new SqlCommand(sql, conn);

                // Add parameters
                cmd.Parameters.AddWithValue("@MachineNumber", machineNumber);
                cmd.Parameters.AddWithValue("@DeviceType", deviceType);
                conn.Open();
                int rows = cmd.ExecuteNonQuery();
                if (rows > 0)
                {
                    isSuccess = true;
                }
                else
                {
                    isSuccess = false;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("An error occurred: " + ex.Message);
            }
            finally
            {
                conn.Close();
            }
            return isSuccess;
        }
    }
}
