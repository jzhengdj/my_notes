using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using FTD2XX_NET;
using System.Threading;
//using System.Runtime.Remoting.Metadata.W3cXsd2001;



namespace testUI
{
    public partial class Form1 : Form
    {
        delegate void StringArgReturningVoidDelegate(string text);
        delegate void GridViewArgReturningVoidDelegate(string[] result_array);

        public Form1()
        {
            InitializeComponent();
            customiseForm();
            GetVirtuCommPort();
            sendMsg(LOGIN_PROD_CMD_BA);

            //measureTask = new Task(keepMeasure);
            //measureThread = new Thread(keepMeasure);
            //measureThread.IsBackground = true;
            
        }

        

        #region constants


        //***************  command list  ******************
        // standard message header
        static readonly byte[] STX_BA = new byte [] { 2, 2, 2, 2, 0, 0, 0 };
        static readonly byte[] LOGIN_PROD_CMD = new byte[] {115, 77, 73, 0, 0, 6, 232, 12, 116, 77};

        // command data field
        //static readonly int[] DeviceIdent_DataLength = new int[] { 0, 0, 1 };


        byte[] sRN_DeviceIdent_BA = get_sRN_BA("DeviceIdent");
        byte[] sRN_CidVersion_BA = get_sRN_BA("CidVersion");
        byte[] sRN_MCAD_BA = get_sRN_BA("MCAD");
        byte[] sRN_MCST_BA = get_sRN_BA("MCST");
        byte[] sRN_MCHS_BA = get_sRN_BA("MCHS"); //StatisticOutDone
        

        byte[] LOGIN_PROD_CMD_BA = get_Whole_CMD_BA(LOGIN_PROD_CMD);


        public struct dataField
        {
            public short fieldLength;
            public string fieldName;
            public string fieldType;
            public string valueString;
            public long valueInt;
        };


        public dataField[] DeviceIdent = new dataField[]
        {
            new dataField() { fieldLength = 0, fieldName = "Name" , fieldType = "FlexString"},
            new dataField() { fieldLength = 0, fieldName = "Version", fieldType = "FlexString"}
        };

        public dataField[] CidVersion = new dataField[]
        {
            new dataField() { fieldLength = 2, fieldName = "Major Version", fieldType = "UInt"},
            new dataField() { fieldLength = 2, fieldName = "Minor Version", fieldType = "UInt"},
            new dataField() { fieldLength = 2, fieldName = "Patch Version", fieldType = "UInt"},
            new dataField() { fieldLength = 4, fieldName = "Build Number", fieldType = "UDInt"},
            new dataField() { fieldLength = 1, fieldName = "Version Classifier", fieldType = "Enum8"}
        };

        public dataField[] MCAD = new dataField[]
        {
            new dataField() { fieldLength = 2, fieldName = "Angle", fieldType = "Int"},
            new dataField() { fieldLength = 2, fieldName = "AveragingDepth", fieldType = "UInt"}
        };

        public dataField[] MCST = new dataField[]
        {
            new dataField() { fieldLength = 2, fieldName = "Angle", fieldType = "Int"},
            new dataField() { fieldLength = 2, fieldName = "AveragingDepth", fieldType = "UInt"},
            new dataField() { fieldLength = 4, fieldName = "DistMin", fieldType = "UDInt"},
            new dataField() { fieldLength = 4, fieldName = "DistMax", fieldType = "UDInt" },
            new dataField() { fieldLength = 4, fieldName = "DistSum", fieldType = "UDInt" },
            new dataField() { fieldLength = 4, fieldName = "DistSquareSum", fieldType = "UDInt" },
            new dataField() { fieldLength = 4, fieldName = "DistSquareSumOverflowCounter", fieldType = "UDInt" },

            new dataField() { fieldLength = 4, fieldName = "LevelMin", fieldType = "UDInt" },
            new dataField() { fieldLength = 4, fieldName = "LevelMax", fieldType = "UDInt" },
            new dataField() { fieldLength = 4, fieldName = "LevelSum", fieldType = "UDInt" },
            new dataField() { fieldLength = 4, fieldName = "LevelSquareSum", fieldType = "UDInt" },
            new dataField() { fieldLength = 4, fieldName = "LevelSquareSumOverflowCounter", fieldType = "UDInt" }
        };

        public dataField[] MCHS = new dataField[] //StatisticOutDone
        {
            new dataField() { fieldLength = 1, fieldName = "Written", fieldType = "Bool"}
        };



        //*************************************************
        #endregion

        int sweepAngle_Start = -500, sweepAngle_End = 500;
        int sampleDepth = 5;

        bool flag_stop_measure = true;

        // Create new instance of the FTDI device class
        FTDI ftdi_handle = new FTDI();

        FTDI.FT_STATUS ftStatus = FTDI.FT_STATUS.FT_OK;
        UInt32 ftdiDeviceCount;

        

        public void measure()
        {

            // send angle and depth requirements
            byte[] temp_bytes = get_sWN_BA("MCAD", MCAD);
            verifySend(temp_bytes);

            // set OutDone false to get new reading.
            MCHS[0].valueInt = 0;
            temp_bytes = get_sWN_BA("MCHS", MCHS);
            verifySend(temp_bytes);

            
            // ensure the reading is ready. (keep pulling the flag until it's 1)
            while (MCHS[0].valueInt == 0)
            {
                Thread.Sleep(100);
                getDataFieldValue(sRN_MCHS_BA, ref MCHS);

            }

            getDataFieldValue(sRN_MCST_BA, ref MCST);
            //printAllDataField(MCST); // print measurement data
            
            processData();
            return;
        }

        public void sweepOnce()
        {
            CT_Clear_dataGridView1();

            MCAD[1].valueInt = sampleDepth;
            for (int angle = sweepAngle_Start; angle <= sweepAngle_End; angle+=100)
            {
                MCAD[0].valueInt = angle;
                measure();
            }
        }


        public void keepMeasure()
        {

            CT_Clear_dataGridView1();
            MCAD[0].valueInt = sweepAngle_Start; // initialise the measuring angle

            while (true)
            {
                
                if (flag_stop_measure)
                    break;
                
                if (MCAD[0].valueInt > sweepAngle_End)
                {
                    MCAD[0].valueInt = sweepAngle_Start;
                    //CT_Clear_dataGridView1();
                }
                if (MCAD[1].valueInt != sampleDepth)
                    MCAD[1].valueInt = sampleDepth;
                measure();
                MCAD[0].valueInt += 100; // increase measuring angle
            }
            // perform cleanup if there is any.
        }

        public void processData()
        {
            long N = MCST[1].valueInt;

            long DistSum = MCST[4].valueInt;
            long DistSquareSum = MCST[5].valueInt;

            long RssiSum = MCST[9].valueInt;
            long RssiSquareSum = MCST[10].valueInt;

            //calculate average 
            
            double avg_dist = (double)DistSum / N;
            double avg_rssi = (double)RssiSum / N;

            //calculate standard Deviation
            
            double s_dist = Math.Sqrt((DistSquareSum - DistSum * DistSum / N) / (N - 1));
            double s_rssi = Math.Sqrt((RssiSquareSum - RssiSum * RssiSum / N) / (N - 1));

            string[] data_array = { String.Format("{0:0.00}", MCST[0].valueInt/100),
                String.Format("{0:0.00}", avg_dist),
                String.Format("{0:0.00}", s_dist),
                String.Format("{0:0.00}", avg_rssi),
                String.Format("{0:0.00}", s_rssi)
            };

            CT_Write_dataGridView1(data_array);
            //CT_Write_listBox2(String.Format("Angle: {0}, Avg_dist: {1:0.0}, Deviation_dist: {2:0.00}, Avg_rssi: {3:0.0}, Deviation_rssi: {4:0.00}",
            //MCST[0].valueInt/100, avg_dist, s_dist, avg_rssi, s_rssi));
            return;
        }
        


        #region basic functions

        public void verifySend(byte[] dataToSend)
        {
            sendMsg(dataToSend);

            // receive the first 8 byte first (header)
            uint header_length = 8;
            byte[] header = new byte[header_length];
            receiveMsg(ref header, header_length);

            // receive the meat of the message, call it message
            uint messageLength = (uint)header[header_length - 1] + 1; // +1 to include CRC byte
            byte[] message = new byte[messageLength];
            receiveMsg(ref message, messageLength);



            // verify the receive message is valid.
            if (getCRC(message) != 0)
            {
                listBox1.Items.Add("Failed CRC check");
                return;
            }

            // verify the receive message is acknowledged.
            if (message[2] != (byte)65)
            {
                listBox2.Items.Add("!!! Not Acknowledged !!!");
                return;
            }

            /*
            // print out the data for debugging
            listBox2.Items.Add(BitConverter.ToString(dataToSend));
            */

            return;

        }

        public void getDataFieldValue(byte[] dataToSend, ref dataField[] dataFormat)
        {
            sendMsg(dataToSend);

            // receive the first 8 byte first (header)
            uint header_length = 8;
            byte[] header = new byte[header_length];
            receiveMsg(ref header, header_length);

            // receive the meat of the message, call it message
            uint messageLength = (uint)header[header_length - 1] + 1; // +1 to include CRC byte
            byte[] message = new byte[messageLength];
            receiveMsg(ref message, messageLength);



            // verify the receive message is valid.
            if (getCRC(message) != 0)
            {
                listBox2.Items.Add("Failed CRC check");
                return;
            }

            

            // verify the receive message is acknowledged.
            if (message[2] != (byte)65)
            {
                listBox2.Items.Add("!!! Not Acknowledged !!!");
                return;
            }

            // get populate the data
            int pos = (int)(dataToSend.Length - header_length - 1);
            
            

            for (int i = 0; i < dataFormat.Length; i++)
            {
                int len = dataFormat[i].fieldLength;
                if (len == 0)
                {
                    // get the content of the flex string and store in the valueString, valueInt don't care.
                    len = (int)message[pos]*256+(int)message[pos+1];

                    if (len == 0) 
                        continue; // no content in the flex string field

                    byte[] sub_message = new byte[len];
                    Array.Copy(message, pos+2, sub_message, 0, len);
                    dataFormat[i].valueString = Encoding.ASCII.GetString(sub_message);

                    // adjust pos
                    pos += len + 2;
                   
                    
                } else {
                    // get the number value. populate valueInt only, valueString don't care.
                    long temp_int = 0;
                    
                    for (int j = 0; j < len; j++)

                        //temp_int = temp_int * 256 + (int)message[pos+j]; // true if all data are uint.
                        if (dataFormat[i].fieldType == "Int")
                            temp_int += (short)(message[pos + j] << 8 * (len - j - 1));
                        else
                            temp_int += (long)(message[pos + j] << 8 * (len - j - 1)); // ignore uint for now, take everything as int.

                    dataFormat[i].valueInt = temp_int;

                    //adjust pos
                    pos += len;
                    
                }
            }


            
        }

        public void printAllDataField(dataField[] dataFormat)
        {
            string dataToShow = "";

            for (int i = 0; i < dataFormat.Length; i++)
            {
                if (dataFormat[i].fieldLength == 0)
                {
                    // store the flex string.
                    dataToShow += String.Format("{0}: {1}    ", dataFormat[i].fieldName, dataFormat[i].valueString);
                }
                else
                {
                    // store the number value.
                    dataToShow += String.Format("{0}: {1}  ", dataFormat[i].fieldName, dataFormat[i].valueInt);
                }
            }

            // print out for debug.
            //listBox2.SelectedIndex = listBox2.Items.Count - 1;
            //listBox2.SelectedIndex = -1;
            CT_Write_listBox2(dataToShow);
            //listBox2.Items.Add(dataToShow);
            return;
        }


        public void sendMsg(byte[] dataToWrite)
        {

            // purge the rest of the message
            ftStatus = ftdi_handle.Purge(FTDI.FT_PURGE.FT_PURGE_RX);
            if (ftStatus != FTDI.FT_STATUS.FT_OK)
            {
                listBox1.Items.Add("Failed to Purge RX buffer (error " + ftStatus.ToString() + ")");
                return;
            }

            // Write string data to the device

            //byte[] dataToWrite = sRN_DeviceIdent_BA;
            UInt32 numBytesWritten = 0;
            // Note that the Write method is overloaded, so can write string or byte array data
            ftStatus = ftdi_handle.Write(dataToWrite, dataToWrite.Length, ref numBytesWritten);
            if (ftStatus != FTDI.FT_STATUS.FT_OK)
            {
                // Wait for a key press
                listBox1.Items.Add("Failed to write to device (error " + ftStatus.ToString() + ")");
                return;
            }

            //listBox2.Items.Add(BitConverter.ToString(dataToWrite));
            return;
        }

        public void receiveMsg(ref byte[] readData, uint res_length) //
        { 
            UInt32 numBytesAvailable = 0;
            UInt32 numberOfTries = 0;
            do
            {
                ftStatus = ftdi_handle.GetRxBytesAvailable(ref numBytesAvailable);
                if (ftStatus != FTDI.FT_STATUS.FT_OK)
                {
                    // Wait for a key press
                    listBox1.Items.Add("Failed to get number of bytes available to read (error " + ftStatus.ToString() + ")");
                    return;
                }
                System.Threading.Thread.Sleep(5);
                numberOfTries++;
            } while (numBytesAvailable < res_length && numberOfTries < 100);

            if (numberOfTries > 99)
            {
                listBox1.Items.Add("No response received from the Device. Please check connection.");
                return;
            }
            
            UInt32 numBytesRead = 0;
            // Note that the Read method is overloaded, so can read string or byte array data
            ftStatus = ftdi_handle.Read(readData, res_length, ref numBytesRead);
            if (ftStatus != FTDI.FT_STATUS.FT_OK)
            {
                // Wait for a key press
                listBox1.Items.Add("Failed to read data (error " + ftStatus.ToString() + ")");
                return;
            }
            
            //listBox2.Items.Add(BitConverter.ToString(readData));

            return;

            /*
            
            //listBox2.Items.Add(GetBytesToString(readData));
            //listBox2.Items.Add(BitConverter.ToString(get_sRN_BA("DeviceIdent")));

            // prepare the byte for the data
            string value = "";
            foreach (byte byt in readData)
                value += String.Format("{0} ", byt);
            listBox2.Items.Add(value);



            // convert the byte array to ASCII code, HOW???
            string[] hexValuesSplit = BitConverter.ToString(readData).Split('-');
            string response = "test";
            foreach (String hex in hexValuesSplit)
            {
                // Convert the number expressed in base-16 to an integer.
                int value = Convert.ToInt32(hex, 16);
                // Get the character corresponding to the integral value.

                listBox2.Items.Add(((char)value).ToString());
                response += ((char)value).ToString();
                listBox2.Items.Add(response);
                //Char.ConvertFromUtf32(value);
                //listBox2.Items.Add(Char.ConvertFromUtf32(value).ToString());
            }
            //listBox2.Items.Add(response);




            // Check the amount of data available to read
            // In this case we know how much data we are expecting, 
            // so wait until we have all of the bytes we have sent.
            UInt32 numBytesAvailable = 0;
            do
            {
                ftStatus = ftdi_handle.GetRxBytesAvailable(ref numBytesAvailable);
                if (ftStatus != FTDI.FT_STATUS.FT_OK)
                {
                    // Wait for a key press
                    listBox1.Items.Add("Failed to get number of bytes available to read (error " + ftStatus.ToString() + ")");
                    return;
                }
            } while (numBytesAvailable < sRN_DeviceIdent_RES_LENGTH);

            // Now that we have the amount of data we want available, read it
            string readData;
            UInt32 numBytesRead = 0;
            // Note that the Read method is overloaded, so can read string or byte array data
            ftStatus = ftdi_handle.Read(out readData, numBytesAvailable, ref numBytesRead);
            if (ftStatus != FTDI.FT_STATUS.FT_OK)
            {
                // Wait for a key press
                listBox1.Items.Add("Failed to read data (error " + ftStatus.ToString() + ")");
                return;
            }
            listBox2.Items.Add(readData);

            // Close our device
            ftStatus = ftdi_handle.Close();
            */


        }

        private void GetVirtuCommPort()
        {
            

            ftdi_handle.ResetPort();

            // Determine the number of FTDI devices connected to the machine
            ftStatus = ftdi_handle.GetNumberOfDevices(ref ftdiDeviceCount);
            
            if (ftStatus == FTDI.FT_STATUS.FT_OK)
            {
                listBox1.Items.Add("Number of FTDI devices: " + ftdiDeviceCount.ToString());
            }
            else
            {
                return;
            }

            if (ftdiDeviceCount == 0)
            {
                button1.Text = "Search for FTDI";
                listBox1.Items.Add("No FTDI device found!");
                return;
            }

            // Allocate storage for device info list
            FTDI.FT_DEVICE_INFO_NODE[] ftdiDeviceList = new FTDI.FT_DEVICE_INFO_NODE[ftdiDeviceCount];

            // Populate our device list
            ftStatus = ftdi_handle.GetDeviceList(ftdiDeviceList);

            if (ftStatus == FTDI.FT_STATUS.FT_OK)
            {
                for (UInt32 i = 0; i < ftdiDeviceCount; i++)
                {
                    listBox1.Items.Add("Device Index: " + i.ToString());
                    listBox1.Items.Add("Flags: " + String.Format("{0:x}", ftdiDeviceList[i].Flags));
                    listBox1.Items.Add("Type: " + ftdiDeviceList[i].Type.ToString());
                    listBox1.Items.Add("ID: " + String.Format("{0:x}", ftdiDeviceList[i].ID));
                    listBox1.Items.Add("Location ID: " + String.Format("{0:x}", ftdiDeviceList[i].LocId));
                    listBox1.Items.Add("Serial Number: " + ftdiDeviceList[i].SerialNumber.ToString());
                    listBox1.Items.Add("Description: " + ftdiDeviceList[i].Description.ToString());
                    listBox1.Items.Add("");
                }
            }

            // Open first device in our list by serial number
            ftStatus = ftdi_handle.OpenBySerialNumber(ftdiDeviceList[0].SerialNumber);
            if (ftStatus != FTDI.FT_STATUS.FT_OK)
            {
                listBox1.Items.Add("Failed to open device (error " + ftStatus.ToString() + ")");
                return;
            }

            // Set up device data parameters
            // Set Baud rate to 230400
            ftStatus = ftdi_handle.SetBaudRate(230400);
            if (ftStatus != FTDI.FT_STATUS.FT_OK)
            {
                // Wait for a key press
                listBox1.Items.Add("Failed to set Baud rate (error " + ftStatus.ToString() + ")");
                return;
            }

            // Set data characteristics - Data bits, Stop bits, Parity
            ftStatus = ftdi_handle.SetDataCharacteristics(FTDI.FT_DATA_BITS.FT_BITS_8, FTDI.FT_STOP_BITS.FT_STOP_BITS_1, FTDI.FT_PARITY.FT_PARITY_EVEN);
            if (ftStatus != FTDI.FT_STATUS.FT_OK)
            {
                listBox1.Items.Add("Failed to set data characteristics (error " + ftStatus.ToString() + ")");
                return;
            }

            
            // Set flow control - set flow control to NONE
            ftStatus = ftdi_handle.SetFlowControl(FTDI.FT_FLOW_CONTROL.FT_FLOW_NONE, 0, 0);
            if (ftStatus != FTDI.FT_STATUS.FT_OK)
            {
                listBox1.Items.Add("Failed to set flow control (error " + ftStatus.ToString() + ")");
                return;
            }

            // Set read timeout to 5 seconds, write timeout to infinite
            ftStatus = ftdi_handle.SetTimeouts(5000, 0);
            if (ftStatus != FTDI.FT_STATUS.FT_OK)
            {
                listBox1.Items.Add("Failed to set timeouts (error " + ftStatus.ToString() + ")");
                return;
            }

            return;
        }

        public static byte[] get_sRN_BA(string cmd)
        {
            // ensure cmd start with "sRN"
            if (cmd[0] != 's') cmd = "sRN "+ cmd;
            // ensure cmd end with space
            if (cmd[cmd.Length - 1] != ' ') cmd += ' ';
            
            byte[] cmd_ba = Encoding.ASCII.GetBytes(cmd);
            return get_Whole_CMD_BA(cmd_ba);
        }

        public static byte[] get_sWN_BA(string cmd, dataField[] parameters)
        {
            // ensure cmd start with "sWN"
            if (cmd[0] != 's') cmd = "sWN " + cmd;
            // ensure cmd end with space
            if (cmd[cmd.Length - 1] != ' ') cmd += ' ';

            int len = cmd.Length;
            for (int i = 0; i < parameters.Length; i++)
            {
                if (parameters[i].fieldLength != 0)
                    len += parameters[i].fieldLength;
                else
                    len += parameters[i].valueString.Length + 2;
            }

            byte[] cmd_ba = new byte[len];
            Encoding.ASCII.GetBytes(cmd).CopyTo(cmd_ba, 0);
            int pos = cmd.Length;

            for (int i = 0; i < parameters.Length; i++)
            {
                len = parameters[i].fieldLength;
                if (len != 0)
                {
                    byte[] temp_bytes = new byte[len];
                    for (int j = 0; j < len; j++)
                        temp_bytes[j] = (byte)(parameters[i].valueInt >> 8 * (len - j - 1));
                    temp_bytes.CopyTo(cmd_ba, pos);
                    pos += len;
                }
                else
                {
                    len = parameters[i].valueString.Length;
                    cmd_ba[pos] = (byte)(len >> 8);
                    cmd_ba[pos + 1] = (byte)len;
                    Encoding.ASCII.GetBytes(parameters[i].valueString).CopyTo(cmd_ba, pos + 2);
                    pos += len + 2;
                }
            }
            

            return get_Whole_CMD_BA(cmd_ba);
        }

        public static byte[] get_Whole_CMD_BA(byte[] cmd_ba)
        {
            byte[] msg_ba = new byte[STX_BA.Length + cmd_ba.Length + 2];

            STX_BA.CopyTo(msg_ba, 0);
            msg_ba[STX_BA.Length] = (byte)cmd_ba.Length;
            cmd_ba.CopyTo(msg_ba, STX_BA.Length + 1);
            msg_ba[STX_BA.Length + cmd_ba.Length + 1] = getCRC(cmd_ba);

            return msg_ba;

        }

        public static byte getCRC(byte[] value)
        {
            byte crc = 0;
            foreach (byte byt in value)
                crc ^= byt;
            return crc;
        }

        


        private void CT_Write_listBox2(string text)
        {
            // InvokeRequired required compares the thread ID of the  
            // calling thread to the thread ID of the creating thread.  
            // If these threads are different, it returns true.  
            if (this.listBox2.InvokeRequired)
            {
                StringArgReturningVoidDelegate d = new StringArgReturningVoidDelegate(CT_Write_listBox2);
                this.Invoke(d, new object[] { text });
            }
            else
            {
                this.listBox2.Items.Add(text);
            }
        }

        private void CT_Write_dataGridView1(string[] result_array)
        {
            // InvokeRequired required compares the thread ID of the  
            // calling thread to the thread ID of the creating thread.  
            // If these threads are different, it returns true.  
            if (this.dataGridView1.InvokeRequired)
            {
                GridViewArgReturningVoidDelegate d = new GridViewArgReturningVoidDelegate(CT_Write_dataGridView1);
                this.Invoke(d, new object[] { result_array } );
            }
            else
            {
                dataGridView1.Rows.Add(result_array);
            }
        }

        private void CT_Clear_dataGridView1()
        {
            // InvokeRequired required compares the thread ID of the  
            // calling thread to the thread ID of the creating thread.  
            // If these threads are different, it returns true.  
            if (this.dataGridView1.InvokeRequired)
            {
                Action d = new Action(CT_Clear_dataGridView1);
                this.Invoke(d);
            }
            else
            {
                do
                {
                    foreach (DataGridViewRow row in dataGridView1.Rows)
                    {
                        try
                        {
                            dataGridView1.Rows.Remove(row);
                        }
                        catch (Exception) { }
                    }
                } while (dataGridView1.Rows.Count > 1);
            }
        }

        #endregion


        private void button1_Click(object sender, EventArgs e)
        {
            getDataFieldValue(sRN_DeviceIdent_BA, ref DeviceIdent);
            printAllDataField(DeviceIdent);
            //byte[] readData = new byte[sRN_DeviceIdent_RES_LENGTH];
            //receiveMsg(ref readData, sRN_DeviceIdent_RES_LENGTH);
            //listBox2.Items.Add(BitConverter.ToString(sRN_CidVersion_BA));
            getDataFieldValue(sRN_CidVersion_BA, ref CidVersion);
            printAllDataField(CidVersion);

            //getDataFieldValue(sRN_MCST_BA, ref MCST);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            MCAD[0].valueInt = sweepAngle_Start;
            MCAD[1].valueInt = sampleDepth;
            Task taskA = Task.Factory.StartNew(() => measure());
        }

        private void button3_Click(object sender, EventArgs e)
        {
            // change angle, sample depth
            
            Task taskA = Task.Factory.StartNew(() => sweepOnce());
            
            /*
            byte[] temp_bytes = get_sWN_BA("MCAD", MCAD);
            

            verifySend(temp_bytes);
            getDataFieldValue(sRN_MCAD_BA, ref MCAD);

            // set OutDone 
            MCHS[0].valueInt = 0;
            temp_bytes = get_sWN_BA("MCHS", MCHS);

            verifySend(temp_bytes);
            */

            //getDataFieldValue(sRN_MCST_BA, ref MCST);
        }

        private void button4_Click(object sender, EventArgs e)
        {
            
            if (flag_stop_measure)
            {
                flag_stop_measure = false;
                Task taskA = Task.Factory.StartNew(() => keepMeasure());
            }
            else
                flag_stop_measure = true;

        }

        private void button6_Click(object sender, EventArgs e)
        {

            do
            {
                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    try
                    {
                        dataGridView1.Rows.Remove(row);
                    }
                    catch (Exception) { }
                }
            } while (dataGridView1.Rows.Count > 1);

        }

        private void button5_Click(object sender, EventArgs e)
        {

            CT_Clear_dataGridView1();

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            sweepAngle_Start = int.Parse(textBox1.Text) * 100;
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            sweepAngle_End = int.Parse(textBox2.Text) * 100;
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            sampleDepth = int.Parse(textBox3.Text);
        }

        

        
        
    }
    

}
