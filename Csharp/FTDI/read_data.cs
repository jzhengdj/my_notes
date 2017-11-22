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

//using System.Runtime.Remoting.Metadata.W3cXsd2001;



namespace testUI
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            GetVirtuCommPort();
        }

        private void alt_open_btn_Click(object sender, EventArgs e)
        {
            
            getMsg(sRN_DeviceIdent_BA, ref DeviceIdent);
            //byte[] readData = new byte[sRN_DeviceIdent_RES_LENGTH];
            //receiveMsg(ref readData, sRN_DeviceIdent_RES_LENGTH);
            //listBox2.Items.Add(sRN_DeviceIdent_RES_LENGTH);
        }


        //***************  command list  ******************
        // standard message header
        static readonly byte[] STX_BA = new byte [] { 2, 2, 2, 2, 0, 0, 0 };


        // command data field
        //static readonly int[] DeviceIdent_DataLength = new int[] { 0, 0, 1 };


        byte[] sRN_DeviceIdent_BA = getReadCMD_BA("DeviceIdent");

        // const uint sRN_DeviceIdent_RES_LENGTH = 44;
        //const string sRN_DeviceIdent = "020202020000001073524E204465766963654964656E742005";



        public struct ColaBFormat
        {
            public int fieldLength;
            public string fieldName;
            public string valueString;
            public int valueInt;

        };

        public ColaBFormat[] DeviceIdent = new ColaBFormat[]
        {
            new ColaBFormat() { fieldLength = 0, fieldName = "Name" },
            new ColaBFormat() { fieldLength = 0, fieldName = "Version" }
        };




        //*************************************************




        // Create new instance of the FTDI device class
        FTDI ftdi_handle = new FTDI();

        FTDI.FT_STATUS ftStatus = FTDI.FT_STATUS.FT_OK;
        UInt32 ftdiDeviceCount;



        public void getMsg(byte[] dataToSend, ref ColaBFormat[] dataFormat)
        {
            sendMsg(dataToSend);

            // receive the first 8 byte first (header)
            uint header_length = 8;
            byte[] header = new byte[header_length];
            receiveMsg(ref header, header_length);

            // receive the meat of the message
            uint messageLength = (uint)header[header_length - 1];
            byte[] message = new byte[messageLength];
            receiveMsg(ref message, messageLength);

            

            // If the system architecture is little-endian (that is, little end first),
            // reverse the byte array.
            //if (BitConverter.IsLittleEndian)
            //    Array.Reverse(message);

            // get populate the data
            int pos = (int)(dataToSend.Length - header_length - 1);
            for (int i = 0; i < dataFormat.Length; i++)
            {
                int len = dataFormat[i].fieldLength;
                if (len == 0)
                {
                    len = BitConverter.ToInt32(message, pos+1);
                    listBox2.Items.Add(pos);
                    listBox2.Items.Add(message[pos]);
                    listBox2.Items.Add(message[pos+1]);
                    listBox2.Items.Add(len);
                    pos += 2;
                    dataFormat[i].valueInt = 0;
                    // todo: convert the data to string format and store in the valueString.
                }
                    


            }

            // purge the rest of the message
            ftStatus = ftdi_handle.Purge(FTDI.FT_PURGE.FT_PURGE_RX);
            if (ftStatus != FTDI.FT_STATUS.FT_OK)
            {
                listBox1.Items.Add("Failed to Purge RX buffer (error " + ftStatus.ToString() + ")");
                return;
            }

            return;
        }


        public void sendMsg(byte[] dataToWrite)
        {
            // Set read timeout to 5 seconds, write timeout to infinite
            ftStatus = ftdi_handle.SetTimeouts(5000, 0);
            if (ftStatus != FTDI.FT_STATUS.FT_OK)
            {
                listBox1.Items.Add("Failed to set timeouts (error " + ftStatus.ToString() + ")");
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
            
            listBox2.Items.Add(BitConverter.ToString(readData));

            return;

            /*
            
            //listBox2.Items.Add(GetBytesToString(readData));
            //listBox2.Items.Add(BitConverter.ToString(getReadCMD_BA("DeviceIdent")));

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
                alt_open_btn.Text = "Search for FTDI";
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
            

            return;
        }

        public static byte[] getReadCMD_BA(string cmd)
        {
            // ensure cmd start with "sRN"
            if (cmd[1] != 's') cmd = "sRN "+ cmd;
            // ensure cmd end with space
            if (cmd[cmd.Length - 1] != ' ') cmd += ' ';
            
            byte[] msg_ba = new byte[STX_BA.Length + cmd.Length + 2];
            byte[] cmd_ba = Encoding.ASCII.GetBytes(cmd);

            STX_BA.CopyTo(msg_ba, 0);
            msg_ba[STX_BA.Length] = (byte)cmd.Length;
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

        /*
        public static byte[] getStringToBytes(string value)
        {
            SoapHexBinary shb = SoapHexBinary.Parse(value);
            return shb.Value;
        }

        
        public static string GetBytesToString(byte[] value)
        {
            SoapHexBinary shb = new SoapHexBinary(value);
            return shb.ToString();
        }
        */

    }
}
