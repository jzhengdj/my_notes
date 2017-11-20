using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.Remoting.Metadata.W3cXsd2001;
using FTD2XX_NET;





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
            sendCommand(sRN_DeviceIdent_ba);
        }


        //***************  command list  ******************
        const string sRN_DeviceIdent = "020202020000001073524E204465766963654964656E742005";
        byte[] sRN_DeviceIdent_ba = GetStringToBytes(sRN_DeviceIdent);
        const int RES_sRN_DeviceIdent_LENGTH = 20;
        //*************************************************




        FTDI.FT_STATUS ftStatus = FTDI.FT_STATUS.FT_OK;

        // Create new instance of the FTDI device class
        FTDI ftdi_handle = new FTDI();
        UInt32 ftdiDeviceCount;




        
        public void sendCommand(byte[] dataToWrite)
        {
            // Set read timeout to 5 seconds, write timeout to infinite
            ftStatus = ftdi_handle.SetTimeouts(5000, 0);
            if (ftStatus != FTDI.FT_STATUS.FT_OK)
            {
                // Wait for a key press
                listBox1.Items.Add("Failed to set timeouts (error " + ftStatus.ToString() + ")");
                return;
            }
            
            // Write string data to the device
            
            //byte[] dataToWrite = sRN_DeviceIdent_ba;
            UInt32 numBytesWritten = 0;
            // Note that the Write method is overloaded, so can write string or byte array data
            ftStatus = ftdi_handle.Write(dataToWrite, dataToWrite.Length, ref numBytesWritten);
            if (ftStatus != FTDI.FT_STATUS.FT_OK)
            {
                // Wait for a key press
                listBox1.Items.Add("Failed to write to device (error " + ftStatus.ToString() + ")");
                return;
            }
            

            
            /*
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
            } while (numBytesAvailable < RES_sRN_DeviceIdent_LENGTH);

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

            return;
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

        public static byte[] GetStringToBytes(string value)
        {
            SoapHexBinary shb = SoapHexBinary.Parse(value);
            return shb.Value;
        }


    }
}
