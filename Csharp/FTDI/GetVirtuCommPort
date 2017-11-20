FTDI.FT_STATUS ftStatus = FTDI.FT_STATUS.FT_OK;
        FTDI Ftdi1 = new FTDI();
        string comstr;
        UInt32 ftdiDeviceCount;

        private void GetVirtuCommPort()
        {
            ftdiDeviceCount = 0;

            Ftdi1.ResetPort();

            // Determine the number of FTDI devices connected to the machine
            ftStatus = Ftdi1.GetNumberOfDevices(ref ftdiDeviceCount);

            // Check status
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
            ftStatus = Ftdi1.GetDeviceList(ftdiDeviceList);

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
            ftStatus = Ftdi1.OpenBySerialNumber(ftdiDeviceList[0].SerialNumber);
            if (ftStatus != FTDI.FT_STATUS.FT_OK)
            {
                listBox1.Items.Add("Failed to open device (error " + ftStatus.ToString() + ")");
                return;
            }
            Ftdi1.GetCOMPort(out comstr);
            listBox1.Items.Add("Com Port: " + comstr);
            ftStatus = Ftdi1.Close();
            alt_open_btn.Text = "Open VIRTU Found On " + comstr;
            if (ftStatus != FTDI.FT_STATUS.FT_OK)
            {
                listBox1.Items.Add("Failed to open device (error " + ftStatus.ToString() + ")");
                return;
            }

        }
