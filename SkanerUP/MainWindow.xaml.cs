using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Windows;
using System.Collections;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Windows.Forms;
using WIA;
using Microsoft.Win32;

namespace SkanerUP
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {

        DeviceManager deviceManager = new DeviceManager();
        ArrayList scanners = new ArrayList();

        public MainWindow()
        {
            InitializeComponent();

            reloadDevices();

            labelBrightnessValue.Content = sliderBrightness.Value;

            //Console.WriteLine(deviceManager.DeviceInfos[i].Properties["Name"].get_Value());

            
        }

        void reloadDevices()
        {
            foreach (DeviceInfo device in deviceManager.DeviceInfos)
            {
                if (device.Type == WiaDeviceType.ScannerDeviceType)
                {
                    devicesComboBox.Items.Add(device.Properties["Name"].get_Value());
                    scanners.Add(device);
                }
            }

            if (!devicesComboBox.Items.IsEmpty)
            {
                devicesComboBox.SelectedIndex = 0;
            }
        }

        void scan(DeviceInfo device)
        {

            if (textBoxFilePath.Text == "")
            {
                System.Windows.MessageBox.Show("No path choosen");
               
            }
            else
            {
                Cursor = System.Windows.Input.Cursors.Wait;
                buttonScan.IsEnabled = false;
                // Connect to the scanner
                var deviceConnection = device.Connect();
                var scannerItem = deviceConnection.Items[1];

                int resolution = 150;
                int width_pixel = 1250;
                int height_pixel = 1700;
                int color_mode = 1;

                int brightnessPercent = Convert.ToInt16(sliderBrightness.Value);
                AdjustScannerSettings(scannerItem, resolution, 0, 0, width_pixel, height_pixel, 0, brightnessPercent, color_mode);

                var imageFile = (ImageFile)scannerItem.Transfer(FormatID.wiaFormatJPEG);

                string path = textBoxFilePath.Text + "\\scan.jpeg";

                if (File.Exists(path))
                {
                    File.Delete(path);
                }

                // Save image !
                imageFile.SaveFile(path);
                imageScanned.Source = new BitmapImage(new Uri(path));
                Cursor = System.Windows.Input.Cursors.Arrow;
                buttonScan.IsEnabled = true;
            }
            
        }

        private void buttonReload_Click(object sender, RoutedEventArgs e)
        {
            reloadDevices();
        }

        /// <summary>
        /// Adjusts the settings of the scanner with the providen parameters.
        /// </summary>
        /// <param name="scannnerItem">Scanner Item</param>
        /// <param name="scanResolutionDPI">Provide the DPI resolution that should be used e.g 150</param>
        /// <param name="scanStartLeftPixel"></param>
        /// <param name="scanStartTopPixel"></param>
        /// <param name="scanWidthPixels"></param>
        /// <param name="scanHeightPixels"></param>
        /// <param name="brightnessPercents"></param>
        /// <param name="contrastPercents">Modify the contrast percent</param>
        /// <param name="colorMode">Set the color mode</param>
        private static void AdjustScannerSettings(IItem scannnerItem, int scanResolutionDPI, int scanStartLeftPixel, int scanStartTopPixel, int scanWidthPixels, int scanHeightPixels, int brightnessPercents, int contrastPercents, int colorMode)
        {
            const string WIA_SCAN_COLOR_MODE = "6146";
            const string WIA_HORIZONTAL_SCAN_RESOLUTION_DPI = "6147";
            const string WIA_VERTICAL_SCAN_RESOLUTION_DPI = "6148";
            const string WIA_HORIZONTAL_SCAN_START_PIXEL = "6149";
            const string WIA_VERTICAL_SCAN_START_PIXEL = "6150";
            const string WIA_HORIZONTAL_SCAN_SIZE_PIXELS = "6151";
            const string WIA_VERTICAL_SCAN_SIZE_PIXELS = "6152";
            const string WIA_SCAN_BRIGHTNESS_PERCENTS = "6154";
            const string WIA_SCAN_CONTRAST_PERCENTS = "6155";
            SetWIAProperty(scannnerItem.Properties, WIA_HORIZONTAL_SCAN_RESOLUTION_DPI, scanResolutionDPI);
            SetWIAProperty(scannnerItem.Properties, WIA_VERTICAL_SCAN_RESOLUTION_DPI, scanResolutionDPI);
            SetWIAProperty(scannnerItem.Properties, WIA_HORIZONTAL_SCAN_START_PIXEL, scanStartLeftPixel);
            SetWIAProperty(scannnerItem.Properties, WIA_VERTICAL_SCAN_START_PIXEL, scanStartTopPixel);
            SetWIAProperty(scannnerItem.Properties, WIA_HORIZONTAL_SCAN_SIZE_PIXELS, scanWidthPixels);
            SetWIAProperty(scannnerItem.Properties, WIA_VERTICAL_SCAN_SIZE_PIXELS, scanHeightPixels);
            SetWIAProperty(scannnerItem.Properties, WIA_SCAN_BRIGHTNESS_PERCENTS, brightnessPercents);
            SetWIAProperty(scannnerItem.Properties, WIA_SCAN_CONTRAST_PERCENTS, contrastPercents);
            SetWIAProperty(scannnerItem.Properties, WIA_SCAN_COLOR_MODE, colorMode);
        }

        /// <summary>
        /// Modify a WIA property
        /// </summary>
        /// <param name="properties"></param>
        /// <param name="propName"></param>
        /// <param name="propValue"></param>
        private static void SetWIAProperty(IProperties properties, object propName, object propValue)
        {
            Property prop = properties.get_Item(ref propName);
            prop.set_Value(ref propValue);
        }

        private void buttonChooseFile_Click(object sender, RoutedEventArgs e)
        {
            FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog();
            DialogResult result = folderBrowserDialog.ShowDialog();

            if (!string.IsNullOrWhiteSpace(folderBrowserDialog.SelectedPath))
            {
                textBoxFilePath.Text = folderBrowserDialog.SelectedPath;
            }
        }

        private void buttonScan_Click(object sender, RoutedEventArgs e)
        {
            scan(scanners[devicesComboBox.SelectedIndex] as DeviceInfo);
        }

        private void Slider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            StringBuilder stB = new StringBuilder();
            stB.Append(sliderBrightness.Value.ToString());
            stB.Append(" %");
            labelBrightnessValue.Content = stB.ToString();
        }
    }
}
