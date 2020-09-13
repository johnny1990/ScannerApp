using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using WIA;
using System.IO;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Interop;
using System.Collections.Specialized;
using System.Web;
using System.Security.Cryptography;
using System.Reflection;
using System.Configuration;
using System.Drawing.Imaging;

namespace ScannerApp
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        List<System.Drawing.Image> obj = new List<System.Drawing.Image>();
        List<Scannerdata> objListSD = new List<Scannerdata>();
        int GlobalSquence = 1;
        int count = 0;
        int SequenceTotal = 1;

        #region Construction
        public MainWindow()
        {
            InitializeComponent();

            _deviceId = "";
            _deviceInfo = FindDevice(_deviceId);
            btnPreview.IsEnabled = false;
            btnNext.IsEnabled = false;
        }
        #endregion

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
          
        }

        #region WIA constants  
        public const int WIA_DPS_DOCUMENT_HANDLING_CAPABILITIES = 3086;
        public const int WIA_DPS_DOCUMENT_HANDLING_STATUS = 3087;
        public const int WIA_DPS_DOCUMENT_HANDLING_SELECT = 3088;

        public const string WIA_FORMAT_JPEG = "{B96B3CAE-0728-11D3-9D7B-0000F81EF32E}";

        public const int FEED_READY = 0x00000001;
        public const int FLATBED_READY = 0x00000002;

        public const uint BASE_VAL_WIA_ERROR = 0x80210000;
        public const uint WIA_ERROR_PAPER_EMPTY = BASE_VAL_WIA_ERROR + 3;
        #endregion

        private string _deviceId;
        private DeviceInfo _deviceInfo;
        private Device _device;


        private DeviceInfo FindDevice(string deviceId)
        {
            DeviceManager manager = new DeviceManager();
            foreach (DeviceInfo info in manager.DeviceInfos)
                if (info.DeviceID == deviceId)
                    return info;

            return null;
        }
        //scan
        public List<System.Drawing.Image> ScanPages(int dpi = 150, double width = 8.27, double height = 11.69)
        {
            Item item = _device.Items[1];

            // configure item  
            SetDeviceItemProperty(ref item, 6146, 2); 
            SetDeviceItemProperty(ref item, 6147, dpi); 
            SetDeviceItemProperty(ref item, 6148, dpi); 
            SetDeviceItemProperty(ref item, 6151, (int)(dpi * width));
            SetDeviceItemProperty(ref item, 6152, (int)(dpi * height));  
            SetDeviceItemProperty(ref item, 4104, 8);  

   

            List<System.Drawing.Image> images = GetPagesFromScanner(ScanSource.DocumentFeeder, item);
            if (images.Count == 0)
            {

                MessageBoxResult dialogResult;
                do
                {
                    List<System.Drawing.Image> singlePage = GetPagesFromScanner(ScanSource.Flatbed, item);
                    images.AddRange(singlePage);
                    dialogResult = MessageBox.Show("Do you want to scan another page?", "ScanToEvernote", MessageBoxButton.YesNo, MessageBoxImage.Question);
                }
                while (dialogResult == MessageBoxResult.Yes);
            }
            return images;
        }

        private List<System.Drawing.Image> GetPagesFromScanner(ScanSource source, Item item)
        {
            SetDeviceProperty(ref _device, 3088, (int)source);

            List<System.Drawing.Image> images = new List<System.Drawing.Image>();

            int handlingStatus = GetDeviceProperty(ref _device, WIA_DPS_DOCUMENT_HANDLING_STATUS);
            if ((source == ScanSource.DocumentFeeder && handlingStatus == FEED_READY) || (source == ScanSource.Flatbed && handlingStatus == FLATBED_READY))
            {
                do
                {
                    ImageFile wiaImage = null;
                    try
                    {
                        wiaImage = item.Transfer(WIA_FORMAT_JPEG);
                    }
                    catch (COMException ex)
                    {
                        if ((uint)ex.ErrorCode == WIA_ERROR_PAPER_EMPTY)
                            break;
                        else
                            throw;
                    }

                    if (wiaImage != null)
                    {

                        System.Diagnostics.Trace.WriteLine(String.Format("Image is {0} x {1} pixels", (float)wiaImage.Width / 150, (float)wiaImage.Height / 150));
                        System.Drawing.Image image = ConvertToImage(wiaImage);
                        images.Add(image);
                    }
                }
                while (source == ScanSource.DocumentFeeder);
            }
            return images;
        }

        private static System.Drawing.Image ConvertToImage(ImageFile wiaImage)
        {
            byte[] imageBytes = (byte[])wiaImage.FileData.get_BinaryData();
            MemoryStream ms = new MemoryStream(imageBytes);

            System.Drawing.Image image = System.Drawing.Image.FromStream(ms);

            return image;
        }

        #region device properties  

        private void SetDeviceProperty(ref Device device, int propertyID, int propertyValue)
        {
            foreach (Property p in device.Properties)
            {
                if (p.PropertyID == propertyID)
                {
                    object value = propertyValue;
                    p.set_Value(ref value);
                    break;
                }
            }
        }

        private int GetDeviceProperty(ref Device device, int propertyID)
        {
            int ret = -1;

            foreach (Property p in device.Properties)
            {
                if (p.PropertyID == propertyID)
                {
                    ret = (int)p.get_Value();
                    break;
                }
            }

            return ret;
        }

        private void SetDeviceItemProperty(ref Item item, int propertyID, int propertyValue)
        {
            foreach (Property p in item.Properties)
            {
                if (p.PropertyID == propertyID)
                {
                    object value = propertyValue;
                    p.set_Value(ref value);
                    break;
                }
            }
        }

        private int GetDeviceItemProperty(ref Item item, int propertyID)
        {
            int ret = -1;

            foreach (Property p in item.Properties)
            {
                if (p.PropertyID == propertyID)
                {
                    ret = (int)p.get_Value();
                    break;
                }
            }

            return ret;
        }

        #endregion

        private delegate void UpdateProgressBarDelegate(System.Windows.DependencyProperty dp, Object value);


        private void Button_Click(object sender, RoutedEventArgs e)
        {
 
            btnAttach.IsEnabled = false;
            btnScan.IsEnabled = false;
            pic_scan.Source = null;
            lbl_message.Content = string.Empty;


            ProgressBar1.Minimum = 0;
            ProgressBar1.Maximum = 1000;
            ProgressBar1.Value = 0;


            double value = 0;

            UpdateProgressBarDelegate updatePbDelegate = new UpdateProgressBarDelegate(ProgressBar1.SetValue);


            do
            {
                value += 1;

                Dispatcher.Invoke(updatePbDelegate,
                    System.Windows.Threading.DispatcherPriority.Background,
                    new object[] { ProgressBar.ValueProperty, value });

            }
            while (ProgressBar1.Value != ProgressBar1.Maximum);
            if (ProgressBar1.Value == ProgressBar1.Maximum)
            {
                ProgressBar1.IsEnabled = true;

                obj = ScanPages();

                foreach (System.Drawing.Image aa in obj)
                {
                    Scannerdata objSD = new Scannerdata();
                    objSD.Sequence = SequenceTotal;
                    objSD.ImageObj = aa;

                    objListSD.Add(objSD);
                    SequenceTotal++;
                }

                Scannerdata first = (from row in objListSD where row.Sequence == GlobalSquence select row).FirstOrDefault();

                LoadImage(first.ImageObj);
            }
            if (obj.Count == 1)
            {
                btnPreview.IsEnabled = false;
                btnNext.IsEnabled = false;
                lbl_message.Visibility = Visibility.Visible;
                lbl_message.Content = "Single page scaning Completed.";
            }
            else if (obj.Count > 1)
            {
                btnNext.IsEnabled = true;
                lbl_message.Visibility = Visibility.Visible;
                lbl_message.Content = "Multiple page scaning Completed.";
            }
            btnAttach.IsEnabled = true;
            btnScan.IsEnabled = true;
            ProgressBar1.Value = 0;

        }

        public byte[] LoadImage(System.Drawing.Image loaddata)
        {
            BitmapImage bi = new BitmapImage();

            bi.BeginInit();

            MemoryStream ms = new MemoryStream();
  

            loaddata.Save(ms, ImageFormat.Jpeg);



            ms.Seek(0, SeekOrigin.Begin);

    
            bi.CacheOption = BitmapCacheOption.OnLoad;

      
            bi.StreamSource = ms;

            bi.EndInit();
            pic_scan.Source = bi;
            bi.Freeze();


            return ms.ToArray();
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            btnNext.IsEnabled = true;
            if (GlobalSquence < SequenceTotal)
            {
                GlobalSquence--;
            }
            Scannerdata first = (from row in objListSD where row.Sequence == GlobalSquence select row).FirstOrDefault();
            LoadImage(first.ImageObj);
            if (GlobalSquence == 1)
            {
                btnPreview.IsEnabled = false;
            }

        }

      
        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            btnPreview.IsEnabled = true;
            if (GlobalSquence < SequenceTotal)
            {
                GlobalSquence++;
            }
            Scannerdata first = (from row in objListSD where row.Sequence == GlobalSquence select row).FirstOrDefault();
            LoadImage(first.ImageObj);
            if (GlobalSquence == SequenceTotal - 1)
            {
                btnNext.IsEnabled = false;
            }

        }

        private void Button_Click_3(object sender, RoutedEventArgs e)
        {
            using (var ms = new MemoryStream())
            {
                var document = new iTextSharp.text.Document(iTextSharp.text.PageSize.LETTER.Rotate(), 0, 0, 0, 0);
                iTextSharp.text.pdf.PdfWriter.GetInstance(document, ms).SetFullCompression();
                document.Open();

                foreach (System.Drawing.Image aa in obj)
                {
                    MemoryStream msimage = new MemoryStream();

                    aa.Save(msimage, ImageFormat.Jpeg);

                    var image = iTextSharp.text.Image.GetInstance(msimage.ToArray());
                    image.ScaleToFit(document.PageSize.Width, document.PageSize.Height);
                    document.Add(image);
                }
                document.Close();

                string Path = ConfigurationManager.AppSettings["uploadfolderpath"].ToString();


                string filename = "C3kycDMS" + DateTime.Now.ToString("yyyy-MM-dd HHmmss") + ".pdf";

                File.WriteAllBytes(Path + filename, ms.ToArray());
                byte[] test = ms.ToArray();


                MessageBox.Show("File Uploaded Successfully", "Success!", MessageBoxButton.OKCancel);
                pic_scan.Source = null;
            }
        }
    }

    enum ScanSource
    {
        DocumentFeeder = 1,
        Flatbed = 2,
    }

    public class Scannerdata
    {
        public int Sequence { get; set; }

        public System.Drawing.Image ImageObj { get; set; }
    }
}
 
       
    


