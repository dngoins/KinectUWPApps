using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices.WindowsRuntime;
using Windows.Foundation;
using Windows.Foundation.Collections;
using Windows.UI.Xaml;
using Windows.UI.Xaml.Controls;
using Windows.UI.Xaml.Controls.Primitives;
using Windows.UI.Xaml.Data;
using Windows.UI.Xaml.Input;
using Windows.UI.Xaml.Media;
using Windows.UI.Xaml.Navigation;
using Windows.Media.Capture;
using Windows.Media.Capture.Frames;
using System.Threading.Tasks;
using Windows.Graphics.Imaging;
using Windows.UI.Xaml.Media.Imaging;
using Windows.Storage;
using Windows.Storage.Pickers;
using Windows.Storage.Streams;

// The Blank Page item template is documented at http://go.microsoft.com/fwlink/?LinkId=402352&clcid=0x409

namespace KinectCameraFeed
{
    /// <summary>
    /// An empty page that can be used on its own or navigated to within a Frame.
    /// </summary>
    public sealed partial class MainPage : Page
    {
        // MF Capture device to start up
        MediaCapture _mediaCapture;

        private enum FileFormat
        {
            Jpeg,
            Png,
            Bmp,
            Tiff,
            Gif
        }

        // Constants that let us decide on which frame source kind we want
        const int COLOR_SOURCE = 0;
        const int IR_SOURCE = 1;
        const int DEPTH_SOURCE = 2;
        const int LONGIR_SOURCE = 3;
        const int BODY_SOURCE = 4;
        const int BODY_INDEX_SOURCE = 5;


        // UWP BitmapSource
        private SoftwareBitmapSource source = new SoftwareBitmapSource();

        // Xaml Bitmap for updating bitmaps in Xaml
        private WriteableBitmap extBitmap;

        public MainPage()
        {
            
            this.InitializeComponent();
            this.InitializeColorFrames();
        }

        protected override void OnNavigatedTo(NavigationEventArgs e)
        {
            base.OnNavigatedTo(e);
        }
        async void InitializeColorFrames()
        {
            // Devices come grouped together as a set of capabilities
            // Find all the media frame source devices - as a collection of groups            
            var allGroups = await MediaFrameSourceGroup.FindAllAsync();

            // Let's filter using Linq, based on a device group with a DisplayName of "Kinect"
            var eligibleColorGroups = allGroups
                .Where(g => g.SourceInfos.FirstOrDefault(info => info.SourceGroup.DisplayName.Contains("Kinect")) != null).ToList();

            // Check to see if we found a device
            if (eligibleColorGroups.Count == 0)
            {
                //No kinect camera found
                return;
            }

            //Only 1 Kinect camera supported so always take the first in the list
            var kinectColorGroup = eligibleColorGroups[0];
            try
            {
                // Initialize MediaCapture with selected group.
                // This can raise an exception if the source no longer exists,
                // or if the source could not be initialized.
                await InitializeMediaCaptureAsync(kinectColorGroup);
            }
            catch (Exception exception)
            {
                //Release any resources if something goes wrong
                await CleanupMediaCaptureAsync();
                return;
            }

            // Let's get a Device Capability, in this case let's get the color source stream
            MediaFrameSourceInfo colorInfo = kinectColorGroup.SourceInfos[COLOR_SOURCE];

            if (colorInfo != null)
            {
                // Access the initialized frame source by looking up the the ID of the source found above.
                // Verify that the Id is present, because it may have left the group while were were
                // busy deciding which group to use.
                MediaFrameSource frameSource = null;
                if (_mediaCapture.FrameSources.TryGetValue(colorInfo.Id, out frameSource))
                {
                    // Create a frameReader based on the color source stream
                    MediaFrameReader frameReader = await _mediaCapture.CreateFrameReaderAsync(frameSource);

                    // Listen for Frame Arrived events
                    frameReader.FrameArrived += FrameReader_FrameArrived;

                    // Start the color source frame processing
                    MediaFrameReaderStartStatus status = await frameReader.StartAsync();

                    // Status checking for logging purposes if needed
                    if (status != MediaFrameReaderStartStatus.Success)
                    {

                    }
                }
                else
                {
                    // Couldn't get the color frame source from the MF Device

                }
            }
            else
            {
                // There's no color source

            }

        }



        private void ProcessColorFrame(MediaFrameReference clrFrame)
        {
            try
            {
                //clrFrame.
                var buffFrame = clrFrame?.BufferMediaFrame;

                // Get the Individual color Frame
                var vidFrame = clrFrame?.VideoMediaFrame;
                {
                    if (vidFrame == null) return;


                    // create a UWP SoftwareBitmap and copy Color Frame into Bitmap
                    SoftwareBitmap sbt = new SoftwareBitmap(vidFrame.SoftwareBitmap.BitmapPixelFormat, vidFrame.SoftwareBitmap.PixelWidth, vidFrame.SoftwareBitmap.PixelHeight);
                    vidFrame.SoftwareBitmap.CopyTo(sbt);

                    // PixelFormat needs to be in 8bit BGRA for Xaml writable bitmap
                    if (sbt.BitmapPixelFormat != BitmapPixelFormat.Bgra8)
                        sbt = SoftwareBitmap.Convert(vidFrame.SoftwareBitmap, BitmapPixelFormat.Bgra8);

                    if (source != null)
                    {
                        // To write out to writable bitmap which will be used with ImageElement, it needs to run
                        // on UI Thread thus we use Dispatcher.RunAsync()...
                        var ignore = Dispatcher.RunAsync(Windows.UI.Core.CoreDispatcherPriority.Normal, () =>
                        {
                            // This code runs on UI Thread
                            // Create the writableBitmap for ImageElement display
                            extBitmap = new WriteableBitmap(sbt.PixelWidth, sbt.PixelHeight);

                            // Copy contents from UWP software Bitmap
                            // There are other ways of doing this instead of the double copy, 1st copy earlier
                            // this is a second copy.
                            sbt.CopyToBuffer(extBitmap.PixelBuffer);
                            extBitmap.Invalidate();

                            // Set the imageElement source
                            var ig = source.SetBitmapAsync(sbt);
                            imgView.Source = source;

                        });

                    }
                }
            }
            catch (Exception ex)
            {

            }
        }



        private void FrameReader_FrameArrived(MediaFrameReader sender, MediaFrameArrivedEventArgs args)
        {
            // Try to get the FrameReference
            using (var frameRef = sender.TryAcquireLatestFrame())
            {
                if (frameRef != null)
                {
                    if (frameRef.SourceKind == MediaFrameSourceKind.Color)
                    {
                        // Process the frame drawing the color pixels onto the screen
                        ProcessColorFrame(frameRef);
                    }
                }
            }
        }

        /// <summary>
        /// Initializes the MediaCapture object with the given source group.
        /// </summary>
        /// <param name="sourceGroup">SourceGroup with which to initialize.</param>
        private async Task InitializeMediaCaptureAsync(MediaFrameSourceGroup sourceGroup)
        {
            if (_mediaCapture != null)
            {
                return;
            }

            // Initialize mediacapture with the source group.
            _mediaCapture = new MediaCapture();
            var settings = new MediaCaptureInitializationSettings
            {
                SourceGroup = sourceGroup,

                // This media capture can share streaming with other apps.
                SharingMode = MediaCaptureSharingMode.SharedReadOnly,

                // Only stream video and don't initialize audio capture devices.
                StreamingCaptureMode = StreamingCaptureMode.Video,

                // Set to CPU to ensure frames always contain CPU SoftwareBitmap images
                // instead of preferring GPU D3DSurface images.
                MemoryPreference = MediaCaptureMemoryPreference.Cpu
            };

            await _mediaCapture.InitializeAsync(settings);
        }

        /// <summary>
        /// Unregisters FrameArrived event handlers, stops and disposes frame readers
        /// and disposes the MediaCapture object.
        /// </summary>
        private async Task CleanupMediaCaptureAsync()
        {
            if (_mediaCapture != null)
            {
                using (var mediaCapture = _mediaCapture)
                {
                    _mediaCapture = null;

                }
            }
        }



        private async Task<StorageFile> WriteableBitmapToStorageFile(WriteableBitmap WB, FileFormat fileFormat, string fileName = "")
        {
            try
            {
                string FileName = string.Empty;
                if (string.IsNullOrEmpty(fileName))
                {
                    FileSavePicker savePicker = new FileSavePicker();
                    savePicker.SuggestedStartLocation = PickerLocationId.PicturesLibrary;

                    // Dropdown of file types the user can save the file as
                    savePicker.FileTypeChoices.Add("jpeg", new List<string>() { ".jpg", ".jpeg" });

                    // Default file name if the user does not type one in or select a file to replace
                    savePicker.SuggestedFileName = "WorkingWithMediaCapture.jpg";
                    fileName = (await savePicker.PickSaveFileAsync()).Name;
                }
                FileName = fileName;

                Guid BitmapEncoderGuid = BitmapEncoder.JpegEncoderId;
                switch (fileFormat)
                {
                    case FileFormat.Jpeg:
                        //  FileName = string.Format("{0}.jpeg", fileName);
                        BitmapEncoderGuid = BitmapEncoder.JpegEncoderId;
                        break;

                    case FileFormat.Png:
                        //  FileName = string.Format("{0}.png", fileName);
                        BitmapEncoderGuid = BitmapEncoder.PngEncoderId;
                        break;

                    case FileFormat.Bmp:
                        //  FileName = string.Format("{0}.bmp", fileName);
                        BitmapEncoderGuid = BitmapEncoder.BmpEncoderId;
                        break;

                    case FileFormat.Tiff:
                        //  FileName = string.Format("{0}.tiff", fileName);
                        BitmapEncoderGuid = BitmapEncoder.TiffEncoderId;
                        break;

                    case FileFormat.Gif:
                        //  FileName = string.Format("{0}.gif", fileName);
                        BitmapEncoderGuid = BitmapEncoder.GifEncoderId;
                        break;
                }

                var file = await Windows.Storage.KnownFolders.PicturesLibrary.CreateFileAsync(FileName, CreationCollisionOption.GenerateUniqueName);
                using (IRandomAccessStream stream = await file.OpenAsync(FileAccessMode.ReadWrite))
                {
                    BitmapEncoder encoder = await BitmapEncoder.CreateAsync(BitmapEncoderGuid, stream);
                    Stream pixelStream = WB.PixelBuffer.AsStream();
                    byte[] pixels = new byte[pixelStream.Length];
                    await pixelStream.ReadAsync(pixels, 0, pixels.Length);

                    encoder.SetPixelData(BitmapPixelFormat.Bgra8, BitmapAlphaMode.Ignore,
                                        (uint)WB.PixelWidth,
                                        (uint)WB.PixelHeight,
                                        96.0,
                                        96.0,
                                        pixels);
                    await encoder.FlushAsync();
                }
                return file;
            }
            catch (Exception ex)
            {
                return null;
                // throw;
            }
        }

        private void btnExternalCapture_Click(object sender, RoutedEventArgs e)
        {
            if (extBitmap != null)
            {
                var ignore = Dispatcher.RunAsync(Windows.UI.Core.CoreDispatcherPriority.Normal, async () =>
                {
                    var file =
                    await WriteableBitmapToStorageFile(extBitmap, FileFormat.Jpeg, "WorkingWithMediaCaptureFrames.jpg");
                    fileLocation.Text = file.Path;
                });

            }
        }
    }
}
