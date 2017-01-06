using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Windows.ApplicationModel;
using Windows.ApplicationModel.AppService;
using Windows.ApplicationModel.Background;
using Windows.ApplicationModel.Resources.Core;
using Windows.ApplicationModel.VoiceCommands;
using Windows.Storage;
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
using System.IO;
using System.Runtime.InteropServices.WindowsRuntime;
using Microsoft.ProjectOxford.Face;
using Microsoft.ProjectOxford.Face.Contract;

namespace KinectCameraFeed.VoiceCommandService
{

    public sealed class CameraFeedVoiceCommandService : IBackgroundTask
    {
        //
        private readonly IFaceServiceClient faceServiceClient = new FaceServiceClient("ENTER YOUR CLIENT GUID FROM COGNITIVE SERVICES HERE");
        private IEnumerable<FaceAttributes> faceAttribs;

        // MF Capture device to start up
        MediaCapture _mediaCapture;

        bool bColorCameraOn = false;
        bool bReceivingColorFrames = false;
        bool bImageCreated = false;
        bool bFaceDetectionComplete = false;

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
                
        // Xaml Bitmap for updating bitmaps in Xaml
        private SoftwareBitmap extBitmap;


        /// <summary>
        /// the service connection is maintained for the lifetime of a cortana session, once a voice command
        /// has been triggered via Cortana.
        /// </summary>
        VoiceCommandServiceConnection voiceServiceConnection;

        /// <summary>
        /// Lifetime of the background service is controlled via the BackgroundTaskDeferral object, including
        /// registering for cancellation events, signalling end of execution, etc. Cortana may terminate the 
        /// background service task if it loses focus, or the background task takes too long to provide.
        /// 
        /// Background tasks can run for a maximum of 30 seconds.
        /// </summary>
        BackgroundTaskDeferral serviceDeferral;

        /// <summary>
        /// ResourceMap containing localized strings for display in Cortana.
        /// </summary>
        ResourceMap cortanaResourceMap;

        /// <summary>
        /// The context for localized strings.
        /// </summary>
        ResourceContext cortanaContext;

        /// <summary>
        /// Get globalization-aware date formats.
        /// </summary>
        DateTimeFormatInfo dateFormatInfo;

        /// <summary>
        /// Background task entrypoint. Voice Commands using the <VoiceCommandService Target="...">
        /// tag will invoke this when they are recognized by Cortana, passing along details of the 
        /// invocation. 
        /// 
        /// Background tasks must respond to activation by Cortana within 0.5 seconds, and must 
        /// report progress to Cortana every 5 seconds (unless Cortana is waiting for user
        /// input). There is no execution time limit on the background task managed by Cortana,
        /// but developers should use plmdebug (https://msdn.microsoft.com/en-us/library/windows/hardware/jj680085%28v=vs.85%29.aspx)
        /// on the Cortana app package in order to prevent Cortana timing out the task during
        /// debugging.
        /// 
        /// Cortana dismisses its UI if it loses focus. This will cause it to terminate the background
        /// task, even if the background task is being debugged. Use of Remote Debugging is recommended
        /// in order to debug background task behaviors. In order to debug background tasks, open the
        /// project properties for the app package (not the background task project), and enable
        /// Debug -> "Do not launch, but debug my code when it starts". Alternatively, add a long
        /// initial progress screen, and attach to the background task process while it executes.
        /// </summary>
        /// <param name="taskInstance">Connection to the hosting background service process.</param>
        public async void Run(IBackgroundTaskInstance taskInstance)
        {

             // UWP BitmapSource
          //   source = new SoftwareBitmapSource();

             serviceDeferral = taskInstance.GetDeferral();

            // Register to receive an event if Cortana dismisses the background task. This will
            // occur if the task takes too long to respond, or if Cortana's UI is dismissed.
            // Any pending operations should be cancelled or waited on to clean up where possible.
            taskInstance.Canceled += OnTaskCanceled;

            var triggerDetails = taskInstance.TriggerDetails as AppServiceTriggerDetails;

            // Load localized resources for strings sent to Cortana to be displayed to the user.
            cortanaResourceMap = ResourceManager.Current.MainResourceMap.GetSubtree("Resources");

            // Select the system language, which is what Cortana should be running as.
            cortanaContext = ResourceContext.GetForViewIndependentUse();

            // Get the currently used system date format
            dateFormatInfo = CultureInfo.CurrentCulture.DateTimeFormat;

            // This should match the uap:AppService and VoiceCommandService references from the 
            // package manifest and VCD files, respectively. Make sure we've been launched by
            // a Cortana Voice Command.
            if (triggerDetails != null && triggerDetails.Name == "CameraFeedVoiceCommandService")
            {
                try
                {
                    voiceServiceConnection =
                        VoiceCommandServiceConnection.FromAppServiceTriggerDetails(
                            triggerDetails);

                    voiceServiceConnection.VoiceCommandCompleted += OnVoiceCommandCompleted;

                    // GetVoiceCommandAsync establishes initial connection to Cortana, and must be called prior to any 
                    // messages sent to Cortana. Attempting to use ReportSuccessAsync, ReportProgressAsync, etc
                    // prior to calling this will produce undefined behavior.
                    VoiceCommand voiceCommand = await voiceServiceConnection.GetVoiceCommandAsync();

                    // Depending on the operation (defined in AdventureWorks:AdventureWorksCommands.xml)
                    // perform the appropriate command.
                    switch (voiceCommand.CommandName)
                    {
                        case "whoIsInMyRoom":
                            // Access the value of the {location} phrase in the voice command
                            string location = voiceCommand.Properties["location"][0];
                            await ProcessKinectFeed(location);
                            break;
                        default:
                            // As with app activation VCDs, we need to handle the possibility that
                            // an app update may remove a voice command that is still registered.
                            // This can happen if the user hasn't run an app since an update.
                            LaunchAppInForeground();
                            break;
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine("Handling Voice Command failed " + ex.ToString());
                }
            }
        }


        private async Task ProcessKinectFeed(string location)
        {
            // Begin loading data to search for the target store. If this operation is going to take a long time,
            // for instance, requiring a response from a remote web service, consider inserting a progress screen 
            // here, in order to prevent Cortana from timing out. 
            InitializeColorFrames();

            string progressScreenString = "Starting up camera feeds for " + location;  //string.Format(cortanaResourceMap.GetValue("ProgressLookingForTripToDest", cortanaContext).ValueAsString, location);
            while (!bReceivingColorFrames)
            {
                await ShowProgressScreen(progressScreenString);
            }

            await ShowProgressScreen("Processing Image from Kinect camera for " + location);

            var file = await WriteableBitmapToStorageFile(extBitmap, FileFormat.Png, "kinectProcess.png");

            while (!bImageCreated)
            {
                await ShowProgressScreen("Sending Image to cognitive services for " + location);
            }

            var faces = await UploadAndDetectFaces(file);
            while (!bFaceDetectionComplete)
            {
                await ShowProgressScreen("Detecting faces for " + location);
            }

            var facesCount = faces.Count();

            VoiceCommandResponse response;
            StringBuilder sb = new StringBuilder();

            var userMessage = new VoiceCommandUserMessage();
            if (faceAttribs != null)
            {
                foreach (var faceAttrib in faceAttribs)
                {
                    sb.Append(string.Format("There is a {0} age {1}", faceAttrib.Gender, faceAttrib.Age));
                }
            }
            sb.Append(string.Format("there are a total of {0} persons in {1}", facesCount, location));
            userMessage.DisplayMessage = sb.ToString();
            response = VoiceCommandResponse.CreateResponse(userMessage);
            await voiceServiceConnection.ReportSuccessAsync(response);
         
            }
        

        /// <summary>
        /// Show a progress screen. These should be posted at least every 5 seconds for a 
        /// long-running operation, such as accessing network resources over a mobile 
        /// carrier network.
        /// </summary>
        /// <param name="message">The message to display, relating to the task being performed.</param>
        /// <returns></returns>
        private async Task ShowProgressScreen(string message)
        {
            var userProgressMessage = new VoiceCommandUserMessage();

            if (bColorCameraOn)
            {
                userProgressMessage.DisplayMessage = userProgressMessage.SpokenMessage = "Kinect Color On: " + message;
            }
            else
            {
                userProgressMessage.DisplayMessage = userProgressMessage.SpokenMessage = message;
            }
            VoiceCommandResponse response = VoiceCommandResponse.CreateResponse(userProgressMessage);
            await voiceServiceConnection.ReportProgressAsync(response);
        }

       

       
        /// <summary>
        /// Provide a simple response that launches the app. Expected to be used in the
        /// case where the voice command could not be recognized (eg, a VCD/code mismatch.)
        /// </summary>
        private async void LaunchAppInForeground()
        {
            var userMessage = new VoiceCommandUserMessage();
            userMessage.SpokenMessage = "Launching Kinect Camera Feed"; //cortanaResourceMap.GetValue("LaunchingAdventureWorks", cortanaContext).ValueAsString;

            var response = VoiceCommandResponse.CreateResponse(userMessage);

            response.AppLaunchArgument = "";

            await voiceServiceConnection.RequestAppLaunchAsync(response);
        }

        /// <summary>
        /// Handle the completion of the voice command. Your app may be cancelled
        /// for a variety of reasons, such as user cancellation or not providing 
        /// progress to Cortana in a timely fashion. Clean up any pending long-running
        /// operations (eg, network requests).
        /// </summary>
        /// <param name="sender">The voice connection associated with the command.</param>
        /// <param name="args">Contains an Enumeration indicating why the command was terminated.</param>
        private void OnVoiceCommandCompleted(VoiceCommandServiceConnection sender, VoiceCommandCompletedEventArgs args)
        {
            if (this.serviceDeferral != null)
            {
                this.serviceDeferral.Complete();
            }
        }

        /// <summary>
        /// When the background task is cancelled, clean up/cancel any ongoing long-running operations.
        /// This cancellation notice may not be due to Cortana directly. The voice command connection will
        /// typically already be destroyed by this point and should not be expected to be active.
        /// </summary>
        /// <param name="sender">This background task instance</param>
        /// <param name="reason">Contains an enumeration with the reason for task cancellation</param>
        private void OnTaskCanceled(IBackgroundTaskInstance sender, BackgroundTaskCancellationReason reason)
        {
            System.Diagnostics.Debug.WriteLine("Task cancelled, clean up");
            if (this.serviceDeferral != null)
            {
                //Complete the service deferral
                this.serviceDeferral.Complete();
            }
        }

        #region "Kinect Logic"

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
                bColorCameraOn = true;
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
                bColorCameraOn = true;
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
                        bColorCameraOn = true;
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
                   extBitmap = new SoftwareBitmap(vidFrame.SoftwareBitmap.BitmapPixelFormat, vidFrame.SoftwareBitmap.PixelWidth, vidFrame.SoftwareBitmap.PixelHeight);
                    vidFrame.SoftwareBitmap.CopyTo(extBitmap);

                    // PixelFormat needs to be in 8bit BGRA for Xaml writable bitmap
                    if (extBitmap.BitmapPixelFormat != BitmapPixelFormat.Bgra8)
                        extBitmap = SoftwareBitmap.Convert(vidFrame.SoftwareBitmap, BitmapPixelFormat.Bgra8);

                    
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
                        bReceivingColorFrames = true;
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



        private async Task<StorageFile> WriteableBitmapToStorageFile(SoftwareBitmap WB, FileFormat fileFormat, string fileName = "")
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

                    // Set the software bitmap
                    encoder.SetSoftwareBitmap(WB);

                    // Set additional encoding parameters, if needed
                    //encoder.BitmapTransform.ScaledWidth = 320;
                    //encoder.BitmapTransform.ScaledHeight = 240;
                    //encoder.BitmapTransform.Rotation = Windows.Graphics.Imaging.BitmapRotation.Clockwise90Degrees;
                    //encoder.BitmapTransform.InterpolationMode = BitmapInterpolationMode.Fant;
                    encoder.IsThumbnailGenerated = true;

                    try
                    {
                        await encoder.FlushAsync();
                    }
                    catch (Exception err)
                    {
                        switch (err.HResult)
                        {
                            case unchecked((int)0x88982F81): //WINCODEC_ERR_UNSUPPORTEDOPERATION
                                                             // If the encoder does not support writing a thumbnail, then try again
                                                             // but disable thumbnail generation.
                                encoder.IsThumbnailGenerated = false;
                                break;
                            default:
                                throw err;
                        }
                    }

                    if (encoder.IsThumbnailGenerated == false)
                    {
                        await encoder.FlushAsync();
                    }
                                        
                }
                bImageCreated = true;
                return file;
            }
            catch (Exception ex)
            {
                bImageCreated = true;
                return null;
                // throw;
            }
        }
        #endregion

        private async Task<FaceRectangle[]> UploadAndDetectFaces(StorageFile file)
        {
            try
            {
                IEnumerable<FaceRectangle> faceRects = null;
                // var ignore = Dispatcher.RunAsync(Windows.UI.Core.CoreDispatcherPriority.Normal, async () =>
                //{
                using (IRandomAccessStream stream = await file.OpenAsync(FileAccessMode.ReadWrite))
                {
                    FaceAttributeType[] faceAttribTypes = new FaceAttributeType[5];
                    faceAttribTypes[0] = FaceAttributeType.Age;
                    faceAttribTypes[1] = FaceAttributeType.FacialHair;
                    faceAttribTypes[2] = FaceAttributeType.Gender;
                    faceAttribTypes[3] = FaceAttributeType.HeadPose;
                    faceAttribTypes[4] = FaceAttributeType.Smile;


                    Stream strm = stream.AsStreamForRead();
                    var faces = await faceServiceClient.DetectAsync(strm, true, false, faceAttribTypes);
                    faceRects = faces.Select(face => face.FaceRectangle);
                    faceAttribs = faces.Select(f => f.FaceAttributes);
                }
                //});

                if (faceAttribs != null)
                {
                    var attribs = faceAttribs.ToArray();

                }

                bFaceDetectionComplete = true;
                return faceRects.ToArray();
            }
            catch (Exception ex)
            {
                bFaceDetectionComplete = true;
                return new FaceRectangle[0];
            }
        }


    }
}
