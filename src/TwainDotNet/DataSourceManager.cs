using System;
using System.Collections.Generic;
using TwainDotNet.TwainNative;
using System.Runtime.InteropServices;
using TwainDotNet.Win32;
using System.Reflection;
using System.Drawing;
using log4net;

namespace TwainDotNet
{
    public class DataSourceManager : IDisposable
    {
        /// <summary>
        /// The logger for this class.
        /// </summary>
        static ILog log = LogManager.GetLogger(typeof(DataSourceManager));

        IWindowsMessageHook _messageHook;
        Event _eventMessage;

        public Identity ApplicationId { get; private set; }
        public DataSource DataSource { get; private set; }
        public bool UseIncrementalMemoryXfer { get; set; } = false;

        public DataSourceManager(Identity applicationId, IWindowsMessageHook messageHook)
        {
            // Make a copy of the identity in case it gets modified
            ApplicationId = applicationId.Clone();

            ScanningComplete += delegate { };
            TransferImage += delegate { };

            _messageHook = messageHook;
            _messageHook.FilterMessageCallback = FilterMessage;
            IntPtr windowHandle = _messageHook.WindowHandle;

            _eventMessage.EventPtr = Marshal.AllocHGlobal(Marshal.SizeOf(typeof(WindowsMessage)));

            // Initialise the data source manager
            TwainResult result = Twain32Native.DsmParent(
                ApplicationId,
                IntPtr.Zero,
                DataGroup.Control,
                DataArgumentType.Parent,
                Message.OpenDSM,
                ref windowHandle);

            if (result == TwainResult.Success)
            {
                //according to the 2.0 spec (2-10) if (applicationId.SupportedGroups
                // | DataGroup.Dsm2) > 0 then we should call DM_Entry(id, 0, DG_Control, DAT_Entrypoint, MSG_Get, wh)
                //right here
                DataSource = DataSource.GetDefault(ApplicationId, _messageHook);
            }
            else
            {
                throw new TwainException("Error initialising DSM: " + result, result);
            }
        }

        ~DataSourceManager()
        {
            Dispose(false);
        }

        /// <summary>
        /// Notification that the scanning has completed.
        /// </summary>
        public event EventHandler<ScanningCompleteEventArgs> ScanningComplete;

        public event EventHandler<TransferImageEventArgs> TransferImage;

        public IWindowsMessageHook MessageHook { get { return _messageHook; } }

        /// <summary>
        /// 偵測有沒有紙，需要先select一個data source。做完會關閉data source。
        /// </summary>
        public bool IsPaperOn()
        {
            DataSource.OpenSource();
            bool bReturn = DataSource.PaperDetectable;
            DataSource.Close();
            return bReturn;
        }

        public void StartScan(ScanSettings settings)
        {
            bool scanning = false;

            try
            {
                _messageHook.UseFilter = true;
                scanning = DataSource.OpenAndScan(settings, UseIncrementalMemoryXfer);
            }
            catch (TwainException e)
            {
                DataSource.Close();
                EndingScan();
                throw e;
            }
            finally
            {
                // Remove the message hook if scan setup failed
                if (!scanning)
                {
                    EndingScan();
                }
            }
        }

        protected IntPtr FilterMessage(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
        {
            if (DataSource.SourceId.Id == 0)
            {
                handled = false;
                return IntPtr.Zero;
            }

            int pos = User32Native.GetMessagePos();

            WindowsMessage message = new WindowsMessage();
            message.hwnd = hwnd;
            message.message = msg;
            message.wParam = wParam;
            message.lParam = lParam;
            message.time = User32Native.GetMessageTime();
            message.x = (short)pos;
            message.y = (short)(pos >> 16);

            Marshal.StructureToPtr(message, _eventMessage.EventPtr, false);
            _eventMessage.Message = 0;

            TwainResult result = Twain32Native.DsEvent(
                ApplicationId,
                DataSource.SourceId,
                DataGroup.Control,
                DataArgumentType.Event,
                Message.ProcessEvent,
                ref _eventMessage);

            if (result == TwainResult.NotDSEvent)
            {
                handled = false;
                return IntPtr.Zero;
            }

            switch (_eventMessage.Message)
            {
                case Message.XFerReady:
                    Exception exception = null;
                    try
                    {
                        if (this.UseIncrementalMemoryXfer)
                        {
                            TransferPicturesIncremental();
                        }
                        else
                        {
                            TransferPictures();
                        }
                    }
                    catch (Exception e)
                    {
                        exception = e;
                    }
                    CloseDsAndCompleteScanning(exception);
                    break;

                case Message.CloseDS:
                case Message.CloseDSOK:
                case Message.CloseDSReq:
                    CloseDsAndCompleteScanning(null);
                    break;

                case Message.DeviceEvent:
                    break;

                default:
                    break;
            }

            handled = true;
            return IntPtr.Zero;
        }

        protected void TransferPicturesIncremental()
        {
            // see http://www.twain.org/wp-content/uploads/2017/03/TWAIN-2.4-Specification.pdf
            // page 4-20
            Console.WriteLine("TransferPicturesIncremental...");

            if (DataSource.SourceId.Id == 0)
            {
                return;
            }

            PendingXfers pendingTransfer = new PendingXfers();
            TwainResult result;
            try
            {
                int recievedBlockCount = 1;
                do
                {
                    pendingTransfer.Count = 0;     // the Twain source will fill this in during DsPendingTransfer                    

                    Console.WriteLine("Get the image info...");
                    // Get the image info
                    ImageInfo imageInfo = new ImageInfo();
                    result = Twain32Native.DsImageInfo(
                        ApplicationId,
                        DataSource.SourceId,
                        DataGroup.Image,
                        DataArgumentType.ImageInfo,
                        Message.Get,
                        imageInfo);

                    if (result != TwainResult.Success)
                    {
                        DataSource.Close();
                        break;
                    }

                    /*Console.WriteLine("Get the image layout...");
                    ImageLayout imageLayout = new ImageLayout();
                    result = Twain32Native.DsImageLayout(
                        ApplicationId,
                        DataSource.SourceId,
                        DataGroup.Image,
                        DataArgumentType.ImageLayout,
                        Message.GetCurrent,
                        imageLayout);

                    if (result != TwainResult.Success)
                    {
                        DataSource.Close();
                        break;
                    }*/

                    // Setup Destination Bitmap
                    Bitmap bitmap = BitmapRenderer.NewBitmapForImageInfo(imageInfo);

                    Console.WriteLine("Setup incremental Memory XFer...");
                    // Setup incremental Memory XFer                                        
                    SetupMemXfer setupMemXfer = new SetupMemXfer();
                    result = Twain32Native.DsSetupMemXfer(
                        ApplicationId,
                        DataSource.SourceId,
                        DataGroup.Control,
                        DataArgumentType.SetupMemXfer,
                        Message.Get,
                        setupMemXfer
                        );

                    if (result != TwainResult.Success)
                    {
                        DataSource.Close();
                        break;
                    }

                    Console.WriteLine("allocate the preferred buffer size...");
                    // allocate the preferred buffer size          
                    // see twain spec pdf, page 4-21          
                    ImageMemXfer imageMemXfer = new ImageMemXfer();
                    try
                    {
                        imageMemXfer.Memory.Flags = MemoryFlags.AppOwns | MemoryFlags.Pointer;
                        // 對於A8 scanner，Preferred = MaxBufSize，太大了，所以我們選小一點的
                        uint length = 100000;
                        if (length > setupMemXfer.MinBufSize && length < setupMemXfer.MaxBufSize)
                            imageMemXfer.Memory.Length = length;//length;
                        else
                            imageMemXfer.Memory.Length = setupMemXfer.MinBufSize;
                        imageMemXfer.Memory.TheMem = Kernel32Native.GlobalAlloc(GlobalAllocFlags.MemFixed, (int)imageMemXfer.Memory.Length * 2);  // 不知道為什麼原本她size寫要*2倍
                        imageMemXfer.Compression = Compression.None;

                        if (imageMemXfer.Memory.TheMem == IntPtr.Zero)
                        {
                            throw new TwainException("error allocating buffer for memory transfer");
                        }

                        long pixels_written = 0;
                        long total_pixels = imageInfo.ImageWidth * imageInfo.ImageLength;
                        int heightTransfered = 0;

                        do
                        {
                            // perform a transfer
                            result = Twain32Native.DsImageMemXfer(
                                ApplicationId,
                                DataSource.SourceId,
                                DataGroup.Image,
                                DataArgumentType.ImageMemXfer,
                                Message.Get,
                                imageMemXfer
                                );

                            if (result == TwainResult.Success || result == TwainResult.XferDone)
                            {
                                // 計算累計的高度
                                heightTransfered += (int)imageMemXfer.Rows;

                                // dibArray是這次Buffer的RGB陣列
                                byte[] dibArray = ShiftPixels(ref imageMemXfer, imageInfo.BitsPerPixel / 8);

                                BitmapRenderer.TransferPixels(bitmap, imageInfo, imageMemXfer);
                                pixels_written += (imageMemXfer.BytesWritten * 8) / imageInfo.BitsPerPixel;
                                double percent_complete = (double)pixels_written / (double)total_pixels;

                                // 算出空白區域的高度，裁切尾端部分
                                /*Bitmap newBmp;
                                if (result == TwainResult.XferDone)
                                {
                                    percent_complete = 1.0;
                                    int blankHeight = GetCropHeight(bitmap, true);
                                    if (blankHeight > 0 && blankHeight < imageInfo.ImageLength)
                                    {
                                        newBmp = cropImage(bitmap, blankHeight);
                                    }
                                    else
                                        newBmp = new Bitmap(bitmap);
                                }
                                else
                                {
                                    newBmp = cropImage(bitmap, bitmap.Height - heightTransfered);
                                }*/
                                Bitmap newBmp = new Bitmap(bitmap);

                                // fire the transfer event
                                TransferImageEventArgs args = new TransferImageEventArgs(newBmp, result != TwainResult.XferDone, (float)percent_complete);
                                TransferImage(this, args);
                                if (!args.ContinueScanning)
                                {
                                    result = TwainResult.XferDone;
                                }
                            }
                            recievedBlockCount++;

                        } while (result == TwainResult.Success);

                    }
                    finally
                    {
                        if (imageMemXfer.Memory.TheMem != IntPtr.Zero)
                        {
                            Kernel32Native.GlobalFree(imageMemXfer.Memory.TheMem);
                            imageMemXfer.Memory.TheMem = IntPtr.Zero;
                        }
                        bitmap?.Dispose();
                    }

                    // End pending transfers
                    result = Twain32Native.DsPendingTransfer(
                        ApplicationId,
                        DataSource.SourceId,
                        DataGroup.Control,
                        DataArgumentType.PendingXfers,
                        Message.EndXfer,
                        pendingTransfer);

                    if (result != TwainResult.Success)
                    {
                        DataSource.Close();
                        break;
                    }
                }
                while (pendingTransfer.Count != 0);
            }
            finally
            {
                // Reset any pending transfers
                result = Twain32Native.DsPendingTransfer(
                    ApplicationId,
                    DataSource.SourceId,
                    DataGroup.Control,
                    DataArgumentType.PendingXfers,
                    Message.Reset,
                    pendingTransfer);
                //DataSource.Close();
            }
        }

        /// <summary>
        /// 把imageMemXfer.Memory內的記憶體向右移動
        /// </summary>
        /// <param name="imageMemXfer"></param>
        /// <param name="bytePerPixel"></param>
        private byte[] ShiftPixels(ref ImageMemXfer imageMemXfer, int bytePerPixel)
        {
            /// 參考:
            /// https://stackoverflow.com/questions/2185944/why-must-stride-in-the-system-drawing-bitmap-constructor-be-a-multiple-of-4
            /// https://stackoverflow.com/questions/47918451/creating-a-bitmap-from-rgb-array-in-c-result-image-difference/

            byte[] managedArray = new byte[(int)imageMemXfer.Memory.Length];
            Marshal.Copy(imageMemXfer.Memory.TheMem, managedArray, 0, (int)imageMemXfer.Memory.Length);
            int bytesPerRow = (int)imageMemXfer.BytesPerRow;

            byte[] newArray = null;
            int? skipByte = null;
            if (bytesPerRow % 4 != 0)
            {
                skipByte = 4 - (bytesPerRow % 4);
                int newArrayLength = (int)imageMemXfer.Memory.Length + skipByte.Value * (int)imageMemXfer.Rows;
                newArray = new byte[newArrayLength];
            }

            for (int i = 0; i < imageMemXfer.Rows; i++)
            {
                // get a single row
                byte[] rowArray = new byte[bytesPerRow];
                Buffer.BlockCopy(managedArray, i * bytesPerRow, rowArray, 0, bytesPerRow);

                // right shift 3 bytes by i
                byte[] tempArray = new byte[bytesPerRow];
                int shiftAmount = i * bytePerPixel;

                // copy back
                if (skipByte.HasValue)
                {
                    Buffer.BlockCopy(rowArray, 0, newArray, i * (bytesPerRow + skipByte.Value), bytesPerRow);
                }
                else
                    Buffer.BlockCopy(rowArray, 0, managedArray, i * bytesPerRow, bytesPerRow);
            }

            if (skipByte.HasValue)
                Marshal.Copy(newArray, 0, imageMemXfer.Memory.TheMem, newArray.Length);
            else
                Marshal.Copy(managedArray, 0, imageMemXfer.Memory.TheMem, (int)imageMemXfer.Memory.Length);

            if (skipByte.HasValue)
                return newArray;
            else
                return managedArray;
        }

        protected void TransferPictures()
        {
            if (DataSource.SourceId.Id == 0)
            {
                return;
            }

            PendingXfers pendingTransfer = new PendingXfers();
            TwainResult result;
            try
            {
                do
                {
                    pendingTransfer.Count = 0;     // the Twain source will fill this in during DsPendingTransfer
                    IntPtr hbitmap = IntPtr.Zero;

                    // Get the image info
                    ImageInfo imageInfo = new ImageInfo();
                    result = Twain32Native.DsImageInfo(
                        ApplicationId,
                        DataSource.SourceId,
                        DataGroup.Image,
                        DataArgumentType.ImageInfo,
                        Message.Get,
                        imageInfo);

                    if (result != TwainResult.Success)
                    {
                        DataSource.Close();
                        break;
                    }
                    Bitmap bitmap;
                    try
                    {

                        // Transfer the image from the device
                        result = Twain32Native.DsImageTransfer(
                            ApplicationId,
                            DataSource.SourceId,
                            DataGroup.Image,
                            DataArgumentType.ImageNativeXfer,
                            Message.Get,
                            ref hbitmap);

                        if (result != TwainResult.XferDone)
                        {
                            DataSource.Close();
                            break;
                        }
                        if (hbitmap == IntPtr.Zero)
                        {
                            throw new TwainException("Transfer complete, but bitmap pointer is still null.");
                        }

                        bitmap = BitmapRenderer.NewBitmapFromHBitmap(hbitmap);
                    }
                    finally
                    {
                        if (hbitmap != IntPtr.Zero)
                        {
                            Kernel32Native.GlobalFree(hbitmap);
                            hbitmap = IntPtr.Zero;
                        }
                    }

                    // End pending transfers
                    result = Twain32Native.DsPendingTransfer(
                        ApplicationId,
                        DataSource.SourceId,
                        DataGroup.Control,
                        DataArgumentType.PendingXfers,
                        Message.EndXfer,
                        pendingTransfer);

                    if (result != TwainResult.Success)
                    {
                        DataSource.Close();
                        break;
                    }

                    // fire the transfer event...
                    TransferImageEventArgs args = new TransferImageEventArgs(bitmap, pendingTransfer.Count != 0, 1.0f);
                    TransferImage(this, args);
                    if (!args.ContinueScanning)
                        break;
                }
                while (pendingTransfer.Count != 0);
            }
            finally
            {
                // Reset any pending transfers
                result = Twain32Native.DsPendingTransfer(
                    ApplicationId,
                    DataSource.SourceId,
                    DataGroup.Control,
                    DataArgumentType.PendingXfers,
                    Message.Reset,
                    pendingTransfer);
            }
        }

        protected void CloseDsAndCompleteScanning(Exception exception)
        {
            EndingScan();
            DataSource.Close();
            try
            {
                ScanningComplete(this, new ScanningCompleteEventArgs(exception));
            }
            catch
            {
            }
        }

        protected void EndingScan()
        {
            _messageHook.UseFilter = false;
        }

        public void SelectSource()
        {
            DataSource.Dispose();
            DataSource = DataSource.UserSelected(ApplicationId, _messageHook);
        }

        public void SelectSource(DataSource dataSource)
        {
            DataSource.Dispose();
            DataSource = dataSource;
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            Marshal.FreeHGlobal(_eventMessage.EventPtr);

            if (disposing)
            {
                DataSource.Dispose();

                IntPtr windowHandle = _messageHook.WindowHandle;

                if (ApplicationId.Id != 0)
                {
                    // Close down the data source manager
                    Twain32Native.DsmParent(
                        ApplicationId,
                        IntPtr.Zero,
                        DataGroup.Control,
                        DataArgumentType.Parent,
                        Message.CloseDSM,
                        ref windowHandle);
                }

                ApplicationId.Id = 0;
            }
        }

        public static ConditionCode GetConditionCode(Identity applicationId, Identity sourceId)
        {
            Status status = new Status();

            Twain32Native.DsmStatus(
                applicationId,
                sourceId,
                DataGroup.Control,
                DataArgumentType.Status,
                Message.Get,
                status);

            return status.ConditionCode;
        }

        public static readonly Identity DefaultApplicationId = new Identity()
        {
            Id = BitConverter.ToInt32(Guid.NewGuid().ToByteArray(), 0),
            Version = new TwainVersion()
            {
                MajorNum = 1,
                MinorNum = 1,
                Language = Language.USA,
                Country = Country.USA,
                Info = Assembly.GetExecutingAssembly().FullName
            },
            ProtocolMajor = TwainConstants.ProtocolMajor,
            ProtocolMinor = TwainConstants.ProtocolMinor,
            SupportedGroups = (int)(DataGroup.Image | DataGroup.Control),
            Manufacturer = "TwainDotNet",
            ProductFamily = "TwainDotNet",
            ProductName = "TwainDotNet",
        };
    }
}
