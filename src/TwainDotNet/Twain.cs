using System;
using System.Collections.Generic;
using System.Text;
using System.Reflection;
using System.Runtime.InteropServices;
using TwainDotNet.TwainNative;
using TwainDotNet.Win32;
using System.Drawing;

namespace TwainDotNet
{
    public class Twain : IDisposable
    {
        private DataSourceManager _dataSourceManager;

        public bool IsIncremental { get; private set; }

        /// <summary>
        /// Twain dll的最外層。當建構子完成，代表進入twain state 3，接下來就可以去open data souse了。
        /// </summary>
        /// <param name="messageHook"></param>
        /// <param name="useIncremental">使用一次性傳輸或是buffered memory mode傳輸</param>
        public Twain(IWindowsMessageHook messageHook, bool useIncremental = true)
        {
            ScanningComplete += delegate { };
            TransferImage += delegate { };

            _dataSourceManager = new DataSourceManager(DataSourceManager.DefaultApplicationId, messageHook);
            IsIncremental = useIncremental;
            _dataSourceManager.UseIncrementalMemoryXfer = useIncremental;
            _dataSourceManager.ScanningComplete += delegate(object sender, ScanningCompleteEventArgs args)
            {
                ScanningComplete(this, args);
            };
            _dataSourceManager.TransferImage += delegate(object sender, TransferImageEventArgs args)
            {
                TransferImage(this, args);
            };
        }

        /// <summary>
        /// Notification that the scanning has completed.
        /// </summary>
        public event EventHandler<ScanningCompleteEventArgs> ScanningComplete;

        public event EventHandler<TransferImageEventArgs> TransferImage;

        /// <summary>
        /// Starts scanning.
        /// </summary>
        public void StartScanning(ScanSettings settings)
        {
            _dataSourceManager.StartScan(settings);
        }

        /// <summary>
        /// Shows a dialog prompting the use to select the source to scan from.
        /// </summary>
        public void SelectSource()
        {
            _dataSourceManager.SelectSource();
        }

        /// <summary>
        /// Selects a source based on the product name string.
        /// </summary>
        /// <param name="sourceName">The source product name.</param>
        public void SelectSource(string sourceName)
        {
            var source = DataSource.GetSource(
                sourceName,
                _dataSourceManager.ApplicationId,
                _dataSourceManager.MessageHook);

            _dataSourceManager.SelectSource(source);
        }

        /// <summary>
        /// Gets the product name for the default source.
        /// </summary>
        public string DefaultSourceName
        {
            get
            {
                using (var source = DataSource.GetDefault(_dataSourceManager.ApplicationId, _dataSourceManager.MessageHook))
                {
                    return source.SourceId.ProductName;
                }
            }
        }

        /// <summary>
        /// Gets a list of source product names.
        /// </summary>
        public IList<string> SourceNames
        {
            get
            {
                var result = new List<string>();
                var sources = DataSource.GetAllSources(
                    _dataSourceManager.ApplicationId,
                    _dataSourceManager.MessageHook);

                foreach (var source in sources)
                {
                    result.Add(source.SourceId.ProductName);
                    source.Dispose();
                }

                return result;
            }
        }

        /// <summary>
        /// 偵測有沒有紙，需要先select一個data source。
        /// </summary>
        /// <returns></returns>
        /// <exception cref="PaperJamException">在data source開啟後，讀取紙張問題，就會擲回</exception>
        public bool IsPaperOn()
        {
            return _dataSourceManager.IsPaperOn();
        }

        public void Dispose()
        {
            ((IDisposable)_dataSourceManager).Dispose();
            _dataSourceManager = null;
            ScanningComplete = null;
            TransferImage = null;
        }
    }
}
