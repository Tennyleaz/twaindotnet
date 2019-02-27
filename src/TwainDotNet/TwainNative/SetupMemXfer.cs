using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;

namespace TwainDotNet.TwainNative
{
    /// <summary>
    /// used with  DG_CONTROL / DAT_SETUPMEMXFER / MSG_GET
    /// typedef struct {
    ///     TW_UINT32 MinBufSize /* Minimum buffer size in bytes */
    ///     TW_UINT32 MaxBufSize /* Maximum buffer size in bytes */
    ///     TW_UINT32 Preferred /* Preferred buffer size in bytes */
    /// } TW_SETUPMEMXFER, FAR* pTW_SETUPMEMXFER;
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 2)]
    public class SetupMemXfer
    {
        /// <summary>
        /// Minimum buffer size in bytes
        /// </summary>
        public uint MinBufSize;
        /// <summary>
        ///  Maximum buffer size in bytes
        /// </summary>
        public uint MaxBufSize;
        /// <summary>
        /// Preferred buffer size in bytes
        /// </summary>
        public uint Preferred;
    }
}
