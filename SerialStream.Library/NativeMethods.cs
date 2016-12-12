using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace WinSerial.Library
{
    internal static class NativeMethods
    {
        [Flags]
        public enum DcbFlags
        {
            None = 0,
            /* If this member is TRUE, binary mode is enabled. Windows does not support nonbinary 
             * mode transfers, so this member must be TRUE. */
            Binary = 1 << 0,
            /* If this member is TRUE, parity checking is performed and errors are reported. */
            ParityFlag = 1 << 1,
            /* If this member is TRUE, the CTS (clear-to-send) signal is monitored for output flow 
             * control. If this member is TRUE and CTS is turned off, output is suspended until CTS 
             * is sent again. */
            OutxCtsFlow = 1 << 2,
            /* If this member is TRUE, the DSR (data-set-ready) signal is monitored for output flow 
             * control. If this member is TRUE and DSR is turned off, output is suspended until DSR 
             * is sent again. */
            OutxDsrFlow = 1 << 3,
            /* The DTR (data-terminal-ready) flow control. This member can be one of the following
             * values.
             * DTR_CONTROL_DISABLE 0x00: Disables the DTR line when the device is opened and 
             *    leaves it disabled.
             * DTR_CONTROL_ENABLE 0x01: Enables the DTR line when the device is opened and leaves 
             *    it on.
             * DTR_CONTROL_HANDSHAKE 0x02 */
            DtrControl1 = 1 << 4,
            DtrControl2 = 1 << 5,
            /* If this member is TRUE, the communications driver is sensitive to the state of the
             * DSR signal. The driver ignores any bytes received, unless the DSR modem input line is
             *  high. */
            DsrSensitivity = 1 << 6,
            /* If this member is TRUE, transmission continues after the input buffer has come within 
             * XoffLim bytes of being full and the driver has transmitted the XoffChar character to 
             * stop receiving bytes. If this member is FALSE, transmission does not continue until 
             * the input buffer is within XonLim bytes of being empty and the driver has transmitted 
             * the XonChar character to resume reception. */
            TXContinueOnXoff = 1 << 7,
            /* Indicates whether XON/XOFF flow control is used during transmission. If this member 
             * is TRUE, transmission stops when the XoffChar character is received and starts again 
             * when the XonChar character is received. */
            OutX = 1 << 8,
            /* Indicates whether XON/XOFF flow control is used during reception. If this member is 
             * TRUE, the XoffChar character is sent when the input buffer comes within XoffLim bytes 
             * of being full, and the XonChar character is sent when the input buffer comes within
             *  XonLim bytes of being empty. */
            InX = 1 << 9,
            /* Indicates whether bytes received with parity errors are replaced with the character 
             * specified by the ErrorChar member. If this member is TRUE and the fParity member is 
             * TRUE, replacement occurs. */
            ErrorCharFlag = 1 << 10,
            /* If this member is TRUE, null bytes are discarded when received. */
            Null = 1 << 11,
            /* The RTS (request-to-send) flow control. This member can be one of the following 
             * values.
             * RTS_CONTROL_DISABLE 0x00 Disables the RTS line when the device is opened and leaves 
             *    it disabled.
             * RTS_CONTROL_ENABLE 0x01 Enables the RTS line when the device is opened and leaves it
             *    on.
             * RTS_CONTROL_HANDSHAKE 0x02 Enables RTS handshaking. The driver raises the RTS line 
             *    when the "type-ahead" (input) buffer is less than one-half full and lowers the RTS 
             *    line when the buffer is more than three-quarters full. If handshaking is enabled, 
             *    it is an error for the application to adjust the line by using the 
             *    EscapeCommFunction function.
             * RTS_CONTROL_TOGGLE 0x03 Specifies that the RTS line will be high if bytes are 
             *    available for transmission. After all buffered bytes have been sent, the RTS line 
             *    will be low. */
            RtsControl1 = 1 << 12,
            RtsControl2 = 1 << 13,
            /* If this member is TRUE, the driver terminates all read and write operations with an 
             * error status if an error occurs. The driver will not accept any further 
             * communications operations until the application has acknowledged the error by calling 
             * the ClearCommError function. */
            AbortOnError = 1 << 14,
            Dummy2 = 1 << 15,
        };

        /// <summary>
        /// Bildet die DCB Struktur zum Einstellen der Verbindungsparameter ab.
        /// Siehe https://msdn.microsoft.com/en-us/library/windows/desktop/aa363214(v=vs.85).aspx
        /// </summary>
        [StructLayout(LayoutKind.Explicit)]
        public struct Dcb
        {
            /* The length of the structure, in bytes. The caller must set this member to 
             * sizeof(DCB). */
            [FieldOffset(0)] public UInt32 Dcblength;
            /* The baud rate at which the communications device operates. This member can be an 
            actual baud rate value (9600, 115200, 460800, ...) */
            [FieldOffset(4)] public UInt32 BaudRate;
            /* Flags */
            [FieldOffset(8)] public DcbFlags Flags;
            /* Flags ende */
            /* Reserved; must be zero. */
            [FieldOffset(12)] public UInt16 wReserved;
            /* The minimum number of bytes in use allowed in the input buffer before flow control 
             * is activated to allow transmission by the sender. This assumes that either XON/XOFF, 
             * RTS, or DTR input flow control is specified in the fInX, fRtsControl, or fDtrControl 
             * members. */
            [FieldOffset(14)] public UInt16 XonLim;
            /* The minimum number of free bytes allowed in the input buffer before flow control is 
             * activated to inhibit the sender. Note that the sender may transmit characters after 
             * the flow control signal has been activated, so this value should never be zero. This 
             * assumes that either XON/XOFF, RTS, or DTR input flow control is specified in the 
             * fInX, fRtsControl, or fDtrControl members. The maximum number of bytes in use allowed 
             * is calculated by subtracting this value from the size, in bytes, of the input 
             * buffer. */
            [FieldOffset(16)] public UInt16 XoffLim;
            /* The number of bits in the bytes transmitted and received. 8 for 8-N-1, ...*/
            [FieldOffset(18)] public byte ByteSize;
            /* The parity scheme to be used. This member can be one of the following values.
             * EVENPARITY  0x2 Even parity.
             * MARKPARITY  0x3 Mark parity.
             * NOPARITY    0x0 No parity.
             * ODDPARITY   0x1 Odd parity.
             * SPACEPARITY 0x4 Space parity. */
            [FieldOffset(19)] public byte Parity;
            /* The number of stop bits to be used. This member can be one of the following values.
             * ONESTOPBIT   0x0 1 stop bit.
             * ONE5STOPBITS 0x1 1.5 stop bits.
             * TWOSTOPBITS  0x2 2 stop bits. */
            [FieldOffset(20)] public byte StopBits;
            /* The value of the XON character for both transmission and reception. */
            [FieldOffset(21)] public byte XonChar;
            /* The value of the XOFF character for both transmission and reception. */
            [FieldOffset(22)] public byte XoffChar;
            /* The value of the character used to replace bytes received with a parity error. */
            [FieldOffset(23)] public byte ErrorChar;
            /* The value of the character used to signal the end of data. */
            [FieldOffset(24)] public byte EofChar;
            /* The value of the character used to signal an event. */
            [FieldOffset(25)] public byte EvtChar;
            /* Reserved; do not use. */
            [FieldOffset(26)] public UInt16 wReserved1;
            public static Dcb Factory
            {
                get
                {
                    Dcb dcb = new Dcb();
                    dcb.Dcblength = (UInt32)Marshal.SizeOf(dcb);
                    dcb.Flags = DcbFlags.Binary;
                    return dcb;
                }
            }
        };  // public struct Dcb

        /// <summary>
        /// Siehe https://msdn.microsoft.com/de-de/library/windows/desktop/aa363190(v=vs.85).aspx
        /// </summary>
        [StructLayout(LayoutKind.Explicit)]
        public struct CommTimeouts {
            /* The maximum time allowed to elapse before the arrival of the next byte on the 
             * communications line, in milliseconds. If the interval between the arrival of any two 
             * bytes exceeds this amount, the ReadFile operation is completed and any buffered data 
             * is returned. A value of zero indicates that interval time-outs are not used. A value 
             * of MAXDWORD, combined with zero values for both the ReadTotalTimeoutConstant and 
             * ReadTotalTimeoutMultiplier members, specifies that the read operation is to return 
             * immediately with the bytes that have already been received, even if no bytes have 
             * been received. */
            [FieldOffset(0)] public UInt32 ReadIntervalTimeout;
            /* The multiplier used to calculate the total time-out period for read operations, in 
             * milliseconds. For each read operation, this value is multiplied by the requested 
             * number of bytes to be read. */
            [FieldOffset(4)] public UInt32 ReadTotalTimeoutMultiplier;
            /* A constant used to calculate the total time-out period for read operations, in 
             * milliseconds. For each read operation, this value is added to the product of the 
             * ReadTotalTimeoutMultiplier member and the requested number of bytes. A value of zero 
             * for both the ReadTotalTimeoutMultiplier and ReadTotalTimeoutConstant members 
             * indicates that total time-outs are not used for read operations. */
            [FieldOffset(8)] public UInt32 ReadTotalTimeoutConstant;
            /* The multiplier used to calculate the total time-out period for write operations, in 
             * milliseconds. For each write operation, this value is multiplied by the number of 
             * bytes to be written. */
            [FieldOffset(12)] public UInt32 WriteTotalTimeoutMultiplier;
            /* A constant used to calculate the total time-out period for write operations, in 
             * milliseconds. For each write operation, this value is added to the product of the 
             * WriteTotalTimeoutMultiplier member and the number of bytes to be written. A value of 
             * zero for both the WriteTotalTimeoutMultiplier and WriteTotalTimeoutConstant members 
             * indicates that total time-outs are not used for write operations. */
            [FieldOffset(16)] public UInt32 WriteTotalTimeoutConstant;
            public static CommTimeouts Factory
            {
                get
                {
                    return new CommTimeouts() {
                        ReadTotalTimeoutConstant = 1000,
                        WriteTotalTimeoutConstant = 1000
                    };
                }
            }
        }; // public struct CommTimeouts

        /* 
         * METHODEN
         */
        [DllImport("kernel32.dll", CharSet = CharSet.Unicode)]
        public static extern IntPtr CreateFile(string lpFileName, UInt32 dwDesiredAccess, 
            UInt32 dwShareMode, IntPtr lpSecurityAttributes, UInt32 dwCreationDisposition, 
            UInt32 dwFlagsAndAttributes, IntPtr hTemplateFile);

        [DllImport("kernel32.dll")]
        public static extern bool GetCommState(IntPtr hFile, ref Dcb lpDCB);

        [DllImport("kernel32.dll")]
        public static extern bool SetCommState(IntPtr hFile, ref Dcb lpDCB);

        [DllImport("kernel32.dll")]
        public static extern bool SetCommTimeouts(IntPtr hFile, [In] ref CommTimeouts lpCommTimeouts);

        [DllImport("kernel32.dll")]
        public static extern bool ReadFile(IntPtr hFile, byte[] lpBuffer, 
            UInt32 nNumberOfBytesToRead, out UInt32 lpNumberOfBytesRead, IntPtr lpOverlapped);

        [DllImport("kernel32.dll")]
        public static extern bool CloseHandle(IntPtr hHandle);

        [DllImport("kernel32.dll")]
        public static extern bool PurgeComm(IntPtr hFile, UInt32 dwFlags);


        /*
         * KONSTANTEN
         */
        public const UInt32 GenericRead = 0x80000000;
        public const UInt32 GenericWrite = 0x40000000;

        /*
         * HILFSMETHODEN
         */
        public static bool IsInvalidHandle(IntPtr handle)
        {
            return handle.ToInt32() == -1;
        }
    }
}
