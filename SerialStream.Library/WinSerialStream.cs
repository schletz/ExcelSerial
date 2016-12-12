using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace WinSerial.Library
{

    public class WinSerialStream : Stream
    {
        private bool disposed = false;

        public WinSerialPort WinSerialPort { get; }
        /// <summary>
        /// Liefert true.
        /// </summary>
        public override bool CanRead { get; } = true;
        /// <summary>
        /// Liefert immer false, da kein wahlfreier Zugriff möglich ist.
        /// </summary>
        public override bool CanSeek { get; } = false;
        /// <summary>
        /// Liefert immer false, da Schreibfunktionen nicht implementiert sind.
        /// </summary>
        public override bool CanWrite { get; } = false;
        public override int ReadTimeout
        {
            set
            {
                NativeMethods.CommTimeouts commTimeouts = NativeMethods.CommTimeouts.Factory;
                commTimeouts.ReadTotalTimeoutConstant = Convert.ToUInt32(value);
                if (!NativeMethods.SetCommTimeouts(WinSerialPort.HComPort, ref commTimeouts))
                    throw new WinSerialException("Fehler beim Setzen des Verbindungstimeouts.");
            }
        }
        public override int WriteTimeout
        {
            set
            {
                NativeMethods.CommTimeouts commTimeouts = NativeMethods.CommTimeouts.Factory;
                commTimeouts.WriteTotalTimeoutConstant = Convert.ToUInt32(value);
                if (!NativeMethods.SetCommTimeouts(WinSerialPort.HComPort, ref commTimeouts))
                    throw new WinSerialException("Fehler beim Setzen des Verbindungstimeouts.");
            }
        }

        public override long Length
        {
            get { throw new NotSupportedException(); }
        }

        public override long Position
        {
            get { throw new NotSupportedException(); }

            set { throw new NotSupportedException(); }
        }

        public override void Flush()
        {
            NativeMethods.PurgeComm(WinSerialPort.HComPort,
                0x0008 |   // Clears the input buffer (if the device driver has one).
                0x0004);   // Clears the output buffer (if the device driver has one).
        }

        public override int Read(byte[] buffer, int offset, int count)
        {
            byte[] readBuffer = new byte[count];
            UInt32 bytesRead;

            if (!NativeMethods.ReadFile(WinSerialPort.HComPort, readBuffer,
                Convert.ToUInt32(count), out bytesRead, IntPtr.Zero))
                throw new WinSerialException("Fehler beim Lesen.");
            Buffer.BlockCopy(readBuffer, 0, buffer, offset, (int)bytesRead);
            return (int)bytesRead;
        }

        public override long Seek(long offset, SeekOrigin origin)
        {
            throw new NotSupportedException();
        }

        public override void SetLength(long value)
        {
            throw new NotSupportedException();
        }

        public override void Write(byte[] buffer, int offset, int count)
        {
            throw new NotSupportedException();
        }

        public WinSerialStream(WinSerialPort winSerialPort)
        {
            this.WinSerialPort = winSerialPort;
        }
        public override void Close()
        {
            WinSerialPort?.Close();
        }
        protected override void Dispose(bool disposing)
        {
            if (disposed)
                return;
            if (disposing)
            {
                // Free any other managed objects here.
            }
            // Free any unmanaged objects here.
            WinSerialPort?.Close();
            disposed = true;
            // Call base class implementation.
            base.Dispose(disposing);
        }
        ~WinSerialStream()
        {
            Dispose(false);
        }
    }
}
