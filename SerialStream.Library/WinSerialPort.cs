using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Permissions;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace WinSerial.Library
{
    /// <summary>
    /// Emumeration für die Anzahl der Stoppbits. Erlaubt sind 1, 1.5 oder 2.
    /// </summary>
    public enum StopBits { One = 0, OnePointFive = 1, Two = 2 }
    public enum Parity { None = 0, Odd = 1, Even = 2, Mark = 3, Space = 4 }
    /// <summary>
    /// Repräsentiert den seriellen Port.
    /// </summary>
    public class WinSerialPort
    {
        /// <summary>
        /// Handle auf den geöffneten seriellen Anschluss. Ist IntPtr.Zero, wenn der Port 
        /// geschlossen ist.
        /// </summary>
        public IntPtr HComPort { get; private set; } = IntPtr.Zero;
        /// <summary>
        /// Der Name des Ports. Erlaubt ist COM1...COM99
        /// </summary>
        public string PortName
        {
            get { return portName; }
            set
            {
                if (Regex.IsMatch(value, @"^COM[1-9][0-9]?$", RegexOptions.IgnoreCase))
                    portName = value;
                else
                    throw new WinSerialException("Der Portname " + value + " ist ungültig.");
            }
        }
        private string portName = "";
        /// <summary>
        /// Die verwendete Baudrate. Diese muss vom Treiber unterstützt werden.
        /// </summary>
        public int BaudRate { get; set; }
        /// <summary>
        /// Die Anzahl der Datenbits. Diese muss vom Treiber unterstützt werden.
        /// </summary>
        public int DataBits { get; set; }
        /// <summary>
        /// Die Anzahl der Stoppbits. Möglich sind 1, 1.5 oder 2.
        /// </summary>
        public StopBits StopBits { get; set; }
        /// <summary>
        /// Die verwendete Parität. Möglich ist None, Odd, Even, Mark oder Space.
        /// </summary>
        public Parity Parity { get; set; }
        /// <summary>
        /// Liefert true, wenn der Port geöffnet ist. False wenn der Port geschlossen ist.
        /// </summary>
        public bool IsOpen { get { return HComPort != IntPtr.Zero; } }
        /// <summary>
        /// Setzt einen String im Format Baudrate/DataBits-Parity-Stoppbits.
        /// Parity kann N, O, E, M, S sein.
        /// Die Stoppbits können 1, 1.5 oder 2 sein.
        /// Beispiel: 115200/8-E-1.5
        /// </summary>
        public string ComSettings
        {
            get
            {
                return String.Format("{0}/{1}-{2}-{3}",
                    BaudRate,
                    DataBits,
                    Parity.ToString()[0],
                    StopBits == StopBits.One ? "1" : 
                    StopBits == StopBits.OnePointFive ? "1.5" : "2");
            }
            set
            {
                Match match = Regex.Match(value.ToUpper(),
                    @"^([0-9]{3,6})\/([0-9]{1,2})-([NOEMS])-(1|1.5|2)$");
                if (!match.Success)
                    throw new WinSerialException("Die Einstellungen für die COM Schnittstelle haben das falsche Format.");

                BaudRate = int.Parse(match.Groups[1].Value);
                DataBits = int.Parse(match.Groups[2].Value);
                switch (match.Groups[3].Value)
                {
                    case "N": Parity = Parity.None; break;
                    case "E": Parity = Parity.Even; break;
                    case "O": Parity = Parity.Odd; break;
                    case "M": Parity = Parity.Mark; break;
                    case "S": Parity = Parity.Space; break;
                }
                switch (match.Groups[4].Value)
                {
                    case "1": StopBits = StopBits.One; break;
                    case "1.5": StopBits = StopBits.OnePointFive; break;
                    case "2": StopBits = StopBits.Two; break;
                }
            }
        }

        private NativeMethods.Dcb dcb;
        private NativeMethods.CommTimeouts commTimeouts;
        /// <summary>
        /// Öffnet den Port und gibt ein Objekt vom Typ WinSerialStream zurück, von dem dann 
        /// gelesen werden kann. Das Timeout wird auf 10 Sekunden gesetzt.
        /// </summary>
        /// <returns></returns>
        public WinSerialStream Open()
        {
            return Open(10000, 10000);
        }
        /// <summary>
        /// Öffnet den Port und gibt ein Objekt vom Typ WinSerialStream zurück, von dem dann 
        /// gelesen werden kann. Das Timeout für Schreibvorgänge wird auf 10 Sekunden gesetzt.
        /// </summary>
        /// <param name="readTimeout">Timeout für Lesevorgänge in ms.</param>
        /// <returns></returns>
        public WinSerialStream Open(int readTimeout)
        {
            return Open(readTimeout, 10000);
        }
        /// <summary>
        /// Öffnet den Port und gibt ein Objekt vom Typ WinSerialStream zurück, von dem dann 
        /// gelesen werden kann.
        /// </summary>
        /// <param name="readTimeout">Timeout für Lesevorgänge in ms.</param>
        /// <param name="writeTimeout">Timeout für Schreibvorgänge in ms.</param>
        /// <returns></returns>
        public WinSerialStream Open(int readTimeout, int writeTimeout)
        {
            dcb = NativeMethods.Dcb.Factory;
            commTimeouts = NativeMethods.CommTimeouts.Factory;

            /* https://msdn.microsoft.com/en-us/library/windows/desktop/aa363858(v=vs.85).aspx */
            HComPort = NativeMethods.CreateFile(@"\\.\" + portName,        // File Name
                NativeMethods.GenericRead | NativeMethods.GenericWrite,    // Access Mode
                0,            // Share Mode: Prevents other processes from opening
                IntPtr.Zero,  // No Security Attributes
                3,            // Opens a file or device, only if it exists.
                0,            // Non overlapping read/write
                IntPtr.Zero   // No template file for attributes
            );

            try
            {
                if (NativeMethods.IsInvalidHandle(HComPort))
                    throw new WinSerialException("Fehler beim Öffnen des Ports.");
                if (!NativeMethods.GetCommState(HComPort, ref dcb))
                    throw new WinSerialException("Fehler beim Lesen der COM Einstellungen.");

                dcb.BaudRate = Convert.ToUInt32(BaudRate);
                dcb.ByteSize = Convert.ToByte(DataBits);
                dcb.StopBits = Convert.ToByte(StopBits);
                dcb.Parity = Convert.ToByte(Parity);
                if (Parity != Parity.None)
                    dcb.Flags |= NativeMethods.DcbFlags.ParityFlag;

                if (!NativeMethods.SetCommState(HComPort, ref dcb))
                    throw new WinSerialException("Fehler beim Setzen der COM Einstellungen.");


            }
            catch (Exception e)
            {
                NativeMethods.CloseHandle(HComPort);
                throw e;
            }

            WinSerialStream stream = new WinSerialStream(this)
            {
                ReadTimeout = readTimeout,
                WriteTimeout = writeTimeout
            };
            return stream;
        }

        /// <summary>
        /// Schließt den Port und setzt den Handle auf IntPtr.Zero.
        /// </summary>
        public void Close()
        {
            NativeMethods.CloseHandle(HComPort);
            HComPort = IntPtr.Zero;
        }

        /// <summary>
        /// Listet alle verfügbaren COM Schnittstellen auf. Die Implementierung wurde von
        /// https://referencesource.microsoft.com/#System/sys/system/io/ports/SerialPort.cs
        /// übernommen.
        /// </summary>
        /// <returns>Ein Array mit allen COM Schnittstellen.</returns>
        public static string[] GetPortNames()
        {
            RegistryKey baseKey = null;
            RegistryKey serialKey = null;

            String[] portNames = null;

            RegistryPermission registryPermission = new RegistryPermission(RegistryPermissionAccess.Read,
                                    @"HKEY_LOCAL_MACHINE\HARDWARE\DEVICEMAP\SERIALCOMM");
            registryPermission.Assert();

            try
            {
                baseKey = Registry.LocalMachine;
                serialKey = baseKey.OpenSubKey(@"HARDWARE\DEVICEMAP\SERIALCOMM", false);

                if (serialKey != null)
                {

                    string[] deviceNames = serialKey.GetValueNames();
                    portNames = new String[deviceNames.Length];

                    for (int i = 0; i < deviceNames.Length; i++)
                        portNames[i] = (string)serialKey.GetValue(deviceNames[i]);
                }
            }
            finally
            {
                if (baseKey != null)
                    baseKey.Close();

                if (serialKey != null)
                    serialKey.Close();

                RegistryPermission.RevertAssert();
            }

            // If serialKey didn't exist for some reason
            if (portNames == null)
                portNames = new String[0];

            Array.Sort(portNames);
            return portNames;
        }
    }
}
