using System;

namespace WinSerial.Library
{
    [Serializable]
    public class WinSerialException : Exception
    {
        public WinSerialException() { }
        public WinSerialException(string message) : base(message) { }
        public WinSerialException(string message, Exception inner) : base(message, inner) { }
        protected WinSerialException(
          System.Runtime.Serialization.SerializationInfo info,
          System.Runtime.Serialization.StreamingContext context) : base(info, context) { }
    }

}
