using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelSerial
{
    /// <summary>
    /// Parsingfunktionen, die eine gelesene Zeile weiterverarbeiten.
    /// </summary>
    public static class BufferParser
    {
        private static byte[] base64CompareChars = Encoding.ASCII.GetBytes("AZaz09+/");

        /// <summary>
        /// Führt eine Base64 Decodierung durch. Dabei wird allerdings nur der String als Zahl im
        /// Basissystem 64 implementiert. Daher weicht diese Implementierung von der "normalen"
        /// Base64 Decodierung, die Füllbytes beachtet, ab.
        /// Der String darf maximal 10 Stellen lang sein, da ein 64 Bit Wert ausgegeben wird.
        /// </summary>
        /// <param name="val">Der Base64 Codierte string, bestehend aus den ASCII Zeichen 
        /// A-Z, a-z, 0-9, + und /</param>
        /// <param name="decoded">Die decodierte, vorzeichenlose 64bit Zahl.</param>
        /// <returns>False wenn der String zu lange oder ungültige Zeichen enthält.</returns>
        public static bool Base64Decode(string val, out ulong decoded)
        {
            decoded = 0;
            if (val.Length > 10) return false;
            foreach (byte zeichen in Encoding.ASCII.GetBytes(val))
            {
                if (zeichen >= base64CompareChars[0] && zeichen <= base64CompareChars[1])
                    decoded = decoded * 64 + (ulong)(zeichen - base64CompareChars[0]);
                else if (zeichen >= base64CompareChars[2] && zeichen <= base64CompareChars[3])
                    decoded = decoded * 64 + (ulong)(zeichen - base64CompareChars[2] + 26);
                else if (zeichen >= base64CompareChars[4] && zeichen <= base64CompareChars[5])
                    decoded = decoded * 64 + (ulong)(zeichen - base64CompareChars[4] + 52);
                else if (zeichen == base64CompareChars[6])
                    decoded = decoded * 64 + 62;
                else if (zeichen == base64CompareChars[7])
                    decoded = decoded * 64 + 63;
                else
                    return false;
            }
            return true;
        }

        /// <summary>
        /// Führt für jede Zeile ggf. eine Base64 Decodierung durch und liefert das Ergebnis als
        /// object Array, damit dies in Excel 1:1 eingefügt werden kann.
        /// </summary>
        /// <param name="buffer">Das Array mit den zu verarbeitenden Zeilen.</param>
        /// <param name="bufferCount">Die Anzahl der Zeilen, die im Array verarbeitet werdne 
        /// sollen.</param>
        /// <param name="base64Decode">Wird dieses Flag gesetzt, so werden die Werte Base64 
        /// decodiert. Ist dies nicht möglich, wird der Originalstring beibehalten.</param>
        /// <returns>Ein object Array mit 1 Spalte.</returns>
        public static object[,] Parse(string[] buffer, int bufferCount, bool base64Decode = false)
        {
            object[,] parsedBuffer = new object[bufferCount, 1];
            ulong decoded;

            for (int zeile = 0; zeile < bufferCount; zeile++)
            {
                if (base64Decode && Base64Decode(buffer[zeile], out decoded))
                {
                    parsedBuffer[zeile, 0] = decoded;
                }
                else
                {
                    parsedBuffer[zeile, 0] = buffer[zeile];
                }
            }
            return parsedBuffer;
        }

        /// <summary>
        /// Trennt jede Zeile anhand eines Trennzeichens auf, führt für jede Zeile ggf. eine 
        /// Base64 Decodierung durch und liefert das Ergebnis als object Array, damit dies in 
        /// Excel 1:1 eingefügt werden kann.
        /// </summary>
        /// <param name="buffer">Das Array mit den zu verarbeitenden Zeilen.</param>
        /// <param name="bufferCount">Die Anzahl der Zeilen, die im Array verarbeitet werdne 
        /// sollen.</param>
        /// <param name="seperator">Das zu verwendende Trennzeichen.</param>
        /// <param name="base64Decode">Wird dieses Flag gesetzt, so werden die Werte Base64 
        /// decodiert. Ist dies nicht möglich, wird der Originalstring beibehalten.</param>
        /// <returns>Ein object Array mit sovielen Spalten, wie es die Zeile mit den meisten
        /// Trennzeichen erfordert.</returns>
        public static object[,] Parse(string[] buffer, int bufferCount,
            char seperator, bool base64Decode = false)
        {
            /* Da die Anzahl der Trennzeichen unterschiedlich sein kann, nehmen wir die maximale
             * Spaltenanzahl. */
            int colCount = (from b in buffer
                            where b != null
                            select b.Count(c => c == seperator)).DefaultIfEmpty(0).Max()+1;

            object[,] parsedBuffer = new object[bufferCount, colCount];
            ulong decoded;

            for (int zeile = 0; zeile < bufferCount; zeile++)
            {
                int col = 0;
                foreach (string colStr in buffer[zeile].Split(seperator))
                {
                    if (base64Decode && Base64Decode(colStr, out decoded))
                        parsedBuffer[zeile, col++] = decoded;
                    else
                        parsedBuffer[zeile, col++] = colStr;
                }

            }
            return parsedBuffer;
        }

        /// <summary>
        /// Trennt jede Zeile nach einer fixen Anzahl von Stellen auf, führt für jede Zeile ggf.
        /// eine  Base64 Decodierung durch und liefert das Ergebnis als object Array, damit dies in 
        /// Excel 1:1 eingefügt werden kann.
        /// Ist der String zu kurz, um alle Spalten zu befüllen, wird ein Leerstring eingefügt.
        /// Es werden nur die Anzahl der Zeichen verarbeitet, die im Längenarray angegeben wurden.
        /// </summary>
        /// <param name="buffer">Das Array mit den zu verarbeitenden Zeilen.</param>
        /// <param name="bufferCount">Die Anzahl der Zeilen, die im Array verarbeitet werdne 
        /// sollen.</param>
        /// <param name="fieldLengths">Ein Array mit den Längen der einzelnen Spalten.</param>
        /// <param name="base64Decode">Wird dieses Flag gesetzt, so werden die Werte Base64 
        /// decodiert. Ist dies nicht möglich, wird der Originalstring beibehalten.</param>
        /// <returns>Ein object Array mit sovielen Spalten, wie es die Zeile mit den meisten
        /// Trennzeichen erfordert.</returns>
        public static object[,] Parse(string[] buffer, int bufferCount,
            int[] fieldLengths, bool base64Decode = false)
        {
            object[,] parsedBuffer = new object[bufferCount, fieldLengths.Length];
            ulong decoded;

            for (int zeile = 0; zeile < bufferCount; zeile++)
            {
                string line = buffer[zeile];
                int lineLength = line.Length;
                int strPos = 0, col = 0;
                foreach (int colLen in fieldLengths)
                {
                    if (strPos < lineLength)
                    {
                        string colStr = line.Substring(strPos,
                            Math.Min(lineLength - strPos, colLen));
                        if (base64Decode && Base64Decode(colStr, out decoded))
                            parsedBuffer[zeile, col] = decoded;
                        else
                            parsedBuffer[zeile, col] = colStr;
                        strPos += colLen;
                    }
                    else
                    {
                        parsedBuffer[zeile, col] = "";
                    }
                    col++;
                }
            }
            return parsedBuffer;
        }
    }
}
