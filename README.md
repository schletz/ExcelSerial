# ExcelSerial
Bietet ein Add In für Microsoft Excel an, um von der seriellen Schnittstelle zu lesen. Zum Auslesen des COM Ports wurde die Bibliothek SerialStream entwickelt. Diese kapselt die Win32 API Funktionen in den Klassen WinSerialPort und WinSerialStream, die von Stream abgeleitet ist.

```c#
WinSerialPort sp = new WinSerialPort() { ComSettings = "480600/8-N-1" };
// 1000 ms Read Timeout
using (StreamReader sr = new StreamReader(sp.Open(1000), Encoding.ASCII, false, 1024))
{
  string line;
  while ((line = sr.ReadLine()) != null) 
  {
      Console.WriteLine(line);
  }
}
```
##Todo
- Port Auto Discovery (Problem dabei: Ports können auch geöffnet werden, wenn kein Gerät angeschlossen ist).
- Aktualisieren der Portliste nach einem bestimmten Intervall, wenn z. B. ein USB Konverter angeschlossen wird.

##Bugs
- Notfication bei Fehlern geht unter Windows 10, bei 8.1 aber nicht.
