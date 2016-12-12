using System;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.IO;
using System.Threading.Tasks;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Tools.Excel;

using WinSerial.Library;

namespace ExcelSerial
{
    public partial class SerialRibbon
    {
        WinSerialPort sp = new WinSerialPort();
        Task readTask, parseTask;

        Worksheet sheet;
        int aktZeile;
        DateTime lastError = DateTime.MinValue;

        private char csvDelimiter
        {
            get
            {
                if (!chkCsv.Checked || txtSeperator.Text.Length == 0) return '\0';
                return txtSeperator.Text[0];
            }
        }
        private int[] fieldsLength
        {
            get
            {
                if (!chkFixLength.Checked || txtLength.Text == "") return new int[] { int.MaxValue };
                return txtLength.Text.Split(' ', '-', '|', ',', ';')
                    .Select(x => int.Parse(x))
                    .ToArray();
            }
        }
        private bool base64Decode
        {
            get
            {
                return chkBase64.Checked;
            }
        }

        private bool recording
        {
            get { return !cmdRec.Enabled; }
            set { cmdRec.Enabled = !value; }
        }

        private void SerialRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            try
            {
                /* Die Liste aller COM Ports in das Dropdownfeld schreiben */
                foreach (string port in WinSerialPort.GetPortNames().OrderBy(p => p))
                {
                    var item = this.Factory.CreateRibbonDropDownItem();
                    item.Label = port;
                    lstPort.Items.Add(item);
                }
            }
            catch (Exception err)
            {
                writeError(err);
            }
        }

        /*
         * EVENTHANDLER
         */
        private void cmdRec_Click(object sender, RibbonControlEventArgs e)
        {
            sheet = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveSheet);
            try
            {
                /* Bei der 2. Zeile im Arbeitsblatt beginnen. */
                aktZeile = 2;
                sp.PortName = lstPort.SelectedItem.Label;
                sp.ComSettings = txtComSettings.Text;
                /* Liest maximal 1000 Datensätze oder führt spätestens nach 1 Sekunde die
                 * Verarbeitung durch. */
                readTask = new Task (()=>bufferedReadAsync(1000, TimeSpan.FromSeconds(1)));
                readTask.Start();
            }
            catch (Exception err)
            {
                recording = false;
                writeError(err);
            }
        }
        private void cmdStop_Click(object sender, RibbonControlEventArgs e)
        {
            recording = false;
        }

        private void chkCsv_Click(object sender, RibbonControlEventArgs e)
        {
            txtSeperator.Enabled = chkCsv.Checked;
            if (chkCsv.Checked)
            {
                chkFixLength.Checked = false;
                txtLength.Enabled = false;
                txtSeperator.Enabled = true;
            }

        }

        private void chkFixLength_Click(object sender, RibbonControlEventArgs e)
        {
            txtLength.Enabled = chkFixLength.Checked;
            if (chkFixLength.Checked)
            {
                chkCsv.Checked = false;
                txtSeperator.Enabled = false;
                txtLength.Enabled = true;
            }

        }

        private void writeRecords(object[,] data)
        {
            try
            {
                sheet.Range[
                    sheet.Cells[aktZeile, 1],
                    sheet.Cells[aktZeile + data.GetLength(0) - 1, data.GetLength(1)]].Value2 =
                    data;
                aktZeile += data.GetLength(0);
                sheet.Cells[aktZeile, 1].Select();
            }
            catch (Exception err)
            {
                throw new Exception("Fehler beim Schreiben in das Arbeitsblatt.", err);
            }
        }

        /*
         * PRIVATE METHODEN
         */
        private void writeError(Exception err)
        {
            try
            {
                // Nur max. alle 3 Sekunden einen Fehler anzeigen.
                if ((DateTime.Now - lastError).TotalSeconds < 3) return;

                lastError = DateTime.Now;
                System.Windows.Forms.NotifyIcon notify = new System.Windows.Forms.NotifyIcon();
                notify.Icon = System.Drawing.SystemIcons.Exclamation;
                notify.BalloonTipText = err.Message +
                    (err.InnerException != null ? Environment.NewLine + err.InnerException.Message : "");
                notify.BalloonTipTitle = "ExcelSerial";
                notify.BalloonTipIcon = System.Windows.Forms.ToolTipIcon.Error;
                notify.Visible = true;
                notify.ShowBalloonTip(3000);
                notify.Dispose();
            }
            catch { }
        }

        private void bufferedReadAsync(int bufferSize, TimeSpan bufferDuration)
        {
            StreamReader sr = null;
            recording = true;

            while (recording)                             // Wird von StopRecord() auf false gesetzt
            {
                try
                {
                    if (!sp.IsOpen)
                    {
                        sr = new StreamReader(
                            sp.Open((int)bufferDuration.TotalMilliseconds), Encoding.ASCII, false, 1024);
                    }

                    string[] buffer = new string[bufferSize];
                    int bufferCount = 0;
                    DateTime start = DateTime.Now;
                    try
                    {
                        while (recording && bufferCount < bufferSize && (DateTime.Now - start) < bufferDuration)
                        {
                            string line = sr.ReadLine();
                            /* Hat es kein Timeout gegeben? */
                            if (line != null)
                            {
                                // Anzahl der Spalten ermitteln. Der Puffer muss zum Schreiben ins Excel
                                // Arbeitsblatt rechteckig sein. Daher wird die maximale Spaltenanzahl für
                                // alle Datensätze des Puffers genommen (kann bei CSV sich ja ändern).
                                buffer[bufferCount++] = line;
                            }
                        }
                    }
                    /* Fehler beim Lesen: Port schließen und neu öffnen. */
                    catch (Exception err)
                    {
                        sr.Close();
                        sp.Close();
                        System.Threading.Thread.Sleep(1000);
                        writeError(err);
                    }
                    // Startet den Verarbeitungstask.
                    parseTask = new Task(() => parseBufferAsync(buffer, bufferCount));
                    parseTask.Start();
                }
                /* Fehler beim Öffnen des Ports */
                catch (Exception err)
                {
                    writeError(err);
                }
            }
            sr?.Close();
            sp.Close();
            /* Nach dem Beenden warten, bis die Daten ins Arbeitsblatt geschrieben wurden. */
            parseTask?.Wait();
        }

        private void parseBufferAsync(string[] buffer, int bufferCount)
        {
            object[,] parsedBuffer;
            if (csvDelimiter != '\0')
            {
                parsedBuffer = BufferParser.Parse(buffer,
                    bufferCount, csvDelimiter, base64Decode);
            }
            else if (fieldsLength.Length > 0)
            {
                parsedBuffer = BufferParser.Parse(buffer,
                    bufferCount, fieldsLength, base64Decode);
            }
            else
            {
                parsedBuffer = BufferParser.Parse(buffer,
                    bufferCount, base64Decode);
            }
            writeRecords(parsedBuffer);
        }
    }
}
