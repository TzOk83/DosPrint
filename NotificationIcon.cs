/*
 * Created by SharpDevelop.
 * User: Tomek
 * Date: 2012-09-25
 * Time: 18:14
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Printing;
using System.IO;
using System.Text;
using System.Threading;
using System.Timers;
using System.Windows.Forms;
using System.Configuration;

namespace DOSprint
{
    public sealed class NotificationIcon
    {
        private NotifyIcon notifyIcon;
        private ContextMenu notificationMenu;

        private StringReader textToPrint;

        private Font printFont, printFontCond, printFontDH;
        private List<formatCodes> lineFormat;

        private int lineNumber;

        private FileSystemWatcher watch;
        private PrintDocument printDocument;
        private System.Timers.Timer timer;
        private PrinterSettings prtSettings;
        private PageSetupDialog pageSetupDialog;
        private PrintDialog prtDialog;
        private PrintController printController;

        #region Initialize icon and menu
        public NotificationIcon()
        {
            notifyIcon = new NotifyIcon();
            notificationMenu = new ContextMenu(InitializeMenu());

            notifyIcon.DoubleClick += IconDoubleClick;
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(NotificationIcon));
            notifyIcon.Icon = (Icon)resources.GetObject("$this.Icon");
            notifyIcon.ContextMenu = notificationMenu;

            watch = new FileSystemWatcher();
            watch.Path = @"C:\";
            watch.NotifyFilter = NotifyFilters.LastAccess | NotifyFilters.LastWrite | NotifyFilters.FileName | NotifyFilters.DirectoryName;
            watch.Filter = "spool.dat";
            watch.Changed += new FileSystemEventHandler(OnChanged);
            watch.Created += new FileSystemEventHandler(OnChanged);
            watch.EnableRaisingEvents = true;

            printDocument = new PrintDocument();
            printDocument.DocumentName = "DOSprint";
            printDocument.OriginAtMargins = true;
            printController = new StandardPrintController();
            printDocument.PrintController = printController;


            timer = new System.Timers.Timer();
            timer.Interval = 1500;
            timer.Elapsed += new ElapsedEventHandler(TimerElapsed);

            prtSettings = new PrinterSettings();
            pageSetupDialog = new PageSetupDialog();
            prtDialog = new PrintDialog();
            pageSetupDialog.EnableMetric = true;

            printDocument.PrinterSettings.PrinterName = Settings.Default.PrinterName;

            pageSetupDialog.PageSettings = printDocument.DefaultPageSettings;
            pageSetupDialog.PageSettings.Margins.Left = Settings.Default.Left;
            pageSetupDialog.PageSettings.Margins.Right = Settings.Default.Right;
            pageSetupDialog.PageSettings.Margins.Top = Settings.Default.Top;
            pageSetupDialog.PageSettings.Margins.Bottom = Settings.Default.Bottom;
            pageSetupDialog.PageSettings.Landscape = Settings.Default.Landscape;
            printDocument.PrintPage += new PrintPageEventHandler(PrintDocumentPrintPage);
            prtDialog.PrinterSettings = printDocument.PrinterSettings;
            prtDialog.Document = printDocument;

            printDocument.DefaultPageSettings = pageSetupDialog.PageSettings;

            notificationMenu.MenuItems[0].Text = printDocument.PrinterSettings.PrinterName;
            notificationMenu.MenuItems[0].Enabled = false;
        }

        private MenuItem[] InitializeMenu()
        {
            MenuItem[] menu = new MenuItem[] {
				new MenuItem(),
				new MenuItem("-"),
				new MenuItem("Ustawiania drukarki...", menuPrinterClick),
				new MenuItem("Ustawiania strony...", menuPageClick),
				new MenuItem("-"),
				new MenuItem("Wyjście", menuExitClick)
			};
            return menu;
        }
        #endregion

        #region Main - Program entry point
        /// <summary>Program entry point.</summary>
        /// <param name="args">Command Line Arguments</param>
        [STAThread]
        public static void Main(string[] args)
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            bool isFirstInstance;
            // Please use a unique name for the mutex to prevent conflicts with other programs
            using (Mutex mtx = new Mutex(true, "DOSprint", out isFirstInstance))
            {
                if (isFirstInstance)
                {
                    NotificationIcon notificationIcon = new NotificationIcon();
                    notificationIcon.notifyIcon.Visible = true;
                    Application.Run();
                    notificationIcon.notifyIcon.Dispose();

                }
                else
                {
                    // The application is already running
                    // TODO: Display message box or change focus to existing application instance
                }
            } // releases the Mutex
        }
        #endregion

        #region Event Handlers
        private void menuPrinterClick(object sender, EventArgs e)
        {
            if (prtDialog.ShowDialog() == DialogResult.OK)
            {
                printDocument.PrinterSettings = prtDialog.PrinterSettings;
                notificationMenu.MenuItems[0].Text = printDocument.PrinterSettings.PrinterName;
                Settings.Default.PrinterName = printDocument.PrinterSettings.PrinterName;
                Settings.Default.Save();
            }
        }

        private void menuPageClick(object sender, EventArgs e)
        {
            if (printDocument.PrinterSettings.IsValid)
            {
                if (pageSetupDialog.ShowDialog() == DialogResult.OK)
                {
                    printDocument.DefaultPageSettings = pageSetupDialog.PageSettings;
                    Settings.Default.Top = printDocument.DefaultPageSettings.Margins.Top;
                    Settings.Default.Bottom = printDocument.DefaultPageSettings.Margins.Bottom;
                    Settings.Default.Left = printDocument.DefaultPageSettings.Margins.Left;
                    Settings.Default.Right = printDocument.DefaultPageSettings.Margins.Right;
                    Settings.Default.Landscape = printDocument.DefaultPageSettings.Landscape;
                    Settings.Default.Save();
                }
            }
            else MessageBox.Show("Nieprawidłowa drukarka.", "DOSprint", MessageBoxButtons.OK, MessageBoxIcon.Error);

        }

        private void menuExitClick(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void IconDoubleClick(object sender, EventArgs e)
        {
            MessageBox.Show("The icon was double clicked");
        }
        #endregion

        void TimerElapsed(object sender, EventArgs e)
        {
            timer.Stop();
            string outString = "";
            lineNumber = 0;
            try
            {
                byte[] inBytes = File.ReadAllBytes(@"c:\spool.dat");
                outString = CreateString(inBytes, out lineFormat);
            }
            catch
            {
                MessageBox.Show("Błąd dostępu do pliku c:\\spool.dat.", "DOSprint", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            textToPrint = new StringReader(outString);
            printFont = new System.Drawing.Font("Lucida Console", 10f);
            printFontCond = new System.Drawing.Font("Lucida Console", 6.5f);
            printFontDH = new System.Drawing.Font("Lucida Console", 16f);
            if (printDocument.PrinterSettings.IsValid) printDocument.Print();
            else MessageBox.Show("Nieprawidłowa drukarka.", "DOSprint", MessageBoxButtons.OK, MessageBoxIcon.Error);
            textToPrint.Close();
        }

        void OnChanged(object source, FileSystemEventArgs e)
        {
            timer.Stop();
            timer.Start();
        }

        string CreateString(byte[] MazArray, out List<formatCodes> FormatLines)
        {
            string result = string.Empty;
            FormatLines = new List<formatCodes>();
            Dictionary<byte, char> dic = new Dictionary<byte, char>();
            dic.Add(0x86, 'ą');
            dic.Add(0x8D, 'ć');
            dic.Add(0x91, 'ę');
            dic.Add(0x92, 'ł');
            dic.Add(0xA4, 'ń');
            dic.Add(0xA2, 'ó');
            dic.Add(0x9E, 'ś');
            dic.Add(0xA6, 'ź');
            dic.Add(0xA7, 'ż');
            dic.Add(0x8F, 'Ą');
            dic.Add(0x95, 'Ć');
            dic.Add(0x90, 'Ę');
            dic.Add(0x9C, 'Ł');
            dic.Add(0xA5, 'Ń');
            dic.Add(0xA3, 'Ó');
            dic.Add(0x98, 'Ś');
            dic.Add(0xA0, 'Ź');
            dic.Add(0xA1, 'Ż');
            dic.Add(0x09, '\t');

            int line = 0;
            FormatLines.Add(new formatCodes());
            byte MazChar;
            bool bold = false, cond = false, dheight = false;
            for (int i = 0; i < MazArray.Length; i++)
            {
                MazChar = MazArray[i];
                if (MazChar == 0x1B) // Esc
                {
                    i++; // next char
                }
                else if (MazChar == 0x0E)
                {
                    dheight = true;
                }
                else if (MazChar == 0x0F)
                {
                    cond = true;
                }
                else if (MazChar == 0x12)
                {
                    cond = false;
                }
                else if (MazChar == 0x0C)
                {
                    FormatLines[line].formFeed = true;
                }
                else if (MazChar == 0x0A)
                {
                    result += Environment.NewLine;
                    FormatLines[line].Bold = bold;
                    FormatLines[line].Cond = cond;
                    FormatLines[line].DHeight = dheight;
                    dheight = false;
                    FormatLines.Add(new formatCodes());
                    line++;
                }
                else if (dic.ContainsKey(MazChar))
                {
                    result += dic[MazChar];
                }
                else if (MazChar > 31 & MazChar < 127)
                {
                    result += System.Text.Encoding.Default.GetString(new byte[] { MazChar });
                }
            }
            int pos = 0;
            for (int i = 0; i < FormatLines.Count; i++)
            {
                if (FormatLines[i].formFeed)
                    pos = i;
            }
            FormatLines[pos].EOF = true;
            return result;
        }

        void PrintDocumentPrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            int count = 0;
            float leftMargin = e.MarginBounds.Left;
            float topMargin = e.MarginBounds.Top;
            float yPos = topMargin;
            string line = null;
            float linesPerPage = 76;

            while ((count < linesPerPage) & !lineFormat[lineNumber].formFeed)
            {
                try
                {
                    line = textToPrint.ReadLine();
                }
                catch { }
                if (line == null)
                {
                    break;
                }
                yPos += printFont.GetHeight(e.Graphics);
                e.Graphics.DrawString(line, lineFormat[lineNumber].Cond ? printFontCond : lineFormat[lineNumber].DHeight ? printFontDH : printFont, Brushes.Black, leftMargin, yPos, new StringFormat());
                count++;
                if (lineNumber + 1 < lineFormat.Count) lineNumber++;
            }
            if (line != null & !lineFormat[lineNumber].EOF)
            {
                lineFormat[lineNumber].formFeed = false;
                e.HasMorePages = true;
            }
        }
    }
    public class formatCodes
    {
        public formatCodes() { }
        public bool Cond, Bold, DHeight, formFeed, EOF;
    }
}
