
// DISCLAIMER:
// This software was developed by U.S. Government employees as part of
// their official duties and is not subject to copyright. No warranty implied
// or intended.
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Threading;
using System.Reflection;

namespace wd2md
{
    public partial class Form1 : Form
    {
        public string wdfilename = ""; /// word filename
        public string mdfilename = ""; /// markdown filename
        public string dialogtitle; /// saved dialog title for restoring
        public string appnamever;  // app name and version string
        string[] ListStyle = new string[] { };
        string[] CodeStyle = new string[] { };
        string[] TitleStyle = new string[] { };
        string[] Heading1 = new string[] { };
        string[] Heading2 = new string[] { };
        string[] Heading3 = new string[] { };
        bool bRunning = false;

        public WordAutomation wd = new WordAutomation();
        public IniReader ini;
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

            Assembly assembly = Assembly.GetExecutingAssembly();
            AssemblyName assemblyName = assembly.GetName();

            appnamever = string.Format("{0} : {1}", assemblyName.Name, assemblyName.Version.ToString());
            this.Text = appnamever;
            string exefolder = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\";
            ini = new IniReader(exefolder + "Config.ini");
            string list = ini.GetSetting("STYLES", "List");
            string[] ListStyle = ini.GetValues("STYLES", "List", ',');
            string[] CodeStyle = ini.GetValues("STYLES", "Code", ',');
            string[] TitleStyle = ini.GetValues("STYLES", "Title", ',');
            string[] Heading1 = ini.GetValues("STYLES", "Heading1", ',');
            string[] Heading2 = ini.GetValues("STYLES", "Heading2", ',');
            string[] Heading3 = ini.GetValues("STYLES", "Heading3", ',');
        }

        /// <summary>
        /// Event handler for selecting word document.
        /// Full path saved.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void label3_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();

            openFileDialog1.InitialDirectory = // Directory.GetCurrentDirectory();
            Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments); ;
            openFileDialog1.Filter = "Word files (*.docx)|*.docx";
            openFileDialog1.FilterIndex = 1;
            openFileDialog1.RestoreDirectory = true;

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    if ((openFileDialog1.OpenFile()) != null)
                    {
                        wdfilename = openFileDialog1.FileName;
                        textBox1.Text = wdfilename;
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: Could not read file from disk. Original error: " + ex.Message);
                }
            }
        }

        /// <summary>
        /// Go button event. Start wd2md process thread.
        /// It it optional to specify a md file name.
        /// Initializes the wd2md process with filenames.
        /// If fails, doesn't start thread.
        /// Inserts a TOC depending on the checkbox.
        /// Prepares progress bar and timer.
        /// Starts a worker thread that runs.
        /// Set as a STA thread or fails when using clipboard (GUI issue?)
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click(object sender, EventArgs e)
        {
            if (bRunning)
            {
                System.Windows.Forms.MessageBox.Show("Already Running!");
                return;
            }
            bRunning = true;
            /// It it optional to specify a md file name
            if (mdfilename.Length == 0)
            {
                mdfilename = Path.GetDirectoryName(wdfilename) + "\\" + Path.GetFileNameWithoutExtension(wdfilename) + ".md";
                textBox2.Text = mdfilename;
            }

            // Initializes the wd2md process with filenames
            wd = new WordAutomation();
            if (wd.Init(wdfilename, mdfilename) < 0)
            {
                bRunning = false;
                return; // init failed
            }
            /// Read ini file for style formats
            /// to match against. If none skip
            try
            {
                wd.bInsertTOC = checkBox1.Checked;

                // Assign styles to match against... if assigned
                if (ListStyle.Length > 0)
                    wd.ListStyle = ListStyle;
                if (CodeStyle.Length > 0)
                    wd.CodeStyle = CodeStyle;
                if (TitleStyle.Length > 0)
                    wd.TitleStyle = TitleStyle;
                if (Heading1.Length > 0)
                    wd.Heading1 = Heading1;
                if (Heading2.Length > 0)
                    wd.Heading2 = Heading2;
                if (Heading3.Length > 0)
                    wd.Heading3 = Heading3;
            }
            catch (Exception ex)
            {
                bRunning = false;
                System.Windows.Forms.MessageBox.Show("Ini file error!");
               return;
            }
            // Prepare progress strip
            toolStripProgressBar1.Minimum = 0;
            toolStripProgressBar1.Maximum = 100;
            toolStripProgressBar1.Step = 1;
            CheckForIllegalCrossThreadCalls = false;
            timer1.Enabled = true;

            // Save dialog title for restoring after done or error
            dialogtitle = this.Text;

            // Spawn thread to convert word into markdown
            Thread workerThread = new Thread(wd.Run);
            // https://social.msdn.microsoft.com/Forums/en-US/95a587f2-5645-44e4-9f8e-bf9f9f1ba48e/get-clipboard-data-from-word-addin-background-thread?forum=vsto
            workerThread.SetApartmentState(System.Threading.ApartmentState.STA);
            // Or it fails on line System.Windows.Forms.Clipboard.SetText("  ");
            workerThread.Start();
            bRunning = true;

        }
        /// <summary>
        /// timer1_Tick is a procedure that runs when wd2md process is running.
        /// It updates the progress bar in the status area.
        /// The progress bar takes number of paragraphs as maximum, assuming it is > 0.
        /// When done the timer is turned off, the progress bar is hidden.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e">timer expired</param>
        private void timer1_Tick(object sender, EventArgs e)
        {
            this.Text = dialogtitle + " - " + wd.doneStatus;
            statusStrip1.Visible = true;
            toolStripProgressBar1.PerformStep();
            if (wd.nParagraphs > 0)
            {
                toolStripProgressBar1.Maximum = wd.nParagraphs;

            }
            /// When done, this timer is turned off, and the progress bar is hidden.
            /// Lots of variables are reset.
            if (wd.isDone())
            {
                timer1.Enabled = false;
                toolStripProgressBar1.Value = 0;
                statusStrip1.Visible = false;
                textBox2.Text = "";
                textBox1.Text = "";
                mdfilename = "";
                wdfilename = "";
                bRunning = false;

            }
            // Update progress bar with current wd2md paragraph number
            // crude way to keep track of progress, but it is long.
            toolStripProgressBar1.Value = wd.nCurParagraph;
        }

        /// <summary>
        /// Markdown file to save to. Optional. Uses modified word file name
        /// if no markdown filename is selected. Allows overwrite.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void label4_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            if (wdfilename.Length != 0)
            {
                saveFileDialog1.InitialDirectory = Path.GetDirectoryName(wdfilename) + "\\";
                //saveFileDialog1.FileName = Path.GetFileNameWithoutExtension(wdfilename) + ".md";
            }
            saveFileDialog1.Filter = "md files (*.md)|*.md";
            saveFileDialog1.FilterIndex = 1;
            saveFileDialog1.RestoreDirectory = true;

            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                if ((saveFileDialog1.OpenFile()) != null)
                {
                    mdfilename = saveFileDialog1.FileName;
                    textBox2.Text = mdfilename;
                }
            }
        }
    }
}
