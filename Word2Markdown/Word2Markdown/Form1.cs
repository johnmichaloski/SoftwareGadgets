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
using System.Globalization;

namespace Word2Markdown
{
    /// <summary>
    /// THe problem is that the images require the clipboard which is part of the GUI thread
    /// in order to make a gif image for markdown. So WordAutomation class MUST run in the GUI thread 
    /// NOT as a background worker thread.  Its unfortunate that the clipboard must be used to
    /// copy the inline picture and convert to gif, but life is too short to figure out 
    /// otherwise. So  WordAutomation class MUST run in the GUI thread for now.
    /// </summary>
    public partial class Form1 : Form
    {
        bool bRunning;
        WordAutomation wd=null;
        FileDialog fileDialog;
        DateTime start_time;
        TimeSpan elapsed_time;
        DateTime start_apptime;
        TimeSpan actual_time;
        float max_seconds;
        System.Timers.Timer timer;

        public Form1()
        {
            InitializeComponent();
            backgroundWorker1.DoWork += new DoWorkEventHandler(backgroundWorker1_DoWork);
            backgroundWorker1.ProgressChanged += new ProgressChangedEventHandler(backgroundWorker1_ProgressChanged);
            backgroundWorker1.RunWorkerCompleted += new RunWorkerCompletedEventHandler(backgroundWorker1_RunWorkerCompleted);
            //timer = new System.Timers.Timer();
            //timer.Interval = 1000;
            //timer.Elapsed += new System.Timers.ElapsedEventHandler(timer_Tick);
            //timer.Start();

             bRunning = false;
            start_time = DateTime.Now;
            start_apptime = DateTime.Now;
            max_seconds = 0;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (!bRunning)
            {
                fileDialog = new System.Windows.Forms.OpenFileDialog();
                fileDialog.Filter = "Word files (*.docx)|";
                var result = fileDialog.ShowDialog();
                if (result == System.Windows.Forms.DialogResult.Cancel)
                    return;
                wd = new WordAutomation();
                wd.filename = fileDialog.FileName;
                wd.filetitle = Path.GetFileNameWithoutExtension(wd.filename);
                wd.bInsertTOC = checkBox1.Checked;
                this.textBox1.Text = fileDialog.FileName;
                bRunning = true;
#if MULTITHREAD
                button1.Text = "Cancel";
                button2.Hide;
                backgroundWorker1 = new BackgroundWorker();
                wd.Init();
                backgroundWorker1.RunWorkerAsync();
                while (!wd.IsRunning())
                    Thread.Sleep(100);
                progressBar1.Maximum = wd.numParagraphs();
                progressBar1.Value = 0;  
#else
                button1.Text = "Cancel";
                progressBar1.Value = 0;
                wd.Init(progress_Callback);
                progressBar1.Maximum = wd.numParagraphs();

                timer1.Interval = 10;
                timer1.Tick += new EventHandler(timer1_Tick);
                timer1.Start();

                //wd.Run(progress_Callback);
#endif
            }

            else if (button1.Text == "Cancel")
            {
                System.Windows.Forms.MessageBox.Show("Abort!");
                if (wd.IsRunning())
                {
                    wd.setRunning(false);
                    timer1.Start();
                    // fixme should be abort to clean up better.
                    while (!wd.IsDone())
                        Thread.Sleep(100);
                }
                System.Windows.Forms.Application.Exit();
            }

            else // if (button1.Text == "Done")
            {
                System.Windows.Forms.Application.Exit();
            }

        }
        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            button1.Text = "Done";
            System.Windows.Forms.Clipboard.SetText("  ");

        }
        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            wd.Run(progress_Callback);
        }
        void empty_Callback (int i)
        {
            backgroundWorker1.ReportProgress(i);        
        }
        void progress_Callback(int i)
        {
#if MULTITHREADED
            backgroundWorker1.ReportProgress(i);
#else
            if (i < 0)
            {
                string status = wd.Status();
                this.label2.Text = status;
                return;
            }
            int tick = i;
            if (tick >= progressBar1.Maximum)
            {
                tick = progressBar1.Maximum;
                button1.Text = "Done";
                progressBar1.Value = tick;
                label2.Text = "Estimated time remaining is 0";
                return;
            }
            float n = tick;

            progressBar1.Value = tick;
            elapsed_time = DateTime.Now - start_time;
            //max_seconds = Math.Max(elapsed_time.Seconds, max_seconds);

            float avg_seconds = (max_seconds * (n - 1) + (float)elapsed_time.Seconds) / n;
            max_seconds = avg_seconds;
            start_time = DateTime.Now;
            int cycles_togo = progressBar1.Maximum - tick;
            string units = " Seconds";
            int togo = (int)(max_seconds * cycles_togo);
            if (togo > 60)
            {
                units = " Minute";
                togo = togo / 60;
                if (togo > 1)
                    units += "s";
            }
            if (togo > 60)
            {
                units = " Hour";
                togo = togo / 3600;
                if (togo > 1)
                    units += "s";
            }
            label2.Text = "Estimated time remaining is " + togo.ToString() + units;

#endif
        }        
        private void backgroundWorker1_ProgressChanged(object sender,
           ProgressChangedEventArgs e)
        {

            if (e.ProgressPercentage < 0)
            {
                string status = wd.Status();
                this.label2.Text = status;
                return;
            }
            int tick = e.ProgressPercentage;
            if (tick >= progressBar1.Maximum)
            {
                tick = progressBar1.Maximum;
                button1.Text = "Done";
                progressBar1.Value = tick;
                label2.Text = "Estimated time remaining is 0";
                return;
            }
            float n = tick;
 
            progressBar1.Value = tick;
            elapsed_time = DateTime.Now - start_time;
            //max_seconds = Math.Max(elapsed_time.Seconds, max_seconds);

            float avg_seconds = (max_seconds * (n - 1) + (float)elapsed_time.Seconds) / n;
            max_seconds= avg_seconds;
            start_time = DateTime.Now;
            int cycles_togo = progressBar1.Maximum-tick;
            string units=" Seconds";
            int togo = (int) (max_seconds * cycles_togo);
            if (togo > 60) 
            {
                units=" Minute";
                togo = togo/60;
                if (togo > 1)
                    units += "s";
            }
            if (togo > 60)
            {
                units = " Hour";
                togo = togo / 3600;
                if (togo > 1)
                    units += "s";
            } 
            label2.Text = "Estimated time remaining is " + togo.ToString() + units;

        }

        private void timer1_Tick(object sender, EventArgs e)
        {
//            actual_time = DateTime.Now - start_apptime;
 //           this.toolStripStatusLabel1.Text = "Actual elapsed time is " + actual_time.ToString(@"hh\:mm\:ss", new CultureInfo("en-US"));
            if (!wd.Run(progress_Callback))  // || bCancel
            {
                timer1.Stop();
                wd.Cleanup();
            }
            
                
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {

        }
    }

}
