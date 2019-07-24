

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

namespace wd2md
{
    public partial class Form2 : Form
    {
        List<MdStyle> lstStyle = new List<MdStyle>();

        //public Dictionary<string, MdStyle> lstStyle = new Dictionary<string, string>();
        public Form2()
        {
            InitializeComponent();
        }

        public void Form2_Load(object sender, EventArgs e)
        {
            lstStyle.Add(new MdStyle()
            {
                mdstyle = "List",
                wdstyles = String.Join(",", Form1.ListStyle)
            });
            lstStyle.Add(new MdStyle()
            {
                mdstyle = "Code",
                wdstyles = String.Join(",", Form1.CodeStyle)
            });
            lstStyle.Add(new MdStyle()
            {
                mdstyle = "Title",
                wdstyles = String.Join(",", Form1.TitleStyle)
            });
            lstStyle.Add(new MdStyle()
            {
                mdstyle = "Heading1",
                wdstyles = String.Join(",", Form1.Heading1)
            });
            lstStyle.Add(new MdStyle()
            {
                mdstyle = "Heading2",
                wdstyles = String.Join(",", Form1.Heading2)
            });
            lstStyle.Add(new MdStyle()
            {
                mdstyle = "Heading3",
                wdstyles = String.Join(",", Form1.Heading3)
            });
            //use binding source to hold dummy data
            BindingSource binding = new BindingSource();
            binding.DataSource = lstStyle;

            //bind datagridview to binding source
            dataGridView1.DataSource = binding;

        }

        private string [] ToStringArray(string s, char sep)
        {
            var vals = new string[] { };
            vals = s.Split(sep);
            vals = (from v in vals
                    select v.Trim()).ToArray();
            return vals;
        }
        private void button1_Click(object sender, EventArgs e)
        {
            // ok - save current styles in use and ini file. 
            // No error checking.
            Form1 form1 = (Form1) Application.OpenForms["Form1"];
            
            foreach (var style in lstStyle)
            {
                if (style.mdstyle == "List")
                {
                    Form1.ListStyle = ToStringArray(style.wdstyles, ',');
                    form1.ini.SetSetting("STYLES", "List", style.wdstyles);
                } 
                else if (style.mdstyle == "Code")
                {
                    Form1.CodeStyle = ToStringArray(style.wdstyles, ',');
                    form1.ini.SetSetting("STYLES", "Code", style.wdstyles);
                }
                else if (style.mdstyle == "Title")
                {
                    Form1.TitleStyle = ToStringArray(style.wdstyles, ',');
                    form1.ini.SetSetting("STYLES", "Title", style.wdstyles);
                }
                else if (style.mdstyle == "Heading1")
                {
                    Form1.Heading1 = ToStringArray(style.wdstyles, ',');
                    form1.ini.SetSetting("STYLES", "Heading1", style.wdstyles);
                }
                else if (style.mdstyle == "Heading2")
                {
                    Form1.Heading2 = ToStringArray(style.wdstyles, ',');
                    form1.ini.SetSetting("STYLES", "Heading2", style.wdstyles);
                } 
                else if (style.mdstyle == "Heading3")
                {
                    Form1.Heading3 = ToStringArray(style.wdstyles, ',');
                    form1.ini.SetSetting("STYLES", "Heading3", style.wdstyles);
                }
            }
            form1.ini.SaveSettings(); 
        }

    }
}
