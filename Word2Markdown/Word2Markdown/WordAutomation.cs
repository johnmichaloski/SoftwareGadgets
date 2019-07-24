using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Word;
using Microsoft.Office.Interop;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools;
using System.IO;
using System.Diagnostics;
using System.Threading;
using System.Drawing;  

namespace Word2Markdown
{

    // Extension methods must be defined in a static class.
    public static class StringExtension
    {
        // This is the extension method.
        // The first parameter takes the "this" modifier
        // and specifies the type for which the method is defined.
        public static string Write(this String str)
        {
            str += str;
            return str;
        }
    }

    public class StringWriter
    {
        public string _str;
        public  void Write( String str)
        {
            _str += str;
        }
        public void Close()
        {

        }
        public string Get()
        {
            return _str;
        }
        public void Set(string str)
        {
             _str=str;
        }
    }

    public class MdStyle
    {
        public string mdstyle { get; set; }
        public string wdstyles { get; set; }
    }
    /// <summary>
    /// TOC contains the information about a table of content entry. 
    /// This includes the string, the numbers for h1, h2, h3.
    /// For example, 1.2 give a number array: 1,2,0.
    /// 1.2.1 gives a number array of 1,2,1.
    /// When no subsection, the number is zero.
    /// A static array is kept to keep track of the
    /// detecting of headings. H1,H2,H3 are also
    /// static methods to create a TOC entry using
    /// the static index array, and adding the section
    /// name string.
    /// </summary>
    public class TOC
    {

        /// <summary>
        /// This is the static heading counter. 
        /// Used by all instances of TOC.
        /// </summary>
        public static int[] h = { 0, 0, 0 };
        /// <summary>
        /// Instance data of a TOC.
        /// </summary>
        public int[] hdr;
        /// <summary>
        ///  String of the header.
        /// </summary>
        public string heading;
        /// <summary>
        /// Constructor. Builds empty TOC.
        /// </summary>
        public TOC()
        {
            heading = "";
            hdr = new int[3];
        }
        /// <summary>
        /// Builds H1 TOC, using static section "odometer".
        /// User provides TOC H1 label as string.
        /// </summary>
        /// <param name="hdr">heading label</param>
        /// <returns> new TOC </returns>
        public static TOC H1(string hdr)
        {
            TOC toc = new TOC();
            toc.heading = hdr;
            ++h[0]; h[1] = 0; h[2] = 0;
            toc.hdr[0] = h[0];
            toc.hdr[1] = 0;
            toc.hdr[2] = 0;
            return toc;
        }
        /// <summary>
        /// Builds H2 TOC, using static section "odometer".
        /// User provides TOC 2 label as string.
        /// </summary>
        /// <param name="hdr">heading2 label </param>
        /// <returns>new TOC</returns>
        public static TOC H2(string hdr)
        {
            TOC toc = new TOC();
            toc.heading = hdr;
            ++h[1]; h[2] = 0;
            toc.hdr[0] = h[0];
            toc.hdr[1] = h[1];
            toc.hdr[2] = 0;
            return toc;
        }
        /// <summary>
        /// Builds H3 TOC, using static section "odometer".
        /// User provides TOC 3 label as string.
        /// </summary>
        /// <param name="hdr">heading2 label </param>
        /// <returns>new TOC</returns>
        public static TOC H3(string hdr)
        {
            ++h[2];
            TOC toc = new TOC();
            toc.heading = hdr;
            toc.hdr[0] = h[0];
            toc.hdr[1] = h[1];
            toc.hdr[2] = h[2];
            return toc;
        }
    }
    /// <summary>
    ///     Handles the word automation and conversion into Markdown.
    /// </summary>
    class WordAutomation
    {

        /// <summary>
        /// Table of contens handling, If enabled by bInsertTOC true,
        /// inserts a placeholder (tocpattern) into the md text.
        /// After the document has been processed, it then inserts
        /// a table of contents where the placeholder was inserted.
        /// </summary>
        private static string tocpattern = "@^v@^V^V@^V@^V@^V@";
        /// <summary>
        /// List of table of contents entries: h1,h2,h3 and equivalents
        /// are saved into toc.
        /// </summary>
        public List<TOC> toc = new List<TOC>();
        /// <summary>
        /// Boolean true if found table of contents in document.
        /// </summary>
        public bool bFoundTOC = false;
        /// <summary>
        /// Boolean true if table of contents should be inserts.
        /// FIXME: if no toc in word, no TOC marker is placed in
        /// document.
        /// </summary>
        public bool bInsertTOC = true;

        /// <summary>
        /// in code progressing
        /// </summary>
        public bool bInCode;

        // These are the styles to match against 
        public string[] ListStyle = { "List Paragraph" };
        public string[] CodeStyle = { "BoxedCode" };
        public string[] TitleStyle = { "Title" };
        public string[] Heading1 = { "Heading 1", "Heading1", "H1" };
        public string[] Heading2 = { "Heading 2" };
        public string[] Heading3 = { "Heading 3" };

        /// <summary>
        /// Fake stream to write output.
        /// </summary>
        StringWriter sw;

        public string CleanNametag(string str)
        {
            return str.Replace(" ", "_");
        }
        /// <summary>
        ///     List of all tables and their ranges.
        /// </summary>
        public List<Range> TablesRanges = new List<Range>();

        /// <summary>
        ///     List of all images and their ranges.
        /// </summary>
        public List<Range> ImageRanges = new List<Range>();

        public Object oMissing = System.Reflection.Missing.Value;

        /// <summary>
        /// The Microsoft XML bibliography.
        /// </summary>
        public XmlBibliography bib = new XmlBibliography();


        /// <summary>
        ///     Word application COM variable.
        /// </summary>
        public Microsoft.Office.Interop.Word.Document oWordDoc;

        /// <summary>
        ///     Word document COM variable.
        /// </summary>
        public Microsoft.Office.Interop.Word.Application oWord;

        /// <summary>
        ///     Filename of the current word file undergoing conversione.
        /// </summary>
        public string filename = "X:\\src\\github\\johnmichaloski\\MTConnectToolbox\\Agents\\ZeissMTConnectAgent\\doc\\ZeissAgentReadme.docx";

        /// <summary>
        ///     Foldername of the current word file undergoing conversion.
        /// </summary>
        public string foldername;

        /// <summary>
        ///     File title of current word file undergoing conversion. Useful for differentiating images.
        /// </summary>
        public string filetitle;

        /// <summary>
        ///     File title of current word file undergoing conversion. Useful for differentiating images.
        /// </summary>
        public string filecopy;

        /// <summary>
        /// Complete references list
        /// </summary>
        public List<string> references = new List<string>();

        /// <summary>
        /// Reference index key into bibliography index. 
        /// Used for multiple citations of the same reference.
        /// </summary>
        public Dictionary<int, int> cited = new Dictionary<int, int>();

        /// <summary>
        /// synchronization variable for words/paragraph loop. 
        /// To abort set brunning to false before done.
        /// </summary>
        private bool bRunning;

        /// <summary>
        /// synchronization variable indicating word to markdown down, either completed or aborted.
        /// </summary>
        private bool bDone;

        /// <summary>
        /// Linu number of progress so far. If < 0 then status message.
        /// </summary>
        public int linenum;

        /// <summary>
        /// mutex varaible for running, done flags
        /// </summary>
        private readonly object syncLock = new object();

        /// <summary>
        /// status string for when linenum < 0
        /// </summary>
        public string status = "";

        /// <summary>
        /// Constructor
        /// </summary>
        public WordAutomation()
        {
            setRunning(false);
            setDone(false);
        }
        public string Status(string s = "")
        {
            string _status;
            lock (syncLock)
            {
                if (s.Length > 0)
                    status = s;
                _status = status;
            }
            return _status;
        }
        public bool IsRunning()
        {
            bool b;
            lock (syncLock)
            {
                b = bRunning;
            }
            return b;
        }
        public int numParagraphs()
        {
            return oWordDoc.Paragraphs.Count;
        }
        public bool IsDone()
        {
            bool b;
            lock (syncLock)
            {
                b = bDone;
            }
            return b;
        }
        public void setRunning(bool b)
        {
            lock (syncLock)
            {
                bRunning = b;
            }
        }
        public void setDone(bool b)
        {
            lock (syncLock)
            {
                bDone = b;
            }
        }


        /// <summary>
        ///     Copied filename of the current word file undergoing conversione.
        /// </summary>
        public void GetTableRanges()
        {
            for (int iCounter = 1; iCounter <= oWordDoc.Tables.Count; iCounter++)
            {
                Range TRange = oWordDoc.Tables[iCounter].Range;
                TablesRanges.Add(TRange);
            }
        }

        /////////////////////////////////////////////////////////////////////////
        /// <summary>
        /// Markdown line feed equivalent. 3 cr/lfs for now.
        /// </summary>
        /// <returns> string containing new lines</returns>
        public string md_newline()
        {
            return Environment.NewLine + Environment.NewLine + Environment.NewLine;
        }

        /////////////////////////////////////////////////////////////////////////
        /// <summary>
        /// Get path of executable app.
        /// </summary>
        /// <returns></returns>
        private string ExePath()
        {
            string path = System.IO.Path.GetDirectoryName(new Uri(System.Reflection.Assembly.GetAssembly(typeof(WordAutomation)).CodeBase).LocalPath);
            // THis returns a file:// URI :(  File.Copy DOES not like.
            //path = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase);
            return path;
        }

        /////////////////////////////////////////////////////////////////////////
        /// <summary>
        /// 
        /// </summary>
        private void WdFindAllSubscripts()
        {
            object missing = System.Reflection.Missing.Value;
            object istrue = true;
            oWordDoc.Select();
            oWordDoc.Range().WholeStory();
            string before, after;
            for (int j = 0; j < 2; j++)
            {
                Range r = oWordDoc.Range();
                r.Find.ClearFormatting();
                r.Find.Replacement.ClearFormatting();
                if (j == 0)
                {
                    r.Find.Font.Subscript = 1;
                    before = "<sub>"; after = "</sub>";
                }
                else
                {
                    r.Find.Font.Superscript = 1;
                    before = "<sup>"; after = "</sup>";
                }
                r.Find.Text = "";
                r.Find.Forward = true;
                r.Find.Wrap = WdFindWrap.wdFindStop; // wdFindStop
                r.Find.Format = true;
                r.Find.MatchCase = false;
                r.Find.MatchWholeWord = false;
                r.Find.MatchWildcards = false;
                r.Find.MatchSoundsLike = false;
                r.Find.MatchAllWordForms = false;
                r.Find.Execute(ref missing, ref missing, ref missing,
                       ref missing, ref missing, ref missing, ref missing,
                       ref missing, ref missing, ref missing, ref missing,
                       ref missing, ref missing, ref missing, ref missing);
                bool bFlag = true;
                while (r.Find.Found && bFlag)
                {
                    System.Diagnostics.Debug.Print("Range =[" + Convert.ToString(r.Start) + "," + Convert.ToString(r.End) + "]");

                    try
                    {
                        Range rend = oWordDoc.Range(r.End, r.End);
                        rend.InsertAfter(after);
                        Range rend1 = oWordDoc.Range(r.End, r.End + after.Length + 1);
                        rend1.Font.Superscript = 0;
                        rend1.Font.Subscript = 0;

                        Range rstart = oWordDoc.Range(r.Start, r.Start);
                        rstart.InsertBefore(before);

                        // reset find after <> added
                        r.Start = r.End + after.Length + 1;

                        r.Find.Execute(ref missing, ref missing, ref missing,
                                       ref missing, ref missing, ref missing, ref missing,
                                       ref missing, ref istrue, ref missing, ref missing,
                                       ref missing, ref missing, ref missing, ref missing);
                    }
                    catch { bFlag = false; }

                }
            }
        }

        /////////////////////////////////////////////////////////////////////////
        /// <summary>
        ///     Converts drawing canvases to inline drawings using cut and paste special.
        /// </summary>
        public void ConvertShapesToInline()
        {
            object missing = System.Reflection.Missing.Value;
            //Sadly this doesn't work with Word 2016, probable mismatch of word interop reference versions
            // Works with VS 2010 C# and Word 2010... word interop reference version 14.0.0.
            // http://forums.asp.net/t/1912406.aspx?How+to+use+two+versions+of+MS+Office+Interop+Assemblies+in+a+C+Application+
            // Ended up going thru document cutting drawing, and pasting special as enhanced metafile... :(
            Microsoft.Office.Interop.Word.Shapes allshapes = oWordDoc.Shapes;
            for (int i = allshapes.Count - 1; i >= 0; i--)
            {
                // Because MS COM numbering is 1..n, not 0..n-1 
                Shape s = allshapes[i + 1];
                s.Select();
                Range r = oWord.Selection.Range;
                //System.Diagnostics.Debug.Print("Range =[" + Convert.ToString(r.Start) + Convert.ToString(r.End) + "]");
                r.Cut();
                object objDataTypeMetafile = Microsoft.Office.Interop.Word.WdPasteDataType.wdPasteEnhancedMetafile;
                object objPlacement = Microsoft.Office.Interop.Word.WdOLEPlacement.wdInLine;
                r.PasteSpecial(ref missing, ref missing,
                    ref objPlacement, ref missing, ref objDataTypeMetafile,
                    ref missing, ref missing);
            }
            //number = oWordDoc.InlineShapes.Count;
            //System.Windows.Forms.MessageBox.Show(number.ToString()); // 0 to begin with
        }

        /////////////////////////////////////////////////////////////////////////
        /// <summary>
        ///     Extracts all the image ranges into the ImageRanges data structure .
        /// </summary>
        public void GetImageRanges()
        {
            foreach (InlineShape shape in oWordDoc.InlineShapes)
            {

                if (shape.Type == Microsoft.Office.Interop.Word.WdInlineShapeType.wdInlineShapePicture)
                {
                    ImageRanges.Add(shape.Range);
                }
                else if (shape.Type == WdInlineShapeType.wdInlineShapeLinkedPicture)
                {
                    ImageRanges.Add(shape.Range);
                }
            }
        }

        /////////////////////////////////////////////////////////////////////////
        /// <summary>
        ///     Saves all the image ranges into the folder images under the current filename folder.
        ///     Uses clipboard to copy and paste into image handler, which saves AS gif.
        /// </summary>
        public void SaveAllImages(string directory)
        {
            int nimages = 1;
            System.IO.Directory.CreateDirectory(foldername + "images");
            // Save all images as gif
            foreach (InlineShape shape in oWordDoc.InlineShapes)
            {

                if (shape.Type == WdInlineShapeType.wdInlineShapePicture)
                {

                    shape.Select();
                    oWord.Selection.CopyAsPicture();
                    SaveClipboardImage(foldername + "images\\" + filetitle + "_image" + nimages.ToString() + ".gif");
                    Thread.Sleep(1000);
                    nimages++;
                }
                else if (shape.Type == WdInlineShapeType.wdInlineShapeLinkedPicture)
                {
                    System.Diagnostics.Debug.WriteLine(shape.Range.Text);
                    shape.Select();
                    oWord.Selection.CopyAsPicture();
                    SaveClipboardImage(foldername + "images\\" + filetitle + "_image" + nimages.ToString() + ".gif");
                    nimages++;
                }
                else
                {
                    shape.Select();
                    oWord.Selection.CopyAsPicture();
                    SaveClipboardImage(foldername + "images\\" + filetitle + "_image" + nimages.ToString() + ".gif");
                    nimages++;
                }
            }
        }
        /////////////////////////////////////////////////////////////////////////
        /// <summary>
        ///     Determine if given paragraph p in found in the ImageRanges.
        /// </summary>
        public string InImage(Paragraph p)
        {
            Boolean bInImage = false;
            string totaltext = "";
            Range r = p.Range;
            int iCounter = 1;
            foreach (Range range in ImageRanges)
            {
                if (range.Start >= r.Start && range.End <= r.End)
                {
                    bInImage = true;
                    break;
                }
                iCounter++;

            }
            if (bInImage)
            {
                // Unclear if figure on new line is appropriate - but in general better than not.
                // Doesn't work totaltext = "\n\n<p align=\"center\"> ![Figure" + iCounter.ToString() + "](./images/image" + iCounter.ToString() + ".gif?raw=true)\n</p>\n";
                //totaltext = "\n\n![Figure" + iCounter.ToString() + "](./images/image" + iCounter.ToString() + ".gif?raw=true)\n";
                //totaltext = "\n\n![Figure" + iCounter.ToString() + "](./images/image" + iCounter.ToString() + ".gif?raw=true)\n\n";
                totaltext = md_newline() + Environment.NewLine + "![Figure" + iCounter.ToString() + "](./images/" + filetitle + "_image" + iCounter.ToString() + ".gif)" + md_newline();
                // FIXME: next line should be centered if figure.
            }
            return totaltext;
        }

        /////////////////////////////////////////////////////////////////////////
        /// <summary>
        ///     Determine if given paragraph p in found in the TablesRanges.
        /// </summary>
        public string InTable(Paragraph p, ref Range tablerange)
        {
            Boolean bInTable;
            string totaltext = "";
            bInTable = false;
            Range r = p.Range;
            int iCounter = 1;
            string currLine;
            foreach (Range range in TablesRanges)
            {
                if (r.Start >= range.Start && r.End <= range.End)
                {
                    bInTable = true;
                    tablerange = range;
                    break;
                }
                iCounter++;

            }
            if (bInTable)
            {
                totaltext += md_newline() + "<TABLE>";
                foreach (Row aRow in oWordDoc.Tables[iCounter].Rows)
                {
                    totaltext += md_newline() + "<TR>";

                    foreach (Cell aCell in aRow.Cells)
                    {
                        currLine = aCell.Range.Text;
                        char[] delimiterChars = { '\r' };
                        string[] words = currLine.Split(delimiterChars);
                        // Fixme  - single line cells only have \r\a
                        currLine = currLine.Replace("\r", "<BR>");

                        // Count the number of paragraphs?
                        totaltext += md_newline() + "<TD>" + currLine + "</TD>";
                    }
                    totaltext += Environment.NewLine + "</TR>";
                }
                totaltext += md_newline() + "</TABLE>";
            }
            return totaltext;
        }

        /////////////////////////////////////////////////////////////////////////
        /// <summary>
        /// Let's remove trailing spaces and comma.
        /// </summary>
        /// <param name="inputString"> string to trim and save </param>
        /// <returns> trimmed string of \r\b\t and space</returns>
        public string TrimEnd(string inputString, ref string trailing)
        {
            trailing = "";
            string delims = " \a\r\n\t";
            int i = inputString.Count() - 1;
            while (i >= 0 && delims.IndexOf(inputString[i]) >= 0)
            {
                trailing = inputString[i] + trailing;
                i--;
            }
            return inputString.Trim(" \r\n\t\a".ToArray());
        }

        /////////////////////////////////////////////////////////////////////////
        /// <summary>
        ///     Looks at each word in a paragraph and formats font if word is bold or underlined.
        /// </summary>
        public string Words(Paragraph p)
        {
            string text = "";
            bool B = false, I = false;
            for (int i = 1; i <= p.Range.Words.Count; i++)
            {
                Range w = p.Range.Words[i];
                bool bTrans = false;
                //if (w.Text == "SaveClipboardImage")
                //{
                //    System.Diagnostics.Debugger.Break();
                //}
                // vba true is -1, false = 0
                if (w.Bold == -1 && !B)
                {
                    B = true;
                    text += "**";
                }
                if (w.Italic == -1 && !I)
                {
                    I = true;
                    text += "_";
                }
                if (w.Italic == 0 && I)
                {
                    I = false;
                    text = text.TrimEnd(text[text.Length - 1]);
                    text += "_";
                    bTrans = true;
                }
                if (w.Bold == 0 && B)
                {
                    B = false;
                    text = text.TrimEnd(text[text.Length - 1]);
                    text += "**";
                    bTrans = true;
                }
                if (bTrans)
                    text += " ";
                text += w.Text;
                //System.Diagnostics.Debug.WriteLine(w.Text);
            }
            string trailing = "";
            text = TrimEnd(text, ref trailing);
            // Now backtrack through all the spaces
            if (I)
            {
                text += "_";
            }
            if (B)
            {
                text += "**";
            }

            text += trailing;
            return text;
        }

        /////////////////////////////////////////////////////////////////////////
        /// <summary>
        /// 
        /// </summary>
        public void DeleteTableOfContents()
        {
            foreach (TableOfContents toc in oWordDoc.TablesOfContents)
            {
                this.bFoundTOC = true;
                //if(++i==1)
                //    toc.Range.InsertBefore("* auto-gen TOC:\r\n {:toc}");
                string tochdr = this.md_newline() + "# Table of Contents" + this.md_newline();
                if (bInsertTOC)
                {
                    tochdr += tocpattern;
                    toc.Range.InsertBefore(tochdr);
                }
                toc.Delete();
            }
        }

        /////////////////////////////////////////////////////////////////////////
        /// <summary>
        /// Generate a table of contents using all the H1,H2,H3 sections
        /// that were found. Creates a string with local reference to href 
        /// in the md file.
        /// </summary>
        /// <param name="text">string containing the TOC </param>
        /// <returns></returns>
        public string generateTOC(string text)
        {
            string tablecontents = "";
            TOC lasttoc = new TOC();
            lasttoc.hdr = new int[] { 0, 0, 0 };


            for (int i = 0; i < toc.Count; i++)
            {
                if (toc[i].hdr[2] > lasttoc.hdr[2])
                {
                    //if(toc[i].hdr[2]==1)
                    //    tablecontents += "\r\n";
                    string num = toc[i].hdr[0].ToString() + "." + toc[i].hdr[1].ToString() + "." + toc[i].hdr[2].ToString();
                    tablecontents += "\r\n\r\n    " + num + "[" + toc[i].heading + "](#" + CleanNametag(toc[i].heading) + ")";
                }
                else if (toc[i].hdr[1] > lasttoc.hdr[1])
                {
                    //if (toc[i].hdr[1] == 1)
                    //    tablecontents += "\r\n";
                    string num = toc[i].hdr[0].ToString() + "." + toc[i].hdr[1].ToString();
                    tablecontents += "\r\n\r\n  " + num + " [" + toc[i].heading + "](#" + CleanNametag(toc[i].heading) + ")";
                }
                else if (toc[i].hdr[0] > lasttoc.hdr[0])
                {
                    tablecontents += "\r\n\r\n" + toc[i].hdr[0].ToString() + " [" + toc[i].heading + "](#" + CleanNametag(toc[i].heading) + ")";
                }
                lasttoc = toc[i];
            }
            Debug.Print(tablecontents);
            return text.Replace(tocpattern, tablecontents);
        }

        /////////////////////////////////////////////////////////////////////////
        /// <summary>
        /// Loops through math fields, and replaces with markdown equivalent. 
        /// At this time, uses codecog translation, becuase can't figure out how to
        /// install mathjax script into markdown.
        /// </summary>
        public void HandleMath()
        {

            OMaths maths = oWordDoc.OMaths;
            //foreach (OMath math in maths)
            for (int i = oWordDoc.OMaths.Count ; i > 0; i--)
            {
                OMath math = maths[i];
                math.ConvertToNormalText();
                // inline = 1 wdOMathInline
                math.Range.Select();
                WdOMathType type = math.Type;
                string text;
                Range rngFieldCode = math.Range;
                text = rngFieldCode.Text;
                //string urlencoded = Uri.EscapeDataString(text);
                //urlencoded=urlencoded.Replace("%20", "&space");
                string urlencoded = text;
                //urlencoded=urlencoded.Replace("\\\\", "\\");
                urlencoded = urlencoded.Replace(" ", "&space;");
                string field = " <img src=\"https://latex.codecogs.com/svg.latex?" + urlencoded
                    + "\" title=\"" + urlencoded + "\" /> ";
                Debug.Print(field);
                Object styleNormal = "Normal";
                oWord.Selection.set_Style(ref styleNormal);
                oWord.Selection.Range.Bold = 0;
                oWord.Selection.Range.Italic = 0;
                oWord.Selection.InsertBefore(field);

                math.Range.Delete();
                math.Remove();
            }
        }

        /////////////////////////////////////////////////////////////////////////
        /// <summary>
        /// Loop through all fields and replace references with markdown equivalent.
        /// 
        /// </summary>
        public void HandleReferenceFields()
        {
            int counter = 0;
            foreach (Field myField in oWordDoc.Fields)
            {
                Range rngFieldCode = myField.Code;
                String fieldText = rngFieldCode.Text;
                fieldText = fieldText.Trim();  // get rid of optional leading space
                Debug.Print("Field code = " + fieldText + "\n");
                if (fieldText.StartsWith("CITATION"))
                {

                    String fieldType = "CITATION";
                    Int32 fieldLen = fieldType.Length;
                    Int32 endMerge = fieldText.IndexOf("\\");
                    Int32 fieldNameLength = fieldText.Length - endMerge;
                    String fieldName = fieldText.Substring(fieldLen, endMerge - fieldLen);
                    fieldName = fieldName.Trim();

                    // Find index of this reference to 
                    int n = this.bib.Index(fieldName);
                    if (n < 0)
                        Debugger.Break();

                    int citation = new List<int>(cited.Keys).IndexOf(n);
                    if (citation == -1)
                    {
                        counter++;
                        cited.Add(n, counter);
                        citation = counter;

                        // Save citation with new reference in bibliography 
                        string bibentry = "\\[" + citation.ToString() + "\\] " + "<a name=\"Reference_" + citation.ToString() + "\"></a>" + this.bib.Entry(fieldName);
                        references.Add(bibentry);
                    }
                    else
                    {
                        // Citation is to existing reference in bibliography. Use this citation.
                        citation = cited[n];
                    }


                    // select field  and replace with markdown field
                    string field = "\\[[" + citation.ToString() + "]" + "(#Reference_" + citation.ToString() + ")\\]";
                    myField.Select();
                    myField.Delete();
                    //oWord.Selection.InsertBefore(field);
                    oWord.Selection.TypeText(field);

                }
                else if (fieldText.StartsWith("BIBLIOGRAPHY"))
                {

                }
                else if (fieldText.StartsWith("SEQ"))
                {

                }
                else if (fieldText.StartsWith("REF"))
                {

                }
                else if (fieldText.StartsWith("HYPERLINK"))
                {

                }
                else
                {
                    //Debugger.Break();

                }

                // SEQ, REF, HYPERLINK
                // e.g.  HYPERLINK "https://codeocean.com/ieee/signup" \t "_blank"
                // SEQ Figure \* ARABIC

                //// CITATION chandrasekaran2013computational \l 1033 
                //if (myField.Type == WdFieldType.wdFieldCitation)
                //{

                //}
                //else if (myField.Type == WdFieldType.wdFieldBibliography)
                //{

                //}
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="callback"></param>
        public int Init()
        {
            // String totaltext = "";

            foldername = Path.GetDirectoryName(filename) + "\\";
            oWord = new Microsoft.Office.Interop.Word.Application();

           try
            {
                filecopy = foldername + "Readme.docx";
                File.Copy(filename, filecopy, true);
            }
            catch (Exception)
            {
                System.Windows.Forms.MessageBox.Show("Backup copy failed - Please close the backup Readme.docx file. Thank you.");
                bRunning = false;
                return -1;
            }
           try
           {
               filename = foldername + "Readme.docx";
               oWordDoc = oWord.Documents.Open(filename);

               SaveAllImages("");
               oWordDoc.Close(WdSaveOptions.wdDoNotSaveChanges);
               oWord.Quit();
           }
           catch
           {
               return -1;
           }
           oWord.Quit();
           return 0;


        }
        /////////////////////////////////////////////////////////////////////////
        /// <summary>
        ///     Pops dialog to retrieve word file to convert. Saves images, all image and table ranges, and
        /// then processes each paragraph. If image or table, handles. Tables are currently handled as HTML.
        /// If other style, mapping is performed.
        /// Output is streamed to Readme.md in the same folder as the oringal work file. 
        /// </summary>
        public void Init(Action<int> callback)
        {
            // String totaltext = "";

            foldername = Path.GetDirectoryName(filename) + "\\";
            oWord = new Microsoft.Office.Interop.Word.Application();

            //oWordDoc = oWord.Documents.Open(filename);
            //MakeCopy("Readme.docx");
            Status("Copying file to working document readme.docx");
            callback(-1); Thread.Sleep(100);

            // Copy file to backup file called readme.doc
            try
            {
                filecopy = foldername + "Readme.docx";
                File.Copy(filename, filecopy, true);
            }
            catch (Exception)
            {
                System.Windows.Forms.MessageBox.Show("Backup copy failed - Please close the backup Readme.docx file. Thank you.");
                bRunning = false;
                return;
            }
            filename = foldername + "Readme.docx";
            oWordDoc = oWord.Documents.Open(filename);

            // Save and convert images first.
            Status("Converting and saving images");
            callback(-1); Thread.Sleep(100);

            //ConvertShapesToInline(); // Word 2010 versus 2016 issue
            GetImageRanges();  // get  but skip saving as uses clipboard
            SaveAllImages("");

            // status resolving citations and bibliography
            Status("Resolving references");
            callback(-1); Thread.Sleep(100);

            int cnt = oWordDoc.Bibliography.Sources.Count;
            Debug.Print(oWordDoc.Bibliography.ToString());

            string bibxml = "<b:Sources xmlns:b=\"http://schemas.openxmlformats.org/officeDocument/2006/bibliography\" xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/bibliography\" SelectedStyle=\"\">";
            foreach (Source source in oWordDoc.Bibliography.Sources)
            {
                bibxml += source.XML;
            }
            bibxml += "</b:Sources>";
            Debug.Print(bibxml);

            Status("Handling math equations");
            callback(-1); Thread.Sleep(100);
            HandleMath();

            //bib.parseXmlFile(@"C:\Users\michalos\AppData\Roaming\Microsoft\Bibliography\Sources.xml");
            bib.parseXml(bibxml);
            Status("Handling reference fields");
            callback(-1); Thread.Sleep(100);
            HandleReferenceFields();


            Status("Handling subscripts");
            callback(-1); Thread.Sleep(100);
            WdFindAllSubscripts();

            sw = new StringWriter();
            //StreamWriter sw = new StreamWriter(foldername + "Readme.md");

            //MAKING THE APPLICATION VISIBLE
            oWord.Visible = true;
            //oWord.Visible = false; // really don't want to mess around with word doc

            Status("Replacing smart quotes");
            callback(-1); Thread.Sleep(100);
            WdReplaceSmartQuotes();

            GetTableRanges();
            DeleteTableOfContents(); // generally doesnt make sense without page numbering

            //In github markdown, any URL (like http://www.github.com/) will be automatically converted into a clickable link
            //ConvertHyperlinks();
            
            setRunning(true);
            linenum = 1;
        }
        public bool Run(Action<int> callback)
        {
            bInCode = false;
            //for (int i = 1; (i <= oWordDoc.Paragraphs.Count) && this.IsRunning(); i++)
            if ((linenum > oWordDoc.Paragraphs.Count) || !this.IsRunning() )
            {
                callback(oWordDoc.Paragraphs.Count);
                return false;
            }
            for(int i=0; i< 1; i++)
            {
                //Thread.Sleep(100);
                Words(oWordDoc.Paragraphs[linenum]);
                callback(linenum);

                // Look for table in paragraph
                Range tablerange = oWordDoc.Paragraphs[linenum].Range;
                string tabletext = InTable(oWordDoc.Paragraphs[linenum], ref  tablerange);
                if (tabletext != "")
                {
                    tabletext = tabletext.Replace("\a", "");
                    //totaltext += tabletext;
                    sw.Write(tabletext);
                    Range r = oWordDoc.Paragraphs[linenum].Range;
                    while (linenum < oWordDoc.Paragraphs.Count && oWordDoc.Paragraphs[linenum + 1].Range.Start < tablerange.End)
                        linenum++;
                    continue;
                }

                // Look for image in paragraph
                string imagetext = InImage(oWordDoc.Paragraphs[linenum]);
                if (imagetext != "")
                {
                    //totaltext += imagetext;
                    sw.Write(imagetext);

                    continue;
                }

                // Now check for style if here
                Style style = oWordDoc.Paragraphs[linenum].get_Style();
                string line = oWordDoc.Paragraphs[linenum].Range.Text.ToString();
                //System.Diagnostics.Debug.WriteLine(line);

                if (Array.IndexOf(TitleStyle, style.NameLocal) >= 0)
                {
                    string title = md_newline() + "# " + oWordDoc.Paragraphs[linenum].Range.Text.ToString() + md_newline() + "----" + Environment.NewLine;
                    sw.Write(title);
                    //totaltext += md_newline() + "# " + oWordDoc.Paragraphs[i].Range.Text.ToString() + md_newline() + "----" + Environment.NewLine;
                    // http://stackoverflow.com/questions/9721944/automatic-toc-in-github-flavoured-markdown
                    //totaltext+="* auto-gen TOC:\r\n{:toc}";  // no longer used by github
                }
                else if (Array.IndexOf(Heading1, style.NameLocal) >= 0)
                {
                    if (!String.IsNullOrEmpty(line.Trim()))
                    {
                        //totaltext += md_newline() + "# " + line;
                        sw.Write(md_newline() + "# <a name=\"" + CleanNametag(line.Trim()) + "\"></a>" + line);
                        toc.Add(TOC.H1(line.Trim()));
                    }
                }
                else if (Array.IndexOf(Heading2, style.NameLocal) >= 0)
                {
                    if (!String.IsNullOrEmpty(line.Trim()))
                    {
                         sw.Write(md_newline() + "## <a name=\"" + CleanNametag(line.Trim()) + "\"></a>" + line);
                        toc.Add(TOC.H2(line.Trim()));
                    }
                }
                else if (Array.IndexOf(Heading3, style.NameLocal) >= 0)
                {
                    if (!String.IsNullOrEmpty(line.Trim()))
                    {
                        //totaltext += md_newline() + "### " + line;
                        sw.Write(md_newline() + "### <a name=\"" + CleanNametag(line.Trim()) + "\"></a> " + line);
                        toc.Add(TOC.H3(line.Trim()));

                    }
                }
                else if (Array.IndexOf(CodeStyle, style.NameLocal) >= 0)
                {
                    if (!bInCode)
                    {
                        //totaltext += md_newline();
                        sw.Write(md_newline());
                    }
                    bInCode = true;
                    //totaltext += md_newline() + "\t" + line;
                    sw.Write(md_newline() + "\t" + line);
                    continue;
                }
                else if (Array.IndexOf(ListStyle, style.NameLocal) >= 0)
                {
                    //totaltext += md_newline(); //  "\r\n ";
                    sw.Write(md_newline());
                    line = line.TrimStart();
                    Range r = oWordDoc.Paragraphs[linenum].Range;
                    if (r.ListFormat.ListType == WdListType.wdListBullet ||
                        r.ListFormat.ListType == WdListType.wdListPictureBullet)
                    {
                        for (int k = 1; k < r.ListFormat.ListLevelNumber; k++)
                        {
                            //totaltext += "\t";
                            sw.Write("\t");
                        }
                        sw.Write("- ");
                        //totaltext += "- ";
                    }
                    else if (r.ListFormat.ListType == WdListType.wdListNoNumbering)
                    {
                        for (int k = 1; k < r.ListFormat.ListLevelNumber; k++)
                        {
                            sw.Write("\t");
                            //totaltext += "\t";
                        }
                    }
                    else
                    {
                        //totaltext += r.ListFormat.ListValue.ToString() + ". ";
                        sw.Write(r.ListFormat.ListValue.ToString() + ". ");
                    }
                    //totaltext += line;
                    sw.Write(line);
                }
                else
                {
                    // Fixme: handle style - center, left, justified, right
                    ParagraphFormat pf = oWordDoc.Paragraphs[linenum].Format;
                    if (pf.Alignment == WdParagraphAlignment.wdAlignParagraphCenter)
                    {
                        //totaltext += md_newline() + "<p align=\"center\">" + md_newline() + Words(oWordDoc.Paragraphs[i]) + md_newline() + "</p>";
                        sw.Write(md_newline() + "<p align=\"center\">" + md_newline() + Words(oWordDoc.Paragraphs[linenum]) + md_newline() + "</p>");
                    }
                    else
                    {
                        // totaltext += md_newline() + Words(oWordDoc.Paragraphs[i]);
                        sw.Write(md_newline() + Words(oWordDoc.Paragraphs[linenum]));

                    }
                }

                bInCode = false;
            }
            linenum++;
            // Done even if not done.
            return true;
        }

        public void Cleanup()
        {

            if (references.Count > 0)
                sw.Write("# References" + this.md_newline());
            for (int i = 0; i < references.Count; i++)
            {
                //totaltext += references[i] + this.md_newline();
                sw.Write(references[i] + this.md_newline());
            }

            if (bInsertTOC)
            {
                 string totaltext = generateTOC(sw.Get());
                 sw.Set(totaltext);
            }

            //sw.Write(totaltext);
            sw.Close();
            StreamWriter md = new StreamWriter(foldername + "Readme.md");
            md.Write(sw.Get());
            md.Close();

            //  You placed a large amount of content on the clipboard HEADACHE...
            //System.Windows.Forms.Clipboard.SetText("  ");
            //oWord.Application.DisplayAlerts =  WdAlertLevel.wdAlertsNone ;  // Doesn't work
            oWordDoc.Close(WdSaveOptions.wdDoNotSaveChanges);
            oWord.Quit();
            File.Delete(filecopy);
            setDone(true);

        }

        /////////////////////////////////////////////////////////////////////////
        /// <summary>
        ///     Saves the clipboard into the given filename as a gif.  Uses System.Drawing.Image
        /// </summary>
        public void SaveClipboardImage(string filename)
        {
            if (System.Windows.Forms.Clipboard.GetDataObject() != null)
            {
                System.Windows.Forms.IDataObject oDataObj = System.Windows.Forms.Clipboard.GetDataObject();
                if (oDataObj.GetDataPresent(System.Windows.Forms.DataFormats.Bitmap))
                {
                    System.Drawing.Image oImgObj = (System.Drawing.Image)oDataObj.GetData(System.Windows.Forms.DataFormats.Bitmap, true);
                    //To Save as Bitmap
                    //oImgObj.Save("c:\\Test.bmp", System.Drawing.Imaging.ImageFormat.Bmp);
                    //To Save as Jpeg
                    oImgObj.Save(filename, System.Drawing.Imaging.ImageFormat.Gif);
                    //To Save as Gif
                    //oImgObj.Save("c:\\Test.gif", System.Drawing.Imaging.ImageFormat.Gif);
                }
            }

        }

        /////////////////////////////////////////////////////////////////////////
        // Herein under this comment are some VBA routines mapped into C#, and untested.
        // http://labs.physics.berkeley.edu//index.php/Doc_To__Converter 
        /////////////////////////////////////////////////////////////////////////

        /// <summary>
        /// Word replace of string with escape prefix
        /// </summary>
        /// <param name="c"></param>
        private void EscapeCharacter(string c)
        {
            WdReplaceString(c, "\\" + c);
        }

        /////////////////////////////////////////////////////////////////////////
        /// <summary>
        /// Escape HTML sensitive characters in word document.
        /// </summary>
        private void EscapeChars()
        {
            EscapeCharacter("*");
            EscapeCharacter("#");
            EscapeCharacter("_");

            EscapeCharacter("{");
            EscapeCharacter("}");
            EscapeCharacter("[");
            EscapeCharacter("]");

            EscapeCharacter("|");
        }

        /////////////////////////////////////////////////////////////////////////
        /// <summary>
        /// use word automation to make a copy of document. Seems untested and wrong.
        /// </summary>
        /// <param name="filename"> filename</param>
        public void WdMakeCopy(string filename)
        {
            string path = foldername + filename;
            oWordDoc.SaveAs2(path);
            oWordDoc = oWord.Documents.Open(path);
        }

        /////////////////////////////////////////////////////////////////////////
        /// <summary>
        /// Left align word document
        /// </summary>
        public void WdLeftAlign()
        {
            oWordDoc.Select(); // empty selection
            oWordDoc.Range().WholeStory();
            oWordDoc.Range().ParagraphFormat.LeftIndent = oWord.InchesToPoints(0);
            oWordDoc.Range().ParagraphFormat.SpaceBeforeAuto = 0;
            oWordDoc.Range().ParagraphFormat.SpaceAfterAuto = 0;
            oWordDoc.Range().ParagraphFormat.FirstLineIndent = oWord.InchesToPoints(0);
            oWordDoc.Range().ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
        }

        /////////////////////////////////////////////////////////////////////////
        /// <summary>
        /// Convert Word hyperlinks  to markdown equivalent
        /// </summary>
        public void ConvertHyperlinks()
        {
            foreach (Hyperlink link in oWordDoc.Hyperlinks)
            {
                if (link.TextToDisplay.Count() < 1)
                    link.Range.InsertBefore(link.Address);
                else
                    link.Range.InsertBefore("![" + link.Address + " " + link.TextToDisplay + "](" + link.Address + ")");
            }
            while (oWordDoc.Hyperlinks.Count > 0)
            {
                foreach (Hyperlink alink in oWordDoc.Hyperlinks)
                    alink.Range.Delete();
            }
        }

        /////////////////////////////////////////////////////////////////////////
        /// <summary>
        /// Word automation to replace string with new string
        /// </summary>
        /// <param name="findStr"> string to find</param>
        /// <param name="replacementStr"> string to replace with</param>
        private void WdReplaceString(string findStr, string replacementStr)
        {
            object missing = System.Type.Missing;
            object replaceAll = WdReplace.wdReplaceAll;

            oWordDoc.Select();
            oWordDoc.Range().WholeStory();
            Range r = oWordDoc.Range();
            r.Find.ClearFormatting();
            r.Find.Text = findStr;
            r.Find.Replacement.Text = replacementStr;
            r.Find.Forward = true;
            r.Find.Wrap = WdFindWrap.wdFindContinue;
            r.Find.Format = false;
            r.Find.MatchCase = false;
            r.Find.MatchWholeWord = false;
            r.Find.MatchWildcards = false;
            r.Find.MatchSoundsLike = false;
            r.Find.MatchAllWordForms = false;
            r.Find.Execute(ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref replaceAll,
                ref missing, ref missing, ref missing, ref missing);
        }

        /////////////////////////////////////////////////////////////////////////
        /// <summary>
        /// Replace word's smart quotes with straight quotes
        /// </summary>
        public void WdReplaceSmartQuotes()
        {
            oWord.Options.AutoFormatAsYouTypeReplaceQuotes = false;
            WdReplaceString("“", "\"");
            WdReplaceString("”", "\"");
            WdReplaceString("‘", "'");
            WdReplaceString("’", "'");
        }

        /////////////////////////////////////////////////////////////////////////
        /// <summary>
        /// Replace word markup of heading to markdown 
        /// </summary>
        /// <param name="styles"> array of styles to change</param>
        /// <param name="headerPrefix">markdown prefix to line</param>
        public void MdReplaceHeading(string[] styles, string headerPrefix)
        {

            oWordDoc.Select();

            Styles allstyles = oWordDoc.Styles;
            int i;
            for (i = 0; i < allstyles.Count; i++)
                if (allstyles[i].NameLocal == styles[0])
                    break;

            oWord.Selection.Find.ClearFormatting();
            oWord.Selection.Find.Format = true;
            oWord.Selection.Find.MatchCase = false;
            oWord.Selection.Find.MatchWholeWord = false;
            oWord.Selection.Find.MatchWildcards = false;
            oWord.Selection.Find.MatchSoundsLike = false;
            oWord.Selection.Find.MatchAllWordForms = false;
            oWord.Selection.Find.Forward = true;
            oWord.Selection.Find.Wrap = WdFindWrap.wdFindContinue;
            oWord.Selection.Find.set_Style(allstyles[i]);

            // fIXME there are multiple selection for each style match.
            while (oWord.Selection.Find.Execute())
            {
                if (oWord.Selection.Text.IndexOf("\r") == 0)
                {
                    // Just process the chunk before any newline characters
                    // We'll pick-up the rest with the next search
                    oWord.Selection.Collapse();
                    oWord.Selection.MoveEndUntil("\r");
                }

                // Don't bother to markup newline characters (prevents a loop, as well)
                if (oWord.Selection.Text != "\n")
                {
                    oWord.Selection.InsertBefore(headerPrefix);
                    oWord.Selection.InsertBefore("\n");
                    oWord.Selection.InsertAfter(headerPrefix);
                }
                oWord.Selection.Range.set_Style(WdBuiltinStyle.wdStyleNormal);
            }
        }

        /////////////////////////////////////////////////////////////////////////
        /// <summary>
        /// Convert Word Heading1 to markdown equivalent
        /// </summary>
        private void ConvertH1()
        {
            MdReplaceHeading(Heading1, "#");
        }

        /////////////////////////////////////////////////////////////////////////
        /// <summary>
        /// Convert Word Heading2 to markdown equivalent
        /// </summary>
        private void ConvertH2()
        {
            MdReplaceHeading(Heading2, "##");
        }

        /////////////////////////////////////////////////////////////////////////
        /// <summary>
        ///  Replace Word heading 3 with markdown heading 3
        /// </summary>
        private void ConvertH3()
        {
            MdReplaceHeading(Heading3, "###");
        }
    }
}
