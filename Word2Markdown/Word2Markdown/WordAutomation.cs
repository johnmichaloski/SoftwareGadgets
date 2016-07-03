using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Word;
using Microsoft.Office.Interop;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools;
using System.IO;


namespace Word2Markdown
{
    /// <summary>
    ///     Handles the word automation and conversion into Markdown.
    /// </summary>
    class WordAutomation
    {
        // These are the styles to match against 
        public string[] ListStyle = { "List Paragraph" };
        public string[] CodeStyle = { "BoxedCode" };
        public string[] TitleStyle = { "Title" };
        public string[] Heading1 = { "Heading 1", "Heading1", "H1" };
        public string[] Heading2 = { "Heading 2" };
        public string[] Heading3 = { "Heading 3" };

        public WordAutomation() { }
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
        ///     Extracts all the table ranges into the TablesRanges data structure .
        /// </summary>
        public void GetTableRanges()
        {
            for (int iCounter = 1; iCounter <= oWordDoc.Tables.Count; iCounter++)
            {
                Range TRange = oWordDoc.Tables[iCounter].Range;
                TablesRanges.Add(TRange);
            }
        }
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
        /// <summary>
        ///     Saves all the image ranges into the folder images under the current filename folder.
        ///     Uses clipboard to copy and paste into image handler, which saves AS JPG.
        /// </summary>
        public void SaveAllImages(string directory)
        {
            int nimages = 1;
            System.IO.Directory.CreateDirectory(foldername + "images");

            // Save all images as 
            foreach (InlineShape shape in oWordDoc.InlineShapes)
            {

                if (shape.Type == WdInlineShapeType.wdInlineShapePicture)
                {
                    shape.Select();
                    oWord.Selection.CopyAsPicture();
                    SaveClipboardImage(foldername + "images\\image" + nimages.ToString() + ".jpg");
                    nimages++;
                }
                else if (shape.Type == WdInlineShapeType.wdInlineShapeLinkedPicture)
                {
                    System.Diagnostics.Debug.WriteLine(shape.Range.Text);
                    shape.Select();
                    oWord.Selection.CopyAsPicture();
                    SaveClipboardImage(foldername + "images\\image" + nimages.ToString() + ".jpg");
                    nimages++;
                }
            }
        }
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
                totaltext = "![Figure" + iCounter.ToString() + "](./images/image" + iCounter.ToString() + ".jpg?raw=true)";
            }
            return totaltext;
        }
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
                totaltext += "\r\n<TABLE>";
                foreach (Row aRow in oWordDoc.Tables[iCounter].Rows)
                {
                    totaltext += "\r\n<TR>";

                    foreach (Cell aCell in aRow.Cells)
                    {
                        currLine = aCell.Range.Text;
                        char[] delimiterChars = { '\r' };
                        string[] words = currLine.Split(delimiterChars);
						// Fixme  - single line cells only have \r\a
                        currLine = currLine.Replace("\r", "<BR>");  

                        // Count the number of paragraphs?
                        totaltext += "\r\n<TD>" + currLine + "</TD>";
                    }
                    totaltext += "\r\n</TR>";
                }
                totaltext += "\r\n</TABLE>";
            }
            return totaltext;
        }
        /// <summary>
        /// Let's remove trailing spaces and comma.
        /// </summary>
        /// <param name="inputString"> string to trim and save </param>
        /// <returns></returns>
 
        public string TrimEnd(string inputString, ref string trailing)
        {
            trailing = "";
            string delims = " \a\r\n\t";
            int i = inputString.Count()-1;
            while (i >= 0 && delims.IndexOf(inputString[i]) >= 0)
            {
                trailing = inputString[i] + trailing;
                i--;
            }
           return inputString.Trim(" \r\n\t\a".ToArray());
         }

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
        /// <summary>
        ///     Pops dialog to retrieve word file to convert. Saves images, all image and table ranges, and
        /// then processes each paragraph. If image or table, handles. Tables are currently handled as HTML.
        /// If other style, mapping is performed.
        /// Output is streamed to Readme.md in the same folder as the oringal work file. 
        /// </summary>
        public void Init()
        {
            String totaltext = "";

            var fileDialog = new System.Windows.Forms.OpenFileDialog();
            fileDialog.Filter = "Word files (*.docx)|";
            var result = fileDialog.ShowDialog();
            if (result == System.Windows.Forms.DialogResult.Cancel)
                return;
            filename = fileDialog.FileName;
            foldername = Path.GetDirectoryName(filename) + "\\";
            oWord = new Microsoft.Office.Interop.Word.Application();

            //oWordDoc = oWord.Documents.Open(filename);
            //MakeCopy("Readme.docx");
            File.Copy(filename, foldername + "Readme.docx",true);
            filename = foldername + "Readme.docx";
            oWordDoc = oWord.Documents.Open(filename);

            StreamWriter sw = new StreamWriter(foldername + "Readme.md");

            //MAKING THE APPLICATION VISIBLE
            oWord.Visible = true;

            SaveAllImages("");
            GetTableRanges();
            GetImageRanges();

            for (int i = 1; i <= oWordDoc.Paragraphs.Count; i++)
            {
                Words(oWordDoc.Paragraphs[i]);

                // Look for table in paragraph
                Range tablerange = oWordDoc.Paragraphs[i].Range;
                string tabletext = InTable(oWordDoc.Paragraphs[i], ref  tablerange );
                if (tabletext != "")
                {
                    totaltext += tabletext;
					Range r = oWordDoc.Paragraphs[i].Range;
					while(i < oWordDoc.Paragraphs.Count &&  oWordDoc.Paragraphs[i+1].Range.Start < tablerange.End ) 
						i++ ;
                    continue;
                }

                // Look for image in paragraph
                string imagetext = InImage(oWordDoc.Paragraphs[i]);
                if (imagetext != "")
                {
                    totaltext += imagetext;
                    continue;
                }

                // Now check for style if here
                Style style = oWordDoc.Paragraphs[i].get_Style();
                string line = oWordDoc.Paragraphs[i].Range.Text.ToString();
                System.Diagnostics.Debug.WriteLine(line);
                //if (line == "Uninstall")
                //{
                //    System.Diagnostics.Debugger.Break();
                //}
                if (Array.IndexOf(TitleStyle, style.NameLocal) >= 0)
                {
                    totaltext += "\r\n#" + oWordDoc.Paragraphs[i].Range.Text.ToString() + " \r\n----\r\n";
                }
                else if (Array.IndexOf(Heading1, style.NameLocal) >= 0)
                {
                    totaltext += "\r\n#" + line;
                }
                else if (Array.IndexOf(Heading2, style.NameLocal) >= 0)
                {
                    totaltext += "\r\n##" + line;
                }
                else if (Array.IndexOf(Heading3, style.NameLocal) >= 0)
                {
                    totaltext += "\r\n###" + line;
                }
                else if (Array.IndexOf(CodeStyle, style.NameLocal) >= 0)
                {
                    totaltext += "\r\n\t" + line;
                }
                else if (Array.IndexOf(ListStyle, style.NameLocal) >= 0)
                {
                    totaltext += "\r\n ";
                    Range r = oWordDoc.Paragraphs[i].Range;
                    if (r.ListFormat.ListType == WdListType.wdListBullet ||
                        r.ListFormat.ListType == WdListType.wdListPictureBullet)
                    {
                        for (int k = 1; k < r.ListFormat.ListLevelNumber; k++)
                            totaltext += "\t";
                        totaltext += "*";
                    }
                    else if (r.ListFormat.ListType == WdListType.wdListNoNumbering)
                    {
                        for (int k = 1; k < r.ListFormat.ListLevelNumber; k++)
                            totaltext += "\t";
                    }
                    else
                    {
                        totaltext += r.ListFormat.ListValue.ToString() + ". ";
                    }
                    totaltext += line;
                }
                else
                {
                    totaltext += "\r\n" + Words(oWordDoc.Paragraphs[i]);
                }
            }
            sw.Write(totaltext);
            sw.Close();
            System.Windows.Forms.Clipboard.SetText(" ");
            oWordDoc.Close();
            oWord.Quit();
        }
        /// <summary>
        ///     Saves the clipboard into the given filename as a jpg.  Uses System.Drawing.Image
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
                    oImgObj.Save(filename, System.Drawing.Imaging.ImageFormat.Jpeg);
                    //To Save as Gif
                    //oImgObj.Save("c:\\Test.gif", System.Drawing.Imaging.ImageFormat.Gif);
                }
            }

        }
        // Herein under this comment are some VBA routines mapped into C#, and untested.
        // http://labs.physics.berkeley.edu//index.php/Doc_To__Converter 
        private void EscapeCharacter(string c)
        {
            ReplaceString(c, "\\" + c);
        }

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
        public void MakeCopy(string filename)
        {
            string path = foldername + filename;
            oWordDoc.SaveAs2(path);
            oWordDoc = oWord.Documents.Open(path);
        }
        public void LeftAlign()
        {
            oWordDoc.Select();
            oWordDoc.Range().WholeStory();
            oWordDoc.Range().ParagraphFormat.LeftIndent = oWord.InchesToPoints(0);
            oWordDoc.Range().ParagraphFormat.SpaceBeforeAuto = 0;
            oWordDoc.Range().ParagraphFormat.SpaceAfterAuto = 0;
            oWordDoc.Range().ParagraphFormat.FirstLineIndent = oWord.InchesToPoints(0);
            oWordDoc.Range().ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
        }
        public void ConvertHyperlinks()
        {
            foreach (Hyperlink link in oWordDoc.Hyperlinks)
            {
                if (link.TextToDisplay.Count() < 1)
                    link.Range.InsertBefore(link.Address);
                else
                    link.Range.InsertBefore("[" + link.Address + " " + link.TextToDisplay + "]");
            }
            while (oWordDoc.Hyperlinks.Count > 0)
            {
                foreach (Hyperlink alink in oWordDoc.Hyperlinks)
                    alink.Range.Delete();
            }
        }
        private void ReplaceString(string findStr, string replacementStr)
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
        public void ReplaceSmartQuotes()
        {
            oWord.Options.AutoFormatAsYouTypeReplaceQuotes = false;
            ReplaceString("“", "\"");
            ReplaceString("”", "\"");
            ReplaceString("‘", "'");
            ReplaceString("’", "'");
        }
        public void ReplaceHeading(string[] styles, string headerPrefix)
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
        private void ConvertH1()
        {
            ReplaceHeading(Heading1, "#");
        }
        private void ConvertH2()
        {
            ReplaceHeading(Heading2, "##");
        }
        private void ConvertH3()
        {
            ReplaceHeading(Heading3, "###");
        }
    }
}
