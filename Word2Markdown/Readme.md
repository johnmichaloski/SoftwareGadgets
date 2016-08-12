
#Word 2 Markdown README 
----

7/2/2016 9:47:00 PM
This document describes the procedure to convert Microsoft Word documents into Markdown documents. There are many variants of markdown, but fortunately it is a rather limited set. 
#Instructions
The code is a C# application. Follow these instructions to generate a readme.md and images folder.
 1. Double click  the Word2Markdown.exe file to run the application 
 2. select a Word document (docx), 
 3. program generates a copy (Readme.docx) in the same folder as the original document, 
 4. the program locates and saves all inline pictures in the document under the folder images/ in the name format images/image#.jpg and will cross reference to this image in the markdown
 5. The program generates a Readme.md document with all the mapings.
#Configuration

The Github web site uses markdown as its readme format to describe a repository. This initial goal of this executable was to produce readable Readme.md that included images.

<TABLE>
<TR>
<TD>Word <BR></TD>
<TD>Markdown<BR></TD>
</TR>
<TR>
<TD>Heading 1, Heading 2, … styles<BR></TD>
<TD>Corresponding number of #<BR></TD>
</TR>
<TR>
<TD>Image<BR></TD>
<TD>Save image into images/image#.jpg<BR>Insert markdown:<BR>![Figure #](./images/image#.jpg?raw=true)<BR></TD>
</TR>
<TR>
<TD>Bold Font<BR></TD>
<TD>** text **<BR></TD>
</TR>
<TR>
<TD>Underline Font<BR></TD>
<TD>_ text _<BR></TD>
</TR>
<TR>
<TD>Code Style<BR></TD>
<TD>\t code line1<BR>\t code line2<BR></TD>
</TR>
<TR>
<TD>Table<BR></TD>
<TD>Unclear, just used html to represent table,<BR>Remaining problems with paragraph count using table<BR>Github:<BR>First Header | Second Header<BR>------------ | -------------<BR>Content from cell 1 | Content from cell 2<BR>Content in the first column | Content in the second column<BR></TD>
</TR>
<TR>
<TD>List<BR></TD>
<TD>Unordered (bullet) -  *<BR>Numbered – #. Etc.<BR></TD>
</TR>
<TR>
<TD>Hyperlink <BR></TD>
<TD>Hyperlink![Text](URL)<BR>e.g., http://github.com - automatic!<BR>[GitHub](http://github.com)<BR></TD>
</TR>
<TR>
<TD>Task Lists<BR></TD>
<TD>Unhandled<BR>- [x] @mentions, #refs, [links](), **formatting**, and <del>tags</del> supported<BR>- [x] list syntax required (any unordered or ordered list supported)<BR>- [x] this is a complete item<BR>- [ ] this is an incomplete item<BR></TD>
</TR>
<TR>
<TD>Strikethrough<BR><BR><BR></TD>
<TD>~~this~~  - this appears crossed out.<BR></TD>
</TR>
</TABLE>

#Modification
The program is a C# windows application. Is uses windows office word interoperability to do the word document manipulation. Originally the program was a VBA program, but saving the images could not be done in a reasonable amount of time. So the program was rewritten into visual studio C# 2010, which was quite simple.
There is only one C# class to perform the Word to Markdown conversion. This class is called WordAutomation and does all the word automation and conversions.  Of interest is that Word styles can vary from document to document.  Under the   WordAutomation class definition are the arrays that define the styles to search for to map heading, code, etc. into the corresponding Markdown equivalent. These straing arrays are currently defines as:
	        public string []  ListStyle = {"List Paragraph"};
	        public string []  CodeStyle = {"BoxedCode"};
	        public string []  TitleStyle = {"Title"};
	        public string []  Heading1 = {"Heading 1", "Heading1", "H1"};
	        public string []  Heading2 = {"Heading 2"};
	        public string []  Heading3 = {"Heading 3"};

You can modify these string and recompile the program  to effect the changes.. Obviously an ini file or .Net config file could be used to modify these mappings.


![Figure1](./images/image1.jpg?raw=true)


#WordAutomation Class Description
Handles the word automation and conversion into Markdown. 
##
##Fields
**_TablesRanges_**
List of all tables and their ranges. 
**_ImageRanges_**
List of all images and their ranges. 
**_oWordDoc_**
Word application COM variable. 
**_oWord_**
Word document COM variable. 
**_filename_**
Filename of the current word file undergoing conversione. 
**_foldername_**
Foldername of the current word file undergoing conversion. 
##
##Methods
**_GetTableRanges_**
Extracts all the table ranges into the TablesRanges data structure . 
**_GetImageRanges_**
Extracts all the image ranges into the ImageRanges data structure . 
**_SaveAllImages(System.String)_**
Saves all the image ranges into the folder images under the current filename folder. Uses clipboard to copy and paste into image handler, which saves AS JPG. 
**_InImage(Microsoft.Office.Interop.Word.Paragraph)_**
Determine if given paragraph p in found in the ImageRanges. 
**_InTable(Microsoft.Office.Interop.Word.Paragraph,System.Int32@)_**
Determine if given paragraph p in found in the TablesRanges. 
**_Words(Microsoft.Office.Interop.Word.Paragraph)_**
Looks at each word in a paragraph and formats font if word is bold or underlined. 

**_Init_**
Pops dialog to retrieve word file to convert. Saves images, all image and table ranges, and then processes each paragraph. If image or table, handles. Tables are currently handled as HTML. If other style, mapping is performed. Output is streamed to Readme.md in the same folder as the oringal work file. 
**_SaveClipboardImage(System.String)_**
Saves the clipboard into the given filename as a jpg. Uses System.Drawing.Image 
