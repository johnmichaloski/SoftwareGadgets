
# Word 2 Markdown README 
----

 DATE  \@ "M/d/yyyy h:mm:ss am/pm"  \* MERGEFORMAT 
This document describes the Windows application to convert Microsoft Word documents into Markdown documents. There are many variants of markdown, only github and sourceforge are supported. Images are converted into gifs and saved to be included in markdown file. Markdown conversion also supports tables, code, table of contents, and removes a lot of non-UTF characters, but not seemingly all.
**Contents**

# Table of Contents


1 [Installation](#Installation)

2 [Operating Instructions](#Operating_Instructions)

3 [Configuration](#Configuration)

  3.1 [Manually configuring](#Manually_configuring)

4 [Modification](#Modification)
# <a name="Installation"></a>Installation
Download the msi script for the wd2md application. It will install the application on a windows platform assuming you have administrator rights. It should download any .Net framework version upgrades, but this implementation uses a dated .net framework. One thing THAT WAS NOT TESTED is converting a variety of Word versions into markdown. Only Word 2010 and Word 2016 were tested.
The installation process is a series of next next next clicks:

![Figure1](./Word2MarkdownReadme_images/Word2MarkdownReadme_image1.gif)



</div>
Next:

![Figure2](./Word2MarkdownReadme_images/Word2MarkdownReadme_image2.gif)



</div>
Next:

![Figure3](./Word2MarkdownReadme_images/Word2MarkdownReadme_image3.gif)



</div>


![Figure4](./Word2MarkdownReadme_images/Word2MarkdownReadme_image4.gif)



</div>


![Figure5](./Word2MarkdownReadme_images/Word2MarkdownReadme_image5.gif)



</div>

To access the program you need to use the start menu (whatever it currently is in windows) :

![Figure6](./Word2MarkdownReadme_images/Word2MarkdownReadme_image6.gif)



</div>


# <a name="Operating_Instructions"></a>Operating Instructions
The code is a C# Windows application. It has only been tested on 32 and 64 bit Windows 7 operating systems.
Follow these instructions to generate a readme.md and images folder.
 1. Double click  the wd2md.exe file to run the application or select Word2Markdown in the Command list:
 
 2. To select a Word document (docx), click on the ellipsis (…) under the work document label. A file selection popup dialogue will appear and you can select the Word file you want converted.

![Figure7](./Word2MarkdownReadme_images/Word2MarkdownReadme_image7.gif)



</div>
 You need to select the markdown type: either github or sourceforge. As a non-standard "standard" there are lots of variations. In the end, only the image embedded was different.
 
 You can then press the GO button and the markdown document will automatically be named the Word document title plus the ".md" extension and placed in the same folder. During the conversions images will get generated and placed in the folder under the markdown document named word document title plus "_image". All images will be placed in this folder.
 Below the word document Word2MarkdownReadme.docx (word document extension is docx) will be converted into markdown, and the file Word2MarkdownReadme.md and the Word2MarkdownReadme_images folder will be created with an images in the documented converted to gif and copied into this folder.
 

![Figure8](./Word2MarkdownReadme_images/Word2MarkdownReadme_image8.gif)



</div>
 
 You can then rename the file to Readme.md, and the image files and folder will still work:
 

![Figure9](./Word2MarkdownReadme_images/Word2MarkdownReadme_image9.gif)



</div>
 
 The wd2md conversion is not fool-proof and there can be problems.  If the wd2md conversion fails or you interrupt the conversion process, the word backup document copy is left open and can hang the following conversions process.
 

![Figure10](./Word2MarkdownReadme_images/Word2MarkdownReadme_image10.gif)



</div>
 
 You will need to start the task manager, click to processes and kill thewinword.exe process (running in the background) to release the backup word copy.

![Figure11](./Word2MarkdownReadme_images/Word2MarkdownReadme_image11.gif)



</div>
 This can be a problem if you are normally using Word for editing.  wd2md not perfect, and won't intentionally harm your system.
 
 3. The wd2md program generates a copy of the Word document (for now Readme.docx) in the same folder as the original word document,  Then the dialog presents a progress bar in the status area at the bottom of the dialog that increments according to the processing of word paragraph.

![Figure12](./Word2MarkdownReadme_images/Word2MarkdownReadme_image12.gif)



</div>
 
 In case you were wondering the 1.0.6528.23997 is the C# automatic versioning. The 6528 refers to the number of days since Jan 1 2000 and the 23997 is the number of seconds since 12AM divided by 2. It will do.
 
 4. When completed a Done dialog pops up, and an OK is required to dismiss the popup and return to the wd2md dialog:

![Figure13](./Word2MarkdownReadme_images/Word2MarkdownReadme_image13.gif)



</div>
 
 
 After the Done ok click, the progress bar is dismissed, and the program is ready to convert another word document:

![Figure14](./Word2MarkdownReadme_images/Word2MarkdownReadme_image14.gif)



</div>
 5. the program locates and saves all inline pictures in the document under the folder images/ in the name format worfile_images/image#.jpg and will cross reference to this image in the markdown
 6. The program generates a Markdown document with all the mappings and embedded images.
# <a name="Configuration"></a>Configuration
The Github web site uses markdown as its readme format to describe a repository. This initial goal of this executable was to produce readable Readme.md that included images.

<table>
<tr>
<td>Word <br></td>
<td>Markdown<br></td>
</tr>
<tr>
<td>Heading 1, Heading 2, … styles<br></td>
<td>Corresponding number of #<br></td>
</tr>
<tr>
<td>Image<br></td>
<td>Save image into images/image#.jpg<br>Insert markdown:<br>![Figure #](./images/image#.jpg?raw=true)<br></td>
</tr>
<tr>
<td>Bold Font<br></td>
<td>** text **<br></td>
</tr>
<tr>
<td>Underline Font<br></td>
<td>_ text _<br></td>
</tr>
<tr>
<td>Code Style<br></td>
<td>\t code line1<br>\t code line2<br></td>
</tr>
<tr>
<td>Table<br></td>
<td>Unclear, just used html to represent table,<br>Remaining problems with paragraph count using table<br>Github:<br>First Header | Second Header<br>------------ | -------------<br>Content from cell 1 | Content from cell 2<br>Content in the first column | Content in the second column<br></td>
</tr>
<tr>
<td>List<br></td>
<td>Unordered (bullet) -  *<br>Numbered – #. Etc.<br></td>
</tr>
<tr>
<td>Hyperlink <br></td>
<td>Hyperlink![Text](URL)<br>e.g., http://github.com - automatic!<br>[GitHub]( HYPERLINK "http://github.com" )<br></td>
</tr>
<tr>
<td>Task Lists<br></td>
<td>Unhandled<br>- [x] @mentions, #refs, [links](), **formatting**, and <del>tags</del> supported<br>- [x] list syntax required (any unordered or ordered list supported)<br>- [x] this is a complete item<br>- [ ] this is an incomplete item<br></td>
</tr>
<tr>
<td>Strikethrough<br><br><br></td>
<td>~~this~~  - this appears crossed out.<br></td>
</tr>
</table>
## <a name="Manually_configuring"></a>Manually configuring
There is the ability to manually configure the wd2md application. The file Config.ini is an "ini" file format (sections with keys and values) as shown below. Only Heading1 has multiple entries delimited by a comma. Note, beginning and trailing spaces are removed; however, space inside the text is not.

	[STYLES]
	List  = List Paragraph
	Code  = BoxedCode 
	Title  =      Title
	Heading1  =  Heading 1,Heading1, H1
	Heading2  =  Heading 2
	Heading3  =  Heading 3
The Config.ini file contains the matching styles that get mapped into the corresponding markdown formats. You can add more styles in the ini file, for example, another Heading2 style:

	Heading2  =  Heading 2, H2
The values are parsed as comma separated values, with extra space around the style names removed.
Below shows the internal name, the markdown equivalent and the Word Style name used. 
<table>
<tr>
<td>Name<br></td>
<td>Markdown<br></td>
<td>Word<br></td>
</tr>
<tr>
<td>List<br></td>
<td><br></td>
<td>List Paragraph <br></td>
</tr>
<tr>
<td>Code<br></td>
<td>Indent by tab<br></td>
<td>BoxedCode (my style name) you can add yours, see configuration.<br></td>
</tr>
<tr>
<td>Title<br></td>
<td><br></td>
<td>Title<br></td>
</tr>
<tr>
<td>Heading1<br></td>
<td>#<br></td>
<td>Heading 1<br></td>
</tr>
<tr>
<td>Heading2<br></td>
<td>##<br></td>
<td>Heading 2<br></td>
</tr>
<tr>
<td>Heading3<br></td>
<td>###<br></td>
<td>Heading 3<br></td>
</tr>
</table>

Thus for example the code style is "BoxedCode" but you can change or add another word style name in the config.ini file that the application reads for style mappings. You will note all the styles in this word document are the styles matched by the wd2md application.
# <a name="Modification"></a>Modification
The program is a C# windows application. Is uses windows office word interoperability to do the word document manipulation. Originally the program was a VBA program, but saving the images could not be done in a reasonable amount of time. So the program was rewritten into visual studio C# 2010, which was quite simple.
There is only one C# class to perform the Word to Markdown conversion. This class is called WordAutomation and does all the word automation and conversions.  Of interest is that Word styles can vary from document to document.  Under the   WordAutomation class definition are the arrays that define the styles to search for to map heading, code, etc. into the corresponding Markdown equivalent. These string arrays are currently defines as:

	        public string []  ListStyle = {"List Paragraph"};
	        public string []  CodeStyle = {"BoxedCode"};
	        public string []  TitleStyle = {"Title"};
	        public string []  Heading1 = {"Heading 1", "Heading1", "H1"};
	        public string []  Heading2 = {"Heading 2"};
	        public string []  Heading3 = {"Heading 3"};

You can modify these strings and recompile the program to effect the changes... Obviously an ini file or .Net config file could be used to modify these mappings.

Errors:
 1. Running word application has backup file open, as the application did not shut down gracefully

![Figure15](./Word2MarkdownReadme_images/Word2MarkdownReadme_image15.gif)



</div>
 2. Need config.ini file in application folder:

![Figure16](./Word2MarkdownReadme_images/Word2MarkdownReadme_image16.gif)



</div>
 3. Wrong install platform version:
 If you click the msi file you get this message box:

![Figure17](./Word2MarkdownReadme_images/Word2MarkdownReadme_image17.gif)



</div>
If you click the setup.exe you get this message box:

![Figure18](./Word2MarkdownReadme_images/Word2MarkdownReadme_image18.gif)



</div>

You need 32  bit install for 32 machines and 64 for 64-bit machines.