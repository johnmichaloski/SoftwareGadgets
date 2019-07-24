using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using System.Diagnostics;

#pragma warning disable 1591
namespace wd2md
{
    public class XmlBibAuthor
    {
        public string first;
        public string middle;
        public string last;
    }
    public class XmlBibAuthorList
    {
        public List<XmlBibAuthor> authors;

        /////////////////////////////////////////////////////////////////////////
        /// <summary>
        /// 
        /// </summary>
        /// <param name="namelist"></param>
        /// <param name="nsm"></param>
        /// <returns></returns>
        public List<XmlBibAuthor> parseAuthors(XmlNode namelist, XmlNamespaceManager nsm)
        {
            XmlNodeList persons = namelist.SelectNodes("b:Author/b:Author/b:NameList/b:Person", nsm);
            authors = new List<XmlBibAuthor>();
            foreach (XmlNode node in persons)
            {
                XmlBibAuthor author = new XmlBibAuthor();
                author.last = null;
                author.first = null;
                XmlNode tag = node.SelectSingleNode("b:Last", nsm);
                if (tag != null)
                {
                    author.last = tag.InnerText;
                    author.last = author.last.Replace("\"","");
                }

                tag = node.SelectSingleNode("b:First", nsm);
                if (tag != null)
                {
                    author.first = tag.InnerText;
                    author.first = author.first.Replace("\"", "");
                }
 
                tag = node.SelectSingleNode("b:Middle", nsm);
                if (tag != null)
                    author.middle = tag.InnerText;

                if (author.first != null || author.last != null)
                    authors.Add(author);
            }
            return authors;
        }

        /////////////////////////////////////////////////////////////////////////
        /// <summary>
        /// Creates a string author list from a list of authors
        /// </summary>
        /// <param name="authors"></param>
        /// <returns></returns>
        public static String Entry(List<XmlBibAuthor> authors)
        {
            String str = "";
            for (int i = 0; i < authors.Count; i++)
            {
                XmlBibAuthor author = authors[i];
                if (author.first != null)
                    str += author.first[0] + ". ";

                if (author.middle != null)
                    str += author.middle[0] + ". ";

                if (author.last != null)
                    str += author.last;

                if ((i + 1) < authors.Count && str.Length>0)
                    str += ", ";

                if ((i + 1) == (authors.Count - 1))
                    str += " and ";
            }

            return str;
        }
    }

    /////////////////////////////////////////////////////////////////////////
    /// <summary>
    /// 
    /// </summary>
    public class XmlBibEntry
    {

        /// <summary>
        /// List of fields of each bib type.
        /// </summary>
        public Dictionary<string, IEnumerable<String>> fields;

        /// <summary>
        /// string dictionary of bib attribute and value
        /// </summary>
        public Dictionary<string, string> bibentries;

        /// <summary>
        /// List of authors
        /// </summary>
        public List<XmlBibAuthor> authors;

        /// <summary>
        /// The tag that identifies a bib entry
        /// </summary>
        public string tag;

        /////////////////////////////////////////////////////////////////////////
        /// <summary>
        /// Constructor. Builds the fields for various bib types.
        /// </summary>
        public XmlBibEntry()
        {
            fields = new Dictionary<string, IEnumerable<String>>();
            bibentries = new Dictionary<string, string>();
            authors = new List<XmlBibAuthor>();
            this.fields.Add("Book", new List<String>(new string[] { 
                "Title",
                "Year", 
                "Month", 
                "Publisher", 
                 "Pages",
               "URL" }));
            this.fields.Add("BookSection", new List<String>(new string[] { 
                "Title",
                "Year", 
                "Month", 
                "Publisher", 
                 "Pages",
                 "BookTitle",
               "URL" }));
            this.fields.Add("JournalArticle", new List<String>(new string[] { 
                "Title", 
                "Year", 
                "Month", 
                "Volume", 
                "Number", 
                "JournalName", 
                "Pages",
                "URL" }));
            this.fields.Add("ArticleInAPeriodical", new List<String>(new string[] { 
                "Title", 
                "Year", 
                "Month", 
                "Volume", 
                "Number", 
                "JournalName", 
                "Pages",
                "URL" }));
            this.fields.Add("ConferenceProceedings", new List<String>(new string[] { 
                "Title", 
                "Year", 
                "Month", 
                "Volume", 
                "Number", 
                "ConferenceName",
                "Pages",
                "URL" }));
            this.fields.Add("Report", new List<String>(new string[] { 
                "Title", 
                "Year", 
                "Month", 
                "Volume", 
                "Number", 
                "ConferenceName",
                "Pages",
                "URL" }));
            this.fields.Add("InternetSite", new List<String>(new string[] { 
                "Title", 
                "Year", 
                "Month", 
                "Volume", 
                "Number", 
                "ConferenceName",
                "Pages",
                "URL" }));
            this.fields.Add("DocumentFromInternetSite", new List<String>());
            this.fields.Add("ElectronicSource", new List<String>());
            this.fields.Add("Art", new List<String>());
            this.fields.Add("SoundRecording", new List<String>());
            this.fields.Add("Performance", new List<String>());
            this.fields.Add("Film", new List<String>());
            this.fields.Add("Interview", new List<String>());
            this.fields.Add("Patent", new List<String>());
            this.fields.Add("Case", new List<String>());
            // Jabref add BIBTEX_HowPublished for URL
            this.fields.Add("Misc", new List<String>(new string[] { "Title", "Year", "Month", "Publisher", "Comments", "URL" }));
        }

        /////////////////////////////////////////////////////////////////////////
        /// <summary>
        /// Builds a string for a bib entry.
        /// </summary>
        /// <returns></returns>
        public string Entry()
        {
            string str = "";
            str += XmlBibAuthorList.Entry( this.authors);
            str += ". ";
            if (bibentries.ContainsKey("Title"))
                str += this.bibentries["Title"] + ". ";
            if (bibentries.ContainsKey("JournalName"))
                str += this.bibentries["JournalName"] + ", ";
            if (bibentries.ContainsKey("ConferenceName"))
                str += "In " + this.bibentries["ConferenceName"] + ", ";
            if (bibentries.ContainsKey("BookTitle"))
                str += "In " + this.bibentries["BookTitle"] + ", ";
            if (bibentries.ContainsKey("Volume"))
                str += this.bibentries["Volume"] ;
            if (bibentries.ContainsKey("Number"))
                str += "("+ this.bibentries["Number"] + "): ";
            if (bibentries.ContainsKey("Pages"))
            {
                string pp = this.bibentries["Pages"].Replace("--", "-");
                str += "pp. " + this.bibentries["Pages"] + ", ";

            }
            if (bibentries.ContainsKey("Publisher"))
                str += this.bibentries["Publisher"] + ", ";
            if (bibentries.ContainsKey("Month"))
            {
                List<String> months = new List<String>(new string[] { 
                "Jan", 
                "Feb", 
                "Mar", 
                "Apr", 
                "May", 
                "Jun",
                "July",
                "Aug",
                "Sept",
                "Oct",
                "Nov",
                "Dec"});

                string month = this.bibentries["Month"];
                int n;
                bool isNumeric = int.TryParse(month, out n);
                if (!isNumeric)
                {
                    str += month + " ";
                }
                else
                {
                    if (n >= 0 && n < 12)
                        str += months[n] + " ";
                }
            }
            if (bibentries.ContainsKey("Year"))
                str += this.bibentries["Year"] + ". ";
            if (bibentries.ContainsKey("URL"))
                str += this.bibentries["URL"] + ". ";

            if (bibentries.ContainsKey("Comments"))
                str += this.bibentries["Comments"] + ". ";
            return str;
        }

    }

    /// <summary>
    /// Keeps track of all bibliography entries in a list collection collection.
    /// </summary>
    public class XmlBibliography
    {
        /// <summary>
        /// Location of bibliography XSD schema
        /// </summary>
        private const string bibliographyNS = @"http://schemas.openxmlformats.org/officeDocument/2006/bibliography";

        /// <summary>
        /// bibligraph xml document
        /// </summary>
        private XmlDocument bibform;


        /// <summary>
        /// Bib entries list.
        /// </summary>
        List<XmlBibEntry> entries = new List<XmlBibEntry>();

        /////////////////////////////////////////////////////////////////////////
        /// <summary>
        ///  constructor. initialize new bib entries list.
        /// </summary>
        public XmlBibliography()
        {
            entries = new List<XmlBibEntry>(); 
        }

        /////////////////////////////////////////////////////////////////////////
        /// <summary>
        /// returns bibentry of matching tag bi bib entries list.
        /// </summary>
        /// <param name="tag"></param>
        /// <returns>bib entry or null if not found</returns>
        public XmlBibEntry Find(string tag)
        {
            XmlBibEntry result = entries.Find(x => x.tag == tag);
            return result;
        }

        /////////////////////////////////////////////////////////////////////////
        /// <summary>
        /// returns index of matching tag bi bib entries list.
        /// </summary>
        /// <param name="tag"></param>
        /// <returns> >=0 index if found or -1 if not found</returns>
        public int Index(string tag)
        {
            int index = entries.FindIndex(x => x.tag == tag);
            if (index >= 0)
                index++;
            return index;
        }

        /////////////////////////////////////////////////////////////////////////
        /// <summary>
        /// produces string for given tag bib entry
        /// </summary>
        /// <param name="tag"> bibliography tag</param>
        /// <returns> string for tag bib entry, empty string if not found</returns>
        public string Entry(string tag)
        {
            XmlBibEntry bib = Find(tag);
            if (bib != null)
                return bib.Entry();
            return "";
        }

        /////////////////////////////////////////////////////////////////////////
        /// <summary>
        /// Generate list of bibliography entries: tag + title + linefeed
        /// </summary>
        /// <returns></returns>
        public string Dump()
        {
            String str = "";
            foreach ( XmlBibEntry entry in entries)
            {
                str += entry.tag + "=";
                if (entry.bibentries.ContainsKey("Title"))
                    str += entry.bibentries["Title"];
                str += Environment.NewLine;

            }
            return str;
        }

        /////////////////////////////////////////////////////////////////////////
        /// <summary>
        /// given an XML string parse the bibliogphy
        /// </summary>
        /// <param name="xml"> xmls string</param>
        public void parseXml(string xml)
        {
            bibform = new XmlDocument();
            bibform.LoadXml(xml);
            parseBibliography();
        }

        /////////////////////////////////////////////////////////////////////////
        /// <summary>
        /// parse a given file that contains xml describing bibliography
        /// </summary>
        /// <param name="bibformIn"></param>
        public void parseXmlFile(string bibformIn)
        {
            XmlDocument bibform = new XmlDocument();
            bibform.Load(bibformIn);
            parseBibliography();
        }

        /// <summary>
        /// using xmldoc parse bibligraphy for bib entries
        /// </summary>
        public void parseBibliography()
        {
            XmlNamespaceManager nsm = new XmlNamespaceManager(bibform.NameTable);
            nsm.AddNamespace("b", bibliographyNS);

            string[] keys = new string[] { "Book", "JournalArticle", "ConferenceProceedings", "Misc" };
            // For every datatag, see if you can find localized content.
            XmlNodeList nl = bibform.SelectNodes("//b:Source", nsm);
            foreach (XmlNode node in nl)
            {
                XmlNode n = node.SelectSingleNode("b:Tag", nsm);
                if (n == null)
                    continue;
                XmlBibEntry bib = new XmlBibEntry();
                bib.tag = n.InnerText;
                n = node.SelectSingleNode("b:SourceType", nsm);
                if (n == null)
                    continue;
                string type = n.InnerText;
                IEnumerable<String> fields = bib.fields[type];

                // break on Misc bib
                //if (type == "Misc")
                //    Debugger.Break();

                foreach (String field in fields)
                {
                    n = node.SelectSingleNode("b:" + field, nsm);



                    if (n != null)
                    {
                        string text = n.InnerText;
                        text = text.Replace("\"", "");
                        bib.bibentries.Add(field, text);
                    }
                }
                XmlBibAuthorList authorlist = new XmlBibAuthorList();
                List<XmlBibAuthor> authors = authorlist.parseAuthors(node, nsm);
                if (authors.Count > 0)
                    bib.authors = authors;
                entries.Add(bib);
            }

        }

    }

}
