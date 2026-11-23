using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Serialization;
using Word = Microsoft.Office.Interop.Word;

namespace XtendedRefs.ComAddIn.Domain.Models.NewModels
{
    [XmlRoot("ChapterList", Namespace = "XtendedRefs.ComAddIn.Chapters")]
    public class ChaptersList 
    {
        public ChaptersList() { }

        [XmlElement("Chapter", Namespace = "XtendedRefs.ComAddIn.Chapters")]
        public List<ChapterInfo> Chapters { get; set; } = new List<ChapterInfo>();
        


    }

    [XmlRoot("Chapter", Namespace = "XtendedRefs.ComAddIn.Chapters")]
    public class ChapterInfo : DocumentAnchor
    {
        public ChapterInfo() { }

        public ChapterInfo(string title, string styleName, int level)
        {
            Title = title;
            StyleName = styleName;
            Level = level;
        }

        [XmlElement("Title")]
        public string Title { get; set; } = string.Empty;

        [XmlElement("StyleName")]
        public string StyleName { get; set; } = string.Empty;

        [XmlElement("Level")]
        public int Level { get; set; }

        [XmlArray("Part")]
        public List<string> NumberParts { get; set; } = new();

        [XmlIgnore]
        public string Number => string.Join(".", NumberParts);
        
        [XmlIgnore]
        public override string DisplayText { get; set; }

        public override string GetNumber(Word.Range range)
        {
            string inputString = range.ListFormat.ListString;
            Match match = Regex.Match(inputString, @"\d+(?:\.\d+)*");

            if (!match.Success)
                return string.Empty;
            List<string> listNum = match.Value.Split('.').ToList();
            return string.Join(".", listNum);
        }

        public override string ToString()
        {
            string result = $"Number: {Number}\nTitle {Title}\nStyle name {StyleName}\n Level {Level}";
            return result;
        }
    }
    
}