using Interop = Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace embedInFile.WordInterop
{
    public class Link
    {
        public const char FIELD_SEPARATOR = ';';
        private string name { get; set; }
        private string URL { get; set; }
        private string Id { get; set; }
        private string contentControlID { get; set; }

        public Link(string name, string URL, string contentControlID)
        {
            this.name = name;
            this.URL = URL;
            this.Id = URL.Substring(31);
            this.contentControlID = contentControlID;
        }

        public string getName() { return name; }
        public string getURL() { return URL; }
        public string getId() { return Id; }
        public string getCCId() { return contentControlID; }

        public bool Equals(Link other)
        {
            return this.name == other.name;
        }
        /// <summary>
        /// Returns the object in string form
        /// </summary>
        /// <returns>name + FIELD_SEPARATOR + URL + FIELD_SEPARATOR + contentControlID</returns>
        public string serialize()
        {
            StringBuilder sb = new StringBuilder();
            sb.Append(name);
            sb.Append(FIELD_SEPARATOR);
            sb.Append(URL);
            sb.Append(FIELD_SEPARATOR);
            sb.Append(contentControlID);
            return sb.ToString();
        }
        /// <summary>
        /// Returns a Link object from its serialized version
        /// </summary>
        /// <param name="serial"></param>
        /// <returns></returns>
        public static Link deserialize(string serial)
        {
            string[] fields = serial.Split(FIELD_SEPARATOR);
            return new Link(fields[0], fields[1], fields[2]);
        }
    }
}
