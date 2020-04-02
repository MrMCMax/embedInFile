using Microsoft.Office.Tools.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;

namespace embedInFile
{
    public interface IDriveConnection
    {
        /// <summary>
        /// Will do all necessary tasks to connect to drive, create folders if necessary and upload the file.
        /// </summary>
        /// <param name="path">path to the file to upload</param>
        /// <param name="documentName">name of the document to insert into</param>
        /// <returns>The public URL linked to the file</returns>
        Task uploadLink(Word.Range cc, string path, string documentName);

        void removeAll();

        void removeLink(string path);

        void listFiles();
    }
}
