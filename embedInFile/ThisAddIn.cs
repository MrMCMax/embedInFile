using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Tools = Microsoft.Office.Tools.Word;
using System.Threading.Tasks;
using System.Windows.Forms;
using Newtonsoft.Json;
using System.IO;
using embedInFile.WordInterop;

namespace embedInFile
{
    public partial class ThisAddIn
    {
        public const string LINKS_PROPERTY = "links";
        public const string DOC_ID_PROPERTY = "drive_doc_ID";
        public const char LIST_SEPARATOR = '#';
        public const char PAIR_SEPARATOR = ',';

        //The connection to drive
        private IDriveConnection driveEmbedding;

        //We have a dictionary with ALL the links currently in the dictionary. 
        //It is only persisted when the document is saved.
        //Receives all events: add and delete.
        Dictionary<int, Dictionary<string, Link>> ccToLinks;
        ISet<int> started;

        private int newDocID;   //Will be used to identify documents during this execution

        /********************/
        /* WORD EVENTS      */
        /********************/
        #region Word Events
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            //Initialize counter for identifying new documents with the persisted value
            string numb = File.ReadAllText(getClickOnceLocation() + "\\" + "IDNumber.txt").Trim();
            newDocID = Int32.Parse(numb);
            //Initialize dictionaries
            ccToLinks = new Dictionary<int, Dictionary<string, Link>>();
            started = new HashSet<int>();
            Globals.Ribbons.Ribbon1.insertButton.Enabled = false;
            Globals.Ribbons.Ribbon1.deleteAllButton.Enabled = false;
            this.Application.DocumentChange += Application_DocumentChange;
            this.Application.DocumentBeforeSave += Application_DocumentBeforeSave;
            this.Application.DocumentBeforeClose += Application_DocumentBeforeClose;
            this.Application.DocumentOpen += Application_DocumentOpen;

            
        }


        private void Application_DocumentOpen(Word.Document Doc)
        {
            //The newly opened document could have a docID or not.
            if (!containsPropertyOnDocument(DOC_ID_PROPERTY, Doc))
            {
                //The doc doesn't have an ID. We create a new one for it.
                int tempdocID = newDocID++;
                File.WriteAllText(getClickOnceLocation() + "\\" + "IDNumber.txt", newDocID + "");
                addStringPropertyOnDocument(DOC_ID_PROPERTY, tempdocID + "", Doc);
            }
            //Here, the document has a docID property. We have to check that it is not duplicated among
            //the already opened documents
            int docID = Int32.Parse(getStringPropertyOnDocument(DOC_ID_PROPERTY, Doc));
            while (started.Contains(docID))
            {
                docID++;
            }
            started.Add(docID);
            //Whether it has changed or not, we update it.
            setStringPropertyOnDocument(DOC_ID_PROPERTY, docID + "", Doc);

            //We create the entry for this document, and its link dictionary at the start.
            ccToLinks.Add(docID, new Dictionary<string, Link>());
            //Add handler
            Doc.ContentControlBeforeDelete += onContentControlBeforeDelete;
            //load links
            loadLinks();
        }

        private void Application_DocumentBeforeClose(Word.Document Doc, ref bool Cancel)
        {
            /*
            int docID = Int32.Parse(getStringPropertyOnDocument(DOC_ID_PROPERTY, Doc));
            if (started.Contains(docID))
            {
                started.Remove(docID);
            }
            */
            MessageBox.Show("Close event");
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            MessageBox.Show("Shutdown event");
        }

        private void Application_DocumentBeforeSave(Word.Document Doc, ref bool SaveAsUI, ref bool Cancel)
        {
            //Due to the primitive operation of properties, we can't just append to the list of links,
            //we have to re-write it entirely. The dictionary that contains the full updated list of links is
            //ccToLinks, so we call a write on it.
            MessageBox.Show($"Save event: ui:{SaveAsUI}, cancel:{Cancel}");
            saveLinks();
        }

        private void Application_DocumentChange()
        {
            Globals.Ribbons.Ribbon1.insertButton.Enabled = true;
            Globals.Ribbons.Ribbon1.deleteAllButton.Enabled = true;
            if (Application.Documents.Count >= 1)
            {
                //This event might appear because of a new document, or because of a change in documents.
                //If it's a new document, it won't have a DOC_ID property, won't trigger open.
                //The document hasn't been opened before
                if (!containsProperty(DOC_ID_PROPERTY))
                {
                    //Set the property, that will only be used during this execution. We have to make sure
                    //that is not repeated in the set of opened documents.
                    int docID = newDocID++;
                    while (started.Contains(docID))
                    {
                        docID++;
                    }
                    started.Add(docID);
                    File.WriteAllText(getClickOnceLocation() + "\\" + "IDNumber.txt", newDocID + "");
                    addStringProperty(DOC_ID_PROPERTY, docID + "");
                    //We create the entry for this document, and its link dictionary at the start.
                    ccToLinks.Add(docID, new Dictionary<string, Link>());
                    //Add handler
                    Application.ActiveDocument.ContentControlBeforeDelete += onContentControlBeforeDelete;
                }
                else
                {
                    //This is either a change document event, or an opened document.
                    //In case it is a change of document, we don't need to do anything.
                    //In case it is an opened document, the event for opening a document will be in charge.
                }
            }
        }

        private void onContentControlBeforeDelete(Word.ContentControl OldContentControl, bool InUndoRedo)
        {
            try
            {
                string Id = OldContentControl.ID;
                Dictionary<string, Link> links = ccToLinks[Int32.Parse(getStringProperty(DOC_ID_PROPERTY))];
                if (links.ContainsKey(Id)) //There's a link to remove
                {
                    MessageBox.Show("Drive content will be deleted", "Content control deleted", MessageBoxButtons.OK);
                    string fileID = links[Id].getId();
                    if (getDriveEmbedding().removeLink(fileID))
                    {
                        links.Remove(Id);
                    }
                } else
                {
                    sayWord("Not found doc id: " + Id);
                }
            } catch  (Exception e)
            {
                sayWord(e.ToString());
            }
        }
        #endregion

        /********************/
        /* WORD  METHODS    */
        /********************/
        #region Word Methods

        #region Document Custom Properties Methods

        /// <summary>
        /// Adds a property to the Application.ActiveDocument
        /// Precondition: the document doesn't have the property
        /// </summary>
        /// <param name="name"></param>
        /// <param name="content"></param>
        public void addStringProperty(string name, string content)
        {
            addStringPropertyOnDocument(name, content, Application.ActiveDocument);
        }

        private void addStringPropertyOnDocument(string name, string content, Word.Document doc)
        {
            doc.CustomDocumentProperties.Add(name, false, Office.MsoDocProperties.msoPropertyTypeString, content);
        }

        /// <summary>
        /// Updates a property of Application.ActiveDocument, or adds it in case it doesn't exist
        /// </summary>
        /// <param name="name"></param>
        /// <param name="content"></param>
        public void setStringProperty(string name, string content)
        {
            setStringPropertyOnDocument(name, content, Application.ActiveDocument);
        }

        private void setStringPropertyOnDocument(string name, string content, Word.Document doc)
        {
            deletePropertyOnDocument(name, doc);
            addStringPropertyOnDocument(name, content, doc);
        }

        /// <summary>
        /// Returns the string value of the given property of Application.ActiveDocument.
        /// Precondition: containsProperty(name)
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        public string getStringProperty(string name)
        {
            return getStringPropertyOnDocument(name, Application.ActiveDocument);
        }

        private string getStringPropertyOnDocument(string name, Word.Document doc)
        {
            return (string)doc.CustomDocumentProperties[name].Value;
        }

        public void deleteProperty(string name)
        {
            deletePropertyOnDocument(name, Application.ActiveDocument);
        }

        private void deletePropertyOnDocument(string name, Word.Document doc)
        {
            if (containsPropertyOnDocument(name, doc))
            {
                doc.CustomDocumentProperties[name].Delete();
            }
        }


        /// <summary>
        /// Checks wether Application.ActiveDocument has a property
        /// </summary>
        /// <param name="name">the name of the property to check</param>
        /// <returns></returns>
        public bool containsProperty(string name)
        {
            return containsPropertyOnDocument(name, Application.ActiveDocument);
        }

        private bool containsPropertyOnDocument(string name, Word.Document doc)
        {
            foreach (Office.DocumentProperty prop in doc.CustomDocumentProperties)
            {
                if (prop.Name == name)
                {
                    return true;
                }
            }
            return false;
        }

        /// <summary>
        /// Writes in the document all the Application.ActiveDocument properties
        /// </summary>
        public void listProperties()
        {
            foreach (Office.DocumentProperty prop in Application.ActiveDocument.CustomDocumentProperties)
            {
                sayWord($"Property: {prop.Name}: {prop.Value}\r\n");
            }
        }

        #endregion

        /// <summary>
        /// Reads the serialized links property and loads it into the dictionary
        /// </summary>
        private void loadLinks()
        {
            if (containsProperty(LINKS_PROPERTY))
            {
                Dictionary<string, Link> temp = ccToLinks[Int32.Parse(getStringProperty(DOC_ID_PROPERTY))];
                string prop = getStringProperty(LINKS_PROPERTY);
                string[] links = prop.Split(LIST_SEPARATOR);
                foreach (string p in links)
                {
                    string[] kv = p.Split(PAIR_SEPARATOR);
                    temp.Add(kv[0], Link.deserialize(kv[1]));
                }
            }
        }

        /// <summary>
        /// Saves the dictionary of links
        /// </summary>
        private void saveLinks()
        {
            Dictionary<string, Link> links = ccToLinks[Int32.Parse(getStringProperty(DOC_ID_PROPERTY))];
            if (links.Count > 0)
            {
                StringBuilder sb = new StringBuilder();
                foreach (KeyValuePair<string, Link> p in links)
                {
                    sb.Append(p.Key + PAIR_SEPARATOR + p.Value.serialize());
                    sb.Append(LIST_SEPARATOR);
                }
                sb.Remove(sb.Length - 1, 1);
                setStringProperty(LINKS_PROPERTY, sb.ToString());
            }
        }

        /// <summary>
        /// Adds a new link to the document registered links, temporarily until it's saved
        /// </summary>
        /// <param name="name">name of the link</param>
        /// <param name="URL">URL of the link</param>
        /// <param name="cc">The content control associated to the link</param>
        public void addNewLink(string name, string URL, Tools.ContentControl cc)
        {
            //A new link has been added. Store it in the ccToLinks dictionary
            Link l = new Link(name, URL, cc.ID);
            try
            {
                ccToLinks[Int32.Parse(getStringProperty(DOC_ID_PROPERTY))].Add(cc.ID, l);
            } catch (Exception e)
            {
                sayWord(e.ToString());
            } finally
            {
                Console.WriteLine("Dolor");
            }
        }
        
        public void listAllLinks()
        {
            foreach (KeyValuePair<string, Link> link in ccToLinks[Int32.Parse(getStringProperty(DOC_ID_PROPERTY))])
            {
                sayWord($"Name: {link.Key}, URL: {link.Value}\r\n");
            }
        }

        public void sayWord(string word)
        {
            //Get current document
            if (Application.Documents.Count >= 1)
            {
                Word.Document doc = Application.ActiveDocument;
                //Get current user selection, if none, the selection is the cursor point
                Word.Selection currentSelection = Application.Selection;
                //Get current user overtype policy. Usually would be false, but could be true.
                //We need it to be false. We store it so that we can restore it at the end.
                bool userOvertype = Application.Options.Overtype;
                if (userOvertype)
                {
                    Application.Options.Overtype = false;
                }
                try
                {
                    //Now we test if the cursor is at insertion mode (no selection):
                    if (currentSelection.Type == Word.WdSelectionType.wdSelectionIP)
                    {
                        //If it is, then insert the text.
                        currentSelection.TypeText(word);
                    }
                    else if (currentSelection.Type == Word.WdSelectionType.wdSelectionNormal)
                    {
                        //The user selected a piece of text. 
                        //Now we need to know if the policy on write over selected text is to replace it.
                        if (Application.Options.ReplaceSelection)
                        {
                            //If it is, "collapse" (delete) the selection (and set the cursor at the beggining)
                            object direction = Word.WdCollapseDirection.wdCollapseStart;
                            currentSelection.Collapse(ref direction);
                        }
                        //Whatever happened with the selection, insert the text
                        currentSelection.TypeText(word);
                    }
                    else { } //Do nothing

                } catch (Exception e)
                {
                    
                }

                //Restore overtype
                Application.Options.Overtype = userOvertype;
            }
        }

        public Word.Document activeDoc()
        {
            if (this.Application.Documents.Count >= 1)
            {
                return this.Application.ActiveDocument;
            } else
            {
                return null;
            }
        }

        private Tools.ContentControl AddTextControlAtSelection(string name)
        {
            Tools.Document vstoDoc = Globals.Factory.GetVstoObject(activeDoc());
            Tools.ContentControl cc = vstoDoc.Controls.AddContentControl(name,
                Word.WdContentControlType.wdContentControlRichText);
            return cc;
        }

        public Tools.ContentControl createDriveContentControl()
        {
            string newName = "cc_" + getNextNumber();
            return AddTextControlAtSelection(newName);
        }

        /// <summary>
        /// Returns a unique number for this document.
        /// </summary>
        /// <returns></returns>
        public int getNextNumber()
        {
            var vars = activeDoc().Variables;
            Word.Variable v = null;
            int number = 0;
            foreach (Word.Variable prop in vars)
            {
                if (prop.Name == "number")
                {
                    v = prop;
                    number = Int32.Parse(prop.Value);
                }
            }
            if (number == 0)
            {
                vars.Add("number", 1);
                return 0;
            } else
            {
                v.Value = "" + (number + 1);
                return number;
            }
        }

        public void addObjectVariable()
        {
            string dict = "72";
            object reference = dict;
            activeDoc().Variables.Add("Dict", ref reference);
        }

        public void listVariables()
        {
            foreach (Word.Variable v in activeDoc().Variables)
            {
                sayWord($"Variable: name={v.Name}, object={v.Value.ToString()}\r\n");
            }
        }

        public string getClickOnceLocation()
        {
            //Get the assembly information
            System.Reflection.Assembly assemblyInfo = System.Reflection.Assembly.GetExecutingAssembly();

            //Location is where the assembly is run from 
            string assemblyLocation = assemblyInfo.Location;

            //CodeBase is the location of the ClickOnce deployment files
            Uri uriCodeBase = new Uri(assemblyInfo.CodeBase);
            string ClickOnceLocation = Path.GetDirectoryName(uriCodeBase.LocalPath.ToString());
            return ClickOnceLocation;
        }

        #endregion

        /********************/
        /* DRIVE METHODS    */
        /********************/
        #region Drive Methods
        public IDriveConnection getDriveEmbedding()
        {
            if (driveEmbedding == null)
            {
                driveEmbedding = new DriveEmbedding(getClickOnceLocation());
            }
            return driveEmbedding;
        }

        public void listFiles()
        {
            getDriveEmbedding().listFiles();
        }

        public void deleteSelected()
        {
            getDriveEmbedding().removeLink(Application.Selection.FormattedText.Text);
        }

        public void uploadLink(string[] paths)
        {
            foreach (string path in paths)
            {
                string name = Path.GetFileName(path);
                Tools.ContentControl cc = Globals.ThisAddIn.createDriveContentControl();

                Task.Factory.StartNew(async () =>
                {
                    string result = await getDriveEmbedding().uploadLink(cc.Range, path, name);
                    cc.LockContents = true;
                    //If the creation was successful, store the link
                    if (!result.Equals(""))
                    {
                        addNewLink(name, result, cc);
                    }
                });
            }
        }

        #endregion

        #region Código generado por VSTO

        /// <summary>
        /// Método necesario para admitir el Diseñador. No se puede modificar
        /// el contenido de este método con el editor de código.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
