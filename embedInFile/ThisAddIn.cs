using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace embedInFile
{
    public partial class ThisAddIn
    {
        
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            Globals.Ribbons.Ribbon1.insertButton.Enabled = false;
            Globals.Ribbons.Ribbon1.deleteAllButton.Enabled = false;
            this.Application.DocumentChange += Application_DocumentChange;
        }

        private void Application_DocumentChange()
        {
            Globals.Ribbons.Ribbon1.insertButton.Enabled = true;
            Globals.Ribbons.Ribbon1.deleteAllButton.Enabled = true;
            if (Application.Documents.Count >= 1)
            {
                this.Application.ActiveDocument.ContentControlBeforeDelete += onContentControlBeforeDelete;
                this.Application.ActiveDocument.ContentControlBeforeContentUpdate += onContentControlBeforeContentUpdate;
            }
        }

        private void onContentControlBeforeContentUpdate(Word.ContentControl ContentControl, ref string Content)
        {
            Task t = new Task(() => {
                var dialogResult = MessageBox.Show("Content is changing", "Content control text changing", MessageBoxButtons.OK);
            });
            t.Start();
        }

        private void onContentControlBeforeDelete(Word.ContentControl OldContentControl, bool InUndoRedo)
        {
            if (OldContentControl.Range.Hyperlinks.Count > 0) //There's a link to remove
            {
                MessageBox.Show("Drive content will be deleted", "Content control deleted", MessageBoxButtons.OK);
                string url = OldContentControl.Range.Hyperlinks[1].Address.Substring(31);
                Globals.Ribbons.Ribbon1.getDriveEmbedding().removeLink(url);
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
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

        public Word.Document getCurrentDocument()
        {
            if (this.Application.Documents.Count >= 1)
            {
                return this.Application.ActiveDocument;
            } else
            {
                return null;
            }
        }

        private ContentControl AddTextControlAtSelection(string name)
        {
            Document vstoDoc = Globals.Factory.GetVstoObject(getCurrentDocument());
            ContentControl cc = vstoDoc.Controls.AddContentControl(name,
                Word.WdContentControlType.wdContentControlRichText);
            return cc;
        }

        public ContentControl createDriveContentControl()
        {
            string newName = "cc_" + getNextNumber();
            return AddTextControlAtSelection(newName);
        }

        public int getNextNumber()
        {
            var vars = getCurrentDocument().Variables;
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
