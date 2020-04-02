using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using System.IO;
using Word = Microsoft.Office.Interop.Word;
using Tools = Microsoft.Office.Tools.Word;

namespace embedInFile
{
    public partial class Ribbon1
    {
        private IDriveConnection driveEmbedding;
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void insertButton_Click(object sender, RibbonControlEventArgs e)
        {
            Word.Document currentDoc;
            if ((currentDoc = Globals.ThisAddIn.getCurrentDocument()) == null)
            {
                return;
            }
            string docName = currentDoc.Name;
            string[] paths = selectFiles();
            if (paths == null) return;
            if (driveEmbedding == null)
            {
                driveEmbedding = new DriveEmbedding();
            }
            foreach (string path in paths)
            {
                Tools.ContentControl cc = Globals.ThisAddIn.createDriveContentControl();
                System.Threading.Tasks.Task.Factory.StartNew(async () =>
                {
                    await driveEmbedding.uploadLink(cc.Range, path, docName);
                    cc.LockContents = true;
                });
            }
            
        }

        private void deleteAll_Click(object sender, RibbonControlEventArgs e)
        {
            //Test write and retrieve
        }

        private string[] selectFiles()
        {
            OpenFileDialog d = new OpenFileDialog();
            d.Title = "Seleccionar archivos";
            d.Multiselect = true;
            d.CheckFileExists = true;
            d.CheckPathExists = true;
            d.Filter = "Audio Files (*.mp3)|*.mp3|All Files (*.*)|*.*";
            string defaultDir = Properties.Settings.Default.defaultDirectory;
            if (defaultDir == null || defaultDir == "")
            {
                defaultDir = d.InitialDirectory;
            }
            d.InitialDirectory = defaultDir;
            if (d.ShowDialog() == DialogResult.OK)
            {
                Properties.Settings.Default.defaultDirectory = Path.GetDirectoryName(d.FileNames[0]);
                return d.FileNames;
            } else
            {
                return null;
            }
        }

        private void debug_list_Click(object sender, RibbonControlEventArgs e)
        {
            if (driveEmbedding == null)
            {
                driveEmbedding = new DriveEmbedding();
            }
            driveEmbedding.listFiles();
        }

        private void deleteSelectedButton_Click(object sender, RibbonControlEventArgs e)
        {
            if (driveEmbedding == null)
            {
                driveEmbedding = new DriveEmbedding();
            }
            driveEmbedding.removeLink(Globals.ThisAddIn.Application.Selection.FormattedText.Text);
        }

        public IDriveConnection getDriveEmbedding()
        {
            if (driveEmbedding == null)
            {
                driveEmbedding = new DriveEmbedding();
            }
            return driveEmbedding;
        }
    }
}
