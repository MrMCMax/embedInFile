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
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }
        /// <summary>
        /// Full functionality described in the API
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void insertButton_Click(object sender, RibbonControlEventArgs e)
        {
            if (Globals.ThisAddIn.activeDoc() == null)
            {
                return;
            }
            string[] paths = selectFiles(); //OpenFileDialog returns the files selected by the user
            if (paths == null) return;
            Globals.ThisAddIn.uploadLink(paths);
        }

        private void deleteAll_Click(object sender, RibbonControlEventArgs e)
        {
            //Test write and retrieve
        }

        /// <summary>
        /// Opens an OpenFileDialog and returns an array of the absolute paths 
        /// to the files that the user has selected. If the user didn't select any file,
        /// it returns null.
        /// </summary>
        /// <returns></returns>
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

        private void listAllLinks_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.listAllLinks();
        }

        private void listFiles_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.listFiles();
        }

        private void deleteSelected_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.deleteSelected();
        }

        private void listAssembly_Click(object sender, RibbonControlEventArgs e)
        {
            //Get the assembly information
            System.Reflection.Assembly assemblyInfo = System.Reflection.Assembly.GetExecutingAssembly();

            //Location is where the assembly is run from 
            string assemblyLocation = assemblyInfo.Location;

            //CodeBase is the location of the ClickOnce deployment files
            Uri uriCodeBase = new Uri(assemblyInfo.CodeBase);
            string ClickOnceLocation = Path.GetDirectoryName(uriCodeBase.LocalPath.ToString());
            Globals.ThisAddIn.sayWord(ClickOnceLocation+"\r\n");
        }

        private void AddProperty_Click(object sender, RibbonControlEventArgs e)
        {
            DialogResult res = MessageBox.Show("Property?", "Property", MessageBoxButtons.YesNo);
            Globals.ThisAddIn.addStringProperty("Test", (res == DialogResult.Yes ? "Yes" : "No"));
        }

        private void ListProperties_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.listProperties();
        }

        private void AddCC_Click(object sender, RibbonControlEventArgs e)
        {
            Tools.ContentControl cc = Globals.ThisAddIn.createDriveContentControl();
            Globals.ThisAddIn.sayWord("ID: " + cc.ID);
        }

        private void Button1_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.addObjectVariable();
        }

        private void ListVariables_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.listVariables();
        }
    }
}
