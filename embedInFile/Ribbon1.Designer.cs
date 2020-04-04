namespace embedInFile
{
    partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Variable del diseñador necesaria.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon1()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Limpiar los recursos que se estén usando.
        /// </summary>
        /// <param name="disposing">true si los recursos administrados se deben desechar; false en caso contrario.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Código generado por el Diseñador de componentes

        /// <summary>
        /// Método necesario para admitir el Diseñador. No se puede modificar
        /// el contenido de este método con el editor de código.
        /// </summary>
        private void InitializeComponent()
        {
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.insertButton = this.Factory.CreateRibbonButton();
            this.deleteAllButton = this.Factory.CreateRibbonButton();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.listAllLinks = this.Factory.CreateRibbonButton();
            this.group3 = this.Factory.CreateRibbonGroup();
            this.listFiles = this.Factory.CreateRibbonButton();
            this.deleteSelected = this.Factory.CreateRibbonButton();
            this.listAssembly = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.group2.SuspendLayout();
            this.group3.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.group2);
            this.tab1.Groups.Add(this.group3);
            this.tab1.Label = "Audio+";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.insertButton);
            this.group1.Items.Add(this.deleteAllButton);
            this.group1.Label = "Este documento";
            this.group1.Name = "group1";
            // 
            // insertButton
            // 
            this.insertButton.Label = "Insertar objeto...";
            this.insertButton.Name = "insertButton";
            this.insertButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.insertButton_Click);
            // 
            // deleteAllButton
            // 
            this.deleteAllButton.Label = "Eliminar links";
            this.deleteAllButton.Name = "deleteAllButton";
            this.deleteAllButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.deleteAll_Click);
            // 
            // group2
            // 
            this.group2.Items.Add(this.listAllLinks);
            this.group2.Label = "Todos los documentos";
            this.group2.Name = "group2";
            // 
            // listAllLinks
            // 
            this.listAllLinks.Label = "Ver todos los links...";
            this.listAllLinks.Name = "listAllLinks";
            this.listAllLinks.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.listAllLinks_Click);
            // 
            // group3
            // 
            this.group3.Items.Add(this.listFiles);
            this.group3.Items.Add(this.deleteSelected);
            this.group3.Items.Add(this.listAssembly);
            this.group3.Label = "Debug";
            this.group3.Name = "group3";
            // 
            // listFiles
            // 
            this.listFiles.Label = "listFiles";
            this.listFiles.Name = "listFiles";
            this.listFiles.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.listFiles_Click);
            // 
            // deleteSelected
            // 
            this.deleteSelected.Label = "deleteSelected";
            this.deleteSelected.Name = "deleteSelected";
            this.deleteSelected.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.deleteSelected_Click);
            // 
            // listAssembly
            // 
            this.listAssembly.Label = "listAssembly";
            this.listAssembly.Name = "listAssembly";
            this.listAssembly.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.listAssembly_Click);
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.group3.ResumeLayout(false);
            this.group3.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton insertButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton deleteAllButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton listAllLinks;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton listFiles;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton deleteSelected;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton listAssembly;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
