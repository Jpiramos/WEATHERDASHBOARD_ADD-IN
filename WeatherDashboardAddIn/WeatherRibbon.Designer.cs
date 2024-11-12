namespace WeatherDashboardAddIn
{
    partial class WeatherRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Variável de designer necessária.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public WeatherRibbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Limpar os recursos que estão sendo usados.
        /// </summary>
        /// <param name="disposing">true se for necessário descartar os recursos gerenciados; caso contrário, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Código gerado pelo Designer de Componentes

        /// <summary>
        /// Método necessário para suporte ao Designer - não modifique 
        /// o conteúdo deste método com o editor de código.
        /// </summary>
        private void InitializeComponent()
        {
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.btnBuscarClima = this.Factory.CreateRibbonButton();
            this.txtCidade = this.Factory.CreateRibbonEditBox();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.group2.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.group2);
            this.tab1.Label = "Clima e Umidade";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.txtCidade);
            this.group1.Label = "Seleção";
            this.group1.Name = "group1";
            // 
            // group2
            // 
            this.group2.Items.Add(this.btnBuscarClima);
            this.group2.Name = "group2";
            // 
            // btnBuscarClima
            // 
            this.btnBuscarClima.Label = "Buscar Clima";
            this.btnBuscarClima.Name = "btnBuscarClima";
            this.btnBuscarClima.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnBuscarClima_Click);
            // 
            // txtCidade
            // 
            this.txtCidade.Label = "Cidade Desejada: ";
            this.txtCidade.Name = "txtCidade";
            this.txtCidade.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.editBox1_TextChanged);
            // 
            // WeatherRibbon
            // 
            this.Name = "WeatherRibbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.WeatherRibbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnBuscarClima;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox txtCidade;
    }



    partial class ThisRibbonCollection
    {
        internal WeatherRibbon WeatherRibbon
        {
            get { return this.GetRibbon<WeatherRibbon>(); }
        }
    }
}
