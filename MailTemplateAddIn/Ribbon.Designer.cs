namespace MailTemplateAddIn
{
	partial class Ribbon : Microsoft.Office.Tools.Ribbon.RibbonBase
	{
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.IContainer components = null;

		public Ribbon()
			: base(Globals.Factory.GetRibbonFactory())
		{
			InitializeComponent();
		}

		/// <summary> 
		/// Clean up any resources being used.
		/// </summary>
		/// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
		protected override void Dispose(bool disposing)
		{
			if (disposing && (components != null))
			{
				components.Dispose();
			}
			base.Dispose(disposing);
		}

		#region Component Designer generated code

		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			this.MailTab = this.Factory.CreateRibbonTab();
			this.MailTemplateGroup = this.Factory.CreateRibbonGroup();
			this.Select = this.Factory.CreateRibbonButton();
			this.MailTab.SuspendLayout();
			this.MailTemplateGroup.SuspendLayout();
			this.SuspendLayout();
			// 
			// MailTab
			// 
			this.MailTab.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
			this.MailTab.ControlId.OfficeId = "TabMail";
			this.MailTab.Groups.Add(this.MailTemplateGroup);
			this.MailTab.Label = "TabMail";
			this.MailTab.Name = "MailTab";
			// 
			// MailTemplateGroup
			// 
			this.MailTemplateGroup.Items.Add(this.Select);
			this.MailTemplateGroup.Label = "Template";
			this.MailTemplateGroup.Name = "MailTemplateGroup";
			// 
			// Select
			// 
			this.Select.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
			this.Select.Description = "Select a template.";
			this.Select.Label = "Select";
			this.Select.Name = "Select";
			this.Select.ShowImage = true;
			this.Select.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Select_Click);
			// 
			// Ribbon
			// 
			this.Name = "Ribbon";
			this.RibbonType = "Microsoft.Outlook.Explorer";
			this.Tabs.Add(this.MailTab);
			this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon_Load);
			this.MailTab.ResumeLayout(false);
			this.MailTab.PerformLayout();
			this.MailTemplateGroup.ResumeLayout(false);
			this.MailTemplateGroup.PerformLayout();
			this.ResumeLayout(false);

		}

		#endregion

		internal Microsoft.Office.Tools.Ribbon.RibbonTab MailTab;
		internal Microsoft.Office.Tools.Ribbon.RibbonGroup MailTemplateGroup;
		internal Microsoft.Office.Tools.Ribbon.RibbonButton Select;
	}

	partial class ThisRibbonCollection
	{
		internal Ribbon Ribbon
		{
			get { return this.GetRibbon<Ribbon>(); }
		}
	}
}
