namespace MailTemplateAddIn
{
	partial class TemplateSelectWindow
	{
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.IContainer components = null;

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

		#region Windows Form Designer generated code

		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			this.TemplateComboBox = new System.Windows.Forms.ComboBox();
			this.VariablesGroupBox = new System.Windows.Forms.GroupBox();
			this.ApplyBtn = new System.Windows.Forms.Button();
			this.SuspendLayout();
			// 
			// TemplateComboBox
			// 
			this.TemplateComboBox.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.TemplateComboBox.FormattingEnabled = true;
			this.TemplateComboBox.Location = new System.Drawing.Point(5, 5);
			this.TemplateComboBox.Name = "TemplateComboBox";
			this.TemplateComboBox.Size = new System.Drawing.Size(710, 21);
			this.TemplateComboBox.TabIndex = 0;
			this.TemplateComboBox.SelectionChangeCommitted += new System.EventHandler(this.TemplateComboBox_SelectionChangeCommitted);
			// 
			// VariablesGroupBox
			// 
			this.VariablesGroupBox.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.VariablesGroupBox.AutoSize = true;
			this.VariablesGroupBox.Location = new System.Drawing.Point(5, 30);
			this.VariablesGroupBox.Name = "VariablesGroupBox";
			this.VariablesGroupBox.Padding = new System.Windows.Forms.Padding(5);
			this.VariablesGroupBox.Size = new System.Drawing.Size(790, 415);
			this.VariablesGroupBox.TabIndex = 1;
			this.VariablesGroupBox.TabStop = false;
			this.VariablesGroupBox.Text = "Variables";
			// 
			// ApplyBtn
			// 
			this.ApplyBtn.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.ApplyBtn.Location = new System.Drawing.Point(720, 4);
			this.ApplyBtn.Name = "ApplyBtn";
			this.ApplyBtn.Size = new System.Drawing.Size(75, 23);
			this.ApplyBtn.TabIndex = 2;
			this.ApplyBtn.Text = "Apply";
			this.ApplyBtn.UseVisualStyleBackColor = true;
			this.ApplyBtn.Click += new System.EventHandler(this.ApplyBtn_Click);
			// 
			// TemplateSelectWindow
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.AutoScroll = true;
			this.ClientSize = new System.Drawing.Size(800, 450);
			this.Controls.Add(this.ApplyBtn);
			this.Controls.Add(this.VariablesGroupBox);
			this.Controls.Add(this.TemplateComboBox);
			this.Name = "TemplateSelectWindow";
			this.Text = "Template Select";
			this.ResumeLayout(false);
			this.PerformLayout();

		}

		#endregion

		private System.Windows.Forms.ComboBox TemplateComboBox;
		private System.Windows.Forms.GroupBox VariablesGroupBox;
		private System.Windows.Forms.Button ApplyBtn;
	}
}