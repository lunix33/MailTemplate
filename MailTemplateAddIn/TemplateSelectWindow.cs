using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace MailTemplateAddIn
{
	public partial class TemplateSelectWindow : Form
	{
		const int padLR = 5;
		const int padTB = 15;
		const int spacing = 25;

		private string templatePath;
		public Outlook.MailItem template;
		private Dictionary<string, TextBox> inputs = new Dictionary<string, TextBox>();

		/// <summary>
		/// Get the file names of the templates in the template directory.
		/// </summary>
		private List<string> TemplateFiles
		{
			get
			{
				// Get all the file names from the template directory.
				string[] files = Directory.GetFiles(templatePath, "*.oft");
				IEnumerable<string> fileNames = from f in files
												select Path.GetFileName(f);

				return fileNames.ToList();
			}
		}

		public TemplateSelectWindow()
		{
			InitializeComponent();

			// Get the path to the template folder.
			string appdata = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
			this.templatePath = Path.Combine(appdata, "Microsoft", "Templates");

			// Set the combobox content.
			this.TemplateComboBox.DataSource = this.TemplateFiles;
			this.TemplateComboBox_SelectionChangeCommitted(this.TemplateComboBox, null);
		}

		/// <summary>
		/// Action executed once a template is selected in the combobox.
		/// </summary>
		/// <param name="sender">The sender object.</param>
		/// <param name="e">The fired event args.</param>
		private void TemplateComboBox_SelectionChangeCommitted(object sender, EventArgs e)
		{
			string file = (string)this.TemplateComboBox.SelectedValue;
			string path = Path.Combine(this.templatePath, file);

			this.template = Globals.ThisAddIn.Application.CreateItemFromTemplate(path);

			List<string> msgParams = this.GetParams(this.template.Subject);
			msgParams.AddRange(this.GetParams(this.template.HTMLBody));
			msgParams = msgParams.Distinct().ToList();
			this.CreateInputs(msgParams);
		}

		/// <summary>
		/// Action fired when the apply button is clicked.
		/// </summary>
		/// <param name="sender">The sender object.</param>
		/// <param name="e">The fired event args.</param>
		private void ApplyBtn_Click(object sender, EventArgs e)
		{
			foreach (string k in this.inputs.Keys)
			{
				string val = (this.inputs[k].Text != "") ?
					this.inputs[k].Text : String.Format("{{{0}}}", k);

				string regstr = String.Format(@"\{{(?:&nbsp;)?:{0}\}}", k);
				Regex reg = new Regex(regstr, RegexOptions.Compiled);

				this.template.Subject = reg.Replace(this.template.Subject, val);
				this.template.HTMLBody = reg.Replace(this.template.HTMLBody, val);
			}

			this.Close();
		}

		/// <summary>
		/// Get the list of parameters in a string.
		/// Parameter format: {:name}
		/// </summary>
		/// <param name="input">The input string.</param>
		/// <returns>The list of parameters.</returns>
		private List<string> GetParams(string input)
		{
			List<string> rtn = new List<string>();
			Regex exp = new Regex(@"\{(?:&nbsp;)?:([^{}]+)\}", RegexOptions.Compiled);
			MatchCollection matches = exp.Matches(input);
			foreach (Match m in matches)
			{
				GroupCollection gc = m.Groups;
				rtn.Add((string)gc[1].Value);
			}

			return rtn;
		}

		/// <summary>
		/// Create the input fields for the parameters.
		/// </summary>
		/// <param name="msgParams">The list of parameters.</param>
		private void CreateInputs(List<string> msgParams)
		{
			this.inputs.Clear();
			this.VariablesGroupBox.Controls.Clear();

			for (int i = 0; i < msgParams.Count; i++)
			{
				string p = msgParams[i];

				TextBox tb = new TextBox
				{
					Name = p,
					Text = String.Format("{{{0}}}", p),
					Dock = DockStyle.Top
				};

				Panel pan = new Panel
				{
					Anchor = (AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right),
					Dock = DockStyle.Top,
					Padding = new Padding(0, 0, 0, 5),
					AutoSize = true,
					AutoSizeMode = AutoSizeMode.GrowOnly
				};

				pan.Controls.Add(tb);
				this.VariablesGroupBox.Controls.Add(pan);
				this.inputs.Add(p, tb);
			}
		}
	}
}
