using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;

namespace MailTemplateAddIn
{
	public partial class Ribbon
	{
		private void Ribbon_Load(object sender, RibbonUIEventArgs e)
		{

		}

		private void Select_Click(object sender, RibbonControlEventArgs e)
		{
			TemplateSelectWindow w = new TemplateSelectWindow();
			w.ShowDialog();
			w.template.Display();
		}
	}
}
