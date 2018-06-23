using DocLibrary;
using System;
using System.Windows.Forms;

namespace Executor
{
	public partial class Form1 : Form
	{
		public Form1()
		{
			InitializeComponent();
		}

		private void button1_Click(object sender, EventArgs e)
		{
			using (var dlg = new OpenFileDialog())
			{
				if (dlg.ShowDialog() == DialogResult.OK)
				{
					Document doc = Document.Create(dlg.FileName);
					doc.Replace("key", "value");
					doc.Save(dlg.FileName);
					doc.Save(dlg.FileName + ".docx");
					doc.Close();
				}
			}
		}
	}
}
