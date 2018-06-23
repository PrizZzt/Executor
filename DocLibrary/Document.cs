using NPOI.OpenXml4Net.OPC;
using NPOI.XWPF.UserModel;
using System.Collections.Generic;
using System.IO;

namespace DocLibrary
{
	public class Document
	{
		private XWPFDocument doc;
		private string FileName;
		private string SavedFileName;

		public bool IsLoaded => doc != null;

		public static Document Create(string fileName)
		{
			var document = new Document();
			document.Load(fileName);
			return document;
		}

		private Document()
		{
		}

		public void Load(string fileName)
		{
			Close();

			FileName = fileName;

			SavedFileName = CreateTempFile();
			File.Copy(FileName, SavedFileName);
			doc = new XWPFDocument(OPCPackage.Open(FileName));
		}

		public void Save(string fileName)
		{
			if (!IsLoaded) return;

			var tempForSave = CreateTempFile();
			using (var fs = new FileStream(tempForSave, FileMode.OpenOrCreate, FileAccess.Write))
				doc.Write(fs);

			doc.Close();
			ReplaceFile(tempForSave, fileName);

			if (FileName == fileName)
			{
				File.Copy(FileName, SavedFileName, true);
			}

			doc = new XWPFDocument(OPCPackage.Open(FileName));
		}

		public void Close()
		{
			if (!IsLoaded) return;

			doc.Close();
			ReplaceFile(SavedFileName, FileName);
			File.Delete(SavedFileName);

			doc = null;
		}

		public void Replace(string key, string value)
		{
			Replace(new Dictionary<string, string> { { key, value } });
		}

		public void Replace(Dictionary<string, string> values)
		{
			if (!IsLoaded) return;

			foreach (var value in values)
			{
				foreach (XWPFParagraph p in doc.Paragraphs)
				{
					if (p.Runs == null) continue;
					CheckParagraph(p, value.Key, value.Value);
				}

				foreach (XWPFTable tbl in doc.Tables)
				{
					foreach (XWPFTableRow row in tbl.Rows)
					{
						foreach (XWPFTableCell cell in row.GetTableCells())
						{
							foreach (XWPFParagraph p in cell.Paragraphs)
							{
								CheckParagraph(p, value.Key, value.Value);
							}
						}
					}
				}
			}
		}

		private void CheckParagraph(XWPFParagraph target, string key, string value)
		{
			for (int i = 0; i < target.Runs.Count; i++)
			{
				if (target.Runs[i].Text == "%%" && i < target.Runs.Count - 2 && target.Runs[i + 2].Text == "%%" && target.Runs[i + 1].Text == key)
				{
					target.Runs[i + 1].SetText(value);
					target.RemoveRun(i + 2);
					target.RemoveRun(i);
				}
			}
		}

		private string CreateTempFile()
		{
			var tempFileName = Path.GetTempFileName();
			if (File.Exists(tempFileName))
				File.Delete(tempFileName);

			return tempFileName;
		}

		private void ReplaceFile(string source, string destination)
		{
			if (File.Exists(destination))
				File.Delete(destination);
			File.Move(source, destination);
		}
	}
}
