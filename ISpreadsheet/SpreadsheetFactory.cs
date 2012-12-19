using System;
using System.IO;
using ISpreadsheet.Implementation.EPPlus;
using ISpreadsheet.Implementation.NPOI;

namespace ISpreadsheet {
	/// <summary>
	/// Workbook factory that returns an IWorkbook from either a file name or a stream
	/// </summary>
	public class SpreadsheetFactory {
		/// <summary>
		/// Returns an IWorkbook from a physical file path
		/// </summary>
		/// <param name="filepath"></param>
		/// <returns></returns>
		public static IWorkbook GetWorkbook(string filepath) {
			string ext = Path.GetExtension(filepath.ToLower());
			switch (ext) {
				case ".xls":
					return NpoiWorkbook.OpenFromFile(filepath);
				case ".xlsx":
					return EPPlusWorkbook.OpenFromFile(filepath);
			}
			throw new ApplicationException("Extension " + ext + " not recognized");
		}

		/// <summary>
		/// Returns an IWorkbook from a stream, providing the logical extension for getting the right implementation
		/// </summary>
		/// <param name="stream"></param>
		/// <param name="extension"></param>
		/// <returns></returns>
		public static IWorkbook GetWorkbook(Stream stream, string extension) {
			if(string.IsNullOrEmpty(extension))
				throw new ApplicationException("To get a workbook from a stream, the extension parameter cannot be empty");
			string ext = extension ?? "xls";
			switch (ext.ToLower().Replace(".", "")) {
				case "xls":
					return NpoiWorkbook.OpenFromStream(stream);
				case "xlsx":
					return EPPlusWorkbook.OpenFromStream(stream);
			}
			throw new ApplicationException("Extension " + ext + " not recognized");
		}

	}
}