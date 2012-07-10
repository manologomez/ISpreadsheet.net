using System;
using System.IO;

namespace ISpreadsheet {
	/// <summary>
	/// Abstract workbook
	/// </summary>
	public interface IWorkbook : IDisposable {
		/// <summary>
		/// Get sheet by name
		/// </summary>
		/// <param name="name"></param>
		/// <returns></returns>
		IWorksheet GetSheet(string name);
		/// <summary>
		/// Get sheet by number (starting in 1)
		/// </summary>
		/// <param name="num"></param>
		/// <returns></returns>
		IWorksheet GetSheet(int num);
		/// <summary>
		/// Available sheets
		/// </summary>
		IWorksheet[] Sheets { get; }
		/// <summary>
		/// Creates a new sheet with the given name
		/// </summary>
		/// <param name="name"></param>
		/// <returns></returns>
		IWorksheet CreateSheet(string name);
		/// <summary>
		/// Saves the contents of the workbook to a file with an optional password
		/// </summary>
		/// <param name="file">Path of the file to save</param>
		/// <param name="password"></param>
		void SaveToFile(string file, string password = "");
		/// <summary>
		/// Saves the contents of the workbook to a stream with an optional password
		/// </summary>
		/// <param name="stream"></param>
		/// <param name="password"></param>
		void SaveToStream(Stream stream, string password = "");
	}
}