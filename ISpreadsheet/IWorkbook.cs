using System;

namespace ISpreadsheet{
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
	}
}