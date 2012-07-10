using System.Collections.Generic;

namespace ISpreadsheet {
	/// <summary>
	/// Abstract worksheet
	/// </summary>
	public interface IWorksheet {
		/// <summary>
		/// Name of the sheet
		/// </summary>
		string Name { get; set; }
		/// <summary>
		/// Get value of a cell using range address notation: A:1, B:5, etc.
		/// </summary>
		/// <param name="address"></param>
		/// <returns></returns>
		object GetCell(string address);
		/// <summary>
		/// Value of a cell using column and row number, 1 based indexed
		/// </summary>
		/// <param name="col"></param>
		/// <param name="row"></param>
		/// <returns></returns>
		object GetCell(int col, int row);
		/// <summary>
		/// Value of a cell using column and row address (A,B...) , 1 based indexed
		/// </summary>
		/// <param name="col"></param>
		/// <param name="row"></param>
		/// <returns></returns>
		object GetCell(string col, int row);
		/// <summary>
		/// Object array from a row
		/// </summary>
		/// <param name="num">Row number (1 based index)</param>
		/// <returns></returns>
		object[] GetRow(int num);
		/// <summary>
		/// Row expressed as a dictionary in which keys are row letters
		/// </summary>
		/// <param name="num">Row number (1 based index)</param>
		/// <returns></returns>
		IDictionary<string, object> GetRowMap(int num);
		/// <summary>
		/// Total number of columns in the sheet
		/// </summary>
		int NumColumns { get; }
		/// <summary>
		/// Total number of rows in the sheet
		/// </summary>
		int NumRows { get; }

		/// <summary>
		/// Get the string value of a cell using range address notation
		/// </summary>
		/// <param name="address"></param>
		/// <returns></returns>
		string GetString(string address);

		/// <summary>
		/// Get the string value of a cell using either column number or letter and a row number
		/// </summary>
		/// <param name="col"></param>
		/// <param name="row"></param>
		/// <returns></returns>
		string GetString(dynamic col, int row);

		/// <summary>
		/// Sets the value of a cell by address
		/// </summary>
		/// <param name="address"></param>
		/// <param name="value"></param>
		/// <returns></returns>
		IWorksheet SetValue(string address, object value);

		/// <summary>
		/// Sets the value of a cell by indexes
		/// </summary>
		/// <param name="col"></param>
		/// <param name="row"></param>
		/// <param name="value"></param>
		/// <returns></returns>
		IWorksheet SetValue(int col, int row, object value);
	}
}