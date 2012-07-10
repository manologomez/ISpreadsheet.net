using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

namespace ISpreadsheet {
	/// <summary>
	/// Useful extension methods to obtain different datatypes from cell values
	/// </summary>
	public static class WorksheetExtensions {
		private static CultureInfo ci = CultureInfo.GetCultureInfo("en-US");

		public static int? GetInt(this IWorksheet hoja, string col, int row) {
			var txt = hoja.GetCell(col, row) ?? "";
			return ParseInt(txt);
		}

		public static int? GetInt(this IWorksheet hoja, string address) {
			var txt = hoja.GetCell(address) ?? "";
			return ParseInt(txt);
		}

		public static float? GetFloat(this IWorksheet hoja, string col, int row) {
			var txt = hoja.GetString(col, row) ?? "";
			return ParseFloat(txt);
		}

		public static float? GetFloat(this IWorksheet hoja, string address) {
			var txt = hoja.GetString(address) ?? "";
			return ParseFloat(txt);
		}

		public static decimal? GetDecimal(this IWorksheet hoja, string col, int row) {
			var txt = hoja.GetString(col, row) ?? "";
			return ParseDecimal(txt);
		}

		public static decimal? GetDecimal(this IWorksheet hoja, string address) {
			var txt = hoja.GetString(address) ?? "";
			return ParseDecimal(txt);
		}

		public static double? GetDouble(this IWorksheet hoja, string col, int row) {
			var txt = hoja.GetString(col, row) ?? "";
			return ParseDouble(txt);
		}

		public static double? GetDouble(this IWorksheet hoja, string address) {
			var txt = hoja.GetString(address) ?? "";
			return ParseDouble(txt);
		}

		public static int? ParseInt(object value) {
			// TODO: test this further
			if (value == null) return null;
			if (value is int) return (int)value;
			int aux;
			if (int.TryParse(value.ToString(), out aux))
				return aux;
			return null;
		}

		public static float? ParseFloat(string txt) {
			if (string.IsNullOrEmpty(txt)) return null;
			float aux;
			if (float.TryParse(txt.Replace(",", "."), NumberStyles.AllowDecimalPoint, ci, out aux))
				return aux;
			return null;
		}

		public static decimal? ParseDecimal(string txt) {
			if (string.IsNullOrEmpty(txt)) return null;
			decimal aux;
			if (decimal.TryParse(txt.Replace(",", "."), NumberStyles.AllowDecimalPoint, ci, out aux))
				return aux;
			return null;
		}

		public static double? ParseDouble(string txt) {
			if (string.IsNullOrEmpty(txt)) return null;
			double aux;
			if (double.TryParse(txt.Replace(",", "."), NumberStyles.AllowDecimalPoint, ci, out aux))
				return aux;
			return null;
		}

		public static DateTime? GetDateTime(this IWorksheet hoja, string address) {
			var txt = hoja.GetString(address) ?? "";
			return TryGetDateTime(txt);
		}

		public static DateTime? GetDateTime(this IWorksheet hoja, string col, int row) {
			var txt = hoja.GetString(col, row) ?? "";
			return TryGetDateTime(txt);
		}

		public static string[] knownDateFormats = new[]{
			                   	"yyyy/MM/dd",
			                   	"dd/MM/yyyy"
			                   	, "dd/MM/yyyy HH:mm"
			                   	, "dd/MM/yyyy H:mm:ss"
			                   	, "dd/MM/yyyy HH:mm:ss"
			                   	, "yyyy-MM-dd"
			                   	, "yyyy-MM-dd HH:mm"
			                   	, "yyyy-MM-dd H:mm:ss"
			                   	, "yyyy-MM-dd HH:mm:ss"
			                   	, "dd-MM-yyyy"
			                   	, "dd-MM-yyyy HH:mm"
			                   	, "dd-MM-yyyy H:mm:ss"
			                   	, "dd-MM-yyyy HH:mm:ss"
			                   	, "dd.MM.yyyy"
			                   	, "dd.MM.yyyy HH:mm"
			                   	, "dd.MM.yyyy H:mm:ss"
			                   	, "dd.MM.yyyy HH:mm:ss"
			                   };

		/// <summary>
		/// Attempts to obtain a DateTime from a string using some well known formats
		/// TODO: Custom formats and culture
		/// </summary>
		/// <param name="dateString"></param>
		/// <returns></returns>
		public static DateTime? TryGetDateTime(string dateString) {
			if (string.IsNullOrEmpty(dateString))
				return null;
			DateTime dt;
			if (DateTime.TryParseExact(dateString, knownDateFormats, ci, DateTimeStyles.None, out dt))
				return dt;
			return null;
		}
	}
}
