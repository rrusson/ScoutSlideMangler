using System;
using System.Data;
using System.Data.OleDb;
using System.IO;
using ExcelDataReader;

namespace ScoutSlideMangler
{
	/// <summary>
	/// Extracts a Data from Excel
	/// </summary>
	public class ExcelReader
	{
		/// <summary>
		/// Extracts a DataTable from Excel
		/// </summary>
		/// <param name="fileName">File to parse</param>
		/// <param name="sheetName">Name of the sheet to return as a DataTable</param>
		/// <returns>Excel sheet converted to a DataTable (1st row headers become field names)</returns>
		public DataTable GetData(string filePath, string sheetName)
		{
			using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
			{
				// Auto-detect format, supports:
				//  - Binary Excel files (2.0-2003 format; *.xls)
				//  - OpenXml Excel files (2007 format; *.xlsx)
				using (IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream))
				{
					var result = reader.AsDataSet();

					return result.Tables[sheetName];
				}
			}
		}

		#region Alt Excel Data Extraction Methods
		/// <summary>
		/// Extracts multiple DataTables from Excel
		/// </summary>
		/// <param name="fileName">File to parse</param>
		/// <returns>Multiple Excel sheets converted to a DataSet</returns>
		public DataSet Parse(string fileName)
		{
			string connectionString = string.Format("provider=Microsoft.Jet.OLEDB.4.0; data source={0};Extended Properties=Excel 8.0;", fileName);

			var data = new DataSet();

			foreach (var sheetName in GetExcelSheetNames(connectionString))
			{
				using (OleDbConnection con = new OleDbConnection(connectionString))
				{
					var dataTable = new DataTable();
					string query = string.Format("SELECT * FROM [{0}]", sheetName);
					con.Open();
					OleDbDataAdapter adapter = new OleDbDataAdapter(query, con);
					adapter.Fill(dataTable);
					data.Tables.Add(dataTable);
				}
			}

			return data;
		}

		/// <summary>
		/// Extracts a DataTable from Excel
		/// </summary>
		/// <param name="fileName">File to parse</param>
		/// <param name="sheetName">Name of the sheet to return as a DataTable</param>
		/// <returns>Excel sheet converted to a DataTable (1st row headers become field names)</returns>
		public DataTable Parse(string fileName, string sheetName)
		{
			string connectionString = string.Format("provider=Microsoft.Jet.OLEDB.4.0; data source={0};Extended Properties=Excel 8.0;", fileName);

			using (OleDbConnection con = new OleDbConnection(connectionString))
			{
				var dataTable = new DataTable();
				string query = string.Format("SELECT * FROM [{0}]", sheetName);
				con.Open();
				OleDbDataAdapter adapter = new OleDbDataAdapter(query, con);
				adapter.Fill(dataTable);
				return dataTable;
			}
		}

		/// <summary>
		/// Gets the names of all tabs in the spreadsheet
		/// </summary>
		/// <param name="connectionString"></param>
		/// <returns></returns>
		private string[] GetExcelSheetNames(string connectionString)
		{
			OleDbConnection con = null;
			DataTable dt = null;
			con = new OleDbConnection(connectionString);
			con.Open();
			dt = con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

			if (dt == null)
			{
				return null;
			}

			String[] excelSheetNames = new String[dt.Rows.Count];
			int i = 0;

			foreach (DataRow row in dt.Rows)
			{
				excelSheetNames[i] = row["TABLE_NAME"].ToString();
				i++;
			}

			return excelSheetNames;
		}
		#endregion
	}
}
