using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.OleDb;
using System.IO;
using ExcelDataReader;

namespace ScoutSlideMangler
{
	public static class ExcelReader
	{
		public static DataSet Parse(string fileName)
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

		public static DataTable Parse(string fileName, string sheetName)
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

		public static DataTable GetData(string filePath, string sheetName)
		{
			using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
			{
				// Auto-detect format, supports:
				//  - Binary Excel files (2.0-2003 format; *.xls)
				//  - OpenXml Excel files (2007 format; *.xlsx)
				using (IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream))
				{

					// Choose one of either 1 or 2:

					// 1. Use the reader methods
					//do
					//{
					//	while (reader.Read())
					//	{
					//		reader.GetDouble(0);
					//	}
					//} while (reader.NextResult());

					// 2. Use the AsDataSet extension method
					var result = reader.AsDataSet();

					return result.Tables[sheetName];
				}
			}
		}


		public static string[] GetExcelSheetNames(string connectionString)
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

	}
}
