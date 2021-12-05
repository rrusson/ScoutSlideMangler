using System;
using System.Collections.Generic;
using System.Data;

namespace ScoutSlideMangler
{
	public class ScoutParser
	{
		private DataTable _data;
		private int _row;

		public void ProcessMeritBadges()
		{
			_data = ExcelReader.GetData(@"C:\Temp\COH.xlsx", "Merit Badge");

			int totalRows = _data.Rows.Count;

			for (int i = 1; i < _data.Rows.Count; i++)
			{
				var scout = new TroopPerson
				{
					FirstName = _data.Rows[i][1].ToString(),
					LastName = _data.Rows[i][0].ToString()
					//FullName = _data.Rows[i][0].ToString()	// If data combines first and last
				};

				Console.WriteLine(scout.LastName + ": " + scout.MeritBadges.Count);
			}

			Console.ReadKey();
		}

		private List<string> GetBadges(string scoutName)
		{
			var badges = new List<string>
			{
				_data.Rows[_row][3].ToString()
			};

			//First row is headers
			string nextScoutName = GetName(_data.Rows[1]);

			while (scoutName == nextScoutName)
			{
				badges.Add(_data.Rows[_row][3].ToString());
				_row++;
				scoutName = GetName(_data.Rows[_row]);
			}

			//nextName = (_row + 1 < _data.Rows.Count ? GetName(_data.Rows[_row + 1]) : "FINISHED");

			return badges;
		}

		private static string GetName(DataRow row)
		{
			return row[1].ToString() + Environment.NewLine + row[0].ToString();
		}
	}
}
