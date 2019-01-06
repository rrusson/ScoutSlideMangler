using System;
using System.Data;

namespace ScoutSlideMangler
{
	public class TroopPersonFactory
	{
		DataTable _data;

		int _currentRow;
		string _currentName;

		public TroopPersonFactory(DataTable data)
		{
			if (data == null || data.Rows.Count < 1)
			{
				throw new DataException("NO DATA FOUND!!!");
			}

			_data = data;

			//Prime values for first record
			_currentRow = 1;
			_currentName = GetNextName(1);
		}


		/*
			FORMAT IN EXCEL IS 4 COLUMNS LIKE THIS, WITH NAME AND BADGES:
			LastName	FirstName	Type	BadgeName
			Davis		Dave		[opt]	Art
											Cooking
											Citizenship in the World
			Jones		John				Nature
											Painting
		*/

		public TroopPerson GetNextScoutFromData()
		{
			if (_currentRow >= _data.Rows.Count)
			{
				return null;
			}

			DataRow row = _data.Rows[_currentRow];

			var scout = new TroopPerson
			{
				FirstName = row[1].ToString().Trim(),
				LastName = row[0].ToString().Trim()
			};

			string nextName;

			// Add Badges to Scout until a new name is encountered
			do
			{
				scout.MeritBadges.Add(_data.Rows[_currentRow][3].ToString());
				_currentRow++;
				nextName = GetNextName(_currentRow);
			} while (nextName == _currentName);

			_currentName = nextName;	//change name to next scout
			return scout;
		}

		internal string GetNextName(int rowNumber)
		{
			if (rowNumber >= _data.Rows.Count)
			{
				return null;
			}

			DataRow row = _data.Rows[rowNumber];
			if (string.IsNullOrWhiteSpace(row[1].ToString()))
			{
				return _currentName;	// Empty rows mean data didn't repeat name of Scout
			}

			return row[1].ToString() + Environment.NewLine + row[0].ToString();
		}

	}
}
