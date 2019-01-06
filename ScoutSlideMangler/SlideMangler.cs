using System;
using System.Data;

namespace ScoutSlideMangler
{
	public class SlideMangler
	{
		/// <summary>
		/// Extracts Excel data and creates slides
		/// </summary>
		public void CreateSlides(string excelFilePath, string sheetName)
		{
			DataTable data = new ExcelReader().GetData(excelFilePath, sheetName);
			var slideMaker = new MeritBadgeSlideCreator(18);
			var scoutFactory = new TroopPersonFactory(data);

			// Get first scout for processing
			TroopPerson scout = scoutFactory.GetNextScoutFromData();

			while (scout != null)   // Exit when we're all out of rows
			{
				slideMaker.SaveSlide(scout);
				Console.WriteLine($"{scout.FirstName} {scout.LastName}: {scout.MeritBadges.Count} badges");
				scout = scoutFactory.GetNextScoutFromData();
			}
		}
	}
}
