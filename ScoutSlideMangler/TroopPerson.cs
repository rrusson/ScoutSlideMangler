using System.Collections.Generic;

namespace ScoutSlideMangler
{
	public class TroopPerson
	{
		private string _fullName;

		public string FirstName { get; set; }


		public string LastName { get; set; }


		public string FullName {
			get => _fullName ?? FirstName + " " + LastName;
			set => _fullName = value;
		}


		public string Position { get; set; }


		public string Rank { get; set; }


		public List<string> MeritBadges { get; set; } = new List<string>();
	}
}
