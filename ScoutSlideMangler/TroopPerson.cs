using System.Collections.Generic;

namespace ScoutSlideMangler
{
	public class TroopPerson
	{
		public string FirstName { get; set; }


		public string LastName { get; set; }


		public string Position { get; set; }


		public string Rank { get; set; }


		public List<string> MeritBadges { get; set; } = new List<string>();
	}
}
