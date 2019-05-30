using System.Collections.Generic;

namespace ScoutSlideMangler
{
	public class MeritBadgeSlideCreator
	{
		private int _slideNumber;

		/// <summary>
		/// Makes a slide based off a big, ugly template
		/// </summary>
		/// <param name="firstSlideNumber">The first slide number to insert/overwrite in the .PPTX</param>
		public MeritBadgeSlideCreator(int firstSlideNumber)
		{
			_slideNumber = firstSlideNumber;
		}

		public void SaveSlide(TroopPerson scout)
		{
			//string template = $"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><p:sld xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id=\"1\" name=\"\"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"0\" cy=\"0\"/><a:chOff x=\"0\" y=\"0\"/><a:chExt cx=\"0\" cy=\"0\"/></a:xfrm></p:grpSpPr><p:sp><p:nvSpPr><p:cNvPr id=\"1111\" name=\"{scout.FirstName}{scout.LastName}\"/><p:cNvSpPr txBox=\"1\"><a:spLocks noGrp=\"1\"/></p:cNvSpPr><p:nvPr><p:ph type=\"title\"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x=\"33866\" y=\"225954\"/><a:ext cx=\"4419601\" cy=\"1198469\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr><a:normAutofit fontScale=\"90000\"/></a:bodyPr><a:lstStyle/><a:p><a:pPr><a:defRPr sz=\"4000\" b=\"1\"><a:latin typeface=\"Copperplate Gothic Bold\"/><a:ea typeface=\"Copperplate Gothic Bold\"/><a:cs typeface=\"Copperplate Gothic Bold\"/><a:sym typeface=\"Copperplate Gothic Bold\"/></a:defRPr></a:pPr><a:r><a:t>{scout.FirstName}</a:t></a:r></a:p><a:p><a:pPr><a:defRPr sz=\"4000\" b=\"1\"><a:latin typeface=\"Copperplate Gothic Bold\"/><a:ea typeface=\"Copperplate Gothic Bold\"/><a:cs typeface=\"Copperplate Gothic Bold\"/><a:sym typeface=\"Copperplate Gothic Bold\"/></a:defRPr></a:pPr><a:r><a:t>{scout.LastName}</a:t></a:r></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id=\"1112\" name=\"Chess…\"/><p:cNvSpPr txBox=\"1\"><a:spLocks noGrp=\"1\"/></p:cNvSpPr><p:nvPr><p:ph type=\"body\" sz=\"half\" idx=\"1\"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x=\"220133\" y=\"1828800\"/><a:ext cx=\"4204855\" cy=\"4724400\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz=\"3400\"/></a:pPr><a:r><a:t>Chess</a:t></a:r></a:p><a:p><a:pPr><a:defRPr sz=\"3400\"/></a:pPr><a:r><a:t>Swimming*</a:t></a:r></a:p></p:txBody></p:sp><p:pic><p:nvPicPr><p:cNvPr id=\"1113\" name=\"r1944.jpg\" descr=\"r1944.jpg\"/><p:cNvPicPr><a:picLocks noChangeAspect=\"1\"/></p:cNvPicPr><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed=\"rId2\"><a:extLst/></a:blip><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x=\"4495800\" y=\"228600\"/><a:ext cx=\"4419600\" cy=\"5908675\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom><a:ln w=\"12700\"><a:miter lim=\"400000\"/></a:ln></p:spPr></p:pic></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr><mc:AlternateContent xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:p14=\"http://schemas.microsoft.com/office/powerpoint/2010/main\"><mc:Choice Requires=\"p14\"><p:transition spd=\"slow\"><p:fade thruBlk=\"1\"/></p:transition></mc:Choice><mc:Fallback xmlns:a14=\"http://schemas.microsoft.com/office/drawing/2010/main\" xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\" xmlns=\"\"><p:transition spd=\"med\"><p:fade/></p:transition></mc:Fallback></mc:AlternateContent></p:sld>";
			string template = $"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><p:sld xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id=\"1\" name=\"\"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"0\" cy=\"0\"/><a:chOff x=\"0\" y=\"0\"/><a:chExt cx=\"0\" cy=\"0\"/></a:xfrm></p:grpSpPr><p:sp><p:nvSpPr><p:cNvPr id=\"1111\" name=\"{scout.FirstName}{scout.LastName}\"/><p:cNvSpPr txBox=\"1\"><a:spLocks noGrp=\"1\"/></p:cNvSpPr><p:nvPr><p:ph type=\"title\"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x=\"33866\" y=\"225954\"/><a:ext cx=\"4419601\" cy=\"1198469\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr><a:normAutofit fontScale=\"90000\"/></a:bodyPr><a:lstStyle/><a:p><a:pPr><a:defRPr sz=\"4000\" b=\"1\"><a:latin typeface=\"Copperplate Gothic Bold\"/><a:ea typeface=\"Copperplate Gothic Bold\"/><a:cs typeface=\"Copperplate Gothic Bold\"/><a:sym typeface=\"Copperplate Gothic Bold\"/></a:defRPr></a:pPr><a:r><a:t>{scout.FirstName}</a:t></a:r></a:p><a:p><a:pPr><a:defRPr sz=\"4000\" b=\"1\"><a:latin typeface=\"Copperplate Gothic Bold\"/><a:ea typeface=\"Copperplate Gothic Bold\"/><a:cs typeface=\"Copperplate Gothic Bold\"/><a:sym typeface=\"Copperplate Gothic Bold\"/></a:defRPr></a:pPr><a:r><a:t>{scout.LastName}</a:t></a:r></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id=\"1112\" name=\"Chess…\"/><p:cNvSpPr txBox=\"1\"><a:spLocks noGrp=\"1\"/></p:cNvSpPr><p:nvPr><p:ph type=\"body\" sz=\"half\" idx=\"1\"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x=\"220133\" y=\"1828800\"/><a:ext cx=\"4204855\" cy=\"4724400\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/>";
			string temp2 = "</p:txBody></p:sp><p:pic><p:nvPicPr><p:cNvPr id=\"1113\" name=\"r1944.jpg\" descr=\"r1944.jpg\"/><p:cNvPicPr><a:picLocks noChangeAspect=\"1\"/></p:cNvPicPr><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed=\"rId2\"><a:extLst/></a:blip><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x=\"4495800\" y=\"228600\"/><a:ext cx=\"4419600\" cy=\"5908675\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom><a:ln w=\"12700\"><a:miter lim=\"400000\"/></a:ln></p:spPr></p:pic></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr><mc:AlternateContent xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:p14=\"http://schemas.microsoft.com/office/powerpoint/2010/main\"><mc:Choice Requires=\"p14\"><p:transition spd=\"slow\"><p:fade thruBlk=\"1\"/></p:transition></mc:Choice><mc:Fallback xmlns:a14=\"http://schemas.microsoft.com/office/drawing/2010/main\" xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\" xmlns=\"\"><p:transition spd=\"med\"><p:fade/></p:transition></mc:Fallback></mc:AlternateContent></p:sld>";

			//string badgeName = "Basket Weaving";
			//string badgeTemplate = $"<a:p><a:r><a:t>{badgeName}</a:t></a:r></a:p>";
			string badges = GetBadges(scout.MeritBadges);
			string slide = string.Format(template, scout.FirstName, scout.LastName) + badges + temp2;

			System.IO.File.WriteAllText($@"C:\Temp\CoH\ppt\slides\slide{_slideNumber}.xml", slide);
			_slideNumber++;
		}


		private string GetBadges(List<string> badges)
		{
			string chunk = "";
			foreach (var badgeName in badges)
			{
				chunk = chunk + $"<a:p><a:r><a:t>{badgeName}</a:t></a:r></a:p>";
			}

			return chunk;
		}
	}
}
