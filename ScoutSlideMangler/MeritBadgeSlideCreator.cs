﻿using System.Collections.Generic;

namespace ScoutSlideMangler
{
	public class MeritBadgeSlideCreator
	{
		private int _slideNumber;

		//private const string template = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><p:sld xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id=\"1\" name=\"\"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"0\" cy=\"0\"/><a:chOff x=\"0\" y=\"0\"/><a:chExt cx=\"0\" cy=\"0\"/></a:xfrm></p:grpSpPr><p:sp><p:nvSpPr><p:cNvPr id=\"1111\" name=\"{scout.FirstName}{scout.LastName}\"/><p:cNvSpPr txBox=\"1\"><a:spLocks noGrp=\"1\"/></p:cNvSpPr><p:nvPr><p:ph type=\"title\"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x=\"33866\" y=\"225954\"/><a:ext cx=\"4419601\" cy=\"1198469\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr><a:normAutofit fontScale=\"90000\"/></a:bodyPr><a:lstStyle/><a:p><a:pPr><a:defRPr sz=\"4000\" b=\"1\"><a:latin typeface=\"Copperplate Gothic Bold\"/><a:ea typeface=\"Copperplate Gothic Bold\"/><a:cs typeface=\"Copperplate Gothic Bold\"/><a:sym typeface=\"Copperplate Gothic Bold\"/></a:defRPr></a:pPr><a:r><a:t>{scout.FirstName}</a:t></a:r></a:p><a:p><a:pPr><a:defRPr sz=\"4000\" b=\"1\"><a:latin typeface=\"Copperplate Gothic Bold\"/><a:ea typeface=\"Copperplate Gothic Bold\"/><a:cs typeface=\"Copperplate Gothic Bold\"/><a:sym typeface=\"Copperplate Gothic Bold\"/></a:defRPr></a:pPr><a:r><a:t>{scout.LastName}</a:t></a:r></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id=\"1112\" name=\"Chess…\"/><p:cNvSpPr txBox=\"1\"><a:spLocks noGrp=\"1\"/></p:cNvSpPr><p:nvPr><p:ph type=\"body\" sz=\"half\" idx=\"1\"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x=\"220133\" y=\"1828800\"/><a:ext cx=\"4204855\" cy=\"4724400\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz=\"3400\"/></a:pPr><a:r><a:t>Chess</a:t></a:r></a:p><a:p><a:pPr><a:defRPr sz=\"3400\"/></a:pPr><a:r><a:t>Swimming*</a:t></a:r></a:p></p:txBody></p:sp><p:pic><p:nvPicPr><p:cNvPr id=\"1113\" name=\"r1944.jpg\" descr=\"r1944.jpg\"/><p:cNvPicPr><a:picLocks noChangeAspect=\"1\"/></p:cNvPicPr><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed=\"rId2\"><a:extLst/></a:blip><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x=\"4495800\" y=\"228600\"/><a:ext cx=\"4419600\" cy=\"5908675\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom><a:ln w=\"12700\"><a:miter lim=\"400000\"/></a:ln></p:spPr></p:pic></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr><mc:AlternateContent xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:p14=\"http://schemas.microsoft.com/office/powerpoint/2010/main\"><mc:Choice Requires=\"p14\"><p:transition spd=\"slow\"><p:fade thruBlk=\"1\"/></p:transition></mc:Choice><mc:Fallback xmlns:a14=\"http://schemas.microsoft.com/office/drawing/2010/main\" xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\" xmlns=\"\"><p:transition spd=\"med\"><p:fade/></p:transition></mc:Fallback></mc:AlternateContent></p:sld>";
		//private const string template = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><p:sld xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id=\"1\" name=\"\"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"0\" cy=\"0\"/><a:chOff x=\"0\" y=\"0\"/><a:chExt cx=\"0\" cy=\"0\"/></a:xfrm></p:grpSpPr><p:sp><p:nvSpPr><p:cNvPr id=\"1111\" ";
		//private const string junk = "<p:cNvSpPr txBox=\"1\"><a:spLocks noGrp=\"1\"/></p:cNvSpPr><p:nvPr><p:ph type=\"title\"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x=\"33866\" y=\"225954\"/><a:ext cx=\"4419601\" cy=\"1198469\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr><a:normAutofit fontScale=\"90000\"/></a:bodyPr><a:lstStyle/><a:p><a:pPr><a:defRPr sz=\"4000\" b=\"1\"><a:latin typeface=\"Copperplate Gothic Bold\"/><a:ea typeface=\"Copperplate Gothic Bold\"/><a:cs typeface=\"Copperplate Gothic Bold\"/><a:sym typeface=\"Copperplate Gothic Bold\"/></a:defRPr></a:pPr><a:r><a:t>{scout.FirstName}</a:t></a:r></a:p><a:p><a:pPr><a:defRPr sz=\"4000\" b=\"1\"><a:latin typeface=\"Copperplate Gothic Bold\"/><a:ea typeface=\"Copperplate Gothic Bold\"/><a:cs typeface=\"Copperplate Gothic Bold\"/><a:sym typeface=\"Copperplate Gothic Bold\"/></a:defRPr></a:pPr><a:r><a:t>{scout.LastName}</a:t></a:r></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id=\"1112\" name=\"Chess…\"/><p:cNvSpPr txBox=\"1\"><a:spLocks noGrp=\"1\"/></p:cNvSpPr><p:nvPr><p:ph type=\"body\" sz=\"half\" idx=\"1\"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x=\"220133\" y=\"1828800\"/><a:ext cx=\"4204855\" cy=\"4724400\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz=\"3400\"/></a:pPr><a:r><a:t>Chess</a:t></a:r></a:p><a:p><a:pPr><a:defRPr sz=\"3400\"/></a:pPr><a:r><a:t>Swimming*</a:t></a:r></a:p></p:txBody></p:sp><p:pic><p:nvPicPr><p:cNvPr id=\"1113\" name=\"r1944.jpg\" descr=\"r1944.jpg\"/><p:cNvPicPr><a:picLocks noChangeAspect=\"1\"/></p:cNvPicPr><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed=\"rId2\"><a:extLst/></a:blip><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x=\"4495800\" y=\"228600\"/><a:ext cx=\"4419600\" cy=\"5908675\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom><a:ln w=\"12700\"><a:miter lim=\"400000\"/></a:ln></p:spPr></p:pic></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr><mc:AlternateContent xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:p14=\"http://schemas.microsoft.com/office/powerpoint/2010/main\"><mc:Choice Requires=\"p14\"><p:transition spd=\"slow\"><p:fade thruBlk=\"1\"/></p:transition></mc:Choice><mc:Fallback xmlns:a14=\"http://schemas.microsoft.com/office/drawing/2010/main\" xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\" xmlns=\"\"><p:transition spd=\"med\"><p:fade/></p:transition></mc:Fallback></mc:AlternateContent></p:sld>";

		public MeritBadgeSlideCreator()
		{
			_slideNumber = 21;
		}

		public void SaveSlide(TroopPerson scout)
		{
			//string template = $"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><p:sld xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id=\"1\" name=\"\"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"0\" cy=\"0\"/><a:chOff x=\"0\" y=\"0\"/><a:chExt cx=\"0\" cy=\"0\"/></a:xfrm></p:grpSpPr><p:sp><p:nvSpPr><p:cNvPr id=\"1111\" name=\"{scout.FirstName}{scout.LastName}\"/><p:cNvSpPr txBox=\"1\"><a:spLocks noGrp=\"1\"/></p:cNvSpPr><p:nvPr><p:ph type=\"title\"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x=\"33866\" y=\"225954\"/><a:ext cx=\"4419601\" cy=\"1198469\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr><a:normAutofit fontScale=\"90000\"/></a:bodyPr><a:lstStyle/><a:p><a:pPr><a:defRPr sz=\"4000\" b=\"1\"><a:latin typeface=\"Copperplate Gothic Bold\"/><a:ea typeface=\"Copperplate Gothic Bold\"/><a:cs typeface=\"Copperplate Gothic Bold\"/><a:sym typeface=\"Copperplate Gothic Bold\"/></a:defRPr></a:pPr><a:r><a:t>{scout.FirstName}</a:t></a:r></a:p><a:p><a:pPr><a:defRPr sz=\"4000\" b=\"1\"><a:latin typeface=\"Copperplate Gothic Bold\"/><a:ea typeface=\"Copperplate Gothic Bold\"/><a:cs typeface=\"Copperplate Gothic Bold\"/><a:sym typeface=\"Copperplate Gothic Bold\"/></a:defRPr></a:pPr><a:r><a:t>{scout.LastName}</a:t></a:r></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id=\"1112\" name=\"Chess…\"/><p:cNvSpPr txBox=\"1\"><a:spLocks noGrp=\"1\"/></p:cNvSpPr><p:nvPr><p:ph type=\"body\" sz=\"half\" idx=\"1\"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x=\"220133\" y=\"1828800\"/><a:ext cx=\"4204855\" cy=\"4724400\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz=\"3400\"/></a:pPr><a:r><a:t>Chess</a:t></a:r></a:p><a:p><a:pPr><a:defRPr sz=\"3400\"/></a:pPr><a:r><a:t>Swimming*</a:t></a:r></a:p></p:txBody></p:sp><p:pic><p:nvPicPr><p:cNvPr id=\"1113\" name=\"r1944.jpg\" descr=\"r1944.jpg\"/><p:cNvPicPr><a:picLocks noChangeAspect=\"1\"/></p:cNvPicPr><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed=\"rId2\"><a:extLst/></a:blip><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x=\"4495800\" y=\"228600\"/><a:ext cx=\"4419600\" cy=\"5908675\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom><a:ln w=\"12700\"><a:miter lim=\"400000\"/></a:ln></p:spPr></p:pic></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr><mc:AlternateContent xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:p14=\"http://schemas.microsoft.com/office/powerpoint/2010/main\"><mc:Choice Requires=\"p14\"><p:transition spd=\"slow\"><p:fade thruBlk=\"1\"/></p:transition></mc:Choice><mc:Fallback xmlns:a14=\"http://schemas.microsoft.com/office/drawing/2010/main\" xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\" xmlns=\"\"><p:transition spd=\"med\"><p:fade/></p:transition></mc:Fallback></mc:AlternateContent></p:sld>";
			//string template = $"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><p:sld xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id=\"1\" name=\"\"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"0\" cy=\"0\"/><a:chOff x=\"0\" y=\"0\"/><a:chExt cx=\"0\" cy=\"0\"/></a:xfrm></p:grpSpPr><p:sp><p:nvSpPr><p:cNvPr id=\"1111\" name=\"{scout.FirstName}{scout.LastName}\"/><p:cNvSpPr txBox=\"1\"><a:spLocks noGrp=\"1\"/></p:cNvSpPr><p:nvPr><p:ph type=\"title\"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x=\"33866\" y=\"225954\"/><a:ext cx=\"4419601\" cy=\"1198469\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr><a:normAutofit fontScale=\"90000\"/></a:bodyPr><a:lstStyle/><a:p><a:pPr><a:defRPr sz=\"4000\" b=\"1\"><a:latin typeface=\"Copperplate Gothic Bold\"/><a:ea typeface=\"Copperplate Gothic Bold\"/><a:cs typeface=\"Copperplate Gothic Bold\"/><a:sym typeface=\"Copperplate Gothic Bold\"/></a:defRPr></a:pPr><a:r><a:t>{scout.FirstName}</a:t></a:r></a:p><a:p><a:pPr><a:defRPr sz=\"4000\" b=\"1\"><a:latin typeface=\"Copperplate Gothic Bold\"/><a:ea typeface=\"Copperplate Gothic Bold\"/><a:cs typeface=\"Copperplate Gothic Bold\"/><a:sym typeface=\"Copperplate Gothic Bold\"/></a:defRPr></a:pPr><a:r><a:t>{scout.LastName}</a:t></a:r></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id=\"1112\" name=\"Chess…\"/><p:cNvSpPr txBox=\"1\"><a:spLocks noGrp=\"1\"/></p:cNvSpPr><p:nvPr><p:ph type=\"body\" sz=\"half\" idx=\"1\"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x=\"220133\" y=\"1828800\"/><a:ext cx=\"4204855\" cy=\"4724400\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/>";
			string templateStart = $"<p:sld xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id=\"1\" name=\"\"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"0\" cy=\"0\"/><a:chOff x=\"0\" y=\"0\"/><a:chExt cx=\"0\" cy=\"0\"/></a:xfrm></p:grpSpPr><p:sp><p:nvSpPr><p:cNvPr id=\"1111\" name=\"{scout.FirstName} {scout.LastName}\"/><p:cNvSpPr txBox=\"1\"><a:spLocks noGrp=\"1\"/></p:cNvSpPr><p:nvPr><p:ph type=\"title\"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x=\"139338\" y=\"225954\"/><a:ext cx=\"4356462\" cy=\"1198469\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr><a:normAutofit/></a:bodyPr><a:lstStyle/><a:p><a:pPr><a:defRPr sz=\"4000\" b=\"1\"><a:latin typeface=\"Copperplate Gothic Bold\"/><a:ea typeface=\"Copperplate Gothic Bold\"/><a:cs typeface=\"Copperplate Gothic Bold\"/><a:sym typeface=\"Copperplate Gothic Bold\"/></a:defRPr></a:pPr><a:r><a:rPr dirty=\"0\"><a:effectLst><a:outerShdw blurRad=\"50800\" dist=\"38100\" dir=\"5400000\" algn=\"t\" rotWithShape=\"0\"><a:prstClr val=\"black\"><a:alpha val=\"40000\"/></a:prstClr></a:outerShdw></a:effectLst></a:rPr><a:t>{scout.FirstName}</a:t></a:r></a:p><a:p><a:pPr><a:defRPr sz=\"4000\" b=\"1\"><a:latin typeface=\"Copperplate Gothic Bold\"/><a:ea typeface=\"Copperplate Gothic Bold\"/><a:cs typeface=\"Copperplate Gothic Bold\"/><a:sym typeface=\"Copperplate Gothic Bold\"/></a:defRPr></a:pPr><a:r><a:rPr dirty=\"0\"><a:effectLst><a:outerShdw blurRad=\"50800\" dist=\"38100\" dir=\"5400000\" algn=\"t\" rotWithShape=\"0\"><a:prstClr val=\"black\"><a:alpha val=\"40000\"/></a:prstClr></a:outerShdw></a:effectLst></a:rPr><a:t>{scout.LastName}</a:t></a:r></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id=\"1112\" name=\"Chess…\"/><p:cNvSpPr txBox=\"1\"><a:spLocks noGrp=\"1\"/></p:cNvSpPr><p:nvPr><p:ph type=\"body\" sz=\"half\" idx=\"1\"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x=\"220133\" y=\"1637211\"/><a:ext cx=\"4204855\" cy=\"5077098\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr><a:normAutofit/></a:bodyPr><a:lstStyle/>";
			string templateEnd = "</p:txBody></p:sp><p:pic><p:nvPicPr><p:cNvPr id=\"1113\" name=\"r1944.jpg\" descr=\"r1944.jpg\"/><p:cNvPicPr><a:picLocks noChangeAspect=\"1\"/></p:cNvPicPr><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed=\"rId3\"><a:extLst/></a:blip><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x=\"4495800\" y=\"228600\"/><a:ext cx=\"4419600\" cy=\"5908675\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom><a:ln w=\"12700\"><a:solidFill><a:schemeClr val=\"tx1\"/></a:solidFill><a:miter lim=\"400000\"/></a:ln><a:effectLst><a:outerShdw blurRad=\"127000\" dist=\"190500\" dir=\"4200000\" algn=\"tl\" rotWithShape=\"0\"><a:prstClr val=\"black\"><a:alpha val=\"25000\"/></a:prstClr></a:outerShdw></a:effectLst></p:spPr></p:pic></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr><mc:AlternateContent xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:p14=\"http://schemas.microsoft.com/office/powerpoint/2010/main\"><mc:Choice Requires=\"p14\"><p:transition spd=\"slow\"><p:fade thruBlk=\"1\"/></p:transition></mc:Choice><mc:Fallback xmlns=\"\" xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\" xmlns:a14=\"http://schemas.microsoft.com/office/drawing/2010/main\"><p:transition spd=\"med\"><p:fade/></p:transition></mc:Fallback></mc:AlternateContent><p:timing><p:tnLst><p:par><p:cTn id=\"1\" dur=\"indefinite\" restart=\"never\" nodeType=\"tmRoot\"><p:childTnLst><p:seq concurrent=\"1\" nextAc=\"seek\"><p:cTn id=\"2\" dur=\"indefinite\" nodeType=\"mainSeq\"><p:childTnLst><p:par><p:cTn id=\"3\" fill=\"hold\"><p:stCondLst><p:cond delay=\"indefinite\"/><p:cond evt=\"onBegin\" delay=\"0\"><p:tn val=\"2\"/></p:cond></p:stCondLst><p:childTnLst><p:par><p:cTn id=\"4\" fill=\"hold\"><p:stCondLst><p:cond delay=\"0\"/></p:stCondLst><p:childTnLst><p:par><p:cTn id=\"5\" presetID=\"42\" presetClass=\"entr\" presetSubtype=\"0\" fill=\"hold\" grpId=\"0\" nodeType=\"afterEffect\"><p:stCondLst><p:cond delay=\"0\"/></p:stCondLst><p:childTnLst><p:set><p:cBhvr><p:cTn id=\"6\" dur=\"1\" fill=\"hold\"><p:stCondLst><p:cond delay=\"0\"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid=\"1112\"><p:txEl><p:pRg st=\"0\" end=\"0\"/></p:txEl></p:spTgt></p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val=\"visible\"/></p:to></p:set><p:animEffect transition=\"in\" filter=\"fade\"><p:cBhvr><p:cTn id=\"7\" dur=\"1000\"/><p:tgtEl><p:spTgt spid=\"1112\"><p:txEl><p:pRg st=\"0\" end=\"0\"/></p:txEl></p:spTgt></p:tgtEl></p:cBhvr></p:animEffect><p:anim calcmode=\"lin\" valueType=\"num\"><p:cBhvr><p:cTn id=\"8\" dur=\"1000\" fill=\"hold\"/><p:tgtEl><p:spTgt spid=\"1112\"><p:txEl><p:pRg st=\"0\" end=\"0\"/></p:txEl></p:spTgt></p:tgtEl><p:attrNameLst><p:attrName>ppt_x</p:attrName></p:attrNameLst></p:cBhvr><p:tavLst><p:tav tm=\"0\"><p:val><p:strVal val=\"#ppt_x\"/></p:val></p:tav><p:tav tm=\"100000\"><p:val><p:strVal val=\"#ppt_x\"/></p:val></p:tav></p:tavLst></p:anim><p:anim calcmode=\"lin\" valueType=\"num\"><p:cBhvr><p:cTn id=\"9\" dur=\"1000\" fill=\"hold\"/><p:tgtEl><p:spTgt spid=\"1112\"><p:txEl><p:pRg st=\"0\" end=\"0\"/></p:txEl></p:spTgt></p:tgtEl><p:attrNameLst><p:attrName>ppt_y</p:attrName></p:attrNameLst></p:cBhvr><p:tavLst><p:tav tm=\"0\"><p:val><p:strVal val=\"#ppt_y+.1\"/></p:val></p:tav><p:tav tm=\"100000\"><p:val><p:strVal val=\"#ppt_y\"/></p:val></p:tav></p:tavLst></p:anim></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par><p:par><p:cTn id=\"10\" fill=\"hold\"><p:stCondLst><p:cond delay=\"1000\"/></p:stCondLst><p:childTnLst><p:par><p:cTn id=\"11\" presetID=\"42\" presetClass=\"entr\" presetSubtype=\"0\" fill=\"hold\" grpId=\"0\" nodeType=\"afterEffect\"><p:stCondLst><p:cond delay=\"0\"/></p:stCondLst><p:childTnLst><p:set><p:cBhvr><p:cTn id=\"12\" dur=\"1\" fill=\"hold\"><p:stCondLst><p:cond delay=\"0\"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid=\"1112\"><p:txEl><p:pRg st=\"1\" end=\"1\"/></p:txEl></p:spTgt></p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val=\"visible\"/></p:to></p:set><p:animEffect transition=\"in\" filter=\"fade\"><p:cBhvr><p:cTn id=\"13\" dur=\"1000\"/><p:tgtEl><p:spTgt spid=\"1112\"><p:txEl><p:pRg st=\"1\" end=\"1\"/></p:txEl></p:spTgt></p:tgtEl></p:cBhvr></p:animEffect><p:anim calcmode=\"lin\" valueType=\"num\"><p:cBhvr><p:cTn id=\"14\" dur=\"1000\" fill=\"hold\"/><p:tgtEl><p:spTgt spid=\"1112\"><p:txEl><p:pRg st=\"1\" end=\"1\"/></p:txEl></p:spTgt></p:tgtEl><p:attrNameLst><p:attrName>ppt_x</p:attrName></p:attrNameLst></p:cBhvr><p:tavLst><p:tav tm=\"0\"><p:val><p:strVal val=\"#ppt_x\"/></p:val></p:tav><p:tav tm=\"100000\"><p:val><p:strVal val=\"#ppt_x\"/></p:val></p:tav></p:tavLst></p:anim><p:anim calcmode=\"lin\" valueType=\"num\"><p:cBhvr><p:cTn id=\"15\" dur=\"1000\" fill=\"hold\"/><p:tgtEl><p:spTgt spid=\"1112\"><p:txEl><p:pRg st=\"1\" end=\"1\"/></p:txEl></p:spTgt></p:tgtEl><p:attrNameLst><p:attrName>ppt_y</p:attrName></p:attrNameLst></p:cBhvr><p:tavLst><p:tav tm=\"0\"><p:val><p:strVal val=\"#ppt_y+.1\"/></p:val></p:tav><p:tav tm=\"100000\"><p:val><p:strVal val=\"#ppt_y\"/></p:val></p:tav></p:tavLst></p:anim></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par><p:par><p:cTn id=\"16\" fill=\"hold\"><p:stCondLst><p:cond delay=\"2000\"/></p:stCondLst><p:childTnLst><p:par><p:cTn id=\"17\" presetID=\"42\" presetClass=\"entr\" presetSubtype=\"0\" fill=\"hold\" grpId=\"0\" nodeType=\"afterEffect\"><p:stCondLst><p:cond delay=\"0\"/></p:stCondLst><p:childTnLst><p:set><p:cBhvr><p:cTn id=\"18\" dur=\"1\" fill=\"hold\"><p:stCondLst><p:cond delay=\"0\"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid=\"1112\"><p:txEl><p:pRg st=\"2\" end=\"2\"/></p:txEl></p:spTgt></p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val=\"visible\"/></p:to></p:set><p:animEffect transition=\"in\" filter=\"fade\"><p:cBhvr><p:cTn id=\"19\" dur=\"1000\"/><p:tgtEl><p:spTgt spid=\"1112\"><p:txEl><p:pRg st=\"2\" end=\"2\"/></p:txEl></p:spTgt></p:tgtEl></p:cBhvr></p:animEffect><p:anim calcmode=\"lin\" valueType=\"num\"><p:cBhvr><p:cTn id=\"20\" dur=\"1000\" fill=\"hold\"/><p:tgtEl><p:spTgt spid=\"1112\"><p:txEl><p:pRg st=\"2\" end=\"2\"/></p:txEl></p:spTgt></p:tgtEl><p:attrNameLst><p:attrName>ppt_x</p:attrName></p:attrNameLst></p:cBhvr><p:tavLst><p:tav tm=\"0\"><p:val><p:strVal val=\"#ppt_x\"/></p:val></p:tav><p:tav tm=\"100000\"><p:val><p:strVal val=\"#ppt_x\"/></p:val></p:tav></p:tavLst></p:anim><p:anim calcmode=\"lin\" valueType=\"num\"><p:cBhvr><p:cTn id=\"21\" dur=\"1000\" fill=\"hold\"/><p:tgtEl><p:spTgt spid=\"1112\"><p:txEl><p:pRg st=\"2\" end=\"2\"/></p:txEl></p:spTgt></p:tgtEl><p:attrNameLst><p:attrName>ppt_y</p:attrName></p:attrNameLst></p:cBhvr><p:tavLst><p:tav tm=\"0\"><p:val><p:strVal val=\"#ppt_y+.1\"/></p:val></p:tav><p:tav tm=\"100000\"><p:val><p:strVal val=\"#ppt_y\"/></p:val></p:tav></p:tavLst></p:anim></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par><p:par><p:cTn id=\"22\" fill=\"hold\"><p:stCondLst><p:cond delay=\"3000\"/></p:stCondLst><p:childTnLst><p:par><p:cTn id=\"23\" presetID=\"42\" presetClass=\"entr\" presetSubtype=\"0\" fill=\"hold\" grpId=\"0\" nodeType=\"afterEffect\"><p:stCondLst><p:cond delay=\"0\"/></p:stCondLst><p:childTnLst><p:set><p:cBhvr><p:cTn id=\"24\" dur=\"1\" fill=\"hold\"><p:stCondLst><p:cond delay=\"0\"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid=\"1112\"><p:txEl><p:pRg st=\"3\" end=\"3\"/></p:txEl></p:spTgt></p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val=\"visible\"/></p:to></p:set><p:animEffect transition=\"in\" filter=\"fade\"><p:cBhvr><p:cTn id=\"25\" dur=\"1000\"/><p:tgtEl><p:spTgt spid=\"1112\"><p:txEl><p:pRg st=\"3\" end=\"3\"/></p:txEl></p:spTgt></p:tgtEl></p:cBhvr></p:animEffect><p:anim calcmode=\"lin\" valueType=\"num\"><p:cBhvr><p:cTn id=\"26\" dur=\"1000\" fill=\"hold\"/><p:tgtEl><p:spTgt spid=\"1112\"><p:txEl><p:pRg st=\"3\" end=\"3\"/></p:txEl></p:spTgt></p:tgtEl><p:attrNameLst><p:attrName>ppt_x</p:attrName></p:attrNameLst></p:cBhvr><p:tavLst><p:tav tm=\"0\"><p:val><p:strVal val=\"#ppt_x\"/></p:val></p:tav><p:tav tm=\"100000\"><p:val><p:strVal val=\"#ppt_x\"/></p:val></p:tav></p:tavLst></p:anim><p:anim calcmode=\"lin\" valueType=\"num\"><p:cBhvr><p:cTn id=\"27\" dur=\"1000\" fill=\"hold\"/><p:tgtEl><p:spTgt spid=\"1112\"><p:txEl><p:pRg st=\"3\" end=\"3\"/></p:txEl></p:spTgt></p:tgtEl><p:attrNameLst><p:attrName>ppt_y</p:attrName></p:attrNameLst></p:cBhvr><p:tavLst><p:tav tm=\"0\"><p:val><p:strVal val=\"#ppt_y+.1\"/></p:val></p:tav><p:tav tm=\"100000\"><p:val><p:strVal val=\"#ppt_y\"/></p:val></p:tav></p:tavLst></p:anim></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn><p:prevCondLst><p:cond evt=\"onPrev\" delay=\"0\"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst><p:nextCondLst><p:cond evt=\"onNext\" delay=\"0\"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst></p:seq></p:childTnLst></p:cTn></p:par></p:tnLst><p:bldLst><p:bldP spid=\"1112\" grpId=\"0\" uiExpand=\"1\" build=\"p\"/></p:bldLst></p:timing></p:sld>";

			//string badgeName = "Basket Weaving";
			//string badgeTemplate = $"<a:p><a:r><a:t>{badgeName}</a:t></a:r></a:p>";
			string templateBadges = GetBadges(scout.MeritBadges);
			string slide = string.Format(templateStart, scout.FirstName, scout.LastName) + templateBadges + templateEnd;

			//string slide = template;
			System.IO.File.WriteAllText($@"C:\Temp\slide{_slideNumber}.xml", slide);
			_slideNumber++;
		}


		private string GetBadges(List<string> badges)
		{
			string chunk = "";
			foreach (var badgeName in badges)
			{
				chunk = chunk + $"<a:p><a:pPr marL=\"227013\" indent=\"-227013\"><a:buSzPct val=\"120000\"/><a:buBlip><a:blip r:embed=\"rId2\"/></a:buBlip></a:pPr>"
					+ "<a:r><a:rPr dirty=\"0\"><a:latin typeface=\"Verdana\" panose=\"020B0604030504040204\" pitchFamily=\"34\" charset=\"0\"/>"
					+ $"<a:ea typeface=\"Verdana\" panose=\"020B0604030504040204\" pitchFamily=\"34\" charset=\"0\"/></a:rPr><a:t>{badgeName}</a:t></a:r></a:p>";
			}
			
			return chunk;
		}
	}
}