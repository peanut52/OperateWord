/**
 * find_content Macro
 * 宏由 ljing 录制，时间: 2022/06/12
 如何在文章中查找特定的内容，如简介
 */
function find_content()
{
	Selection.Find.Wrap = wdFindContinue;
	Selection.Find.Wrap = wdFindContinue;
	(obj=>{
		obj.Text = "简介";
		obj.Forward = true;
		obj.Wrap = wdFindContinue;
		obj.MatchCase = false;
		obj.MatchByte = true;
		obj.MatchWildcards = false;
		obj.MatchWholeWord = false;
		obj.MatchFuzzy = false;
		obj.Replacement.Text = "";
	})(Selection.Find);
	(obj=>{
		obj.Style = "";
		obj.Highlight = wdUndefined;
		(obj=>{
			obj.Style = "";
			obj.Highlight = wdUndefined;
		})(obj.Replacement);
	})(Selection.Find);
	Selection.Find.Execute(undefined, undefined, undefined, undefined, undefined, undefined, undefined, undefined, undefined, undefined, wdReplaceNone, undefined, undefined, undefined, undefined);
	(obj=>{
		obj.Size = 14;
		obj.SizeBi = 14;
	})(Selection.Font);
	Selection.Font.Name = "仿宋";
	Selection.SetRange(441, 441);

}
/**
 * changeColor Macro
 * 宏由 ljing 录制，时间: 2022/06/12
 设置字体，字号
 */
function changeColor()
{
	Selection.SetRange(16, 438);
	Selection.Font.Name = "宋体";
	(obj=>{
		obj.Size = 14;
		obj.SizeBi = 14;
	})(Selection.Font);

}
/**
 * chinese_and_english_font Macro
 * 宏由 ljing 录制，时间: 2022/06/12
 同时设置中文和英文字体字号
 */
function chinese_and_english_font()
{
	Selection.WholeStory();
	ActiveDocument.Range(0, 16).Start = 16;
	ActiveDocument.Range(16, 16).End = 16;
	(obj=>{
		obj.Underline = wdUnderlineNone;
		obj.EmphasisMark = wdEmphasisMarkNone;
		obj.Hidden = 0;
		obj.Shadow = 0;
		obj.Outline = 0;
		obj.Emboss = 0;
		obj.Engrave = 0;
		obj.Scaling = 100;
		obj.Scaling = 100;
		obj.NameFarEast = "等线 Light";
		obj.NameBi = "宋体";
		obj.NameFarEast = "宋体";
		obj.Bold = 0;
		obj.Size = 14;
		obj.NameAscii = "Times New Roman";
		obj.NameFarEast = "宋体";
		obj.NameAscii = "Times New Roman";
		obj.NameOther = "Times New Roman";
		obj.Bold = 0;
		obj.Size = 14;
	})(Selection.Font);

}
/**
 * paragraph_distance Macro
 * 宏由 ljing 录制，时间: 2022/06/12
 段前与段后间距
 */
function paragraph_distance()
{
	Selection.SetRange(0, 0);
	Selection.EndKey(wdLine, wdExtend);
	(obj=>{
		obj.LineUnitBefore = 0.500000;
		obj.LineUnitBefore = 1;
		obj.SpaceAfter = 16;
		obj.SpaceAfter = 1;
		obj.LineSpacingRule = wdLineSpaceMultiple;
		obj.LineSpacing = 36;
		obj.LineSpacing = 30;
		obj.LineSpacing = 24;
		obj.LineUnitBefore = 1;
		obj.SpaceAfter = 1;
		obj.LineSpacingRule = wdLineSpaceMultiple;
		obj.LineSpacing = 24;
		obj.DisableLineHeightGrid = 0;
		obj.ReadingOrder = wdReadingOrderLtr;
		obj.AutoAdjustRightIndent = -1;
		obj.WidowControl = 0;
		obj.KeepWithNext = -1;
		obj.KeepTogether = -1;
		obj.PageBreakBefore = 0;
		obj.FarEastLineBreakControl = -1;
		obj.WordWrap = -1;
		obj.HangingPunctuation = -1;
		obj.HalfWidthPunctuationOnTopOfLine = 0;
		obj.AddSpaceBetweenFarEastAndAlpha = -1;
		obj.AddSpaceBetweenFarEastAndDigit = -1;
		obj.BaseLineAlignment = wdBaselineAlignAuto;
	})(Selection.ParagraphFormat);

}
/**
 * yemei Macro
 * 宏由 ljing 录制，时间: 2022/06/12
 插入所有页页眉，也就是标题居中，并设置了下划线，奇偶页全部相同
 */
function yemei()
{
	Selection.SetRange(8, 8);
	ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader;
	Selection.SetRange(9993, 9993);
	Selection.PasteAndFormat(wdUseDestinationStylesRecovery);
	Selection.HomeKey(wdLine, wdExtend);
	Selection.SetRange(10001, 10001);
	Selection.EndKey(wdLine, wdMove);
	Selection.HomeKey(wdLine, wdExtend);
	(obj=>{
		obj.EnableFirstPageInSection = true;
		obj.EnableOtherPagesInSection = true;
		obj.ApplyPageBordersToAllSections();
	})(ActiveDocument.Sections.Item(1).Borders);
	(obj=>{
		obj.LineStyle = wdLineStyleNone;
		obj.Visible = false;
	})(Selection.ParagraphFormat.Borders.Item(wdBorderTop));
	(obj=>{
		obj.LineStyle = wdLineStyleNone;
		obj.Visible = false;
	})(Selection.ParagraphFormat.Borders.Item(wdBorderLeft));
	(obj=>{
		obj.Visible = true;
		obj.LineStyle = wdLineStyleThinThickSmallGap;
		obj.LineWidth = wdLineWidth150pt;
	})(Selection.ParagraphFormat.Borders.Item(wdBorderBottom));
	(obj=>{
		obj.LineStyle = wdLineStyleNone;
		obj.Visible = false;
	})(Selection.ParagraphFormat.Borders.Item(wdBorderRight));
	(obj=>{
		obj.LineStyle = wdLineStyleNone;
		obj.Visible = false;
	})(Selection.ParagraphFormat.Borders.Item(wdBorderHorizontal));
	(obj=>{
		obj.DistanceFromLeft = 4;
		obj.DistanceFromRight = 4;
		obj.DistanceFromTop = 1;
		obj.DistanceFromBottom = 1;
	})(Selection.ParagraphFormat.Borders);
	(obj=>{
		obj.DefaultBorderLineStyle = wdLineStyleThinThickSmallGap;
		obj.DefaultBorderLineWidth = wdLineWidth150pt;
		obj.DefaultBorderColor = wdColorBlack;
	})(Options);
	Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter;
	Selection.SetRange(9994, 9994);
	ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument;
	Selection.SetRange(85, 85);

}
/**
 * addPage Macro
 * 宏由 ljing 录制，时间: 2022/06/12
 添加页码
 */
function addPage()
{
	Selection.HomeKey(wdStory, wdMove);
	Selection.SetRange(441, 441);
	Selection.MoveDown(wdLine, 6, wdMove);
	Selection.SetRange(596, 596);
	ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageFooter;
	Selection.SetRange(10009, 10009);
	(obj=>{
		obj.Shapes.AddTextbox(msoTextOrientationHorizontal, 0, 0, 144, 144, ActiveDocument.Sections.Item(1).Footers.Item(wdHeaderFooterPrimary).Range);
		(obj=>{
			obj.Fill.Visible = msoFalse;
			obj.Line.Visible = msoFalse;
			(obj=>{
				obj.AutoSize = 1;
				obj.WordWrap = 0;
				obj.MarginLeft = 0;
				obj.MarginRight = 0;
				obj.MarginTop = 0;
				obj.MarginBottom = 0;
				obj.Orientation = msoTextOrientationHorizontal;
			})(obj.TextFrame);
			obj.RelativeHorizontalPosition = wdRelativeHorizontalPositionMargin;
			obj.Left = -999995;
			obj.RelativeVerticalPosition = wdRelativeVerticalPositionParagraph;
			obj.Top = 0;
			obj.WrapFormat.Type = wdWrapNone;
			(obj=>{
				obj.Text = "X";
				obj.Fields.Add(ActiveDocument.Sections.Item(1).Footers.Item(wdHeaderFooterPrimary).Shapes.Item(1).TextFrame.TextRange, wdFieldPage, "", true);
			})(obj.TextFrame.TextRange);
		})(obj.Shapes.Item(1));
		(obj=>{
			obj.NumberStyle = wdPageNumberStyleArabic;
			obj.RestartNumberingAtSection = true;
			obj.StartingNumber = 1;
		})(obj.PageNumbers);
	})(ActiveDocument.Sections.Item(1).Footers.Item(wdHeaderFooterPrimary));
	Selection.SetRange(10009, 10009);
	ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument;
	Selection.SetRange(497, 497);

}
/**
 * odd_and_single_page Macro
 * 宏由 ljing 录制，时间: 2022/06/12
 设置页眉页脚奇偶页不同的方法
 */
function odd_and_single_page()
{
	Selection.SetRange(0, 15);
	Selection.Copy();
	Selection.SetRange(6, 6);
	ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader;
	Selection.SetRange(9993, 9993);
	Selection.PasteAndFormat(wdUseDestinationStylesRecovery);
	(obj=>{
		obj.OddAndEvenPagesHeaderFooter = 1;
		obj.DifferentFirstPageHeaderFooter = 0;
	})(ActiveDocument.Range(10008, 10008).Range.PageSetup);
	(obj=>{
		(obj=>{
			obj.LineStyle = wdLineStyleNone;
			obj.Color = wdColorBlack;
		})(obj.Headers.Item(wdHeaderFooterPrimary).Range.Paragraphs.Item(1).Format.Borders.Item(wdBorderBottom));
		(obj=>{
			obj.LineStyle = wdLineStyleNone;
			obj.Color = wdColorBlack;
		})(obj.Headers.Item(wdHeaderFooterEvenPages).Range.Paragraphs.Item(1).Format.Borders.Item(wdBorderBottom));
	})(ActiveDocument.Sections.Item(1));
	ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader;
	Selection.WholeStory();
	(obj=>{
		obj.EnableFirstPageInSection = true;
		obj.EnableOtherPagesInSection = true;
		obj.ApplyPageBordersToAllSections();
	})(ActiveDocument.Sections.Item(1).Borders);
	(obj=>{
		obj.LineStyle = wdLineStyleNone;
		obj.Visible = false;
	})(Selection.ParagraphFormat.Borders.Item(wdBorderTop));
	(obj=>{
		obj.LineStyle = wdLineStyleNone;
		obj.Visible = false;
	})(Selection.ParagraphFormat.Borders.Item(wdBorderLeft));
	(obj=>{
		obj.Visible = true;
		obj.LineStyle = wdLineStyleSingle;
		obj.LineWidth = wdLineWidth050pt;
	})(Selection.ParagraphFormat.Borders.Item(wdBorderBottom));
	(obj=>{
		obj.LineStyle = wdLineStyleNone;
		obj.Visible = false;
	})(Selection.ParagraphFormat.Borders.Item(wdBorderRight));
	(obj=>{
		obj.LineStyle = wdLineStyleNone;
		obj.Visible = false;
	})(Selection.ParagraphFormat.Borders.Item(wdBorderHorizontal));
	(obj=>{
		obj.DistanceFromLeft = 4;
		obj.DistanceFromRight = 4;
		obj.DistanceFromTop = 1;
		obj.DistanceFromBottom = 1;
	})(Selection.ParagraphFormat.Borders);
	(obj=>{
		obj.DefaultBorderLineStyle = wdLineStyleThinThickSmallGap;
		obj.DefaultBorderLineWidth = wdLineWidth150pt;
		obj.DefaultBorderColor = wdColorBlack;
	})(Options);
	Selection.SetRange(9994, 9994);
	ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument;
	Selection.SetRange(283, 283);
	Selection.MoveDown(wdLine, 43, wdMove);
	Selection.MoveUp(wdLine, 39, wdMove);
	Selection.SetRange(640, 640);
	ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader;
	Selection.SetRange(10011, 10011);
	Selection.TypeText("v");
	Selection.TypeBackspace();
	Selection.TypeText("摘要");
	Selection.SetRange(10013, 10013);
	Selection.WholeStory();
	Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter;
	Selection.SetRange(10011, 10011);
	ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument;
	Selection.SetRange(675, 675);

}
function main（）{
	addPage();
}