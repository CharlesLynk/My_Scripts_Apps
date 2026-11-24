using System.Collections.Generic;
using Xunit;

namespace MbrohDrawingIndexCreator.Tests;

public sealed class OcrTextCleanupTests
{
	[Theory]
	[InlineData("G16", "G-16")]
	[InlineData("G-I6", "G-16")]
	[InlineData("G-l6", "G-16")]
	[InlineData("6-13", "G-13")]
	[InlineData("i G-14 |", "G-14")]
	[InlineData("| Y-36 |", "Y-36")]
	[InlineData("RW-5-02", "RW-S-02")]
	[InlineData("G - 16", "G-16")]
	[InlineData("D-Ol", "D-01")]
	public void CleanSheetId_RepairsCommonOcrMistakes(string raw, string expected)
	{
		Assert.Equal(expected, OcrTextCleanup.CleanSheetId(raw, expectHyphen: true));
	}

	[Theory]
	[InlineData("G-16", true)]
	[InlineData("RW-S-02", true)]
	[InlineData("A101", true)]
	[InlineData("G-R", false)]
	[InlineData("ST-M-0", true)]
	[InlineData("G-00", true)]
	[InlineData("G-10", true)]
	[InlineData("", false)]
	[InlineData("WEST", false)]
	public void IsPlausibleSheetId_AcceptsOnlyIdShapes(string id, bool expected)
	{
		Assert.Equal(expected, OcrTextCleanup.IsPlausibleSheetId(id));
	}

	[Theory]
	[InlineData("DESIGN CRITERIA |", "DESIGN CRITERIA I")]
	[InlineData("SECTIONS AND DETAILS Ill", "SECTIONS AND DETAILS III")]
	[InlineData("SECTIONS AND DETAILS Vil", "SECTIONS AND DETAILS VII")]
	[InlineData("[ AGGING CONVENTIONS]", "AGGING CONVENTIONS")]
	[InlineData("MECHANICAL\nLEGEND AND SYMBOLS", "MECHANICAL LEGEND AND SYMBOLS")]
	public void CleanTitle_RemovesDebrisAndFixesRomanNumerals(string raw, string expected)
	{
		Assert.Equal(expected, OcrTextCleanup.CleanTitle(raw));
	}

	[Fact]
	public void CleanTitle_KeepsRealNumbers()
	{
		Assert.Equal("UNIT 1", OcrTextCleanup.CleanTitle("UNIT 1"));
	}

	[Fact]
	public void Vocabulary_SnapsNearMissDisciplinesAndTitleWords()
	{
		OcrVocabulary vocabulary = OcrVocabulary.CreateDefault();

		Assert.Equal("GENERAL", vocabulary.SnapDiscipline(OcrTextCleanup.CleanDiscipline("GE NERAL")));
		Assert.Equal("ABBREVIATIONS", vocabulary.SnapTitle("ABREVIAIIONS"));
		Assert.Equal("SYMBOLS, LEGENDS & ABBREVIATIONS", vocabulary.SnapTitle("SYMBOLS, LEGENDS & ABBREVIAIIONS"));
	}

	[Fact]
	public void Vocabulary_LearnsFromTheSetAndLeavesUnknownWordsAlone()
	{
		OcrVocabulary vocabulary = OcrVocabulary.CreateDefault();
		vocabulary.AddDiscipline("OZONATION");
		vocabulary.AddTitle("OZONATED WATER METER VAULT");

		Assert.Equal("OZONATION", vocabulary.SnapDiscipline("OZONATI0N"));
		Assert.Equal("OZONATED WATER METER VAULT", vocabulary.SnapTitle("OZONATED WATER METER VAULT"));
		Assert.Equal("QUUX FROBNICATE", vocabulary.SnapTitle("QUUX FROBNICATE"));
	}
}
