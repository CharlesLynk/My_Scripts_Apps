using System.Collections.Generic;
using System.Linq;
using Xunit;

namespace MbrohDrawingIndexCreator.Tests;

// The rows below are real OCR output from the Bachman set (see the run log), so these tests pin the repairs that matter in practice.
public sealed class IndexAnalyzerTests
{
	private static IndexRow Ocr(string id, double idConfidence, string discipline, string title, double titleConfidence = 90)
	{
		return new IndexRow
		{
			Id = id,
			IdOcr = true,
			IdConfidence = idConfidence,
			Discipline = discipline,
			DisciplineOcr = true,
			DisciplineConfidence = 96,
			Title = title,
			TitleOcr = true,
			TitleConfidence = titleConfidence
		};
	}

	private static IndexRow Text(string id, string discipline, string title) => new IndexRow { Id = id, Discipline = discipline, Title = title };

	private static IndexRow WithCandidates(IndexRow row, params (string text, double confidence)[] candidates)
	{
		row.IdCandidates = candidates.Select(c => new ReadingCandidate(c.text, c.confidence)).ToList();
		return row;
	}

	private static IndexRow WithTitleCandidates(IndexRow row, params (string text, double confidence)[] candidates)
	{
		row.TitleCandidates = candidates.Select(c => new ReadingCandidate(c.text, c.confidence)).ToList();
		return row;
	}
	private static List<IndexRow> Analyze(params IndexRow[] rows)
	{
		List<IndexRow> list = rows.ToList();
		IndexAnalyzer.Analyze(list, OcrVocabulary.CreateDefault());
		return list;
	}

	[Fact]
	public void DisciplineAndTitleReadInEachOthersPlacesAreSwapped()
	{
		IndexRow swapped = Ocr("ST-M-02", 90, "BASINS UPPER PLAN", "MECHANICAL");
		swapped.DisciplineCandidates = new List<ReadingCandidate> { new ReadingCandidate("BASINS UPPER PLAN", 85) };
		swapped.TitleCandidates = new List<ReadingCandidate> { new ReadingCandidate("MECHANICAL", 90) };
		List<IndexRow> rows = Analyze(Text("ST-M-01", "MECHANICAL", "BASINS LOWER PLAN"), swapped, Text("ST-M-03", "MECHANICAL", "BASINS SECTIONS"));

		Assert.Equal("MECHANICAL", rows[1].Discipline);
		Assert.Equal("BASINS UPPER PLAN", rows[1].Title);
		Assert.Contains("swapped", string.Join(" ", rows[1].Reasons));
	}

	[Fact]
	public void ASheetNumberExampleFlagsNumbersOfAnotherFormat()
	{
		List<IndexRow> list = new List<IndexRow>
		{
			Ocr("HH-A-01", 95, "ARCHITECTURAL", "ROOF PLAN"),
			Ocr("HH-A-02", 95, "ARCHITECTURAL", "FLOOR PLAN"),
			Ocr("HHA03", 95, "ARCHITECTURAL", "SECTIONS")
		};
		IndexAnalyzer.Analyze(list, OcrVocabulary.CreateDefault(), "WT-A-104");

		Assert.False(list[0].IdUncertain);
		Assert.True(list[2].IdUncertain);
		Assert.Contains("format you gave", string.Join(" ", list[2].Reasons));
		Assert.Equal("A-A-0", SetProfile.Signature("HH-A-01"));
		Assert.Equal(SetProfile.Signature("HH-A-01"), SetProfile.Signature("WT-A-104"));
	}

	[Fact]
	public void ALoneZeroEndingIsRepairedWhenTheSetNeverUsesZero()
	{
		List<IndexRow> input = new List<IndexRow>();
		for (int i = 1; i <= 10; i++)
		{
			input.Add(Text($"G-{i:00}", "GENERAL", $"GENERAL SHEET {i}"));
		}
		input.Add(Ocr("HLS-A-0", 90, "ARCHITECTURAL", "HYDRATED LIME PLAN"));
		input.Add(Ocr("HLS-A-02", 95, "ARCHITECTURAL", "HYDRATED LIME SECTIONS"));
		input.Add(Ocr("HLS-A-03", 95, "ARCHITECTURAL", "HYDRATED LIME DETAILS"));
		List<IndexRow> rows = Analyze(input.ToArray());

		Assert.Equal("HLS-A-01", rows[10].Id);
	}

	[Fact]
	public void WordsAreCorrectedFromWhatTheirNeighboursPairWithElsewhere()
	{
		OcrVocabulary vocabulary = OcrVocabulary.CreateDefault();
		List<string> titles = new List<string>();
		for (int i = 0; i < 8; i++)
		{
			titles.Add("RAW WATER METER VAULT PLANS");
			titles.Add("HEADHOUSE 2ND FLOOR PLAN");
			titles.Add("WASHWATER TANK PIPING PLAN");
		}
		vocabulary.LearnFromTitles(titles);

		Assert.Equal("RAW WATER METER VAULT PLANS", vocabulary.RefineTitle("RAW WAFER METER VAULT PLANS", out _, out bool guessed));
		Assert.True(guessed);
		Assert.Equal("HEADHOUSE 2ND FLOOR PLAN", vocabulary.RefineTitle("HEADHOUSE CND FLOOR PLAN", out _));
		Assert.Equal("WASHWATER TANK PIPING PLAN", vocabulary.RefineTitle("WASHWATER ANK PIPING PLAN", out _));
		Assert.Equal("WASHWATER TANK PIPING PLAN", vocabulary.RefineTitle("WASHWATER IT TANK PIPING PLAN", out _));
	}

	[Fact]
	public void MisreadDisciplinesAreNotLearnedAsNewOnes()
	{
		OcrVocabulary vocabulary = OcrVocabulary.CreateDefault();
		vocabulary.LearnDisciplines(new Dictionary<string, int> { ["ARCHIT ECT URAL"] = 5, ["CIVIL/YARD"] = 20, ["CIVIL/Y ARD"] = 6, ["CIVIL DETAILS"] = 9 });

		Assert.True(vocabulary.ContainsDiscipline("CIVIL/YARD"));
		Assert.True(vocabulary.ContainsDiscipline("CIVIL DETAILS"));
		Assert.False(vocabulary.ContainsDiscipline("CIVIL/Y ARD"));
		Assert.False(vocabulary.ContainsDiscipline("ARCHIT ECT URAL"));
		Assert.Equal("ARCHITECTURAL", vocabulary.SnapDiscipline("ARCHIT ECT URAL"));
	}

	[Fact]
	public void ADisciplineIsInferredFromALoneLetterSeriesOfTheSet()
	{
		List<IndexRow> input = new List<IndexRow>();
		for (int i = 1; i <= 5; i++)
		{
			input.Add(Text($"Y-{i:00}", "CIVIL/YARD", $"PIPING PLAN {i}"));
		}
		IndexRow scan = Ocr("SHI-Y-01", 95, "", "LAGOON VALVE VAULT");
		scan.DisciplineConfidence = 0;
		input.Add(scan);
		input.Add(Ocr("SHI-Y-02", 95, "CIVIL/YARD", "LAGOON SLUDGE DRAIN"));
		List<IndexRow> rows = Analyze(input.ToArray());

		Assert.Equal("CIVIL/YARD", rows[5].Discipline);
	}

	[Fact]
	public void ACommaSpeckAfterAWordTheSetNeverWritesWithOneIsRemoved()
	{
		OcrVocabulary vocabulary = OcrVocabulary.CreateDefault();
		List<string> titles = Enumerable.Repeat("SECONDARY TREATMENT BASINS PLAN", 8).Concat(new[] { "SECONDARY, TREATMENT BASINS PLAN", "CHEMICAL TANK SCHEDULE, PLAN" }).ToList();
		vocabulary.LearnFromTitles(titles);

		Assert.Equal("SECONDARY TREATMENT BASINS PLAN", vocabulary.RefineTitle("SECONDARY, TREATMENT BASINS PLAN", out _));
		Assert.Equal("CHEMICAL TANK SCHEDULE, PLAN", vocabulary.RefineTitle("CHEMICAL TANK SCHEDULE, PLAN", out _));
	}

	[Fact]
	public void DisciplinesAreRepairedWordByWordLikeTitles()
	{
		OcrVocabulary vocabulary = OcrVocabulary.CreateDefault();

		Assert.Equal("CIVIL/YARD", vocabulary.RefineDiscipline("CNIL/Y ARD"));
		Assert.Equal("CIVIL DETAILS", vocabulary.RefineDiscipline("CWIL DETAILS"));
		Assert.Equal("CIVIL DETAILS", vocabulary.RefineDiscipline("CIVIL DETAILS"));
		Assert.Equal("MECHANICAL", vocabulary.RefineDiscipline("MECHANICAL"));
	}

	[Fact]
	public void ASheetThatStartsANewSeriesIsNotStretchedIntoTheOldOne()
	{
		List<IndexRow> rows = Analyze(
			Ocr("SHI-G-01", 95, "GENERAL", "SURVEY CONTROL"),
			Ocr("SHI-G-02", 95, "GENERAL", "EROSION CONTROL PLAN"),
			Ocr("SHI-D-01", 60, "", "DEMOLITION PLAN"),
			Ocr("SHI-Y-01", 95, "CIVIL/YARD", "OVERALL SITE PLAN"),
			Ocr("SHI-Y-02", 95, "CIVIL/YARD", "LOW SERVICE PUMP STATION"));

		Assert.Equal("SHI-D-01", rows[2].Id);
	}

	[Fact]
	public void SequenceFillsTheNumbersBetweenTwoConsecutiveRuns()
	{
		List<IndexRow> rows = Analyze(
			Ocr("ST-M-0", 38, "MECHANICAL", "BASINS LOWER PLAN"),
			Ocr("ST-M-02", 60, "MECHANICAL", "BASINS UPPER PLAN"),
			Ocr("ST-M-03", 68, "MECHANICAL", "BASINS ENLARGED PLANS I"),
			Ocr("ST-M-04", 66, "MECHANICAL", "BASINS ENLARGED PLANS II"),
			Ocr("ST-M-10", 14, "MECHANICAL", "BASINS NO.5 AND 6 ENLARGED PLANS I"),
			Ocr("ST-M-11", 23, "MECHANICAL", "BASINS NO.5 AND 6 ENLARGED PLANS II"),
			Ocr("ST-WH2", 53, "MECHANICAL", "BASINS NO.5 AND 6 SECTIONS"),
			Ocr("ST-WH3", 48, "MECHANICAL", "BASINS BLOWDOWN VALVE VAULT"),
			Ocr("ST-M-14", 14, "MECHANICAL", "BASINS DETAILS"),
			Ocr("ST-M-15", 9, "MECHANICAL", "BASINS PROCESS FLOW DIAGRAM"));

		Assert.Equal(new[] { "ST-M-01", "ST-M-02", "ST-M-03", "ST-M-04", "ST-M-10", "ST-M-11", "ST-M-12", "ST-M-13", "ST-M-14", "ST-M-15" }.Length, rows.Count);
		Assert.Equal("ST-M-01", rows[0].Id);
		Assert.Equal("ST-M-12", rows[6].Id);
		Assert.Equal("ST-M-13", rows[7].Id);
		Assert.All(rows, r => Assert.False(r.IdUncertain));
	}

	[Fact]
	public void SequenceExtendsFromTheLastGoodSheetWhenReadingsAreClose()
	{
		List<IndexRow> rows = Analyze(
			Ocr("ST-D-09", 74, "DEMOLITION", "BASINS DEMO PLAN"),
			Ocr("ST-D-10", 25, "DEMOLITION", "BASINS DEMO PLAN"),
			Ocr("ST-D-11", 78, "DEMOLITION", "BASINS ENLARGED PLAN"),
			Ocr("ST-DIZ", 34, "DEMOLITION", "BASINS DEMOLITION SECTIONS I"),
			Ocr("ST-DIS", 36, "DEMOLITION", "BASINS DEMOLITION SECTIONS II"),
			Ocr("ST-D-14", 7, "DEMOLITION", "BASINS DEMO SECTIONS I"),
			Ocr("ST-D-5", 31, "DEMOLITION", "BASINS DEMO SECTIONS II"),
			Ocr("ST-D-16", 18, "DEMOLITION", "BASINS DEMO SECTIONS"),
			Ocr("ST-D-1", 0, "DEMOLITION", "BASINS DEMOLITION PHOTOS"));

		Assert.Equal(new[] { "ST-D-09", "ST-D-10", "ST-D-11", "ST-D-12", "ST-D-13", "ST-D-14", "ST-D-15", "ST-D-16", "ST-D-17" }, rows.Select(r => r.Id));
	}

	[Fact]
	public void SequenceFillsAMisreadNumberBetweenTwoTrustedNeighbours()
	{
		List<IndexRow> rows = Analyze(
			Ocr("HH-A-19", 39, "ARCHITECTURAL", "ELEVATIONS I"),
			Ocr("HH-A-20", 87, "ARCHITECTURAL", "ELEVATIONS II"),
			Ocr("AHA-8", 24, "ARCHITECTURAL", "MILLWORK SECTION & DETAILS I"),
			Ocr("HH-A-22", 63, "ARCHITECTURAL", "MILLWORK SECTION & DETAILS II"),
			Ocr("HH-A-23", 54, "ARCHITECTURAL", "TRAILER FACILITIES I"),
			Ocr("HH-A-24", 68, "ARCHITECTURAL", "TRAILER FACILITIES II"));

		Assert.Equal("HH-A-21", rows[2].Id);
		Assert.False(rows[2].IdUncertain);
	}

	[Fact]
	public void ADuplicateNumberInASeriesIsCorrected()
	{
		List<IndexRow> rows = Analyze(
			Ocr("HH-P-02", 78, "PLUMBING", "BASEMENT FLOOR MOD PLAN"),
			Ocr("HH-P-05", 55, "PLUMBING", "GROUND FLOOR MOD PLAN"),
			Ocr("HH-P-04", 88, "PLUMBING", "2ND FLOOR PLAN MOD PLAN"),
			Ocr("HH-P-05", 35, "PLUMBING", "ROOF PLAN"));

		Assert.Equal(new[] { "HH-P-02", "HH-P-03", "HH-P-04", "HH-P-05" }, rows.Select(r => r.Id));
	}

	[Fact]
	public void ASheetNumberThatFitsNeitherTheNumberingNorTheSetIsFlagged()
	{
		List<IndexRow> rows = Analyze(
			Text("G-10", "GENERAL", "NOTES"),
			Ocr("RMP-A-0", 48, "ARCHITECTURAL", "FACILITY FLOOR PLAN"),
			Text("Y-01", "CIVIL", "SITE"));

		Assert.True(rows[1].IdUncertain);
	}

	[Fact]
	public void ASeriesLetterThatContradictsTheDisciplineIsCorrected()
	{
		List<IndexRow> rows = Analyze(
			Ocr("OW-M-01", 54, "MECHANICAL", "METER VAULT NO.1 PLANS"),
			Ocr("OW-M-02", 72, "MECHANICAL", "METER VAULT NO.1 SECTIONS"),
			Ocr("OW-M-03", 59, "MECHANICAL", "METER VAULT NO.2 PLANS"),
			Ocr("OW-M-04", 78, "MECHANICAL", "METER VAULT NO.2 SECTIONS"),
			Text("RW-S-03", "STRUCTURAL", "RAW WATER METER VAULT SECTIONS AND DETAILS"),
			Ocr("RW-N-01", 54, "MECHANICAL", "RAW WATER METER VAULT PLANS"),
			Ocr("RW-N-02", 50, "MECHANICAL", "RAW WATER METER VAULT SECTIONS"));

		Assert.Equal("RW-M-01", rows[5].Id);
		Assert.Equal("RW-M-02", rows[6].Id);
		Assert.True(rows[5].IdUncertain);
	}

	[Fact]
	public void ADisciplineOcrCouldNotReadIsInferredFromTheSeriesLetter()
	{
		List<IndexRow> rows = Analyze(
			Ocr("HH-M-01", 80, "MECHANICAL", "ANALYZER LAYOUT"),
			Ocr("HH-M-02", 85, "MECHANICAL", "ANALYZER PIPING"),
			Ocr("HH-M-03", 86, "MECHANICAL", "ANALYZER PIPING DETAILS"),
			Ocr("HH-M-04", 88, "", "ANALYZER SECTIONS"));

		Assert.Equal("MECHANICAL", rows[3].Discipline);
		Assert.True(rows[3].DisciplineUncertain);
	}

	[Theory]
	[InlineData("OZONATED WATER METER VAULT NO./ PLANS", "OZONATED WATER METER VAULT NO.1 PLANS")]
	[InlineData("RAPID MIX BASIN NQ.2 DEMOLITION SECTION", "RAPID MIX BASIN NO.2 DEMOLITION SECTION")]
	[InlineData("BASINS NO.I THRU 4 ENLARGED PLAN", "BASINS NO.1 THRU 4 ENLARGED PLAN")]
	[InlineData("BASINS NO.LAND 2 UPPER MECHANICAL PLANS", "BASINS NO.1 AND 2 UPPER MECHANICAL PLANS")]
	[InlineData("BASINS NO.ETHRU 4 SECTIONS II", "BASINS NO.1 THRU 4 SECTIONS II")]
	[InlineData("LIME SLURRY_T TROUGH PROFILE", "LIME SLURRY T TROUGH PROFILE")]
	[InlineData("RAPID MIX PUMP. FACILITY ROOF PLAN", "RAPID MIX PUMP FACILITY ROOF PLAN")]
	public void TitleRefinementFixesNumbersAndStrayMarks(string raw, string expected)
	{
		OcrVocabulary vocabulary = OcrVocabulary.CreateDefault();

		Assert.Equal(expected, vocabulary.RefineTitle(raw, out _));
	}

	[Fact]
	public void WordsSplitByAGapAreJoinedUsingWordsTheSetRepeats()
	{
		OcrVocabulary vocabulary = OcrVocabulary.CreateDefault();
		List<string> titles = new List<string>();
		for (int i = 0; i < 12; i++)
		{
			titles.Add("HEADHOUSE ROOF PLAN");
			titles.Add("BACKWASH PUMP STATION PLAN");
			titles.Add("SECONDARY TREATMENT BASINS DETAILS");
		}
		titles.Add("HE ADHOUSE ROOF PLAN");
		titles.Add("BACKWASH PUMP ST ATION SECTION");
		titles.Add("SECONDARY TREATMENT BASINS DET AILS");
		vocabulary.LearnFromTitles(titles);

		Assert.Equal("HEADHOUSE ROOF PLAN", vocabulary.RefineTitle("HE ADHOUSE ROOF PLAN", out bool changed));
		Assert.True(changed);
		Assert.Equal("BACKWASH PUMP STATION SECTION", vocabulary.RefineTitle("BACKWASH PUMP ST ATION SECTION", out _));
		Assert.Equal("SECONDARY TREATMENT BASINS DETAILS", vocabulary.RefineTitle("SECONDARY TREATMENT BASINS DET AILS", out _));
		Assert.Equal("BACKWASH PUMP STATION PLAN", vocabulary.RefineTitle("BACKWASH PUMP STATION PLAN", out bool unchanged));
		Assert.False(unchanged);
	}

	[Fact]
	public void AFragmentPlusTheTailOfAFrequentWordIsReplacedByTheWord()
	{
		OcrVocabulary vocabulary = OcrVocabulary.CreateDefault();
		List<string> titles = new List<string>();
		for (int i = 0; i < 12; i++)
		{
			titles.Add("HEADHOUSE ROOF PLAN");
		}
		titles.Add("FATE ADHOUSE ROOF PLAN");
		titles.Add("ALE ADHOUSE FLOOR PLAN");
		vocabulary.LearnFromTitles(titles);

		Assert.Equal("HEADHOUSE ROOF PLAN", vocabulary.RefineTitle("FATE ADHOUSE ROOF PLAN", out bool changed));
		Assert.True(changed);
		Assert.Equal("HEADHOUSE FLOOR PLAN", vocabulary.RefineTitle("ALE ADHOUSE FLOOR PLAN", out _));
	}

	[Fact]
	public void MisspelledRepeatedWordsAreCorrectedFromTheSet()
	{
		OcrVocabulary vocabulary = OcrVocabulary.CreateDefault();
		vocabulary.LearnFromTitles(Enumerable.Repeat("CLEARWELL ENLARGED PLANS", 8).Concat(new[] { "CLEARWELL FNLARGED PLANS" }));

		Assert.Equal("CLEARWELL ENLARGED PLANS", vocabulary.RefineTitle("CLEARWELL FNLARGED PLANS", out _));
	}

	[Fact]
	public void TitlesWithOddMarksOrLowConfidenceAreFlagged()
	{
		List<IndexRow> rows = Analyze(
			Text("G-01", "GENERAL", "SITE PLAN"),
			Ocr("G-02", 90, "GENERAL", "SECTIONS / DETAILS", 90),
			Ocr("G-03", 90, "GENERAL", "SITE PLAN", 40),
			Ocr("G-04", 90, "GENERAL", "SITE DETAILS", 92));

		Assert.True(rows[1].TitleUncertain);
		Assert.True(rows[2].TitleUncertain);
		Assert.False(rows[3].TitleUncertain);
	}

	[Fact]
	public void ARomanNumeralBetweenTwoThatAreTwoApartIsCorrected_AndABrokenSeriesIsFlagged()
	{
		List<IndexRow> rows = Analyze(
			Ocr("A-01", 90, "ARCHITECTURAL", "EXTERIOR DETAILS II"),
			Ocr("A-02", 90, "ARCHITECTURAL", "EXTERIOR DETAILS II"),
			Ocr("A-03", 90, "ARCHITECTURAL", "EXTERIOR DETAILS IV"),
			Ocr("A-04", 90, "ARCHITECTURAL", "FINISH SCHEDULE III"),
			Ocr("A-05", 90, "ARCHITECTURAL", "FINISH SCHEDULE I"));

		Assert.Equal("EXTERIOR DETAILS III", rows[1].Title);
		Assert.True(rows[4].TitleUncertain);
	}

	[Fact]
	public void ANewSeriesOfAnotherDisciplineIsNotAbsorbedIntoThePreviousOne()
	{
		List<IndexRow> rows = Analyze(
			Ocr("RMP-A-05", 79, "ARCHITECTURAL", "FACILITY EXTERIOR DETAILS"),
			Ocr("RMP-A-06", 79, "ARCHITECTURAL", "FACILITY SCHEDULES AND DOOR DETAILS"),
			Ocr("RMP-M-01", 42, "MECHANICAL", "FACILITY PROCESS SCHEMATIC"),
			Ocr("RMP-M-02", 91, "MECHANICAL", "FACILITY MECHANICAL PLAN"),
			Ocr("RMP-M-03", 68, "MECHANICAL", "FACILITY SECTIONS"));

		Assert.Equal("RMP-M-01", rows[2].Id);
	}

	[Fact]
	public void ADifferentDisciplineBlocksPrefixSnappingToTheNeighbours()
	{
		List<IndexRow> rows = Analyze(
			Ocr("ST-D-09", 88, "DEMOLITION", "BASINS DEMO PLAN"),
			Ocr("ST-D-10", 90, "DEMOLITION", "BASINS DEMO PLAN"),
			Ocr("ST-D-11", 90, "DEMOLITION", "BASINS DEMO PLAN"),
			Ocr("PT-G-0", 9, "GENERAL", "BASIN COMPLEX MECHANICAL PLAN"));

		Assert.StartsWith("PT-G", rows[3].Id);
		Assert.True(rows[3].IdUncertain);
	}

	[Fact]
	public void AGuessedTitleWordIsFlaggedButAPatternFixIsNot()
	{
		List<IndexRow> rows = Analyze(
			Text("A-01", "ARCHITECTURAL", "HYDRATED LIME SYSTEM EXTERIOR ELEVATIONS"),
			Text("A-02", "ARCHITECTURAL", "HYDRATED LIME SYSTEM EXTERIOR DETAILS"),
			Text("A-03", "ARCHITECTURAL", "HYDRATED LIME SYSTEM EXTERIOR SECTIONS"),
			Ocr("A-04", 90, "ARCHITECTURAL", "HYDRATED LIME SYSTEM FXTERIQR ELEVATIONS", 80),
			Ocr("A-05", 90, "ARCHITECTURAL", "HYDRATED LIME SYSTEM NO./ PLANS", 80));

		Assert.Equal("HYDRATED LIME SYSTEM EXTERIOR ELEVATIONS", rows[3].Title);
		Assert.True(rows[3].TitleUncertain);
		Assert.Equal("HYDRATED LIME SYSTEM NO.1 PLANS", rows[4].Title);
		Assert.False(rows[4].TitleUncertain, string.Join("; ", rows[4].Reasons));
	}

	[Fact]
	public void TheMisreadNoQ_IsFixed()
	{
		Assert.Equal("CLEARWELL NO.1 SECTIONS III", OcrVocabulary.CreateDefault().RefineTitle("CLEARWELL NOQ.! SECTIONS III", out _));
	}
	[Fact]
	public void TheReadingThatFitsTheSetsNumberingWinsOverAMoreConfidentOneThatDoesNot()
	{
		List<IndexRow> rows = new List<IndexRow>();
		for (int i = 1; i <= 10; i++)
		{
			rows.Add(Text("G-" + i.ToString("00"), "GENERAL", "NOTES SHEET " + i));
		}
		rows.Add(WithCandidates(Ocr("G-0", 80, "GENERAL", "NOTES SHEET"), ("G-0", 80), ("G-11", 55)));
		rows.Add(Text("G-12", "GENERAL", "NOTES SHEET 12"));

		List<IndexRow> result = Analyze(rows.ToArray());

		Assert.Equal("G-11", result[10].Id);
		Assert.False(result[10].IdUncertain);
	}

	[Fact]
	public void ASetThatNumbersFromZeroOrInSingleDigitsIsNotMarkedWrongForIt()
	{
		List<IndexRow> rows = new List<IndexRow>();
		for (int i = 0; i <= 9; i++)
		{
			rows.Add(Text("P-" + i, "PLANT", "PLAN " + i));
		}
		rows.Add(Ocr("P-10", 90, "PLANT", "PLAN 10"));
		rows.Add(Ocr("P-0", 90, "PLANT", "PLAN 0"));

		List<IndexRow> result = Analyze(rows.ToArray());

		Assert.Equal("P-10", result[10].Id);
		Assert.False(result[10].IdUncertain);
		Assert.Equal("P-0", result[11].Id);
	}

	[Fact]
	public void ASequenceCanPickAnAlternativeReadingInsteadOfInventingANumber()
	{
		List<IndexRow> rows = new List<IndexRow>();
		for (int i = 1; i <= 9; i++)
		{
			rows.Add(Text("ST-M-" + i.ToString("00"), "MECHANICAL", "PLAN " + i));
		}
		rows.Add(WithCandidates(Ocr("ST-W-1", 40, "MECHANICAL", "PLAN 10"), ("ST-W-1", 40), ("ST-M-10", 22)));
		rows.Add(Text("ST-M-11", "MECHANICAL", "PLAN 11"));

		List<IndexRow> result = Analyze(rows.ToArray());

		Assert.Equal("ST-M-10", result[9].Id);
		Assert.Equal(0, result[9].IdFixDistance);
	}

	[Fact]
	public void ARomanNumeralIsTakenFromTheReadingThatContinuesTheSeries()
	{
		List<IndexRow> rows = Analyze(
			Text("A-01", "ARCHITECTURAL", "EXTERIOR DETAILS I"),
			WithTitleCandidates(Ocr("A-02", 90, "ARCHITECTURAL", "EXTERIOR DETAILS I"), ("EXTERIOR DETAILS I", 90), ("EXTERIOR DETAILS II", 80)),
			Text("A-03", "ARCHITECTURAL", "EXTERIOR DETAILS III"));

		Assert.Equal("EXTERIOR DETAILS II", rows[1].Title);
	}

	[Fact]
	public void PunctuationTheSetItselfUsesIsNotFlagged()
	{
		List<IndexRow> rows = new List<IndexRow>();
		for (int i = 1; i <= 6; i++)
		{
			rows.Add(Text("D-" + i.ToString("00"), "GENERAL", "TAP/SAMPLE LINE " + i));
		}
		rows.Add(Ocr("D-07", 90, "GENERAL", "TAP/SAMPLE LINE 7", 90));

		List<IndexRow> result = Analyze(rows.ToArray());

		Assert.False(result[6].TitleUncertain);
	}
	[Fact]
	public void TextLayerRowsAreNeverChangedOrFlagged()
	{
		List<IndexRow> rows = Analyze(
			Text("G-01", "GENERAL", "COVER SHEET /"),
			Text("G-07", "GENERAL", "NOTES"),
			Text("G-02", "GENERAL", "PLAN"));

		Assert.Equal("G-07", rows[1].Id);
		Assert.All(rows, r => Assert.False(r.IdUncertain || r.TitleUncertain || r.DisciplineUncertain));
	}

	[Fact]
	public void AReadingThatIsOnlyPartOfAnotherDoesNotWinOnConfidenceAlone()
	{
		// Real candidates (Balanced preset): the prefix "RM" was lost by the variants that happened to score most confidently.
		List<IndexRow> rows = Analyze(
			WithCandidates(Ocr("RMB-D-0", 78, "DEMOLITION", "RAPID MIX BASINS NO.1 DEMOLITION PLAN"), ("RMB-D-01", 42), ("RMB-D-0", 78), ("RMB-D-01", 45), ("RMB-D-01", 55)),
			WithCandidates(Ocr("B-D-02", 93, "DEMOLITION", "RAPID MIX BASIN NO.1 DEMOLITION SECTION"), ("RMB-D-02", 76), ("RMB-D-02", 38), ("B-D-02", 93)),
			Text("RMB-D-03", "DEMOLITION", "RAPID MIX BASIN NO.2 DEMOLITION SECTION"),
			Text("RMB-M-03", "MECHANICAL", "RAPID MIX BASINS NO.1 AND 2 SECTIONS"),
			Text("RMB-M-04", "MECHANICAL", "RAPID MIX BASIN NO.3 MECHANICAL PLAN"),
			Text("RMB-M-05", "MECHANICAL", "RAPID MIX PIPING DETAILS"));
		Assert.Equal("RMB-D-01", IdPattern.Normalize(rows[0].Id));
		Assert.Equal("RMB-D-02", IdPattern.Normalize(rows[1].Id));
	}
	[Fact]
	public void ANewSeriesStartingHereBeatsAStretchedOldSeries()
	{
		// "HS-M-01" is what the reader saw (one alternative), and the next sheet is HS-M-02. Extending CL-M-09 to CL-M-10 would be invented.
		List<IndexRow> rows = Analyze(
			Text("CL-M-08", "MECHANICAL", "CLEARWELL NO.2 PLAN"),
			Text("CL-M-09", "MECHANICAL", "CLEARWELL NO.2 PLAN & SECTIONS"),
			WithCandidates(Ocr("HS-W-01", 37, "MECHANICAL", "HIGH SERVICE PUMP STATION PLAN"), ("HS-M-O!", 0), ("HS--O1", 13), ("HS-W-01", 22), ("HS-W-O1", 37)),
			Text("HS-M-02", "MECHANICAL", "HIGH SERVICE PS COOLING WATER SECTION"),
			Text("HS-M-03", "MECHANICAL", "HIGH SERVICE PS SECTIONS"));
		Assert.Equal("HS-M-01", IdPattern.Normalize(rows[2].Id));
	}

	[Fact]
	public void ANoIsReadAsOneWhenTheSetNumbersInDigits()
	{
		List<IndexRow> rows = Analyze(
			Text("OW-M-01", "MECHANICAL", "OZONATED WATER METER VAULT NO.1 PLANS"),
			Text("OW-M-03", "MECHANICAL", "OZONATED WATER METER VAULT NO.2 PLANS"),
			Text("OW-M-04", "MECHANICAL", "OZONATED WATER METER VAULT NO.3 PLANS"),
			Ocr("OW-M-02", 80, "MECHANICAL", "OZONATED WATER METER VAULT NO.I SECTIONS"));
		Assert.Equal("OZONATED WATER METER VAULT NO.1 SECTIONS", rows[3].Title);
	}

	[Fact]
	public void ANoIIsKeptWhenTheSetUsesRomanNumerals()
	{
		List<IndexRow> rows = Analyze(
			Text("A-01", "GENERAL", "PLANT UNIT NO.II PLAN"),
			Text("A-02", "GENERAL", "PLANT UNIT NO.III PLAN"),
			Text("A-03", "GENERAL", "PLANT UNIT NO.IV PLAN"),
			Ocr("A-04", 80, "GENERAL", "PLANT UNIT NO.I PLAN"));
		Assert.Equal("PLANT UNIT NO.I PLAN", rows[3].Title);
	}

	[Fact]
	public void StrayMarksInsideWordsAreMended()
	{
		List<IndexRow> rows = Analyze(
			Text("G-01", "GENERAL", "OVERALL SITE PLAN"),
			Text("G-02", "GENERAL", "ENLARGED SITE PLAN"),
			Ocr("G-03", 80, "GENERAL", "OVERALL S/TE PLAN"),
			Ocr("G-04", 80, "GENERAL", "RAPID MIX. BASIN PLAN"));
		Assert.Equal("OVERALL SITE PLAN", rows[2].Title);
	}

	[Fact]
	public void APrefixOtherSheetsShareBeatsOneSeenOnThisSheetAlone()
	{
		List<IndexRow> rows = Analyze(
			WithCandidates(Ocr("BW-M-01", 55, "MECHANICAL", "WASHWATER TANK PLAN"), ("BW-M-01", 55), ("BW-M-01", 40)),
			WithCandidates(Ocr("BW-M-02", 60, "MECHANICAL", "WASHWATER TANK SECTION"), ("BW-M-02", 60), ("BW-M-02", 50)),
			WithCandidates(Ocr("BN-M-03", 62, "MECHANICAL", "WASHWATER TANK DETAILS"), ("BN-M-03", 62), ("BW-M-03", 41)),
			WithCandidates(Ocr("BW-M-04", 58, "MECHANICAL", "WASHWATER TANK PIPING"), ("BW-M-04", 58), ("BW-M-04", 44)));
		Assert.Equal("BW-M-03", IdPattern.Normalize(rows[2].Id));
	}

	[Fact]
	public void AMisreadNumeralIsTakenFromTheSeriesAroundIt()
	{
		List<IndexRow> rows = Analyze(
			Ocr("HLS-S-03", 70, "STRUCTURAL", "HYDRATED LIME SYSTEM SECTIONS AND DETAILS II"),
			Ocr("HLS-S-04", 70, "STRUCTURAL", "HYDRATED LIME SYSTEM SECTIONS AND DETAILS III"),
			Ocr("HLS-S-05", 70, "STRUCTURAL", "HYDRATED LIME SYSTEM SECTIONS AND DETAILS WV"),
			Ocr("HLS-S-06", 70, "STRUCTURAL", "HYDRATED LIME SYSTEM SECTIONS AND DETAILS V"));
		Assert.Equal("HYDRATED LIME SYSTEM SECTIONS AND DETAILS IV", rows[2].Title);
	}

	[Fact]
	public void AWordAtTheEndOfATitleIsNotMistakenForANumeral()
	{
		List<IndexRow> rows = Analyze(
			Ocr("A-01", 70, "GENERAL", "CLEARWELL PLAN I"),
			Ocr("A-02", 70, "GENERAL", "CLEARWELL PLAN II"),
			Ocr("A-03", 70, "GENERAL", "CLEARWELL PLAN ENLARGED"));
		Assert.Equal("CLEARWELL PLAN ENLARGED", rows[2].Title);
	}

	[Fact]
	public void ABlankPageIsLeftOutOfTheNumberingAndFlagged()
	{
		IndexRow blank = new IndexRow { Blank = true, IdOcr = false };
		List<IndexRow> rows = Analyze(
			Text("SHI-Y-01", "CIVIL", "LIME PLAN"),
			blank,
			Text("SHI-Y-02", "CIVIL", "LIME PROFILE"));
		Assert.Equal(string.Empty, rows[1].Id);
		Assert.True(rows[1].TitleUncertain);
		Assert.Contains("blank", string.Join(" ", rows[1].Reasons));
	}

	[Fact]
	public void AWordMostSheetsReadSplitIsJoinedWhenSeveralSheetsReadItWhole()
	{
		// The reader split "HEADHOUSE" as "HE ADHOUSE" on most sheets; a few readings of other sheets had it whole.
		List<IndexRow> rows = new List<IndexRow>();
		for (int i = 1; i <= 8; i++)
		{
			IndexRow row = Ocr($"HH-A-0{i}", 60, "ARCHITECTURAL", "HE ADHOUSE ROOF PLAN");
			row.TitleCandidates = i <= 4
				? new List<ReadingCandidate> { new ReadingCandidate("HE ADHOUSE ROOF PLAN", 85), new ReadingCandidate("HEADHOUSE ROOF PLAN", 70) }
				: new List<ReadingCandidate> { new ReadingCandidate("HE ADHOUSE ROOF PLAN", 85) };
			rows.Add(row);
		}
		List<IndexRow> analyzed = Analyze(rows.ToArray());
		Assert.All(analyzed, r => Assert.Equal("HEADHOUSE ROOF PLAN", r.Title));
	}

	[Fact]
	public void LookAlikeLettersAreCorrectedAgainstTheWordList()
	{
		OcrVocabulary vocabulary = OcrVocabulary.CreateDefault();
		Assert.True(vocabulary.Dictionary.Count > 100000, "the word list was not found next to the tests");

		// One confusable letter: corrected with confidence. Two: corrected as a guess.
		Assert.Equal("OZONE BUILDING PLAN", vocabulary.RefineTitle("QZONE BUILDING PLAN", out bool changedOne, out bool guessedOne));
		Assert.True(changedOne);
		Assert.False(guessedOne);
		Assert.Equal("HYDRATED LIME SYSTEM", vocabulary.RefineTitle("AYDRAIED LIME SYSTEM", out _, out bool guessedTwo));
		Assert.True(guessedTwo);
	}

	[Fact]
	public void RealWordsAndRecurringNamesAreLeftAlone()
	{
		OcrVocabulary vocabulary = OcrVocabulary.CreateDefault();
		vocabulary.LearnFromTitles(Enumerable.Repeat("BACHMAN WATER TREATMENT PLANT", 6));

		Assert.Equal("BACHMAN WATER TREATMENT PLANT", vocabulary.RefineTitle("BACHMAN WATER TREATMENT PLANT", out bool changed));
		Assert.False(changed);
		Assert.Equal("CLEARWELL ENLARGED PLAN", vocabulary.RefineTitle("CLEARWELL ENLARGED PLAN", out _));
	}

	[Fact]
	public void ARepeatedMisreadingIsNotTrustedJustBecauseItRepeats()
	{
		List<IndexRow> rows = new List<IndexRow>();
		for (int i = 1; i <= 6; i++)
		{
			rows.Add(Ocr($"OZ-A-0{i}", 70, "ARCHITECTURAL", $"QZONE BUILDING PLAN {new string('I', i <= 3 ? i : 1)}"));
		}
		rows.Add(Text("OZ-A-07", "ARCHITECTURAL", "OZONE BUILDING SECTIONS"));
		List<IndexRow> analyzed = Analyze(rows.ToArray());
		Assert.All(analyzed.Take(6), r => Assert.True(r.Title.StartsWith("OZONE BUILDING PLAN"), r.Title + " :: " + string.Join("; ", r.Reasons)));
	}

	[Fact]
	public void ANumberAfterNoReadAsALetterIsTakenFromTheNeighbours()
	{
		List<IndexRow> rows = Analyze(
			Ocr("B-01", 80, "MECHANICAL", "RAPID MIX BASIN NO.1 PLAN"),
			Ocr("B-02", 80, "MECHANICAL", "RAPID MIX BASIN NO.W PLAN"),
			Ocr("B-03", 80, "MECHANICAL", "RAPID MIX BASIN NO.3 PLAN"));
		Assert.Equal("RAPID MIX BASIN NO.2 PLAN", rows[1].Title);
		Assert.False(rows[1].TitleUncertain, rows[1].Title + " :: " + string.Join("; ", rows[1].Reasons));
	}

	[Fact]
	public void ANoWithNoNumberAndOtherBrokenConventionsAreFlagged()
	{
		List<IndexRow> rows = Analyze(
			Text("C-01", "CIVIL", "SITE PLAN"),
			Ocr("C-02", 90, "CIVIL", "CLEARWELL NO. PLAN", 95),
			Ocr("C-03", 90, "CIVIL", "CLEARWELL K PLAN", 95),
			Ocr("C-04", 90, "CIVIL", "CLEARWELL PL4NT SECTIONS", 95),
			Ocr("C-05", 90, "CIVIL", "SITE PLANN", 95));
		Assert.Contains("number", string.Join(" ", rows[1].Reasons));
		Assert.Contains("stray single letter", string.Join(" ", rows[2].Reasons));
		Assert.True(rows[3].TitleUncertain, rows[3].Title + " :: " + string.Join("; ", rows[3].Reasons));
		Assert.True(rows[4].TitleUncertain, rows[4].Title + " :: " + string.Join("; ", rows[4].Reasons));
	}

	[Fact]
	public void ANumberLostBeforeThruIsTakenFromTheOtherSheetsWithTheSameWording()
	{
		List<IndexRow> rows = Analyze(
			Ocr("ST-D-01", 80, "DEMOLITION", "SECONDARY TREATMENT BASINS NO.1 THRU 4 DEMO PLAN"),
			Ocr("ST-D-02", 80, "DEMOLITION", "SECONDARY TREATMENT BASINS NO.1 THRU 4 UPPER PLAN"),
			Ocr("ST-D-03", 80, "DEMOLITION", "SECONDARY TREATMENT BASINS NO. THRU 4 ENLARGED PLAN"),
			Ocr("ST-D-04", 80, "DEMOLITION", "SECONDARY TREATMENT BASINS NO.J THRU 4 SECTIONS"));
		Assert.Equal("SECONDARY TREATMENT BASINS NO.1 THRU 4 ENLARGED PLAN", rows[2].Title);
		Assert.Equal("SECONDARY TREATMENT BASINS NO.1 THRU 4 SECTIONS", rows[3].Title);
	}

	[Fact]
	public void RomanNumeralsAndRepeatedWordsAreJudgedSeparately()
	{
		List<IndexRow> rows = Analyze(
			Text("Y-01", "CIVIL", "SITE PLAN AREA V"),
			Ocr("Y-02", 90, "CIVIL", "SITE PLAN AREA V", 95),
			Ocr("Y-03", 90, "CIVIL", "OZONE BUILDING BUILDING DETAILS", 95));
		Assert.DoesNotContain("stray single letter", string.Join(" ", rows[1].Reasons));
		Assert.Contains("repeated", string.Join(" ", rows[2].Reasons));
	}
}