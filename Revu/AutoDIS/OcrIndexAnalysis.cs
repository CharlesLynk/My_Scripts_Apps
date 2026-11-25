using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace MbrohDrawingIndexCreator
{
	// One line of the assembled index, as the analysis sees it.
	internal sealed class IndexRow
	{
		internal string Label = string.Empty;
		internal string Id = string.Empty;
		internal string Discipline = string.Empty;
		internal string Title = string.Empty;
		internal bool IdOcr;
		// A blank page: left out of the numbering and pattern checks, so nothing is invented for it.
		internal bool Blank;
		internal bool DisciplineOcr;
		internal bool TitleOcr;
		internal double IdConfidence = 100;
		internal double DisciplineConfidence = 100;
		internal double TitleConfidence = 100;
		internal bool IdUncertain;
		internal bool DisciplineUncertain;
		internal bool TitleUncertain;
		internal List<string> Reasons = new List<string>();
		internal bool IdChanged;
		internal bool TitleChanged;
		// True when a title word was replaced by a similar-looking one (a guess), as opposed to a pattern fix such as "NO./" -> "NO.1".
		internal bool TitleGuessed;
		internal bool DisciplineChanged;
		internal IdStatus IdState = IdStatus.Unchecked;
		// Every reading the reader produced for an OCR-read field (raw text). The analysis cleans them, scores each against the set's
		// own patterns and keeps the best fit. Empty for fields read from the text layer.
		internal List<ReadingCandidate> IdCandidates = new List<ReadingCandidate>();
		internal List<ReadingCandidate> DisciplineCandidates = new List<ReadingCandidate>();
		internal List<ReadingCandidate> TitleCandidates = new List<ReadingCandidate>();
		// The cleaned alternatives, kept so a later check (numbering, numeral series) can switch to one that fits.
		internal List<string> IdOptions = new List<string>();
		internal List<(string title, double confidence)> TitleOptions = new List<(string, double)>();
		// How far OCR's own reading was from the number the sequence put in its place (4 = blank or unrecognisable).
		internal int IdFixDistance;
	}

	internal enum IdStatus
	{
		Unchecked,
		// Read directly (text layer), or confident, or numbered consecutively with its neighbours.
		Anchor,
		// Filled in because the sheets on both sides number consecutively across it.
		InferredBothSides,
		// Filled in from one neighbour and close to what OCR read.
		InferredOneSide,
		Unresolved
	}

	internal sealed class IndexAnalysisSummary
	{
		internal int IdsRepaired;
		internal int TitlesRefined;
		internal int DisciplinesFilled;
		internal int FlaggedRows;
	}

	internal static class IdPattern
	{
		private static readonly Regex Numbered = new Regex(@"^(?<prefix>(?:.*-)?[A-Z]*?)(?<number>\d{1,4})$", RegexOptions.Compiled);

		internal static string Normalize(string id) => (id ?? string.Empty).Replace('\u2013', '-').Replace('\u2014', '-').Trim();

		internal static bool TryParse(string id, out string prefix, out int number, out int width)
		{
			Match m = Numbered.Match(Normalize(id));
			if (!m.Success || !int.TryParse(m.Groups["number"].Value, out number))
			{
				prefix = string.Empty;
				number = 0;
				width = 0;
				return false;
			}
			prefix = m.Groups["prefix"].Value;
			width = m.Groups["number"].Value.Length;
			return true;
		}

		internal static string Format(string prefix, int number, int width) => prefix + number.ToString().PadLeft(width, '0');

		// Distance ignoring hyphens, so "STWH2" and "ST-M-12" compare as the letters and digits they are.
		internal static int Distance(string a, string b, int limit)
		{
			return OcrTextCleanup.EditDistance(Regex.Replace(Normalize(a).ToUpperInvariant(), "[^A-Z0-9]", string.Empty), Regex.Replace(Normalize(b).ToUpperInvariant(), "[^A-Z0-9]", string.Empty), limit);
		}

		// The discipline letter of a series prefix: the last letter group ("ST-M-" -> "M", "RW-S-" -> "S").
		internal static string? TailLetter(string prefix)
		{
			Match m = Regex.Match(prefix, @"([A-Z]+)-?$");
			if (!m.Success)
			{
				return null;
			}
			string[] segments = prefix.TrimEnd('-').Split('-');
			// A lone one-letter series ("Y-01") is as good a discipline letter as the last group of "ST-M-01".
				return segments.Length >= 2 || segments[0].Length == 1 ? segments[^1] : null;
		}
	}

	// A first read of the set is never perfect, but the set itself is full of patterns: sheet numbers run in sequence within a
	// series, series share a discipline, titles repeat words, and numerals count up. This pass uses those patterns to repair
	// what is clearly wrong and to flag what does not fit. Pure logic, no PDF or OCR involved.
	internal static class IndexAnalyzer
	{
		private const double ConfidentIdLimit = 85.0;

		private const double LowConfidenceIdLimit = 60.0;

		internal static IndexAnalysisSummary Analyze(IList<IndexRow> allRows, OcrVocabulary vocabulary, string? idExample = null)
		{
			IndexAnalysisSummary summary = new IndexAnalysisSummary();
			// Blank pages take no part: the sheets either side of one number consecutively across it.
			foreach (IndexRow blank in allRows.Where(r => r.Blank))
			{
				blank.TitleUncertain = true;
				blank.Reasons.Add("the page looks blank");
			}
			List<IndexRow> rows = allRows.Where(r => !r.Blank).ToList();
			// 1. What does this set look like? (embedded text and confident OCR)  2. Of each OCR field's readings, keep the best fit.
			SetProfile profile = SetProfile.Build(rows);
			profile.UseExample(idExample);
			vocabulary.LearnFromTitles(rows.Select(r => r.Title));
			vocabulary.LearnReadings(rows.Select(r => r.TitleOcr && r.TitleCandidates.Count > 0 ? r.TitleCandidates.Select(c => OcrTextCleanup.CleanTitle(c.Text)) : new[] { r.Title }));
			ChooseReadings(rows, vocabulary, profile);
			FixMisplacedFields(rows, vocabulary);
			profile = SetProfile.Build(rows);
			profile.UseExample(idExample);
			RefineTitles(rows, vocabulary, summary);
			ResolveIds(rows, profile, summary);
			RepairNumbersAfterNo(rows, vocabulary);
			CheckRomanNumerals(rows, vocabulary);
			CheckDisciplines(rows, vocabulary, summary);
			FlagTitles(rows, vocabulary, profile);
			FlagRepeatedTitles(rows);
			FlagIdsOffFormat(rows, profile);
			summary.FlaggedRows = allRows.Count(r => r.IdUncertain || r.DisciplineUncertain || r.TitleUncertain);
			return summary;
		}

		// With a sheet number example from the user, an OCR-read number of a different format is flagged.
		private static void FlagIdsOffFormat(IList<IndexRow> rows, SetProfile profile)
		{
			if (!profile.HasExample)
			{
				return;
			}
			foreach (IndexRow row in rows.Where(r => r.IdOcr && !profile.MatchesExample(r.Id)))
			{
				row.IdUncertain = true;
				row.Reasons.Add("sheet number does not match the format you gave");
			}
		}

		// Fields whose areas overlap, or sheets drawn with the discipline and the title in the other places, read each other's text. A title
		// that is exactly a discipline the set uses is not a title: the discipline is taken from it, and the title from the next best
		// reading of either field that is not a discipline.
		private static void FixMisplacedFields(IList<IndexRow> rows, OcrVocabulary vocabulary)
		{
			foreach (IndexRow row in rows.Where(r => r.DisciplineOcr && r.TitleOcr))
			{
				string titleKey = OcrTextCleanup.CleanDiscipline(row.Title);
				if (titleKey.Length == 0 || !vocabulary.ContainsDiscipline(titleKey))
				{
					continue;
				}
				string oldDiscipline = row.Discipline;
				bool disciplineKnown = oldDiscipline.Length > 0 && vocabulary.ContainsDiscipline(oldDiscipline);
				if (string.Equals(oldDiscipline, titleKey, StringComparison.Ordinal))
				{
					disciplineKnown = true;
				}
				List<(string title, double confidence)> alternatives = row.TitleOptions
					.Concat(row.DisciplineCandidates.Select(c => (title: vocabulary.RefineTitle(vocabulary.SnapTitle(OcrTextCleanup.CleanTitle(c.Text)), out _), confidence: c.Confidence)))
					.Where(o => o.title.Length >= 4 && o.title.Any(char.IsLetter) && !vocabulary.ContainsDiscipline(OcrTextCleanup.CleanDiscipline(o.title)))
					.OrderByDescending(o => o.confidence).ToList();
				if (!disciplineKnown)
				{
					row.Discipline = titleKey;
					row.DisciplineConfidence = row.TitleConfidence;
					row.DisciplineChanged = true;
				}
				if (alternatives.Count > 0)
				{
					row.Title = alternatives[0].title;
					row.TitleConfidence = alternatives[0].confidence;
					row.TitleChanged = true;
					row.Reasons.Add(disciplineKnown ? "the title area held the discipline; the title was taken from another reading" : "discipline and title were in each other's places; swapped");
				}
				else
				{
					row.Title = !disciplineKnown && oldDiscipline.Length > 0 ? OcrTextCleanup.CleanTitle(oldDiscipline) : row.Title;
					row.TitleUncertain = true;
					row.Reasons.Add("the title reads like a discipline name; no other title text was found");
				}
			}
		}

		// Combines the reader's confidence with how well each reading fits the set: a sheet number that matches the set's shapes and
		// series, a discipline the set uses, a title made of words the set uses. A confident reading that does not fit loses to a
		// slightly less confident one that does.
		private static void ChooseReadings(IList<IndexRow> rows, OcrVocabulary vocabulary, SetProfile profile)
		{
			Dictionary<string, int> prefixRows = CountPrefixRows(rows, profile);
			foreach (IndexRow row in rows)
			{
				if (row.IdOcr && row.IdCandidates.Count > 0)
				{
					// Each distinct reading, with how many independent passes produced it.
					List<(string text, double confidence, int support)> options = row.IdCandidates.Select(c => (text: OcrTextCleanup.CleanSheetId(c.Text, profile.ExpectHyphen), confidence: c.Confidence))
						.GroupBy(o => o.text).Select(g => (text: g.Key, confidence: g.Max(o => o.confidence), support: g.Count())).ToList();
					(string text, double confidence, int support) best = options.OrderByDescending(o => ScoreIdReading(o, options, profile, prefixRows)).ThenByDescending(o => o.confidence).First();
					row.Id = best.text;
					row.IdConfidence = best.confidence;
					row.IdOptions = options.Select(o => o.text).Where(o => o.Length > 0).ToList();
				}
				if (row.DisciplineOcr && row.DisciplineCandidates.Count > 0)
				{
					List<(string text, double confidence)> options = row.DisciplineCandidates.Select(c => (text: vocabulary.SnapDiscipline(vocabulary.RefineDiscipline(OcrTextCleanup.CleanDiscipline(c.Text))), confidence: c.Confidence))
						.GroupBy(o => o.text).Select(g => (text: g.Key, confidence: g.Max(o => o.confidence))).ToList();
					(string text, double confidence) best = options.OrderByDescending(o => 0.5 * (o.confidence / 100.0) + 0.5 * (vocabulary.ContainsDiscipline(o.text) ? 1.0 : 0.0) - (o.text.Length == 0 ? 1.0 : 0.0)).ThenByDescending(o => o.confidence).First();
					row.Discipline = best.confidence < 35.0 && !vocabulary.ContainsDiscipline(best.text) ? string.Empty : best.text;
					row.DisciplineConfidence = best.confidence;
				}
				if (row.TitleOcr && row.TitleCandidates.Count > 0)
				{
					List<(string text, double confidence)> options = row.TitleCandidates
						.Select(c => (text: vocabulary.RefineTitle(vocabulary.SnapTitle(OcrTextCleanup.CleanTitle(c.Text)), out _), confidence: c.Confidence))
						.GroupBy(o => o.text).Select(g => (text: g.Key, confidence: g.Max(o => o.confidence))).ToList();
					(string text, double confidence) best = options.OrderByDescending(o => 0.5 * (o.confidence / 100.0) + 0.5 * TitleFit(o.text, vocabulary, profile) - (o.text.Length == 0 ? 1.0 : 0.0)).ThenByDescending(o => o.confidence).First();
					row.Title = best.text;
					row.TitleConfidence = best.confidence;
					row.TitleOptions = options.Where(o => o.text.Length > 0).ToList();
				}
			}
		}

		// Confidence and fit to the set count most, plus a little for every extra pass that read the same text. A reading that is only part
		// of another one (a series prefix with its first letters lost, "RW-M-0" for "RW-M-01") is scored down: readers drop characters at
		// the edges far more often than they invent them, and a short reading tends to be more confident simply because it is short.
		private static double ScoreIdReading((string text, double confidence, int support) option, List<(string text, double confidence, int support)> all, SetProfile profile, Dictionary<string, int> prefixRows)
		{
			if (string.IsNullOrWhiteSpace(option.text))
			{
				return -1.0;
			}
			double score = 0.35 * (option.confidence / 100.0) + 0.40 * profile.IdFit(option.text) + 0.10 * Math.Min(option.support - 1, 3);
			// A series prefix that other sheets also show is far likelier than one that appears on this sheet alone (a misread is
			// idiosyncratic; a series has several sheets). This sheet itself is not counted.
			if (IdPattern.TryParse(option.text, out string prefix, out _, out _) && prefixRows.TryGetValue(prefix.ToUpperInvariant(), out int sheets))
			{
				score += 0.20 * Math.Min(Math.Max(0, sheets - 1), 4) / 4.0;
			}
			string squeezed = Squeeze(option.text);
			foreach ((string text, double confidence, int support) other in all)
			{
				string otherSqueezed = Squeeze(other.text);
				if (otherSqueezed.Length > squeezed.Length && squeezed.Length >= 2 && IdPattern.TryParse(other.text, out _, out _, out _)
					&& (otherSqueezed.EndsWith(squeezed, StringComparison.Ordinal) || otherSqueezed.StartsWith(squeezed, StringComparison.Ordinal)))
				{
					score -= 0.30;
					break;
				}
			}
			return score;
		}

		// For each series prefix, how many sheets show it in at least one of their readings (text-layer sheets count double).
		private static Dictionary<string, int> CountPrefixRows(IList<IndexRow> rows, SetProfile profile)
		{
			Dictionary<string, int> counts = new Dictionary<string, int>(StringComparer.Ordinal);
			foreach (IndexRow row in rows)
			{
				IEnumerable<string> texts = row.IdOcr && row.IdCandidates.Count > 0 ? row.IdCandidates.Select(c => OcrTextCleanup.CleanSheetId(c.Text, profile.ExpectHyphen)) : new[] { row.Id };
				int weight = row.IdOcr ? 1 : 2;
				foreach (string prefix in texts.Select(t => IdPattern.TryParse(t, out string pf, out _, out _) ? pf.ToUpperInvariant() : string.Empty).Where(pf => pf.Length > 0).Distinct())
				{
					counts[prefix] = counts.GetValueOrDefault(prefix) + weight;
				}
			}
			return counts;
		}

		private static string Squeeze(string id) => Regex.Replace(IdPattern.Normalize(id).ToUpperInvariant(), "[^A-Z0-9]", string.Empty);

		// Share of the title's longer words that the set (or the built-in list) knows, less a penalty for each mark the set never uses.
		private static double TitleFit(string title, OcrVocabulary vocabulary, SetProfile profile)
		{
			string[] words = Regex.Matches(title.ToUpperInvariant(), "[A-Z]{4,}").Select(m => m.Value).ToArray();
			double known = words.Length == 0 ? 0.5 : (double)words.Count(w => vocabulary.Dictionary.Count > 0 ? vocabulary.IsKnownForFit(w) : (vocabulary.IsKnownWord(w) || vocabulary.ObservedCount(w) >= 2)) / words.Length;
			int odd = title.Count(c => !profile.AllowsTitleCharacter(c));
			return Math.Max(0.0, known - 0.2 * odd);
		}
		private static void RefineTitles(IList<IndexRow> rows, OcrVocabulary vocabulary, IndexAnalysisSummary summary)
		{
			vocabulary.LearnFromTitles(rows.Select(r => r.Title));
			foreach (IndexRow row in rows.Where(r => r.TitleOcr))
			{
				string refined = vocabulary.RefineTitle(row.Title, out bool changed, out bool guessed);
				if (changed)
				{
					row.Title = refined;
					row.TitleChanged = true;
					row.TitleGuessed = guessed;
					summary.TitlesRefined++;
				}
			}
		}

		// ---- Sheet numbers -------------------------------------------------------------------------------------------

		private sealed class Series
		{
			internal string Prefix = string.Empty;
			internal int Width;
		}

		private static void ResolveIds(IList<IndexRow> rows, SetProfile profile, IndexAnalysisSummary summary)
		{
			int n = rows.Count;
			string[] ids = rows.Select(r => IdPattern.Normalize(r.Id)).ToArray();
			(string prefix, int number, int width)?[] parsed = new (string, int, int)?[n];
			for (int i = 0; i < n; i++)
			{
				parsed[i] = IdPattern.TryParse(ids[i], out string p, out int num, out int w) ? (p, num, w) : null;
			}
			bool[] anchor = new bool[n];
			for (int i = 0; i < n; i++)
			{
				if (!rows[i].IdOcr)
				{
					anchor[i] = true;
				}
				else if (parsed[i].HasValue && rows[i].IdConfidence >= ConfidentIdLimit && profile.PrefixFits(parsed[i]!.Value.prefix) && !(profile.ZeroNumbersUnexpected && parsed[i]!.Value.number == 0 && parsed[i]!.Value.width == 1))
				{
					anchor[i] = true;
				}
				if (i + 1 < n && parsed[i].HasValue && parsed[i + 1].HasValue && profile.PrefixFits(parsed[i]!.Value.prefix) && parsed[i]!.Value.prefix == parsed[i + 1]!.Value.prefix
					&& parsed[i + 1]!.Value.number == parsed[i]!.Value.number + 1)
				{
					anchor[i] = true;
					anchor[i + 1] = true;
				}
			}
			Dictionary<string, Dictionary<string, int>> letterByDiscipline = LearnDisciplineLetters(rows, parsed, anchor);

			int start = 0;
			while (start < n)
			{
				if (anchor[start])
				{
					rows[start].IdState = IdStatus.Anchor;
					start++;
					continue;
				}
				int end = start;
				while (end + 1 < n && !anchor[end + 1])
				{
					end++;
				}
				ResolveRun(rows, ids, parsed, anchor, start, end, summary);
				start = end + 1;
			}
			FixSeriesLetters(rows, ids, parsed, letterByDiscipline, profile, summary);
			for (int i = 0; i < n; i++)
			{
				if (rows[i].IdOcr && rows[i].IdState == IdStatus.Unresolved)
				{
					FixPrefixFromNeighbours(rows, ids, parsed, anchor, i, letterByDiscipline, summary);
				}
			}
			for (int i = 0; i < n; i++)
			{
				rows[i].Id = ids[i] == IdPattern.Normalize(rows[i].Id) ? rows[i].Id : ids[i];
				FlagId(rows[i]);
			}
		}

		// The discipline of a row when it can be relied on (read from the text layer, or confidently by OCR); otherwise null.
		private static string? TrustedDiscipline(IndexRow row)
		{
			if (string.IsNullOrWhiteSpace(row.Discipline) || (row.DisciplineOcr && row.DisciplineConfidence < 70.0))
			{
				return null;
			}
			return row.Discipline.ToUpperInvariant();
		}

		// A sheet of a different (confidently read) discipline than the series it would be filled into starts a series of its own.
		private static bool DisciplineAllows(IndexRow row, string? seriesDiscipline)
		{
			string? own = TrustedDiscipline(row);
			return own == null || seriesDiscipline == null || own == seriesDiscipline;
		}

		private static void ResolveRun(IList<IndexRow> rows, string[] ids, (string prefix, int number, int width)?[] parsed, bool[] anchor, int start, int end, IndexAnalysisSummary summary)
		{
			int n = rows.Count;
			string? leftDiscipline = start > 0 ? TrustedDiscipline(rows[start - 1]) : null;
			string? rightDiscipline = end + 1 < n ? TrustedDiscipline(rows[end + 1]) : null;
			(string prefix, int number, int width)? left = start > 0 && parsed[start - 1].HasValue ? parsed[start - 1] : null;
			(string prefix, int number, int width)? right = end + 1 < n && parsed[end + 1].HasValue ? parsed[end + 1] : null;
			if (left.HasValue && right.HasValue && left.Value.prefix == right.Value.prefix && right.Value.number - left.Value.number == end + 1 - (start - 1)
				&& Enumerable.Range(start, end - start + 1).All(k => DisciplineAllows(rows[k], leftDiscipline ?? rightDiscipline)))
			{
				for (int k = start; k <= end; k++)
				{
					string expected = IdPattern.Format(left.Value.prefix, left.Value.number + (k - (start - 1)), Math.Max(left.Value.width, right.Value.width));
					Apply(rows, ids, parsed, k, expected, IdStatus.InferredBothSides, summary);
				}
				return;
			}
			// Each neighbour offers a continuation of its own numbering for the unread sheets next to it. Where the two offers meet on the same
			// sheet, the one closer to what the reader actually saw wins (a new series starting right here beats a stretched old one).
			Dictionary<int, (string expected, int distance)> fromLeft = new Dictionary<int, (string, int)>();
			if (left.HasValue)
			{
				for (int k = start; k <= end; k++)
				{
					string expected = IdPattern.Format(left.Value.prefix, left.Value.number + (k - (start - 1)), left.Value.width);
					if (StartsOwnSeries(rows[k], parsed[k], left.Value.prefix) || !DisciplineAllows(rows[k], leftDiscipline) || !CloseEnough(rows[k], ids[k], expected))
					{
						break;
					}
					fromLeft[k] = (expected, string.IsNullOrWhiteSpace(ids[k]) ? 4 : MinDistance(rows[k], ids[k], expected));
				}
			}
			Dictionary<int, (string expected, int distance)> fromRight = new Dictionary<int, (string, int)>();
			if (right.HasValue)
			{
				for (int k = end; k >= start; k--)
				{
					int number = right.Value.number - (end + 1 - k);
					if (number < 1)
					{
						break;
					}
					string expected = IdPattern.Format(right.Value.prefix, number, right.Value.width);
					if (StartsOwnSeries(rows[k], parsed[k], right.Value.prefix) || !DisciplineAllows(rows[k], rightDiscipline) || !CloseEnough(rows[k], ids[k], expected))
					{
						break;
					}
					fromRight[k] = (expected, string.IsNullOrWhiteSpace(ids[k]) ? 4 : MinDistance(rows[k], ids[k], expected));
				}
			}
			for (int k = start; k <= end; k++)
			{
				bool hasLeft = fromLeft.TryGetValue(k, out (string expected, int distance) l);
				bool hasRight = fromRight.TryGetValue(k, out (string expected, int distance) r);
				if (hasLeft && (!hasRight || l.distance <= r.distance))
				{
					Apply(rows, ids, parsed, k, l.expected, IdStatus.InferredOneSide, summary);
				}
				else if (hasRight)
				{
					Apply(rows, ids, parsed, k, r.expected, IdStatus.InferredOneSide, summary);
				}
			}			for (int k = start; k <= end; k++)
			{
				if (rows[k].IdState == IdStatus.Unchecked)
				{
					rows[k].IdState = IdStatus.Unresolved;
				}
			}
		}

		// A reading that is the first sheet of a series (number 1) with a different prefix than the neighbour offering to continue its own
		// numbering is a new series starting here (SHI-D-01 between SHI-G-02 and SHI-Y-01), not a misread continuation: the reading stays.
		private static bool StartsOwnSeries(IndexRow row, (string prefix, int number, int width)? own, string offeredPrefix)
		{
			return own.HasValue && own.Value.number == 1 && row.IdConfidence >= 40.0 && !string.Equals(own.Value.prefix, offeredPrefix, StringComparison.Ordinal);
		}

		// The closest any of the readings of this sheet number comes to the value the sequence predicts (0 when one of the reader's
		// alternatives is exactly it: then the sequence only chooses between readings, it invents nothing).
		private static int MinDistance(IndexRow row, string current, string expected)
		{
			int best = IdPattern.Distance(current, expected, 3);
			foreach (string option in row.IdOptions)
			{
				best = Math.Min(best, IdPattern.Distance(option, expected, 3));
			}
			return best;
		}

		// OCR's own reading has to be near what the sequence predicts (a blank or lone "0" is always replaceable).
		private static bool CloseEnough(IndexRow row, string read, string expected)
		{
			if (string.IsNullOrWhiteSpace(read))
			{
				return true;
			}
			int distance = MinDistance(row, read, expected);
			// Something with no digits at all is not a sheet number, so it may differ more from the prediction.
			return distance <= 2 || ((row.IdConfidence < 40.0 || !Regex.IsMatch(read, @"\d")) && distance <= 3);
		}

		private static void Apply(IList<IndexRow> rows, string[] ids, (string prefix, int number, int width)?[] parsed, int k, string expected, IdStatus status, IndexAnalysisSummary summary)
		{
			if (!string.Equals(ids[k], expected, StringComparison.Ordinal))
			{
				rows[k].IdFixDistance = string.IsNullOrWhiteSpace(ids[k]) ? 4 : MinDistance(rows[k], ids[k], expected);
				rows[k].IdChanged = true;
				summary.IdsRepaired++;
				rows[k].Reasons.Add($"sheet number {(string.IsNullOrWhiteSpace(ids[k]) ? "(blank)" : ids[k])} -> {expected} from the numbering of its neighbours");
			}
			ids[k] = expected;
			parsed[k] = IdPattern.TryParse(expected, out string p, out int num, out int w) ? (p, num, w) : null;
			rows[k].IdState = status;
		}

		// Which series letter goes with which discipline, learned from the sheets that are already trustworthy
		// ("M" -> MECHANICAL, "D" -> DEMOLITION). Used to settle a misread letter and to name a discipline OCR could not read.
		private static Dictionary<string, Dictionary<string, int>> LearnDisciplineLetters(IList<IndexRow> rows, (string prefix, int number, int width)?[] parsed, bool[] anchor)
		{
			Dictionary<string, Dictionary<string, int>> byDiscipline = new Dictionary<string, Dictionary<string, int>>(StringComparer.Ordinal);
			for (int i = 0; i < rows.Count; i++)
			{
				if (!anchor[i] || !parsed[i].HasValue || string.IsNullOrWhiteSpace(rows[i].Discipline) || (rows[i].DisciplineOcr && rows[i].DisciplineConfidence < 70.0))
				{
					continue;
				}
				string? letter = IdPattern.TailLetter(parsed[i]!.Value.prefix);
				if (letter == null || letter.Length != 1)
				{
					continue;
				}
				string discipline = rows[i].Discipline.ToUpperInvariant();
				if (!byDiscipline.TryGetValue(discipline, out Dictionary<string, int>? letters))
				{
					letters = new Dictionary<string, int>(StringComparer.Ordinal);
					byDiscipline[discipline] = letters;
				}
				letters[letter] = letters.GetValueOrDefault(letter) + 1;
			}
			return byDiscipline;
		}

		private static string? LetterFor(Dictionary<string, Dictionary<string, int>> letterByDiscipline, string discipline)
		{
			if (!letterByDiscipline.TryGetValue(discipline.ToUpperInvariant(), out Dictionary<string, int>? letters))
			{
				return null;
			}
			int total = letters.Values.Sum();
			KeyValuePair<string, int> top = letters.OrderByDescending(p => p.Value).First();
			return total >= 3 && top.Value * 10 >= total * 6 ? top.Key : null;
		}

		// A run of OCR-read sheets that number consecutively but whose series letter disagrees with their (confidently read)
		// discipline: the letter was misread ("RW-N-01" for "RW-M-01" on MECHANICAL sheets). The whole series takes the discipline's letter.
		private static void FixSeriesLetters(IList<IndexRow> rows, string[] ids, (string prefix, int number, int width)?[] parsed, Dictionary<string, Dictionary<string, int>> letterByDiscipline, SetProfile profile, IndexAnalysisSummary summary)
		{
			int n = rows.Count;
			int i = 0;
			while (i < n)
			{
				if (!parsed[i].HasValue || !rows[i].IdOcr)
				{
					i++;
					continue;
				}
				int j = i;
				while (j + 1 < n && parsed[j + 1].HasValue && rows[j + 1].IdOcr && profile.PrefixFits(parsed[i]!.Value.prefix) && parsed[j + 1]!.Value.prefix == parsed[i]!.Value.prefix && parsed[j + 1]!.Value.number == parsed[j]!.Value.number + 1)
				{
					j++;
				}
				if (j > i && Enumerable.Range(i, j - i + 1).All(k => rows[k].IdConfidence < ConfidentIdLimit))
				{
					string prefix = parsed[i]!.Value.prefix;
					string? tail = IdPattern.TailLetter(prefix);
					List<string> votes = Enumerable.Range(i, j - i + 1)
						.Where(k => !string.IsNullOrWhiteSpace(rows[k].Discipline) && (!rows[k].DisciplineOcr || rows[k].DisciplineConfidence >= 70.0))
						.Select(k => LetterFor(letterByDiscipline, rows[k].Discipline)).Where(l => l != null).Select(l => l!).ToList();
					string? wanted = votes.GroupBy(l => l).OrderByDescending(g => g.Count()).Select(g => g.Key).FirstOrDefault();
					bool tailKnownForDiscipline = tail != null && Enumerable.Range(i, j - i + 1).Any(k => letterByDiscipline.TryGetValue(rows[k].Discipline.ToUpperInvariant(), out Dictionary<string, int>? letters) && letters.GetValueOrDefault(tail) >= 3);
					if (wanted != null && tail != null && tail.Length == 1 && tail != wanted && votes.Count * 2 >= j - i + 1 && !tailKnownForDiscipline)
					{
						string swapped = prefix.Substring(0, prefix.TrimEnd('-').Length - 1) + wanted + "-";
						for (int k = i; k <= j; k++)
						{
							string repaired = IdPattern.Format(swapped, parsed[k]!.Value.number, parsed[k]!.Value.width);
							rows[k].IdChanged = true;
							rows[k].IdUncertain = true;
							summary.IdsRepaired++;
							rows[k].Reasons.Add($"series letter {tail} -> {wanted} to match the discipline");
							ids[k] = repaired;
							parsed[k] = (swapped, parsed[k]!.Value.number, parsed[k]!.Value.width);
						}
					}
				}
				i = j + 1;
			}
		}
		// An unresolved sheet number whose series prefix does not match the surrounding series: take the neighbours' prefix when it
		// is a character or two away, unless the sheet's own discipline says it belongs to a different series.
		private static void FixPrefixFromNeighbours(IList<IndexRow> rows, string[] ids, (string prefix, int number, int width)?[] parsed, bool[] anchor, int i, Dictionary<string, Dictionary<string, int>> letterByDiscipline, IndexAnalysisSummary summary)
		{
			if (!parsed[i].HasValue)
			{
				return;
			}
			(string prefix, int number, int width) own = parsed[i]!.Value;
			string? disciplineLetter = string.IsNullOrWhiteSpace(rows[i].Discipline) ? null : LetterFor(letterByDiscipline, rows[i].Discipline);
			Dictionary<string, int> nearby = new Dictionary<string, int>();
			for (int j = Math.Max(0, i - 6); j <= Math.Min(rows.Count - 1, i + 6); j++)
			{
				if (j != i && anchor[j] && parsed[j].HasValue)
				{
					string prefix = parsed[j]!.Value.prefix;
					nearby[prefix] = nearby.GetValueOrDefault(prefix) + 1;
				}
			}
			string? ownLetter = IdPattern.TailLetter(own.prefix);
			string? chosen = null;
			// The discipline names the series letter: use it when the prefix, with that letter, is a known neighbouring series.
			if (disciplineLetter != null && ownLetter != null && ownLetter.Length == 1 && ownLetter != disciplineLetter)
			{
				string swapped = own.prefix.Substring(0, own.prefix.TrimEnd('-').Length - 1) + disciplineLetter + "-";
				if (nearby.ContainsKey(swapped) || rows[i].IdConfidence < LowConfidenceIdLimit)
				{
					chosen = swapped;
				}
			}
			if (chosen == null && !nearby.ContainsKey(own.prefix))
			{
				int limit = own.prefix.Length >= 5 ? 2 : (own.prefix.Length >= 4 ? 1 : (rows[i].IdConfidence < 50.0 ? 1 : 0));
				chosen = nearby.Where(p => p.Value >= 2 && OcrTextCleanup.EditDistance(own.prefix, p.Key, limit) <= limit)
					.OrderBy(p => OcrTextCleanup.EditDistance(own.prefix, p.Key, limit)).ThenByDescending(p => p.Value)
					.Select(p => p.Key).FirstOrDefault();
				if (chosen != null && disciplineLetter != null && IdPattern.TailLetter(chosen) != disciplineLetter)
				{
					chosen = null;
				}
			}
			if (chosen != null && own.number == 1)
			{
				// A first sheet is a misread of the neighbouring series only if the next sheet carries that series on (RW-N-01 before RW-M-02);
				// otherwise it may be a series of its own (SHI-D-01 between SHI-G-02 and SHI-Y-01).
				bool nextContinues = i + 1 < rows.Count && parsed[i + 1].HasValue && parsed[i + 1]!.Value.prefix == chosen && parsed[i + 1]!.Value.number == 2;
				if (!nextContinues)
				{
					chosen = null;
				}
			}
			if (chosen != null)
			{
				string? seriesDiscipline = Enumerable.Range(0, rows.Count).Where(j => j != i && anchor[j] && parsed[j].HasValue && parsed[j]!.Value.prefix == chosen)
					.Select(j => TrustedDiscipline(rows[j])).Where(d => d != null).GroupBy(d => d).OrderByDescending(g => g.Count()).Select(g => g.Key).FirstOrDefault();
				if (!DisciplineAllows(rows[i], seriesDiscipline))
				{
					chosen = null;
				}
			}
			if (chosen != null && chosen != own.prefix)
			{
				string repaired = IdPattern.Format(chosen, own.number, own.width);
				rows[i].IdChanged = true;
				summary.IdsRepaired++;
				rows[i].Reasons.Add($"sheet number series {own.prefix} -> {chosen} to match the discipline and neighbouring sheets");
				ids[i] = repaired;
				parsed[i] = (chosen, own.number, own.width);
			}
		}

		private static void FlagId(IndexRow row)
		{
			if (!row.IdOcr || string.IsNullOrWhiteSpace(row.Id))
			{
				return;
			}
			if (row.IdState == IdStatus.Anchor || row.IdState == IdStatus.InferredBothSides)
			{
				return;
			}
			if (row.IdState == IdStatus.InferredOneSide)
			{
				if (row.IdConfidence < LowConfidenceIdLimit && row.IdFixDistance >= 2)
				{
					row.IdUncertain = true;
					row.Reasons.Add("sheet number inferred from one neighbour only");
				}
				return;
			}
			row.IdUncertain = true;
			row.Reasons.Add(IdPattern.TryParse(row.Id, out _, out _, out _)
				? "sheet number does not continue the numbering around it or match the set's format"
				: "sheet number could not be read completely");
		}

		// ---- Disciplines ---------------------------------------------------------------------------------------------

		private static void CheckDisciplines(IList<IndexRow> rows, OcrVocabulary vocabulary, IndexAnalysisSummary summary)
		{
			(string prefix, int number, int width)?[] parsed = rows.Select(r => IdPattern.TryParse(IdPattern.Normalize(r.Id), out string p, out int num, out int w) ? ((string, int, int)?)(p, num, w) : null).ToArray();
			bool[] trusted = rows.Select(r => !r.IdOcr || r.IdState != IdStatus.Unresolved).ToArray();
			Dictionary<string, Dictionary<string, int>> letterByDiscipline = LearnDisciplineLetters(rows, parsed, trusted);
			for (int i = 0; i < rows.Count; i++)
			{
				IndexRow row = rows[i];
				string? letter = parsed[i].HasValue ? IdPattern.TailLetter(parsed[i]!.Value.prefix) : null;
				if (string.IsNullOrWhiteSpace(row.Discipline))
				{
					string? inferred = letter != null && letter.Length == 1 ? DisciplineFor(letterByDiscipline, letter) : null;
					inferred ??= letter != null ? DisciplineBetweenNeighbours(rows, parsed, i, letter) : null;
					if (inferred != null && row.DisciplineOcr)
					{
						row.Discipline = inferred;
						row.DisciplineChanged = true;
						row.DisciplineUncertain = true;
						row.Reasons.Add($"discipline inferred as {inferred} from the sheet number series");
						summary.DisciplinesFilled++;
					}
					continue;
				}
				if (!row.DisciplineOcr)
				{
					continue;
				}
				bool known = vocabulary.ContainsDiscipline(row.Discipline);
				string? expectedLetter = LetterFor(letterByDiscipline, row.Discipline);
				bool contradictsSeries = expectedLetter != null && letter != null && letter.Length == 1 && letter != expectedLetter && trusted[i];
				if ((!known && row.DisciplineConfidence < 80.0) || contradictsSeries)
				{
					row.DisciplineUncertain = true;
					row.Reasons.Add(contradictsSeries ? "discipline does not match the sheet number series" : "discipline not seen elsewhere in the set");
				}
			}
		}

		// A sheet with no discipline between two sheets of the same series that agree on one ("CIVIL/YARD" before and after) takes it. A series that
		// changes discipline part-way (CIVIL/YARD, then CIVIL DETAILS) is left alone where the neighbours disagree.
		private static string? DisciplineBetweenNeighbours(IList<IndexRow> rows, (string prefix, int number, int width)?[] parsed, int index, string letter)
		{
			string? Nearest(int step)
			{
				for (int j = index + step; j >= 0 && j < rows.Count; j += step)
				{
					if (!parsed[j].HasValue || IdPattern.TailLetter(parsed[j]!.Value.prefix) != letter)
					{
						return null;
					}
					if (!string.IsNullOrWhiteSpace(rows[j].Discipline))
					{
						return rows[j].Discipline;
					}
				}
				return null;
			}
			string? before = Nearest(-1);
			string? after = Nearest(1);
			return before != null && string.Equals(before, after, StringComparison.OrdinalIgnoreCase) ? before : null;
		}

		private static string? DisciplineFor(Dictionary<string, Dictionary<string, int>> letterByDiscipline, string letter)
		{
			List<(string discipline, int count, int total)> hits = new List<(string, int, int)>();
			foreach (KeyValuePair<string, Dictionary<string, int>> pair in letterByDiscipline)
			{
				if (pair.Value.TryGetValue(letter, out int count))
				{
					hits.Add((pair.Key, count, pair.Value.Values.Sum()));
				}
			}
			if (hits.Count == 0)
			{
				return null;
			}
			(string discipline, int count, int total) best = hits.OrderByDescending(h => h.count).First();
			int allWithLetter = hits.Sum(h => h.count);
			return best.count >= 3 && best.count * 10 >= allWithLetter * 8 ? best.discipline : null;
		}

		// ---- Titles --------------------------------------------------------------------------------------------------

		private static readonly Regex RomanTail = new Regex(@"^(?<base>.*\S)\s+(?<roman>[IVX]{1,5})$", RegexOptions.Compiled);

		private static bool TryRoman(string title, out string baseText, out int value)
		{
			Match m = RomanTail.Match(title ?? string.Empty);
			string roman = m.Success ? m.Groups["roman"].Value : string.Empty;
			if (!m.Success || ToRoman(FromRoman(roman)) != roman)
			{
				baseText = string.Empty;
				value = 0;
				return false;
			}
			baseText = m.Groups["base"].Value;
			value = FromRoman(roman);
			return true;
		}

		internal static string ToRoman(int value)
		{
			(int, string)[] map = { (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I") };
			string result = string.Empty;
			foreach ((int numeral, string text) in map)
			{
				while (value >= numeral)
				{
					result += text;
					value -= numeral;
				}
			}
			return result;
		}

		internal static int FromRoman(string roman)
		{
			int total = 0;
			int previous = 0;
			foreach (char c in roman.ToUpperInvariant().Reverse())
			{
				int v = c switch { 'I' => 1, 'V' => 5, 'X' => 10, _ => 0 };
				total += v < previous ? -v : v;
				previous = Math.Max(previous, v);
			}
			return total;
		}

		private static bool SameBase(string a, string b) => OcrTextCleanup.EditDistance(a.ToUpperInvariant(), b.ToUpperInvariant(), 2) <= 2;

		private static readonly Regex NumeralShapedTail = new Regex(@"^(?<base>.*\S)\s+(?<tail>[A-Z0-9!|/]{1,5})$", RegexOptions.Compiled);

		// A last word that is not a numeral and not a word of the set ("IV" read as "WV"), on a sheet whose neighbours have the same title
		// with numerals: it is the next numeral of the series when it is only a character or two away from it, because a series counts up.
		private static void RepairMisreadNumeral(IList<IndexRow> rows, int i, OcrVocabulary vocabulary)
		{
			IndexRow row = rows[i];
			Match m = NumeralShapedTail.Match(row.Title ?? string.Empty);
			if (!m.Success || vocabulary.IsKnownWord(m.Groups["tail"].Value) || vocabulary.ObservedCount(m.Groups["tail"].Value) >= 3)
			{
				return;
			}
			string baseText = m.Groups["base"].Value;
			string tail = m.Groups["tail"].Value;
			int expected = 0;
			if (i > 0 && TryRoman(rows[i - 1].Title, out string previousBase, out int previous) && SameBase(baseText, previousBase))
			{
				expected = previous + 1;
				if (i + 1 < rows.Count && TryRoman(rows[i + 1].Title, out string nextBase, out int next) && SameBase(baseText, nextBase) && next <= previous)
				{
					return;
				}
			}
			else if (i + 1 < rows.Count && TryRoman(rows[i + 1].Title, out string followingBase, out int following) && SameBase(baseText, followingBase) && following > 1)
			{
				expected = following - 1;
			}
			if (expected < 1 || expected > 30)
			{
				return;
			}
			string numeral = ToRoman(expected);
			if (OcrTextCleanup.EditDistance(tail, numeral, 2) <= 2 && tail.Length <= numeral.Length + 1)
			{
				row.Title = baseText + " " + numeral;
				row.TitleChanged = true;
				row.Reasons.Add($"numeral \"{tail}\" read as {numeral} because it continues the series");
			}
		}

		// Sheets of a series count up (I, II, III). A numeral that falls between two neighbours that are two apart is filled in;
		// one that breaks the counting is flagged. Repeated numerals are left alone (sets do have them) but are flagged when OCR-read.
		private static void CheckRomanNumerals(IList<IndexRow> rows, OcrVocabulary vocabulary)
		{
			for (int i = 0; i < rows.Count; i++)
			{
				IndexRow row = rows[i];
				if (row.TitleOcr && !TryRoman(row.Title, out _, out _))
				{
					RepairMisreadNumeral(rows, i, vocabulary);
				}
				if (!row.TitleOcr || !TryRoman(row.Title, out string baseText, out int value))
				{
					continue;
				}
				bool hasPrevious = i > 0 && TryRoman(rows[i - 1].Title, out string previousBase, out int previousValue) && SameBase(baseText, previousBase);
				bool hasNext = i + 1 < rows.Count && TryRoman(rows[i + 1].Title, out string nextBase, out int nextValue) && SameBase(baseText, nextBase);
				int previous = 0;
				int next = 0;
				if (hasPrevious)
				{
					TryRoman(rows[i - 1].Title, out _, out previous);
				}
				if (hasNext)
				{
					TryRoman(rows[i + 1].Title, out _, out next);
				}
				if (hasPrevious && hasNext && next - previous == 2 && value != previous + 1)
				{
					row.Title = baseText + " " + ToRoman(previous + 1);
					row.TitleChanged = true;
					row.Reasons.Add("numeral corrected to continue the series");
				}
				else if ((hasPrevious && value < previous) || (hasNext && value > next) || (hasPrevious && value == previous && !(hasNext && next == value + 1)))
				{
					// One of the reader's other readings may carry the numeral that continues the series.
					string? alternative = row.TitleOptions.Select(o => o.title).FirstOrDefault(o => TryRoman(o, out string optionBase, out int optionValue) && SameBase(optionBase, baseText)
						&& ((hasPrevious && optionValue == previous + 1 && (!hasNext || optionValue <= next)) || (!hasPrevious && hasNext && optionValue == next - 1)));
					if (alternative != null)
					{
						row.Title = alternative;
						row.TitleChanged = true;
						row.Reasons.Add("numeral taken from another reading because it continues the series");
					}
					else
					{
						row.TitleUncertain = true;
						row.Reasons.Add("numeral does not follow the series around it");
					}
				}
			}
		}


		private static void FlagTitles(IList<IndexRow> rows, OcrVocabulary vocabulary, SetProfile profile)
		{
			foreach (IndexRow row in rows.Where(r => r.TitleOcr))
			{
				if (row.Title.Contains("[DESCRIPTION NOT FOUND]", StringComparison.Ordinal) || string.IsNullOrWhiteSpace(row.Title))
				{
					continue;
				}
				if (row.TitleConfidence < 65.0)
				{
					row.TitleUncertain = true;
					row.Reasons.Add("title read with low confidence");
				}
				if (row.Title.Any(c => !profile.AllowsTitleCharacter(c)) || HasOddPeriod(row.Title))
				{
					row.TitleUncertain = true;
					row.Reasons.Add("title has marks that titles normally do not");
				}
				if (row.TitleGuessed && (row.TitleConfidence < 85.0 || vocabulary.Dictionary.Count > 0))
				{
					row.TitleUncertain = true;
					row.Reasons.Add("title words were corrected automatically");
				}
				if (vocabulary.Dictionary.Count > 0)
				{
					// With a word list, any word that is neither a real word nor consistent across the set is suspect, whatever the confidence.
					string[] misspelled = Regex.Matches(row.Title.ToUpperInvariant(), "[A-Z]{4,}").Select(m => m.Value).Where(w => !vocabulary.IsKnownForFit(w)).Distinct().ToArray();
					if (misspelled.Length > 0)
					{
						row.TitleUncertain = true;
						row.Reasons.Add("possible misspelling (" + string.Join(", ", misspelled.Take(3)) + ")");
					}
				}
				else
				{
					string[] unknown = Regex.Matches(row.Title.ToUpperInvariant(), "[A-Z]{5,}").Select(m => m.Value)
						.Where(w => !vocabulary.IsKnownWord(w) && vocabulary.ObservedCount(w) < 2).Distinct().ToArray();
					if (unknown.Length > 0 && row.TitleConfidence < 75.0)
					{
						row.TitleUncertain = true;
						row.Reasons.Add("title contains a word not seen elsewhere (" + string.Join(", ", unknown.Take(3)) + ")");
					}
				}
				// Standard writing conventions: "NO." is followed by a number; single stray letters, letters mixed with digits inside one
				// word, and a title that starts or ends with a mark are signs of a misread.
				if (Regex.IsMatch(row.Title, @"\bNO(?!\.?\s*\d)"))
				{
					row.TitleUncertain = true;
					row.Reasons.Add("\"NO.\" is not followed by a number");
				}
				if (Regex.IsMatch(row.Title, @"(?<![\w.'/&-])[B-HJ-UWYZ](?![\w.'/&-])"))
				{
					row.TitleUncertain = true;
					row.Reasons.Add("a stray single letter in the title");
				}
				if (Regex.IsMatch(row.Title, @"\b([A-Z]{3,})\s+\1\b"))
				{
					row.TitleUncertain = true;
					row.Reasons.Add("a word is repeated in the title");
				}
				if (Regex.IsMatch(row.Title, @"\b[A-Z]{2,}\d+[A-Z]{2,}\b|\b\d+[A-Z]{3,}\d*\b(?<!\d(ST|ND|RD|TH))"))
				{
					row.TitleUncertain = true;
					row.Reasons.Add("letters and digits are mixed inside a word");
				}
				if (Regex.IsMatch(row.Title.Trim(), @"^[^\w(""']|[^\w.)""']$"))
				{
					row.TitleUncertain = true;
					row.Reasons.Add("the title starts or ends with a stray mark");
				}
			}
		}

		// A series numbers its sheets, so two neighbouring sheets of one series with the very same title and no numeral are a sign that one of
		// them was misread or that a numeral was lost.
		private static void FlagRepeatedTitles(IList<IndexRow> rows)
		{
			for (int i = 1; i < rows.Count; i++)
			{
				IndexRow previous = rows[i - 1];
				IndexRow row = rows[i];
				if (!(row.TitleOcr || previous.TitleOcr) || string.IsNullOrWhiteSpace(row.Title) || row.Title.Contains('[') || !string.Equals(row.Title, previous.Title, StringComparison.OrdinalIgnoreCase))
				{
					continue;
				}
				if (IdPattern.TryParse(row.Id, out string prefix, out _, out _) && IdPattern.TryParse(previous.Id, out string previousPrefix, out _, out _) && prefix == previousPrefix && !TryRoman(row.Title, out _, out _) && !Regex.IsMatch(row.Title, @"\d"))
				{
					row.TitleUncertain = true;
					row.Reasons.Add("same title as the previous sheet of the series, with no number or numeral to tell them apart");
				}
			}
		}

		private static readonly Regex NumberAfterNo = new Regex(@"\bNO\.\s*(?<n>\d{1,2})(?!\d)", RegexOptions.Compiled);

		// "NO." followed by something that is not a number: a look-alike letter or two ("NO.J", "NO. II"), or nothing before THRU/AND.
		private static readonly Regex BrokenAfterNo = new Regex(@"\bNO\.\s*(?:(?<x>[A-Z!|/]{1,3})(?=\s|$)|(?<x>)(?=(?:THRU|THROUGH|AND|&)\b))", RegexOptions.Compiled);

		private static int? NumberAfterNoOf(string title, out string template)
		{
			Match m = NumberAfterNo.Match(title ?? string.Empty);
			if (!m.Success)
			{
				template = string.Empty;
				return null;
			}
			template = NumberAfterNo.Replace(title!, "NO.#");
			return int.Parse(m.Groups["n"].Value);
		}

		// The two words before "NO.", the marker, and the two words after the number: sheets of one series that say "BASINS NO.1 THRU 4"
		// share a key whatever the rest of their titles say.
		private static string NoKey(string title)
		{
			string[] words = Regex.Replace(Regex.Replace(title ?? string.Empty, @"\bNO\.\s*(?:\d{1,2}|[A-Z!|/]{1,3}(?=\s|$)|(?=(?:THRU|THROUGH|AND|&)\b))", "NO.# "), @"\s+", " ").Trim().Split(' ');
			int at = Array.IndexOf(words, "NO.#");
			if (at < 0)
			{
				return string.Empty;
			}
			return string.Join(" ", words.Skip(Math.Max(0, at - 2)).Take(Math.Min(2, at)).Concat(new[] { "NO.#" }).Concat(words.Skip(at + 1).Take(2)));
		}

		// "NO.1", "NO.2", "NO.3" ... : a number after NO. that was read as a letter, or lost, is taken from the rest of the set. First, if
		// every other sheet with the same wording around the number ("BASINS NO.# THRU 4") has the same number, that is it. Otherwise the
		// neighbouring titles decide: the number between two neighbours two apart, the same as both neighbours, or a look-alike digit
		// (I, Z, S, B ...) that the neighbours make possible. Anything less clear is left for the flag.
		private static void RepairNumbersAfterNo(IList<IndexRow> rows, OcrVocabulary vocabulary)
		{
			Dictionary<char, int> lookAlike = new Dictionary<char, int> { ['I'] = 1, ['!'] = 1, ['|'] = 1, ['/'] = 1, ['L'] = 1, ['Z'] = 2, ['S'] = 5, ['B'] = 8, ['G'] = 6 };
			Dictionary<string, Dictionary<int, int>> byKey = new Dictionary<string, Dictionary<int, int>>(StringComparer.Ordinal);
			foreach (IndexRow row in rows)
			{
				if (NumberAfterNoOf(row.Title, out _) is int number && NoKey(row.Title) is { Length: > 0 } key)
				{
					if (!byKey.TryGetValue(key, out Dictionary<int, int>? counts))
					{
						counts = new Dictionary<int, int>();
						byKey[key] = counts;
					}
					counts[number] = counts.GetValueOrDefault(number) + 1;
				}
			}
			for (int i = 0; i < rows.Count; i++)
			{
				IndexRow row = rows[i];
				if (!row.TitleOcr)
				{
					continue;
				}
				Match broken = BrokenAfterNo.Match(row.Title ?? string.Empty);
				if (!broken.Success)
				{
					continue;
				}
				string guess = broken.Groups["x"].Value;
				// "NO. AND ..." style words are not broken numbers.
				if (guess.Length >= 2 && vocabulary.IsRealWord(guess))
				{
					continue;
				}
				int? chosen = null;
				string key = NoKey(row.Title);
				if (key.Length > 0 && byKey.TryGetValue(key, out Dictionary<int, int>? sameWording) && sameWording.Values.Sum() >= 2 && sameWording.Count == 1)
				{
					chosen = sameWording.Keys.First();
				}
				if (!chosen.HasValue)
				{
					string template = BrokenAfterNo.Replace(row.Title, "NO.#");
					int? previous = i > 0 ? NumberAfterNoOf(rows[i - 1].Title, out _) : null;
					string previousShape = i > 0 ? NumberAfterNo.Replace(rows[i - 1].Title ?? string.Empty, "NO.#") : string.Empty;
					int? next = i + 1 < rows.Count ? NumberAfterNoOf(rows[i + 1].Title, out _) : null;
					string nextShape = i + 1 < rows.Count ? NumberAfterNo.Replace(rows[i + 1].Title ?? string.Empty, "NO.#") : string.Empty;
					if (previous.HasValue && !SameBase(template, previousShape))
					{
						previous = null;
					}
					if (next.HasValue && !SameBase(template, nextShape))
					{
						next = null;
					}
					HashSet<int> possible = new HashSet<int>();
					if (previous.HasValue)
					{
						possible.Add(previous.Value);
						possible.Add(previous.Value + 1);
					}
					if (next.HasValue)
					{
						possible.Add(next.Value);
						possible.Add(next.Value - 1);
					}
					if (previous.HasValue && next.HasValue && next.Value - previous.Value == 2)
					{
						chosen = previous.Value + 1;
					}
					else if (previous.HasValue && next.HasValue && previous.Value == next.Value)
					{
						chosen = previous.Value;
					}
					else if (guess.Length == 1 && lookAlike.TryGetValue(guess[0], out int digit) && possible.Contains(digit))
					{
						chosen = digit;
					}
				}
				if (chosen.HasValue)
				{
					row.Title = Regex.Replace(BrokenAfterNo.Replace(row.Title, "NO." + chosen.Value), @"NO\.(\d+)(?=[A-Z])", "NO.$1 ");
					row.TitleChanged = true;
					row.Reasons.Add($"number after NO. read as {chosen.Value} from the numbering of the other sheets");
				}
			}
		}
		private static bool HasOddPeriod(string title)
		{
			foreach (string token in title.Split(' ', StringSplitOptions.RemoveEmptyEntries))
			{
				if (!token.Contains('.'))
				{
					continue;
				}
				string core = token.Trim('.', ',', '(', ')');
				if (Regex.IsMatch(token, @"^\d+(\.\d+)?$|^NOS?\.[\d,&A-Z-]*$", RegexOptions.IgnoreCase) || Regex.IsMatch(core, @"^(MISC|APPROX|ELEV|SECT|TYP|DWG|DET|SCHED|ADDL|ST|MT|FT|INC|DEMO|MOD)$", RegexOptions.IgnoreCase))
				{
					continue;
				}
				return true;
			}
			return false;
		}
	}
}
