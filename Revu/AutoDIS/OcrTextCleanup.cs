using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace MbrohDrawingIndexCreator
{
	// Clean-up of raw OCR text for sheet IDs, disciplines and titles. Pure functions, so they are unit-tested
	// without needing Tesseract or a PDF.
	internal static class OcrTextCleanup
	{
		private static readonly Regex PlausibleSheetId = new Regex(@"^(?=[A-Z0-9.-]*\d)[A-Z0-9]+(?:[-.][A-Z0-9]+)*$", RegexOptions.Compiled);

		internal static bool IsPlausibleSheetId(string id) => PlausibleSheetId.IsMatch(id);

		// Letters that OCR commonly returns for a digit, and vice versa.
		private static char ToDigit(char c) => c switch
		{
			'I' or 'L' or '|' or '!' => '1',
			'O' => '0',
			'S' => '5',
			'B' => '8',
			'Z' => '2',
			_ => c
		};

		private static char ToLetter(char c) => c switch
		{
			'1' => 'I',
			'0' => 'O',
			'5' => 'S',
			'8' => 'B',
			'6' => 'G',
			_ => c
		};

		private static bool LooksNumeric(char c) => char.IsDigit(c) || c is 'I' or 'L' or '|' or '!' or 'O' or 'S' or 'B' or 'Z';

		// Returns the sheet ID with "-" separators, e.g. "G16" -> "G-16" and "G-I6" -> "G-16".
		internal static string CleanSheetId(string raw, bool expectHyphen)
		{
			string[] tokens = (raw ?? string.Empty).Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries);
			if (tokens.Length > 1)
			{
				// Oval-arc fragments come back as stray tokens ("i G-14 |"); prefer the one token that looks like a sheet ID.
				// Tokens that are only letters or only digits are pieces of a spaced-out ID ("G - 16"), so those are joined instead.
				string? single = tokens.Select(token => CleanSheetIdCore(token, expectHyphen))
					.Where(id => IsPlausibleSheetId(id) && id.Any(char.IsLetter) && id.Any(char.IsDigit))
					.OrderByDescending(id => id.Length).FirstOrDefault();
				if (single != null)
				{
					return single;
				}
			}			return CleanSheetIdCore(raw ?? string.Empty, expectHyphen);
		}

		private static string CleanSheetIdCore(string raw, bool expectHyphen)
		{
			string s = raw.Trim().ToUpperInvariant();
			s = Regex.Replace(s, "[\u2010-\u2015\u2212_~]", "-");
			s = Regex.Replace(s, @"\s*-\s*", "-");
			// Border and oval-arc fragments read as bars at either end of the field.
			s = Regex.Replace(s, @"^[\s|!\\/.,:;'`]+|[\s|!\\/.,:;'`]+$", string.Empty);
			s = Regex.Replace(s, @"\s+", string.Empty);
			s = Regex.Replace(s, @"[^A-Z0-9\-./|!]", string.Empty);
			s = s.Trim('-', '.', '/', '|', '!');
			if (s.Length == 0)
			{
				return string.Empty;
			}
			List<string> segments = s.Split('-', StringSplitOptions.RemoveEmptyEntries).ToList();
			if (segments.Count == 1 && expectHyphen)
			{
				Match exact = Regex.Match(segments[0], @"^([A-Z]{1,4})(\d{1,4})$");
				Match loose = Regex.Match(segments[0], @"^([A-Z]{1,4}?)([0-9ILOSBZ|!]{1,4})$");
				Match split = exact.Success ? exact : loose;
				if (split.Success)
				{
					segments = new List<string> { split.Groups[1].Value, split.Groups[2].Value };
				}
			}
			for (int i = 0; i < segments.Count; i++)
			{
				string segment = segments[i];
				bool isLast = i == segments.Count - 1;
				if (isLast && (segments.Count > 1 || segment.All(LooksNumeric)))
				{
					int numeric = segment.Count(LooksNumeric);
					if (numeric >= segment.Length * 0.75)
					{
						segment = new string(segment.Select(ToDigit).ToArray());
					}
				}
				else if (!isLast && (segment.Length == 1 || segment.Count(char.IsLetter) >= segment.Count(char.IsDigit)))
				{
					segment = new string(segment.Select(ToLetter).ToArray());
				}
				segments[i] = segment.Replace("|", string.Empty).Replace("!", string.Empty);
			}
			return string.Join("-", segments.Where(x => x.Length > 0));
		}

		private static readonly char[] TitleNoiseCharacters = "[]{}|_~`\"\\^<>".ToCharArray();

		// Removes border debris and maps a lone "|" (a Roman numeral I in "PLAN I") to "I".
		internal static string CleanTitle(string raw)
		{
			string flat = Regex.Replace(raw ?? string.Empty, @"\s*[\r\n]+\s*", " ");
			List<string> tokens = new List<string>();
			foreach (string token in flat.Split(' ', StringSplitOptions.RemoveEmptyEntries))
			{
				string t = token;
				if (t.Length <= 4 && t.All(c => c is '|' or 'l' or '!' or 'I') && t.Any(c => c != 'I') && tokens.Count > 0 && tokens[^1].Any(char.IsLetter))
				{
					tokens.Add(new string('I', t.Length));
					continue;
				}
				if (t.All(c => TitleNoiseCharacters.Contains(c)))
				{
					continue;
				}
				t = t.Replace('|', 'I').Trim(TitleNoiseCharacters);
				if (t.Length > 0)
				{
					tokens.Add(t);
				}
			}
			if (tokens.Count >= 2 && tokens[^1].Length <= 5 && tokens[^1].All(c => "IVXivxl|!".Contains(c)) && tokens[^2].Any(char.IsLetter)
				&& (tokens[^1].Length >= 2 || tokens[^1].Any(c => "VXvx".Contains(c))))
			{
				tokens[^1] = new string(tokens[^1].Select(c => c switch { 'l' or 'i' or '|' or '!' => 'I', 'v' => 'V', 'x' => 'X', _ => c }).ToArray());
			}
			return string.Join(" ", tokens);
		}

		internal static string CleanDiscipline(string raw)
		{
			string s = Regex.Replace(raw ?? string.Empty, @"\s*[\r\n]+\s*", " ").Trim(' ', '\t', '|', '[', ']', '_', '.', ',');
			s = Regex.Replace(s.ToUpperInvariant(), @"[^A-Z0-9 &/]", string.Empty);
			s = Regex.Replace(s, @"\s{2,}", " ").Trim();
			return s == "1&C" ? "I&C" : s;
		}

		internal static int EditDistance(string a, string b, int limit)
		{
			if (Math.Abs(a.Length - b.Length) > limit)
			{
				return limit + 1;
			}
			int[] previous = new int[b.Length + 1];
			int[] current = new int[b.Length + 1];
			for (int j = 0; j <= b.Length; j++)
			{
				previous[j] = j;
			}
			for (int i = 1; i <= a.Length; i++)
			{
				current[0] = i;
				int rowMinimum = current[0];
				for (int j = 1; j <= b.Length; j++)
				{
					int cost = a[i - 1] == b[j - 1] ? 0 : 1;
					current[j] = Math.Min(Math.Min(current[j - 1] + 1, previous[j] + 1), previous[j - 1] + cost);
					rowMinimum = Math.Min(rowMinimum, current[j]);
				}
				if (rowMinimum > limit)
				{
					return limit + 1;
				}
				(previous, current) = (current, previous);
			}
			return previous[b.Length];
		}
	}

	// Words and disciplines that are known to occur in the set being indexed (taken from sheets whose text layer was
	// read directly) plus a small built-in list, used to repair near-miss OCR such as "GE NERAL" or "ABREVIAIIONS".
	internal sealed class OcrVocabulary
	{
		private static readonly string[] DefaultDisciplines =
		{
			"GENERAL", "CIVIL", "STRUCTURAL", "ARCHITECTURAL", "MECHANICAL", "ELECTRICAL", "INSTRUMENTATION", "I&C",
			"PROCESS", "HVAC", "PLUMBING", "LANDSCAPE", "SITE", "TELECOMMUNICATIONS", "FIRE PROTECTION", "GEOTECHNICAL",
			"SURVEY", "DEMOLITION", "UTILITIES", "TRAFFIC", "STORMWATER", "PIPING", "SPECIFICATIONS"
		};

		private static readonly string[] DefaultTitleWords =
		{
			"ABBREVIATIONS", "ABBREVIATION", "LEGEND", "LEGENDS", "SYMBOLS", "GENERAL", "NOTES", "PLAN", "PLANS", "DETAILS",
			"DETAIL", "SECTION", "SECTIONS", "ELEVATION", "ELEVATIONS", "SCHEDULE", "SCHEDULES", "DIAGRAM", "DIAGRAMS",
			"PROCESS", "HYDRAULIC", "PROFILE", "MECHANICAL", "ELECTRICAL", "STRUCTURAL", "ARCHITECTURAL", "OVERALL", "INDEX",
			"DRAWING", "DRAWINGS", "COVER", "CRITERIA", "DESIGN", "CONVENTIONS", "INFORMATION", "PIPING", "EQUIPMENT",
			"FOUNDATION", "BUILDING", "TREATMENT", "CONTROL", "POWER", "LIGHTING", "GROUNDING", "INSTRUMENTATION",
			"TAGGING", "RIGGING", "DEMOLITION", "EROSION", "SEDIMENT", "GRADING", "DRAINAGE", "UTILITY", "STANDARD",
			"TYPICAL", "CONSTRUCTION", "SEQUENCE", "SPECIAL", "INSPECTIONS", "LOCATION", "VICINITY", "EXISTING",
			"PROPOSED", "SURVEY", "CONTROLS", "ARRANGEMENT", "PLATFORM", "STAIRS", "LAYOUT", "REINFORCING", "CONCRETE",
			"ROOFING", "FINISH", "DOORS", "WINDOWS", "LOUVERS", "VENTILATION", "PLUMBING", "FIXTURES", "CONDUIT",
			"PANEL", "ONE-LINE", "SINGLE", "MOTOR", "CONTROL", "ANALYSIS", "REQUIREMENTS", "CODE", "PROJECT", "WATER",
			"QUALITY", "IMPROVEMENTS", "SYSTEM", "SYSTEMS", "TANK", "BASIN", "FILTER", "BLOWER", "PUMP", "STATION",
			"SITE", "PIPE", "YARD", "ROOF", "FLOOR", "WALL", "WALLS", "VALVE", "VAULT", "TRENCH", "CHEMICAL", "STORAGE", "SCHEMATIC",
			"ENLARGED", "UPPER", "LOWER", "FIRST", "SECOND", "THIRD", "AREA", "LEVEL", "PAVING", "FACILITY", "FACILITIES", "TEMPORARY",
			"INTERIOR", "EXTERIOR", "STRUCTURES", "ROOM", "PARKING", "FENCE", "LANDSCAPE", "SURFACE"
		};

		private readonly HashSet<string> _disciplines = new HashSet<string>(StringComparer.Ordinal);
		private readonly Dictionary<string, int> _words = new Dictionary<string, int>(StringComparer.Ordinal);

		// Words that are certainly right: the built-in drawing words and the words of text-layer titles. Words merely seen on several OCR-read
		// sheets are not trusted (a misreading repeated on a whole series would otherwise pass for a word).
		private readonly HashSet<string> _trusted = new HashSet<string>(StringComparer.Ordinal);

		// The real-word list (English plus drawing terms); empty unless set, so tests and tools opt in.
		internal OcrDictionary Dictionary { get; set; } = OcrDictionary.Empty;

		internal static OcrVocabulary CreateDefault()
		{
			OcrVocabulary vocabulary = new OcrVocabulary { Dictionary = OcrDictionary.Default };
			foreach (string discipline in DefaultDisciplines)
			{
				vocabulary._disciplines.Add(discipline);
			}
			foreach (string word in DefaultTitleWords)
			{
				vocabulary._words[word] = 3;
				vocabulary._trusted.Add(word);
			}
			return vocabulary;
		}

		internal int DisciplineCount => _disciplines.Count;

		internal void AddDiscipline(string discipline)
		{
			string clean = OcrTextCleanup.CleanDiscipline(discipline);
			if (clean.Length >= 2)
			{
				_disciplines.Add(clean);
			}
		}

		internal void AddTitle(string title, int weight = 3)
		{
			foreach (Match match in Regex.Matches((title ?? string.Empty).ToUpperInvariant(), "[A-Z]{4,}"))
			{
				_words.TryGetValue(match.Value, out int count);
				_words[match.Value] = count + weight;
				_trusted.Add(match.Value);
			}
		}

		internal bool ContainsDiscipline(string clean) => _disciplines.Contains(clean);

		// True when a word of four or more letters in the text is not a real word (a reading error, a name, or lettering the reader garbled).
		internal bool ContainsNonWord(string text)
		{
			return Regex.Matches((text ?? string.Empty).ToUpperInvariant(), "[A-Z]{4,}").Any(m => !IsRealWord(m.Value));
		}

		// Disciplines the reader produced confidently on several sheets join the known ones, most frequent first. A reading that is a small
		// edit (or only a gap) away from a discipline already known is the same discipline misread ("ARCHIT ECT URAL"), not a new one.
		internal void LearnDisciplines(IEnumerable<KeyValuePair<string, int>> counts)
		{
			foreach (KeyValuePair<string, int> pair in counts.Where(p => p.Value >= 2 && p.Key.Length >= 2).OrderByDescending(p => p.Value).ThenBy(p => p.Key.Count(c => c == ' ')))
			{
				string key = pair.Key.Replace(" ", string.Empty);
				int limit = key.Length >= 8 ? 2 : 1;
				if (!_disciplines.Any(d => OcrTextCleanup.EditDistance(key, d.Replace(" ", string.Empty), limit) <= limit))
				{
					_disciplines.Add(pair.Key);
				}
			}
		}

		// A discipline read from an image is a word or two, so the title repairs apply to its words: a word torn in two ("Y ARD"), a look-alike
		// letter, two strokes read as one ("CNIL"). A slash keeps its two sides apart (CIVIL/YARD).
		internal string RefineDiscipline(string clean)
		{
			if (clean.Length == 0 || Dictionary.Count == 0)
			{
				return clean;
			}
			List<string> parts = new List<string>();
			foreach (string part in clean.Split('/'))
			{
				List<string> tokens = part.Split(' ', StringSplitOptions.RemoveEmptyEntries).ToList();
				tokens = MergeSplitWords(tokens);
				tokens = CorrectSpelling(tokens, out _);
				parts.Add(string.Join(' ', tokens));
			}
			return string.Join("/", parts);
		}

		internal string SnapDiscipline(string clean)
		{
			if (clean.Length == 0 || _disciplines.Contains(clean))
			{
				return clean;
			}
			string key = clean.Replace(" ", string.Empty);
			int limit = key.Length <= 3 ? 1 : (key.Length * 2 + 4) / 5;
			string? best = null;
			int bestDistance = int.MaxValue;
			bool tie = false;
			foreach (string candidate in _disciplines)
			{
				int distance = OcrTextCleanup.EditDistance(key, candidate.Replace(" ", string.Empty), limit);
				if (distance > limit)
				{
					continue;
				}
				if (distance < bestDistance)
				{
					best = candidate;
					bestDistance = distance;
					tie = false;
				}
				else if (distance == bestDistance)
				{
					tie = true;
				}
			}
			return best != null && !tie ? best : clean;
		}

		private readonly Dictionary<string, int> _observed = new Dictionary<string, int>(StringComparer.Ordinal);
		// How many titles of the set contain each pair of neighbouring words (set after LearnFromTitles).
		private readonly Dictionary<string, int> _titlePairs = new Dictionary<string, int>(StringComparer.Ordinal);
		// For words followed by a number in some titles (LAGOON 2): how often by a digit, and how often by a Roman numeral of two or more letters.
		private readonly Dictionary<string, int> _wordThenDigit = new Dictionary<string, int>(StringComparer.Ordinal);
		private readonly Dictionary<string, int> _wordThenRoman = new Dictionary<string, int>(StringComparer.Ordinal);
		// How many titles have a comma straight after the word.
		private readonly Dictionary<string, int> _wordThenComma = new Dictionary<string, int>(StringComparer.Ordinal);

		// Abbreviations that legitimately end in a period inside a title.
		private static readonly HashSet<string> AbbreviationsWithPeriod = new HashSet<string>(StringComparer.Ordinal)
		{
			"NO", "NOS", "MISC", "APPROX", "ELEV", "SECT", "TYP", "DWG", "DET", "SCHED", "ADDL", "ST", "MT", "FT", "INC", "DEMO", "MOD"
		};

		private int _noDigit;
		private int _noRoman;

		// True when the set numbers things "NO.1", "NO.2" in digits (never in Roman numerals), so a "NO.I" is a misread 1.
		internal bool NumbersAfterNoAreDigits => _noDigit >= 3 && _noRoman == 0;

		internal bool IsKnownWord(string upperWord) => _words.ContainsKey(upperWord);

		// A real word: in the word list or trusted by the set.
		internal bool IsRealWord(string upperWord) => Dictionary.Contains(upperWord) || _trusted.Contains(upperWord);

		// Real, or consistent enough across the set to be a name or term of its own (BACHMAN) rather than a repeated misreading: it must be
		// seen on several sheets and must not be one confusable letter away from a real word (QZONE).
		internal bool IsKnownForFit(string upperWord) => IsRealWord(upperWord) || (ObservedCount(upperWord) >= 3 && Dictionary.Suggest(upperWord, OcrDictionary.ConfusionCost + 0.05).Count == 0);

		internal int WordWeight(string upperWord) => _words.GetValueOrDefault(upperWord);

		// How many sheets' titles contain the word (set after LearnFromTitles).
		internal int ObservedCount(string upperWord) => _observed.GetValueOrDefault(upperWord);

		// Words that recur across the set's own titles are, by repetition, almost certainly spelled right: they join the
		// vocabulary, and any other word that is a near miss of one of them can be corrected.
		internal void LearnFromTitles(IEnumerable<string> titles, int minimumSheets = 3)
		{
			_observed.Clear();
			_titlePairs.Clear();
			_wordThenDigit.Clear();
			_wordThenRoman.Clear();
			_wordThenComma.Clear();
			_noDigit = 0;
			_noRoman = 0;
			foreach (string title in titles)
			{
				string upper = (title ?? string.Empty).ToUpperInvariant();
				if (Regex.IsMatch(upper, @"\bNO\.\s*\d"))
				{
					_noDigit++;
				}
				// A lone "I" is left out: it is as likely a misread 1 as a numeral.
				if (Regex.IsMatch(upper, @"\bNO\.\s*(?:II|III|IV|V|VI|VII|VIII|IX|X)\b"))
				{
					_noRoman++;
				}
				string[] letterWords = Regex.Matches(upper, "[A-Z]+").Select(m => m.Value).ToArray();
				foreach (string commaWord in Regex.Matches(upper, "([A-Z]{3,}),").Select(m => m.Groups[1].Value).Distinct())
				{
					_wordThenComma[commaWord] = _wordThenComma.GetValueOrDefault(commaWord) + 1;
				}
				foreach (Match m in Regex.Matches(upper, @"\b([A-Z]{3,})\s+(\d+)\b"))
				{
					_wordThenDigit[m.Groups[1].Value] = _wordThenDigit.GetValueOrDefault(m.Groups[1].Value) + 1;
				}
				foreach (Match m in Regex.Matches(upper, @"\b([A-Z]{3,})\s+(?:II|III|IV|VI|VII|VIII|IX)\b"))
				{
					_wordThenRoman[m.Groups[1].Value] = _wordThenRoman.GetValueOrDefault(m.Groups[1].Value) + 1;
				}
				foreach (string pair in Enumerable.Range(0, Math.Max(0, letterWords.Length - 1)).Select(k => letterWords[k] + " " + letterWords[k + 1]).Distinct())
				{
					_titlePairs[pair] = _titlePairs.GetValueOrDefault(pair) + 1;
				}
				foreach (string word in letterWords.Distinct())
				{
					_observed[word] = _observed.GetValueOrDefault(word) + 1;
				}
			}
			foreach (KeyValuePair<string, int> pair in _observed)
			{
				if (pair.Key.Length >= 4 && pair.Value >= minimumSheets)
				{
					_words[pair.Key] = Math.Max(_words.GetValueOrDefault(pair.Key), pair.Value);
				}
			}
		}

		// Only all-capital words are touched, and only when exactly one known word is clearly closest.
		internal string SnapTitle(string title)
		{
			if (string.IsNullOrEmpty(title))
			{
				return title;
			}
			StringBuilder result = new StringBuilder();
			foreach (string token in title.Split(' '))
			{
				if (result.Length > 0)
				{
					result.Append(' ');
				}
				result.Append(SnapToken(token));
			}
			return result.ToString();
		}

		// Cleans a whole OCR title using what is known about the set: stray marks, "NO." numbers, words split by a gap
		// ("DET AILS") and near-miss spellings. Returns the title unchanged when nothing needed fixing.
		internal string RefineTitle(string title, out bool changed) => RefineTitle(title, out changed, out _);

		// guessed: a word was replaced by a similar-looking known word (a likelier but unproven fix).
		internal string RefineTitle(string title, out bool changed, out bool guessed)
		{
			string original = title ?? string.Empty;
			string s = Regex.Replace(original, @"\s*_\s*", " ");
			// "NO.1": the O is often read as Q, and the 1 as a slash, I, L, E, pipe or comma.
			s = Regex.Replace(s, @"\bNO[Q0]\.", "NO.");
			s = Regex.Replace(s, @"\bN[Q0O][.,]", "NO.");
			s = Regex.Replace(s, @"\bN[Q0](?=\s|\d|$)", "NO.");
			s = Regex.Replace(s, @"\bNO\.\s*[/|!,;]+(?=\s|[A-Z]|$)", "NO.1 ");
			s = Regex.Replace(s, @"\bNO\.\s*[IL](?=AND|THRU)", "NO.1 ");
			s = Regex.Replace(s, @"\bNO\.\s*E(?=THRU)", "NO.1 ");
			s = Regex.Replace(s, @"\bNO\.\s*[IL](?=\s+(?:AND|THRU|THROUGH|&)\b)", "NO.1");
			if (NumbersAfterNoAreDigits)
			{
				s = Regex.Replace(s, @"\bNO[I|!](?=\s|&|$)", "NO.1");
			}
			s = Regex.Replace(s, @"\bNO\.(?=[A-Z]{2,})", "NO. ");
			s = Regex.Replace(s, @"\bNO\.\s+(?=\d)", "NO.");
			if (NumbersAfterNoAreDigits)
			{
				s = Regex.Replace(s, @"\bNO\.\s*[IJL](?=\s|$)", "NO.1");
			}
			s = Regex.Replace(s, @"\s{2,}", " ").Trim();			List<string> tokens = s.Split(' ', StringSplitOptions.RemoveEmptyEntries).ToList();
			for (int i = 0; i < tokens.Count; i++)
			{
				// A stray period after a word. Short words are only trusted to be plain words when the set uses them often.
				Match m = Regex.Match(tokens[i], @"^([A-Z]{2,})\.$");
				if (m.Success && !AbbreviationsWithPeriod.Contains(m.Groups[1].Value) && (m.Groups[1].Length >= 4 || ObservedCount(m.Groups[1].Value) >= 3))
				{
					tokens[i] = m.Groups[1].Value;
				}
				// A comma that is only a speck after a word the set otherwise never writes with one (SECONDARY, TREATMENT).
				Match comma = Regex.Match(tokens[i], "^([A-Z]{3,}),$");
				if (comma.Success && _wordThenComma.GetValueOrDefault(comma.Groups[1].Value) <= 1 && ObservedCount(comma.Groups[1].Value) >= 5)
				{
					tokens[i] = comma.Groups[1].Value;
				}
				// A letter I read as a slash or bar inside a word ("S/TE" for "SITE").
				Match slash = Regex.Match(tokens[i], @"^([A-Z]+)[/|!]([A-Z]+)$");
				if (slash.Success)
				{
					foreach (char letter in new[] { 'I', 'L' })
					{
						string mended = slash.Groups[1].Value + letter + slash.Groups[2].Value;
						if (IsKnownWord(mended) || ObservedCount(mended) >= 2)
						{
							tokens[i] = mended;
							break;
						}
					}
				}
			}
			for (int i = 1; i < tokens.Count; i++)
			{
				// "LAGOON I" where the set writes "LAGOON 2": a 1 read as the letter I.
				if (tokens[i] is "I" or "L" or "|" && Regex.IsMatch(tokens[i - 1], "^[A-Z]{3,}$") && _wordThenDigit.GetValueOrDefault(tokens[i - 1]) >= 2 && _wordThenRoman.GetValueOrDefault(tokens[i - 1]) == 0)
				{
					tokens[i] = "1";
				}
			}
			tokens = MergeSplitWords(tokens);
			tokens = CorrectSpelling(tokens, out bool spellingGuessed);
			tokens = CorrectFromContext(tokens, ref spellingGuessed);
			string beforeSnap = string.Join(" ", tokens);
			string refined = SnapTitle(beforeSnap);
			guessed = spellingGuessed || !string.Equals(refined, beforeSnap, StringComparison.Ordinal);
			changed = !string.Equals(refined, original, StringComparison.Ordinal);
			return refined;
		}

		private readonly Dictionary<string, int> _wordRows = new Dictionary<string, int>(StringComparer.Ordinal);
		private readonly Dictionary<string, int> _pairRows = new Dictionary<string, int>(StringComparer.Ordinal);

		// Notes, over every reading the reader produced for every sheet (not just the chosen one), how many sheets contain each word and
		// each pair of neighbouring words. Readers split words far more often than they join two different words, so a word that shows up
		// joined on several sheets is the real spelling even when most sheets read it split (a whole series of "HE ADHOUSE").
		internal void LearnReadings(IEnumerable<IEnumerable<string>> readingsPerSheet)
		{
			_wordRows.Clear();
			_pairRows.Clear();
			foreach (IEnumerable<string> readings in readingsPerSheet)
			{
				HashSet<string> words = new HashSet<string>(StringComparer.Ordinal);
				HashSet<string> pairs = new HashSet<string>(StringComparer.Ordinal);
				foreach (string reading in readings)
				{
					string[] tokens = Regex.Matches((reading ?? string.Empty).ToUpperInvariant(), "[A-Z]+").Select(m => m.Value).ToArray();
					foreach (string token in tokens)
					{
						words.Add(token);
					}
					for (int i = 0; i + 1 < tokens.Length; i++)
					{
						pairs.Add(tokens[i] + " " + tokens[i + 1]);
					}
				}
				foreach (string word in words)
				{
					_wordRows[word] = _wordRows.GetValueOrDefault(word) + 1;
				}
				foreach (string pair in pairs)
				{
					_pairRows[pair] = _pairRows.GetValueOrDefault(pair) + 1;
				}
			}
		}

		private bool JoinedFormIsReal(string first, string second)
		{
			string joined = first + second;
			return joined.Length >= 6 && _wordRows.GetValueOrDefault(joined) >= 3 && _wordRows.GetValueOrDefault(joined) * 4 >= _pairRows.GetValueOrDefault(first + " " + second);
		}

		// Words that are not real words are compared with real ones that differ only by letters OCR confuses (QZONE -> OZONE, AYDRAIED ->
		// HYDRATED). One confusable letter is corrected with confidence; two are corrected but reported as a guess; anything less clear is left
		// for the checks to flag. A name that recurs across the set (BACHMAN) is only touched for a single confusable letter.
		private List<string> CorrectSpelling(List<string> tokens, out bool guessed)
		{
			guessed = false;
			if (Dictionary.Count == 0)
			{
				return tokens;
			}
			List<string> result = new List<string>(tokens.Count);
			foreach (string token in tokens)
			{
				Match m = Regex.Match(token, @"^([^A-Za-z]*)([A-Z]{4,})([^A-Za-z]*)$");
				if (!m.Success || IsRealWord(m.Groups[2].Value))
				{
					result.Add(token);
					continue;
				}
				string core = m.Groups[2].Value;
				double limit = ObservedCount(core) >= 3 ? OcrDictionary.ConfusionCost + 0.05 : (core.Length >= 6 ? 2 * OcrDictionary.ConfusionCost + 0.05 : OcrDictionary.ConfusionCost + 0.05);
				IReadOnlyList<(string Word, double Cost)> suggestions = Dictionary.Suggest(core, limit);
				if (suggestions.Count == 0)
				{
					// Two strokes read as one letter (or the reverse): CNIL is CIVIL. Only one reachable real word, or one the set clearly uses most.
					IReadOnlyList<string> rewritten = Dictionary.SuggestRewrites(core);
					string? pick = rewritten.Count == 1 ? rewritten[0] : rewritten.Count > 1 && rewritten.Count(w => ObservedCount(w) > 0) == 1 ? rewritten.First(w => ObservedCount(w) > 0) : null;
					if (pick != null)
					{
						result.Add(m.Groups[1].Value + pick + m.Groups[3].Value);
						guessed |= ObservedCount(pick) < 3;
					}
					else
					{
						result.Add(token);
					}
					continue;
				}
				// Several equally good words: the one the set itself uses most, if only one of them is used at all.
				List<(string Word, double Cost)> best = suggestions.Where(s => s.Cost <= suggestions[0].Cost + 1e-9).ToList();
				if (best.Count > 1)
				{
					List<(string Word, double Cost)> used = best.Where(b => ObservedCount(b.Word) > 0 || _trusted.Contains(b.Word)).ToList();
					if (used.Count != 1)
					{
						result.Add(token);
						continue;
					}
					best = used;
				}
				if (best[0].Cost > OcrDictionary.ConfusionCost + 0.05)
				{
					guessed = true;
				}
				result.Add(m.Groups[1].Value + best[0].Word + m.Groups[3].Value);
			}
			return result;
		}

		private int PairCount(string? first, string? second) => first == null || second == null ? 0 : _titlePairs.GetValueOrDefault(first + " " + second);

		private static string? LetterWord(IReadOnlyList<string> tokens, int index) => index >= 0 && index < tokens.Count && Regex.IsMatch(tokens[index], "^[A-Z]{2,}$") ? tokens[index] : null;

		// What the neighbouring words say. A word seen on a sheet or two whose pairing with its neighbours appears nowhere else, while a
		// commoner word one edit away pairs with those neighbours on several sheets, is that commoner word misread (WAFER METER -> WATER METER,
		// since WATER METER is a pair the set uses); a stray one- or two-letter token whose removal makes its neighbours a known pair is noise;
		// "CND FLOOR" is a 2ND whose digit was read as a letter. Every change is a guess and is flagged.
		private List<string> CorrectFromContext(List<string> tokens, ref bool guessed)
		{
			if (_titlePairs.Count == 0)
			{
				return tokens;
			}
			List<string> result = new List<string>(tokens);
			for (int i = 0; i < result.Count; i++)
			{
				string token = result[i];
				if (Regex.IsMatch(token, "^[CZQ]ND$") && (LetterWord(result, i + 1) != null))
				{
					result[i] = "2ND";
					guessed = true;
					continue;
				}
				string? previous = LetterWord(result, i - 1);
				string? next = LetterWord(result, i + 1);
				if (Regex.IsMatch(token, "^[A-Z]{1,2}$") && !Regex.IsMatch(token, "^[IVX]+$") && ObservedCount(token) < 3 && previous != null && next != null && PairCount(previous, next) >= 2)
				{
					result.RemoveAt(i);
					i--;
					guessed = true;
					continue;
				}
				if (!Regex.IsMatch(token, "^[A-Z]{3,}$") || ObservedCount(token) > 2 || (previous == null && next == null))
				{
					continue;
				}
				if (PairCount(previous, token) > 0 || PairCount(token, next) > 0)
				{
					continue;
				}
				string? best = null;
				int bestSupport = 0;
				bool tie = false;
				foreach (KeyValuePair<string, int> word in _observed)
				{
					if (word.Value < Math.Max(4, 3 * ObservedCount(token) + 1) || word.Key == token || Math.Abs(word.Key.Length - token.Length) > 1)
					{
						continue;
					}
					int allowed = token.Length >= 8 ? 2 : 1;
					if (OcrTextCleanup.EditDistance(token, word.Key, allowed) > allowed || !IsKnownForFit(word.Key))
					{
						continue;
					}
					int support = Math.Max(PairCount(previous, word.Key), PairCount(word.Key, next));
					if (support < 2)
					{
						continue;
					}
					if (support > bestSupport)
					{
						best = word.Key;
						bestSupport = support;
						tie = false;
					}
					else if (support == bestSupport)
					{
						tie = true;
					}
				}
				if (best != null && !tie)
				{
					result[i] = best;
					guessed = true;
				}
			}
			return result;
		}
		private List<string> MergeSplitWords(List<string> tokens)
		{
			List<string> result = new List<string>();
			for (int i = 0; i < tokens.Count; i++)
			{
				// A real word split in two where at least one piece is not a word ("HE ADHOUSE", "DET AILS").
				if (Dictionary.Count > 0 && i + 1 < tokens.Count && Regex.IsMatch(tokens[i], "^[A-Z]{1,10}$") && Regex.IsMatch(tokens[i + 1], "^[A-Z]{1,10}$")
					&& IsRealWord(tokens[i] + tokens[i + 1]) && (!IsRealWord(tokens[i]) || !IsRealWord(tokens[i + 1])))
				{
					result.Add(tokens[i] + tokens[i + 1]);
					i++;
					continue;
				}
				if (i + 1 < tokens.Count && Regex.IsMatch(tokens[i], "^[A-Z]{1,10}$") && Regex.IsMatch(tokens[i + 1], "^[A-Z]{1,10}$") && JoinedFormIsReal(tokens[i], tokens[i + 1]))
				{
					result.Add(tokens[i] + tokens[i + 1]);
					i++;
					continue;
				}
				if (i + 1 < tokens.Count && Regex.IsMatch(tokens[i], "^[A-Z]{1,8}$") && Regex.IsMatch(tokens[i + 1], "^[A-Z]{2,10}$"))
				{
					string merged = tokens[i] + tokens[i + 1];
					int weightOfPieces = Math.Max(ObservedCount(tokens[i]), ObservedCount(tokens[i + 1]));
					string? word = FindWord(merged, merged.Length >= 8 ? 1 : 0);
					bool bothAreWords = tokens[i].Length >= 3 && IsKnownWord(tokens[i]) && IsKnownWord(tokens[i + 1]);
					if (word != null && !bothAreWords && merged.Length >= 6 && WordWeight(word) >= 3 && WordWeight(word) * 2 >= weightOfPieces * 3)
					{
						result.Add(word);
						i++;
						continue;
					}
				}
				// A word the reader has torn into a stray fragment plus its tail or head ("FATE ADHOUSE" for HEADHOUSE): the piece that is not a
				// word is the end (or start) of a word the set uses constantly.
				if (Dictionary.Count > 0 && i + 1 < tokens.Count && Regex.IsMatch(tokens[i], "^[A-Z]{1,4}$") && Regex.IsMatch(tokens[i + 1], "^[A-Z]{5,12}$") && !IsRealWord(tokens[i + 1]))
				{
					string? whole = FindFrequentWordEnding(tokens[i + 1], tokens[i].Length);
					if (whole != null)
					{
						result.Add(whole);
						i++;
						continue;
					}
				}
				if (Dictionary.Count > 0 && i + 1 < tokens.Count && Regex.IsMatch(tokens[i], "^[A-Z]{5,12}$") && Regex.IsMatch(tokens[i + 1], "^[A-Z]{1,4}$") && !IsRealWord(tokens[i]))
				{
					string? whole = FindFrequentWordStarting(tokens[i], tokens[i + 1].Length);
					if (whole != null)
					{
						result.Add(whole);
						i++;
						continue;
					}
				}
				result.Add(tokens[i]);
			}
			return result;
		}

		private string? FindFrequentWordEnding(string tail, int fragmentLength)
		{
			int slack = tail.Length >= 6 ? 1 : 0;
			return FindFrequentWord(tail, word => word.Length - tail.Length >= 1 && word.Length - tail.Length <= fragmentLength + 1
				&& OcrTextCleanup.EditDistance(word.Substring(word.Length - tail.Length), tail, slack) <= slack);
		}

		private string? FindFrequentWordStarting(string head, int fragmentLength)
		{
			int slack = head.Length >= 6 ? 1 : 0;
			return FindFrequentWord(head, word => word.Length - head.Length >= 1 && word.Length - head.Length <= fragmentLength + 1
				&& OcrTextCleanup.EditDistance(word.Substring(0, head.Length), head, slack) <= slack);
		}

		// The commonest set word that fits, only when it is a real, frequently seen word and exactly one such word fits.
		private string? FindFrequentWord(string piece, Func<string, bool> fits)
		{
			string? found = null;
			foreach (KeyValuePair<string, int> word in _words)
			{
				if (word.Value >= 5 && word.Value >= 2 * ObservedCount(piece) && IsKnownForFit(word.Key) && fits(word.Key))
				{
					if (found != null)
					{
						return null;
					}
					found = word.Key;
				}
			}
			return found;
		}

		private string? FindWord(string candidate, int maxDistance)
		{
			if (_words.ContainsKey(candidate))
			{
				return candidate;
			}
			if (maxDistance == 0)
			{
				return null;
			}
			string? best = null;
			int bestWeight = 0;
			foreach (KeyValuePair<string, int> word in _words)
			{
				if (OcrTextCleanup.EditDistance(candidate, word.Key, maxDistance) <= maxDistance && word.Value > bestWeight)
				{
					best = word.Key;
					bestWeight = word.Value;
				}
			}
			return best;
		}

		private string SnapToken(string token)
		{
			Match m = Regex.Match(token, @"^([^A-Za-z]*)([A-Z]{4,})([^A-Za-z]*)$");
			if (!m.Success || _words.ContainsKey(m.Groups[2].Value))
			{
				return token;
			}
			string core = m.Groups[2].Value;
			// A real word is never snapped to a look-alike that merely recurs in the set.
			if (IsRealWord(core))
			{
				return token;
			}
			// A word that itself recurs in the set is trusted; only rarer spellings are corrected, and only to clearly commoner words.
			int ownCount = ObservedCount(core);
			// A long word seen once is far likelier a garbled frequent word (FAYDRAIED for HYDRATED) than a name of its own.
			int limit = core.Length >= 9 && ownCount <= 1 ? 3 : (core.Length >= 8 ? 2 : 1);
			int minimumWeight = Math.Max(core.Length >= 5 ? 1 : 5, ownCount * 2);
			string? best = null;
			int bestDistance = int.MaxValue;
			int bestCount = 0;
			bool tie = false;
			foreach (KeyValuePair<string, int> word in _words)
			{
				if (word.Value < minimumWeight || (Dictionary.Count > 0 && !IsKnownForFit(word.Key)))
				{
					continue;
				}
				int distance = OcrTextCleanup.EditDistance(core, word.Key, limit);
				if (distance > limit)
				{
					continue;
				}
				// Plural/singular and similar one-letter endings are real words, not misreads.
				if (Math.Abs(core.Length - word.Key.Length) == 1 && (core.StartsWith(word.Key, StringComparison.Ordinal) || word.Key.StartsWith(core, StringComparison.Ordinal)))
				{
					continue;
				}
				if (distance < bestDistance || (distance == bestDistance && word.Value > bestCount))
				{
					best = word.Key;
					bestDistance = distance;
					bestCount = word.Value;
					tie = false;
				}
				else if (distance == bestDistance && word.Value == bestCount)
				{
					tie = true;
				}
			}
			return best != null && !tie ? m.Groups[1].Value + best + m.Groups[3].Value : token;
		}
	}
}