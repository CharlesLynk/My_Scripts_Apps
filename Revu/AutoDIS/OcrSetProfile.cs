using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace MbrohDrawingIndexCreator
{
	// One reading of a field, as produced by one pass of the reader (a particular image variant and model).
	internal readonly record struct ReadingCandidate(string Text, double Confidence, bool Large = false);

	// What the drawing set itself shows to be normal: the shapes its sheet numbers take, which series prefixes and number widths
	// occur, and which punctuation appears in its titles. Built from everything that can be trusted (embedded text and confident or
	// self-consistent OCR) rather than from rules about what a drawing number "should" be, so a client's quirky convention is learned,
	// not rejected. When the readings of a field disagree, the one that fits the set's profile best is chosen.
	internal sealed class SetProfile
	{
		private const int MinimumSamples = 8;

		private readonly Dictionary<string, int> _idShapes = new Dictionary<string, int>(StringComparer.Ordinal);
		private readonly Dictionary<string, int> _prefixShapes = new Dictionary<string, int>(StringComparer.Ordinal);
		private readonly Dictionary<string, int> _prefixCounts = new Dictionary<string, int>(StringComparer.Ordinal);
		private readonly Dictionary<string, HashSet<int>> _widthsByPrefix = new Dictionary<string, HashSet<int>>(StringComparer.Ordinal);
		private readonly Dictionary<char, int> _titleCharacters = new Dictionary<char, int>();
		private int _topShape;
		private int _hyphenated;

		internal int IdSamples { get; private set; }

		internal int TextTitleCount { get; private set; }

		// Whether sheet numbers in this set are written with hyphens ("G-16") rather than run together ("G16").
		internal bool ExpectHyphen => _exampleHyphen ?? (IdSamples == 0 || _hyphenated * 2 >= IdSamples);

		private string? _exampleSignature;
		private bool? _exampleHyphen;

		// The user may type one sheet number as an example of the format. Only its kind of characters matters (letters, digits, separators),
		// not their number, so "HH-A-01" and "WT-A-104" are the same format.
		internal static string Signature(string text)
		{
			StringBuilder signature = new StringBuilder();
			foreach (char c in Shape(text.Trim()))
			{
				if (!(c is 'A' or '0') || signature.Length == 0 || signature[signature.Length - 1] != c)
				{
					signature.Append(c);
				}
			}
			return signature.ToString();
		}

		internal bool HasExample => _exampleSignature != null;

		internal void UseExample(string? example)
		{
			if (string.IsNullOrWhiteSpace(example) || !example.Any(char.IsLetterOrDigit))
			{
				_exampleSignature = null;
				_exampleHyphen = null;
				return;
			}
			_exampleSignature = Signature(example);
			_exampleHyphen = example.Contains('-') || example.Contains('\u2013');
		}

		internal bool MatchesExample(string id) => _exampleSignature == null || string.IsNullOrWhiteSpace(id) || Signature(IdPattern.Normalize(id)) == _exampleSignature;

		internal bool HasEnoughIds => IdSamples >= MinimumSamples;

		// See Build: sheet numbers ending in a lone 0 are not something this set uses.
		internal bool ZeroNumbersUnexpected { get; private set; }

		internal static string Shape(string text)
		{
			char[] shape = IdPattern.Normalize(text).ToUpperInvariant().Select(c => char.IsLetter(c) ? 'A' : char.IsDigit(c) ? '0' : c).ToArray();
			return new string(shape);
		}

		// Builds the profile from the rows that can be relied on. Embedded-text rows always count; OCR rows count when they were read
		// confidently; if that leaves too little, OCR rows that number consecutively with a neighbour are added.
		internal static SetProfile Build(IList<IndexRow> rows, double confidentLimit = 85.0)
		{
			SetProfile profile = new SetProfile();
			List<IndexRow> trusted = rows.Where(r => !r.IdOcr || r.IdConfidence >= confidentLimit).ToList();
			foreach (IndexRow row in trusted)
			{
				profile.AddId(row.Id);
			}
			if (!profile.HasEnoughIds)
			{
				for (int i = 0; i + 1 < rows.Count; i++)
				{
					if (trusted.Contains(rows[i]) || !IdPattern.TryParse(rows[i].Id, out string p1, out int n1, out _) || !IdPattern.TryParse(rows[i + 1].Id, out string p2, out int n2, out _))
					{
						continue;
					}
					if (p1 == p2 && n2 == n1 + 1)
					{
						profile.AddId(rows[i].Id);
					}
				}
			}
			// Sheet numbers read from the text layer show what the set really uses: if none of them is a 0 and most are two digits wide or more,
			// an OCR reading ending in a lone 0 ("RW-M-0") has most likely lost the thin 1 of "01".
			List<(int number, int width)> textNumbers = rows.Where(r => !r.IdOcr && IdPattern.TryParse(r.Id, out _, out _, out _))
				.Select(r => { IdPattern.TryParse(r.Id, out _, out int number, out int width); return (number, width); }).ToList();
			profile.ZeroNumbersUnexpected = textNumbers.Count >= 8 && textNumbers.All(t => t.number != 0) && textNumbers.Count(t => t.width >= 2) * 2 > textNumbers.Count;			foreach (IndexRow row in rows.Where(r => !r.TitleOcr && !string.IsNullOrWhiteSpace(r.Title)))
			{
				profile.TextTitleCount++;
				foreach (char c in row.Title.Distinct())
				{
					profile._titleCharacters[c] = profile._titleCharacters.GetValueOrDefault(c) + 1;
				}
			}
			return profile;
		}

		private void AddId(string id)
		{
			if (!IdPattern.TryParse(id, out string prefix, out _, out int width))
			{
				return;
			}
			IdSamples++;
			string shape = Shape(id);
			_idShapes[shape] = _idShapes.GetValueOrDefault(shape) + 1;
			_topShape = Math.Max(_topShape, _idShapes[shape]);
			string prefixShape = Shape(prefix);
			_prefixShapes[prefixShape] = _prefixShapes.GetValueOrDefault(prefixShape) + 1;
			_prefixCounts[prefix] = _prefixCounts.GetValueOrDefault(prefix) + 1;
			if (!_widthsByPrefix.TryGetValue(prefix, out HashSet<int>? widths))
			{
				widths = new HashSet<int>();
				_widthsByPrefix[prefix] = widths;
			}
			widths.Add(width);
			if (id.Contains('-') || id.Contains('\u2013'))
			{
				_hyphenated++;
			}
		}

		// Can this prefix be the start of a sheet-number series in this set? Judged by its shape against the shapes seen; with too little
		// evidence anything made of letters, digits and hyphens passes.
		internal bool PrefixFits(string prefix)
		{
			if (string.IsNullOrEmpty(prefix) || !Regex.IsMatch(prefix, @"^[A-Z0-9]+(?:-[A-Z0-9]+)*-?$"))
			{
				return false;
			}
			if (!HasEnoughIds)
			{
				return !ExpectHyphen || prefix.EndsWith('-') || !prefix.Contains('-');
			}
			return _prefixCounts.ContainsKey(prefix) || _prefixShapes.GetValueOrDefault(Shape(prefix)) >= Math.Max(2, IdSamples / 33);
		}

		// 0 (does not look like anything in this set) to 1 (looks like the set's most common sheet number).
		internal double IdFit(string cleanedId)
		{
			double learned = LearnedIdFit(cleanedId);
			if (_exampleSignature == null)
			{
				return learned;
			}
			return MatchesExample(cleanedId) ? 0.5 + 0.5 * learned : 0.2 * learned;
		}

		private double LearnedIdFit(string cleanedId)
		{
			if (!IdPattern.TryParse(cleanedId, out string prefix, out _, out int width))
			{
				return 0.15;
			}
			if (!HasEnoughIds)
			{
				return 0.5;
			}
			double shape = _topShape == 0 ? 0 : (double)_idShapes.GetValueOrDefault(Shape(cleanedId)) / _topShape;
			bool prefixKnown = _prefixCounts.ContainsKey(prefix);
			double widthFit = _widthsByPrefix.TryGetValue(prefix, out HashSet<int>? widths) ? (widths.Contains(width) ? 1.0 : 0.0) : 0.5;
			return 0.45 * shape + 0.35 * (prefixKnown ? 1.0 : (PrefixFits(prefix) ? 0.4 : 0.0)) + 0.2 * widthFit;
		}

		private static readonly string AlwaysAllowedPunctuation = " &-\u2013,()'\"#:;%+.";

		// Letters, digits and ordinary title punctuation are always fine. Anything else is accepted only if the set's own embedded-text
		// titles use it (a client that writes "TAP/SAMPLE" in its titles is not flagged for a slash).
		internal bool AllowsTitleCharacter(char c)
		{
			if (char.IsLetterOrDigit(c) || AlwaysAllowedPunctuation.Contains(c))
			{
				return true;
			}
			return TextTitleCount >= 5 && _titleCharacters.GetValueOrDefault(c) >= 2;
		}
	}
}
