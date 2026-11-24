using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace MbrohDrawingIndexCreator
{
	// A word list (an ordinary English list plus drawing terms) used to tell a real word from an OCR misreading and to suggest what a
	// misread word probably was. Suggestions come from letters OCR is known to confuse (O/Q, I/T/L/1, A/H, V/N/W ...), so
	// "QZONE" suggests "OZONE" and "AYDRAIED" suggests "HYDRATED".
	internal sealed class OcrDictionary
	{
		internal const double ConfusionCost = 0.4;

		// Letters that are mistaken for one another when lettering is thin, italic or scanned. Each string is one group.
		private static readonly string[] ConfusionGroups = { "OQD0", "ILTJ1!|", "VUYW", "NMHA", "EF", "CG6", "S5", "Z2", "B8", "PR" };

		private static readonly Dictionary<char, char> GroupOf = BuildGroupMap();

		private readonly HashSet<string> _words;
		private readonly Lazy<Dictionary<string, List<string>>> _bySkeleton;

		private OcrDictionary(HashSet<string> words)
		{
			_words = words;
			_bySkeleton = new Lazy<Dictionary<string, List<string>>>(BuildIndex, true);
		}

		internal static OcrDictionary Empty { get; } = new OcrDictionary(new HashSet<string>(StringComparer.Ordinal));

		private static readonly Lazy<OcrDictionary> DefaultInstance = new Lazy<OcrDictionary>(LoadDefault, true);

		// The word lists shipped with the program (dictionary\*.txt); empty if they are missing, in which case only the set's own words count.
		internal static OcrDictionary Default => DefaultInstance.Value;

		internal int Count => _words.Count;

		internal static OcrDictionary FromWords(IEnumerable<string> words)
		{
			return new OcrDictionary(new HashSet<string>(words.Select(w => w.Trim().ToUpperInvariant()).Where(w => w.Length > 0), StringComparer.Ordinal));
		}

		private static OcrDictionary LoadDefault()
		{
			string folder = Path.Combine(AppContext.BaseDirectory, "dictionary");
			if (!Directory.Exists(folder))
			{
				return Empty;
			}
			HashSet<string> words = new HashSet<string>(StringComparer.Ordinal);
			foreach (string file in Directory.GetFiles(folder, "*.txt"))
			{
				foreach (string line in File.ReadLines(file))
				{
					string word = line.Trim();
					if (word.Length >= 2 && word.Length <= 24 && word[0] != '#')
					{
						words.Add(word.ToUpperInvariant());
					}
				}
			}
			return new OcrDictionary(words);
		}

		internal bool Contains(string upperWord) => _words.Contains(upperWord);

		// Real words that differ from the token only by confusable letters (same length), cheapest first. The token itself is never returned.
		internal IReadOnlyList<(string Word, double Cost)> Suggest(string token, double maxCost)
		{
			if (token.Length < 4 || _words.Count == 0)
			{
				return Array.Empty<(string, double)>();
			}
			if (!_bySkeleton.Value.TryGetValue(Skeleton(token), out List<string>? candidates))
			{
				return Array.Empty<(string, double)>();
			}
			List<(string Word, double Cost)> found = new List<(string, double)>();
			foreach (string word in candidates)
			{
				if (word == token)
				{
					continue;
				}
				double cost = 0.0;
				for (int i = 0; i < word.Length && cost <= maxCost; i++)
				{
					cost += SubstitutionCost(token[i], word[i]);
				}
				if (cost <= maxCost + 1e-9)
				{
					found.Add((word, cost));
				}
			}
			return found.OrderBy(f => f.Cost).ThenBy(f => f.Word, StringComparer.Ordinal).ToList();
		}

		// Letters that are drawn as a pair of strokes get read as one, and the other way round: an N is the I and V of "CIVIL" run together,
		// a W is two Vs, an M is an N and an I. Each rewrite stands for something OCR really does, in either direction.
		private static readonly (string Read, string Meant)[] Rewrites =
		{
			("N", "IV"), ("N", "VI"), ("W", "VV"), ("W", "IV"), ("W", "VI"), ("M", "IVI"), ("M", "NI"), ("M", "IN"), ("H", "II"), ("U", "LI"), ("D", "CL"), ("RN", "M"), ("CI", "O"), ("LI", "U")
		};

		// Real words reachable from the token by undoing one such rewrite (CNIL -> CIVIL, CWIL -> CIVIL, RNETER -> METER). The token itself is
		// never returned.
		internal IReadOnlyList<string> SuggestRewrites(string token)
		{
			if (token.Length < 3 || _words.Count == 0)
			{
				return Array.Empty<string>();
			}
			HashSet<string> found = new HashSet<string>(StringComparer.Ordinal);
			foreach ((string read, string meant) in Rewrites)
			{
				for (int from = token.IndexOf(read, StringComparison.Ordinal); from >= 0; from = token.IndexOf(read, from + 1, StringComparison.Ordinal))
				{
					string variant = token.Substring(0, from) + meant + token.Substring(from + read.Length);
					if (variant != token && _words.Contains(variant))
					{
						found.Add(variant);
					}
				}
			}
			return found.OrderBy(w => w, StringComparer.Ordinal).ToList();
		}

		internal static double SubstitutionCost(char a, char b)
		{
			if (a == b)
			{
				return 0.0;
			}
			return GroupOf.TryGetValue(a, out char ga) && GroupOf.TryGetValue(b, out char gb) && ga == gb ? ConfusionCost : 1.0;
		}

		// The word with every letter replaced by its confusion group, so look-alike words share a key.
		internal static string Skeleton(string word)
		{
			return string.Create(word.Length, word, (span, w) =>
			{
				for (int i = 0; i < w.Length; i++)
				{
					span[i] = GroupOf.TryGetValue(w[i], out char group) ? group : w[i];
				}
			});
		}

		private static Dictionary<char, char> BuildGroupMap()
		{
			Dictionary<char, char> map = new Dictionary<char, char>();
			foreach (string group in ConfusionGroups)
			{
				foreach (char c in group)
				{
					map[c] = group[0];
				}
			}
			return map;
		}

		private Dictionary<string, List<string>> BuildIndex()
		{
			Dictionary<string, List<string>> index = new Dictionary<string, List<string>>(StringComparer.Ordinal);
			foreach (string word in _words)
			{
				if (word.Length < 4 || !word.All(char.IsLetter))
				{
					continue;
				}
				string key = Skeleton(word);
				if (!index.TryGetValue(key, out List<string>? list))
				{
					list = new List<string>(2);
					index[key] = list;
				}
				list.Add(word);
			}
			return index;
		}
	}
}
