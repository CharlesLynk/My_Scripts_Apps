using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Text.RegularExpressions;
using Tesseract;

namespace MbrohDrawingIndexCreator
{
	internal enum OcrFieldKind
	{
		SheetId,
		Discipline,
		Title
	}

	internal readonly record struct OcrFieldResult(string Text, double Confidence, int Variant, bool Large = false);

	// Ways of preparing a field image before it is read. CAD lettering is often hairline-thin and italic: thickening keeps thin
	// strokes (a thin italic "II" is otherwise often dropped), de-slanting straightens italics.
	internal enum OcrVariant
	{
		Plain = 0,
		Thick = 1,
		Deslant = 2,
		ThickDeslant = 3
	}

	// How hard to try on one field. The OCR presets (OcrPreset.cs) map onto these.
	// FirstPass: every variant listed is read and the best-scoring reading wins (two cheap reads catch most mistakes).
	// MaxRetries: if that result is still doubtful, up to this many of the other variants are tried, stopping at the first good one.
	internal readonly record struct OcrReadOptions(IReadOnlyList<OcrVariant> FirstPass, int MaxRetries);

	// Reads one title-block field from a cropped image.
	// No character whitelist is used (switching it per field made every read much slower); OcrTextCleanup repairs the text instead.
	internal static class OcrFieldRecognizer
	{
		private const double ItalicSlope = 0.2;

		// Printed captions that sit beside a field ("DWG", "NO.") and are never its content.
		private static readonly HashSet<string> CaptionWords = new HashSet<string>(StringComparer.Ordinal) { "DWG", "NO", "NUMBER", "DRAWING", "SHEET", "TITLE", "DATE", "SCALE", "DISCIPLINE", "REV", "REVISION", "PID", "CONTRACT" };

		private static readonly OcrVariant[] RetryOrder = { OcrVariant.Thick, OcrVariant.Deslant, OcrVariant.ThickDeslant, OcrVariant.Plain };

		// Confidence (0-100) at which a plausible reading is accepted without trying other image variants.
		internal static double AcceptConfidence(OcrFieldKind kind) => kind == OcrFieldKind.Title ? 70.0 : 85.0;

		// Capital letters of about this many pixels read best (44 beat 34 and 54 against hand-read sheets). The reader brings every line to its own working height, so a finer drawing
		// than this does not help; letters taller than TargetLetterHeight * ShrinkTolerance are shrunk to it.
		internal static int TargetLetterHeight { get; set; } = 44;
		private const double ShrinkTolerance = 1.3;

		// Tools and experiments can switch the shrinking off (0) or change the height.

		private static long _heightSum;
		private static int _heightCount;
		private static int _shrunkCount;
		
		internal static void ResetScaleStats()
		{
			Interlocked.Exchange(ref _heightSum, 0);
			Interlocked.Exchange(ref _heightCount, 0);
			Interlocked.Exchange(ref _shrunkCount, 0);
		}

		internal static string DescribeScaleStats()
		{
			int count = Volatile.Read(ref _heightCount);
			if (count == 0)
			{
				return "no letter heights measured";
			}
			return $"letter height before reading averaged {Volatile.Read(ref _heightSum) / (double)count:0} px over {count} image(s) ({Volatile.Read(ref _shrunkCount)} shrunk to about {TargetLetterHeight} px)";
		}

		private static GrayImage Normalize(GrayImage image, OcrFieldKind kind)
		{
			if (TargetLetterHeight <= 0)
			{
				return image;
			}
			GrayImage result = image.NormalizeLineHeight(TargetLetterHeight, ShrinkTolerance, out int before, out bool shrunk);
			if (before > 0)
			{
				Interlocked.Add(ref _heightSum, before);
				Interlocked.Increment(ref _heightCount);
				if (shrunk)
				{
					Interlocked.Increment(ref _shrunkCount);
				}
			}
			return result;
		}

		private static GrayImage Prepare(GrayImage cleaned, OcrVariant variant) => variant switch
		{
			OcrVariant.Thick => cleaned.Thicken(),
			OcrVariant.Deslant => cleaned.Deslant(ItalicSlope),
			OcrVariant.ThickDeslant => cleaned.Thicken().Deslant(ItalicSlope),
			_ => cleaned
		};

		internal static OcrFieldResult Read(TesseractEngine engine, GrayImage field, OcrFieldKind kind, OcrReadOptions options) => ReadAll(engine, field, kind, options, out _);

		// A single-line field (sheet number, discipline) is cut from a generous area around where it should be, so lettering drawn a little
		// high, low or to one side is still inside; the line nearest the expected place is the one read. anchorFraction is that place
		// from the top of the cut (0.5 = the middle).

		// As Read, but also returns every distinct reading the passes produced. The index analysis chooses between them using what the
		// rest of the drawing set looks like, which settles cases where the most confident reading is not the one that fits.
		internal static OcrFieldResult ReadAll(TesseractEngine engine, GrayImage field, OcrFieldKind kind, OcrReadOptions options, out List<OcrFieldResult> all, double anchorFraction = 0.5)
		{
			List<OcrFieldResult> collected = new List<OcrFieldResult>();
			all = collected;
			GrayImage lined = field.RemoveLongLines(0.25, 0.85, kind != OcrFieldKind.Title);
			GrayImage cleaned = Normalize((kind == OcrFieldKind.Title ? lined.RemoveFrameBands().KeepTextLines(false).RemoveEdgeBars() : lined.SelectNearestLine(anchorFraction).KeepMainCluster()).CropToInk(), kind);
			OcrFieldResult? best = null;
			foreach (OcrVariant variant in options.FirstPass)
			{
				OcrFieldResult candidate = Recognize(engine, Prepare(cleaned, variant), kind, (int)variant);
				collected.Add(candidate);
				if (best == null || Score(candidate, kind) > Score(best.Value, kind))
				{
					best = candidate;
				}
			}
			if (best == null)
			{
				best = Recognize(engine, cleaned, kind, 0);
				collected.Add(best.Value);
			}
			if (options.MaxRetries > 0 && !IsAcceptable(best.Value, kind))
			{
				int tries = 0;
				foreach (OcrVariant variant in RetryOrder)
				{
					if (tries >= options.MaxRetries)
					{
						break;
					}
					if (options.FirstPass.Contains(variant))
					{
						continue;
					}
					tries++;
					OcrFieldResult candidate = Recognize(engine, Prepare(cleaned, variant), kind, (int)variant);
					collected.Add(candidate);
					if (Score(candidate, kind) > Score(best.Value, kind))
					{
						best = candidate;
					}
					if (IsAcceptable(best.Value, kind))
					{
						break;
					}
				}
			}
			// A single-line field whose nearest line did not read as the field may be one of its neighbours' lines (a discipline drawn above its
			// box, a sheet number beside a "DWG NO." caption): the next nearest lines are read too and the analysis picks by fit.
			if (kind != OcrFieldKind.Title && !IsAcceptable(best.Value, kind))
			{
				for (int rank = 1; rank <= 2; rank++)
				{
					GrayImage? other = lined.SelectLine(anchorFraction, rank)?.CropToInk();
					other = other == null ? null : Normalize(other.KeepMainCluster().CropToInk(), kind);
					if (other == null)
					{
						break;
					}
					OcrFieldResult candidate = Recognize(engine, other, kind, (int)OcrVariant.Plain);
					if (CaptionWords.Contains(Regex.Replace(candidate.Text.ToUpperInvariant(), "[^A-Z]", string.Empty)))
					{
						continue;
					}
					collected.Add(candidate);
					if (Score(candidate, kind) > Score(best.Value, kind))
					{
						best = candidate;
					}
				}
			}
			return best.Value;
		}
		private static OcrFieldResult Recognize(TesseractEngine engine, GrayImage image, OcrFieldKind kind, int variant)
		{
			PageSegMode mode = kind == OcrFieldKind.Title ? PageSegMode.SingleBlock : PageSegMode.SingleLine;
			using Pix pix = Pix.LoadFromMemory(image.Pad(24).ToBmp());
			using Page page = engine.Process(pix, mode);
			string text = (page.GetText() ?? string.Empty).Trim();
			return new OcrFieldResult(text, page.GetMeanConfidence() * 100.0, variant);
		}

		internal static bool IsAcceptable(OcrFieldResult result, OcrFieldKind kind)
		{
			return result.Confidence >= AcceptConfidence(kind) && IsPlausible(result.Text, kind);
		}

		private static double Score(OcrFieldResult result, OcrFieldKind kind)
		{
			return result.Confidence + (IsPlausible(result.Text, kind) ? 20.0 : 0.0) - (string.IsNullOrWhiteSpace(result.Text) ? 100.0 : 0.0);
		}

		internal static bool IsPlausible(string text, OcrFieldKind kind)
		{
			if (string.IsNullOrWhiteSpace(text))
			{
				return false;
			}
			switch (kind)
			{
				case OcrFieldKind.SheetId:
					return OcrTextCleanup.IsPlausibleSheetId(OcrTextCleanup.CleanSheetId(text, true));
				case OcrFieldKind.Discipline:
					return Regex.IsMatch(text.Trim(), "^[A-Za-z][A-Za-z &/]{2,30}$");
				default:
					int letters = 0;
					foreach (char c in text)
					{
						if (char.IsLetter(c))
						{
							letters++;
						}
					}
					return letters >= 3 && letters >= text.Length * 0.6;
			}
		}
	}
}