using System;
using System.Collections.Generic;

namespace MbrohDrawingIndexCreator
{
	internal enum OcrPreset
	{
		Fast,
		Balanced,
		Thorough
	}

	// The one OCR setting the user sees. Everything technical (image variants, retries, the second look with the large model)
	// is decided here, per preset.
	//
	// Every preset reads each sheet with the fast model. Measured on stroke-font and text sheets, the fast model matches the
	// large one once titles are read thickened and de-slanted, and it is about 8 times quicker. The large model is therefore
	// kept for a second look at only the fields that are still doubtful after the whole index has been checked for patterns
	// (Thorough), which costs seconds rather than minutes.
	//
	// Three presets only. Measured on the Bachman set (380 sheets), reading every field four ways at twice the resolution and giving
	// the large model a four-way second look (the old Ultra) took twice as long as Balanced and was no more accurate against hand-read
	// sheets, so the presets that differed only by a few percent of resolution or one retry were dropped.
	internal sealed record OcrPresetSettings(OcrPreset Preset, string DisplayName, string Description, bool SecondLook, int TitlePasses, int MaxRetries)
	{
		private static readonly OcrVariant[] SingleTitle = { OcrVariant.ThickDeslant };
		private static readonly OcrVariant[] DoubleTitle = { OcrVariant.Thick, OcrVariant.ThickDeslant };
		private static readonly OcrVariant[] Untouched = { OcrVariant.Plain };

		// Titles are always read from thickened strokes. The small sheet-ID and discipline fields read better untouched first
		// and only get the other variants when that first read is doubtful.
		internal OcrReadOptions OptionsFor(OcrFieldKind kind, bool secondLook = false)
		{
			if (kind == OcrFieldKind.Title)
			{
				return new OcrReadOptions(TitlePasses >= 2 ? DoubleTitle : SingleTitle, MaxRetries);
			}
			return new OcrReadOptions(Untouched, MaxRetries);
		}
	}

	internal static class OcrPresets
	{
		internal const OcrPreset Default = OcrPreset.Balanced;

		// Since the field images are brought to a fixed letter height before reading (OcrFieldRecognizer.Normalize), a finer drawing only adds
		// detail to thin or small lettering and costs little: measured against hand-read sheets, accuracy rose slowly from 300 to 800 DPI.
		internal static double FirstPassDpi(OcrPreset preset) => preset switch
		{
			OcrPreset.Fast => 300.0,
			OcrPreset.Balanced => 450.0,
			_ => 600.0
		};

		// Vector sheets: how many times the reader's working letter height the lettering is drawn at (supersampling before the shrink). Measured
		// on hand-read sheets, 1.0 already read as well as fixed 450-800 DPI; the higher presets only keep some headroom for thin lettering.
		internal static double Oversample(OcrPreset preset, bool secondLook) => secondLook ? 2.0 : preset switch
		{
			OcrPreset.Fast => 1.0,
			OcrPreset.Balanced => 1.25,
			_ => 1.5
		};

		// The second look redraws the doubtful sheets finer still.
		internal static double SecondLookDpi(OcrPreset preset) => Math.Max(FirstPassDpi(preset), 800.0);

		// Page drawing stops scaling beyond a few simultaneous renders, so more only slow each other down.
		internal const int MaxConcurrentRenders = 4;
		private static readonly OcrPresetSettings[] All =
		{
			new OcrPresetSettings(OcrPreset.Fast, "Fast", "Reads titles once and gives hard-to-read sheet numbers one more try. Best for clean title blocks with plain lettering, or for a slower computer.", false, 1, 1),
			new OcrPresetSettings(OcrPreset.Balanced, "Balanced (recommended)", "Good accuracy at high speed for most drawing sets. Titles are read two ways and hard-to-read fields get several more tries.", false, 2, 3),
			new OcrPresetSettings(OcrPreset.Thorough, "Thorough", "Balanced, then takes a second, more detailed look at every entry still in doubt using a larger reading model. Adds a little time; choose it for scanned drawings or sets with thin or italic lettering throughout.", true, 2, 3)
		};

		internal static IReadOnlyList<OcrPresetSettings> Presets => All;

		internal static OcrPresetSettings Get(OcrPreset preset) => All[(int)preset];

		// Saved settings from before the presets were reduced to three keep a sensible meaning.
		internal static OcrPreset ParseOrDefault(string? text)
		{
			switch ((text ?? string.Empty).Trim().ToLowerInvariant())
			{
				case "fastest":
					return OcrPreset.Fast;
				case "quality":
				case "ultra":
					return OcrPreset.Thorough;
				default:
					return Enum.TryParse(text, ignoreCase: true, out OcrPreset preset) && Enum.IsDefined(preset) ? preset : Default;
			}
		}
	}
}
