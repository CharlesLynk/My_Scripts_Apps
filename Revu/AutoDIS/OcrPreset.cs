using System;
using System.Collections.Generic;

namespace MbrohDrawingIndexCreator
{
	internal enum OcrPreset
	{
		Fastest,
		Fast,
		Balanced,
		Quality,
		Ultra
	}

	// The one OCR setting the user sees. Everything technical (image variants, retries, the second look with the large model)
	// is decided here, per preset.
	//
	// Every preset reads each sheet with the fast model. Measured on stroke-font and text sheets, the fast model matches the
	// large one once titles are read thickened and de-slanted, and it is about 8 times quicker. The large model is therefore
	// kept for a second look at only the fields that are still doubtful after the whole index has been checked for patterns
	// (Quality and Ultra), which costs seconds rather than minutes.
	internal sealed record OcrPresetSettings(OcrPreset Preset, string DisplayName, string Description, bool SecondLook, int TitlePasses, int MaxRetries, bool AllVariants)
	{
		private static readonly OcrVariant[] SingleTitle = { OcrVariant.ThickDeslant };
		private static readonly OcrVariant[] DoubleTitle = { OcrVariant.Thick, OcrVariant.ThickDeslant };
		private static readonly OcrVariant[] Everything = { OcrVariant.Plain, OcrVariant.Thick, OcrVariant.Deslant, OcrVariant.ThickDeslant };
		private static readonly OcrVariant[] Untouched = { OcrVariant.Plain };

		// Titles are always read from thickened strokes. The small sheet-ID and discipline fields read better untouched first
		// and only get the other variants when that first read is doubtful. The second look reads every variant (Ultra).
		internal OcrReadOptions OptionsFor(OcrFieldKind kind, bool secondLook = false)
		{
			if (AllVariants || (secondLook && Preset == OcrPreset.Ultra))
			{
				return new OcrReadOptions(Everything, 0);
			}
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
			OcrPreset.Fastest => 300.0,
			OcrPreset.Fast => 400.0,
			OcrPreset.Balanced => 450.0,
			OcrPreset.Quality => 600.0,
			_ => 800.0
		};

		// Vector sheets: how many times the reader's working letter height the lettering is drawn at (supersampling before the shrink). Measured
		// on hand-read sheets, 1.0 already read as well as fixed 450-800 DPI; the higher presets only keep some headroom for thin lettering.
		internal static double Oversample(OcrPreset preset, bool secondLook) => secondLook ? 2.0 : preset switch
		{
			OcrPreset.Fastest => 1.0,
			OcrPreset.Fast => 1.0,
			OcrPreset.Balanced => 1.25,
			OcrPreset.Quality => 1.5,
			_ => 2.0
		};

		// The second look redraws the doubtful sheets finer still.
		internal static double SecondLookDpi(OcrPreset preset) => Math.Max(FirstPassDpi(preset), 800.0);

		// Page drawing stops scaling beyond a few simultaneous renders, so more only slow each other down.
		internal const int MaxConcurrentRenders = 4;
		private static readonly OcrPresetSettings[] All =
		{
			new OcrPresetSettings(OcrPreset.Fastest, "Fastest", "Reads titles once and gives hard-to-read sheet numbers one more try. Best for clean title blocks with plain lettering.", false, 1, 1, false),
			new OcrPresetSettings(OcrPreset.Fast, "Fast", "Reads titles two ways and gives hard-to-read fields up to two more tries.", false, 2, 2, false),
			new OcrPresetSettings(OcrPreset.Balanced, "Balanced (recommended)", "Good accuracy at high speed for most drawing sets. Titles are read two ways and hard-to-read fields get several more tries.", false, 2, 3, false),
			new OcrPresetSettings(OcrPreset.Quality, "Quality", "Balanced, then takes a second, more detailed look at every entry still in doubt using a larger reading model. Adds a little time; choose it for scanned drawings or sets with thin or italic lettering throughout.", true, 2, 3, false),
			new OcrPresetSettings(OcrPreset.Ultra, "Ultra", "Every field is read several ways and the best result kept, then entries still in doubt get the larger model's detailed look, also read several ways. Slowest; for the most difficult scans.", true, 2, 0, true)
		};

		internal static IReadOnlyList<OcrPresetSettings> Presets => All;

		internal static OcrPresetSettings Get(OcrPreset preset) => All[(int)preset];

		internal static OcrPreset ParseOrDefault(string? text) => Enum.TryParse(text, ignoreCase: true, out OcrPreset preset) && Enum.IsDefined(preset) ? preset : Default;
	}
}
