using System;
using System.Drawing;
using System.IO;
using Tesseract;
using Xunit;

#pragma warning disable CA1416 // Windows-only System.Drawing is intentional; the app and tests run on Windows.

namespace MbrohDrawingIndexCreator.Tests;

public sealed class OcrTests : IDisposable
{
	private readonly string _folder = Path.Combine(Path.GetTempPath(), "AutoDIS.Tests", Guid.NewGuid().ToString("N"));

	public OcrTests()
	{
		Directory.CreateDirectory(_folder);
	}

	public void Dispose()
	{
		try
		{
			Directory.Delete(_folder, recursive: true);
		}
		catch (IOException)
		{
		}
	}

	[Theory]
	[InlineData("eng", 4_000_000)]
	[InlineData("eng_best", 15_000_000)]
	public void ShippedModel_IsPresentAndNotTruncated(string model, long minimumBytes)
	{
		string file = Path.Combine(OcrSupport.DefaultTessdataDirectory, model + ".traineddata");

		Assert.True(File.Exists(file), $"{file} is missing; the OCR fallback cannot work without it.");
		Assert.True(new FileInfo(file).Length >= minimumBytes, $"{file} looks truncated.");
	}

	[Fact]
	public void BestModel_StartsAndReadsRenderedText()
	{
		OcrEngineResult result = OcrSupport.TryCreateEngine(OcrSupport.DefaultTessdataDirectory, OcrSupport.BestModel);
		using TesseractEngine? engine = result.Engine;

		Assert.NotNull(engine);
		Assert.Equal(OcrSupport.BestModel, result.ModelUsed);
		Assert.False(result.UsedFallbackModel);

		using Pix pix = RenderText("DESIGN CRITERIA");
		using Page page = engine!.Process(pix, PageSegMode.SingleLine);
		Assert.Contains("CRITERIA", page.GetText().ToUpperInvariant());
	}

	[Fact]
	public void MissingTessdataFolder_ReportsTheFolderInsteadOfFailingSilently()
	{
		string missing = Path.Combine(_folder, "no-such-folder");

		OcrEngineResult result = OcrSupport.TryCreateEngine(missing, OcrSupport.BestModel);

		Assert.Null(result.Engine);
		Assert.Contains(missing, result.Message);
	}

	[Fact]
	public void MissingPreferredModel_FallsBackToTheOtherModel()
	{
		string folder = Path.Combine(_folder, "tessdata");
		Directory.CreateDirectory(folder);
		File.Copy(Path.Combine(OcrSupport.DefaultTessdataDirectory, "eng.traineddata"), Path.Combine(folder, "eng.traineddata"));

		OcrEngineResult result = OcrSupport.TryCreateEngine(folder, OcrSupport.BestModel);
		using TesseractEngine? engine = result.Engine;

		Assert.NotNull(engine);
		Assert.Equal(OcrSupport.FastModel, result.ModelUsed);
		Assert.True(result.UsedFallbackModel);
		Assert.Contains("eng_best.traineddata is missing", result.Message);
	}

	[Fact]
	public void EmptyTessdataFolder_ReportsBothModelsMissing()
	{
		string folder = Path.Combine(_folder, "empty");
		Directory.CreateDirectory(folder);

		OcrEngineResult result = OcrSupport.TryCreateEngine(folder, OcrSupport.BestModel);

		Assert.Null(result.Engine);
		Assert.Contains("eng_best.traineddata is missing", result.Message);
		Assert.Contains("eng.traineddata is missing", result.Message);
	}

	[Fact]
	public void RunStats_FlagAProblemOnlyWhenOcrWasNeededAndFellShort()
	{
		OcrRunStats noOcrNeeded = new OcrRunStats { EngineAvailable = false };
		Assert.False(noOcrNeeded.HasProblem);

		OcrRunStats engineMissing = new OcrRunStats { EngineAvailable = false };
		engineMissing.AddNeedingOcr();
		Assert.True(engineMissing.HasProblem);

		OcrRunStats allRead = new OcrRunStats { EngineAvailable = true };
		allRead.AddNeedingOcr();
		allRead.AddAttempted();
		Assert.False(allRead.HasProblem);

		OcrRunStats lowConfidence = new OcrRunStats { EngineAvailable = true };
		lowConfidence.AddNeedingOcr();
		lowConfidence.AddAttempted();
		lowConfidence.AddLowConfidence("page 3: G-0R");
		Assert.True(lowConfidence.HasProblem);
		Assert.Single(lowConfidence.LowConfidenceSheets);

		OcrRunStats someMissing = new OcrRunStats { EngineAvailable = true };
		someMissing.AddNeedingOcr();
		someMissing.AddAttempted();
		someMissing.AddStillMissing();
		Assert.True(someMissing.HasProblem);
	}

	[Fact]
	public void EngineSet_LoadsTheModelThePresetNeeds()
	{
		using OcrEngineSet? best = OcrSupport.TryCreateEngineSet(OcrSupport.DefaultTessdataDirectory, useBestModel: true, out string? reason);
		Assert.NotNull(best);
		Assert.Null(reason);
		Assert.Equal(OcrSupport.BestModel, best!.Description);

		using OcrEngineSet? fast = OcrSupport.TryCreateEngineSet(OcrSupport.DefaultTessdataDirectory, useBestModel: false, out _);
		Assert.NotNull(fast);
		Assert.Equal(OcrSupport.FastModel, fast!.Description);
	}

	[Fact]
	public void EngineSet_FallsBackToTheOtherModelAndSaysSo()
	{
		string folder = Path.Combine(_folder, "onlyfast");
		Directory.CreateDirectory(folder);
		File.Copy(Path.Combine(OcrSupport.DefaultTessdataDirectory, "eng.traineddata"), Path.Combine(folder, "eng.traineddata"));

		using OcrEngineSet? set = OcrSupport.TryCreateEngineSet(folder, useBestModel: true, out string? reason);

		Assert.NotNull(set);
		Assert.Equal(OcrSupport.FastModel, set!.Description);
		Assert.Contains("eng_best.traineddata is missing", reason);
	}

	[Fact]
	public void EngineSet_IsNullWhenNoModelCanBeLoaded()
	{
		OcrEngineSet? set = OcrSupport.TryCreateEngineSet(Path.Combine(_folder, "nothing"), useBestModel: true, out string? reason);

		Assert.Null(set);
		Assert.False(string.IsNullOrEmpty(reason));
	}

	[Fact]
	public void Presets_AreOrderedFromFastToThoroughWithBalancedAsTheDefault()
	{
		Assert.Equal(3, OcrPresets.Presets.Count);
		Assert.Equal(OcrPreset.Balanced, OcrPresets.Default);
		Assert.Equal(OcrPreset.Balanced, OcrPresets.ParseOrDefault(null));
		Assert.Equal(OcrPreset.Balanced, OcrPresets.ParseOrDefault("nonsense"));
		Assert.Equal(OcrPreset.Thorough, OcrPresets.ParseOrDefault("thorough"));
		// Settings saved before the presets were reduced to three.
		Assert.Equal(OcrPreset.Fast, OcrPresets.ParseOrDefault("Fastest"));
		Assert.Equal(OcrPreset.Thorough, OcrPresets.ParseOrDefault("Quality"));
		Assert.Equal(OcrPreset.Thorough, OcrPresets.ParseOrDefault("Ultra"));
		for (int i = 0; i < OcrPresets.Presets.Count; i++)
		{
			Assert.Equal((OcrPreset)i, OcrPresets.Presets[i].Preset);
			Assert.False(string.IsNullOrWhiteSpace(OcrPresets.Presets[i].Description));
		}
	}

	[Fact]
	public void Presets_GetMoreThoroughStepByStep()
	{
		OcrPresetSettings fast = OcrPresets.Get(OcrPreset.Fast);
		OcrPresetSettings balanced = OcrPresets.Get(OcrPreset.Balanced);
		OcrPresetSettings thorough = OcrPresets.Get(OcrPreset.Thorough);

		Assert.True(fast.MaxRetries < balanced.MaxRetries);
		Assert.True(fast.TitlePasses < balanced.TitlePasses);
		Assert.False(fast.SecondLook);
		Assert.False(balanced.SecondLook);
		Assert.True(thorough.SecondLook);
	}
	[Fact]
	public void Presets_ReadTitlesFromThickenedStrokesButSmallFieldsUntouchedFirst()
	{
		OcrPresetSettings balanced = OcrPresets.Get(OcrPreset.Balanced);

		Assert.DoesNotContain(OcrVariant.Plain, balanced.OptionsFor(OcrFieldKind.Title).FirstPass);
		Assert.Contains(OcrVariant.ThickDeslant, balanced.OptionsFor(OcrFieldKind.Title).FirstPass);
		Assert.Equal(new[] { OcrVariant.Plain }, balanced.OptionsFor(OcrFieldKind.SheetId).FirstPass);
		Assert.Equal(new[] { OcrVariant.Plain }, balanced.OptionsFor(OcrFieldKind.Discipline).FirstPass);
	}
	private static Pix RenderText(string text)
	{
		using Bitmap bitmap = new Bitmap(900, 160);
		using (Graphics graphics = Graphics.FromImage(bitmap))
		using (Font font = new Font("Arial", 64f, FontStyle.Regular, GraphicsUnit.Pixel))
		{
			graphics.Clear(Color.White);
			graphics.DrawString(text, font, Brushes.Black, 20f, 30f);
		}
		using MemoryStream stream = new MemoryStream();
		bitmap.Save(stream, System.Drawing.Imaging.ImageFormat.Png);
		return Pix.LoadFromMemory(stream.ToArray());
	}
}
