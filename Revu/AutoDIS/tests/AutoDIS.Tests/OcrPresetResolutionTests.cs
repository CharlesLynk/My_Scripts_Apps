using Xunit;

namespace MbrohDrawingIndexCreator.Tests;

public sealed class OcrPresetResolutionTests
{
	[Fact]
	public void PictureResolutionAndVectorOversamplingGrowWithThePreset()
	{
		Assert.True(OcrPresets.FirstPassDpi(OcrPreset.Ultra) > OcrPresets.FirstPassDpi(OcrPreset.Fastest));
		Assert.True(OcrPresets.SecondLookDpi(OcrPreset.Quality) >= OcrPresets.FirstPassDpi(OcrPreset.Quality));
		Assert.True(OcrPresets.Oversample(OcrPreset.Ultra, false) > OcrPresets.Oversample(OcrPreset.Balanced, false));
		Assert.True(OcrPresets.Oversample(OcrPreset.Balanced, false) > OcrPresets.Oversample(OcrPreset.Fastest, false));
		Assert.True(OcrPresets.Oversample(OcrPreset.Quality, true) >= 2.0);
	}
}
