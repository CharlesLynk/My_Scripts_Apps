using Xunit;

namespace MbrohDrawingIndexCreator.Tests;

public sealed class OcrPresetResolutionTests
{
	[Fact]
	public void PictureResolutionAndVectorOversamplingGrowWithThePreset()
	{
		Assert.True(OcrPresets.FirstPassDpi(OcrPreset.Thorough) > OcrPresets.FirstPassDpi(OcrPreset.Fast));
		Assert.True(OcrPresets.SecondLookDpi(OcrPreset.Thorough) >= OcrPresets.FirstPassDpi(OcrPreset.Thorough));
		Assert.True(OcrPresets.Oversample(OcrPreset.Thorough, false) > OcrPresets.Oversample(OcrPreset.Balanced, false));
		Assert.True(OcrPresets.Oversample(OcrPreset.Balanced, false) > OcrPresets.Oversample(OcrPreset.Fast, false));
		Assert.True(OcrPresets.Oversample(OcrPreset.Thorough, true) >= 2.0);
	}
}
