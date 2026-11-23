using System;
using System.IO;
using Xunit;

namespace MbrohDrawingIndexCreator.Tests;

public sealed class FileHygieneTests : IDisposable
{
	private readonly string _folder = Path.Combine(Path.GetTempPath(), "AutoDIS.Tests", Guid.NewGuid().ToString("N"));

	public FileHygieneTests()
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

	[Fact]
	public void PartialPath_IsBesideFinalFile_AndEndsWithPartial()
	{
		string finalPath = Path.Combine(_folder, "Index.pdf");

		string partial = FileHygiene.PartialPathFor(finalPath);

		Assert.Equal(_folder, Path.GetDirectoryName(partial));
		Assert.StartsWith("Index.pdf.", Path.GetFileName(partial));
		Assert.EndsWith(".partial", partial);
	}

	[Fact]
	public void Commit_ReplacesExistingFile_AndRemovesPartial()
	{
		string finalPath = Path.Combine(_folder, "Index.pdf");
		string partial = Path.Combine(_folder, "Index.pdf.partial");
		File.WriteAllText(finalPath, "old");
		File.WriteAllText(partial, "new");

		FileHygiene.Commit(partial, finalPath);

		Assert.Equal("new", File.ReadAllText(finalPath));
		Assert.False(File.Exists(partial));
	}

	[Fact]
	public void Commit_WhenDestinationIsLocked_ThrowsAndKeepsPartial()
	{
		string finalPath = Path.Combine(_folder, "Index.pdf");
		string partial = Path.Combine(_folder, "Index.pdf.partial");
		File.WriteAllText(finalPath, "old");
		File.WriteAllText(partial, "new");

		using (new FileStream(finalPath, FileMode.Open, FileAccess.Read, FileShare.None))
		{
			// Windows reports a locked destination as IOException or UnauthorizedAccessException depending on the call.
			Exception? error = Record.Exception(() => FileHygiene.Commit(partial, finalPath));
			Assert.True(error is IOException or UnauthorizedAccessException, $"Unexpected: {error}");
		}

		Assert.True(File.Exists(partial));
		Assert.Equal("new", File.ReadAllText(partial));
	}

	[Fact]
	public void SweepStale_RemovesOnlyOldMatchingFiles()
	{
		string oldMatch = Path.Combine(_folder, "AutoDIS_old.pdf");
		string newMatch = Path.Combine(_folder, "AutoDIS_new.pdf");
		string oldOther = Path.Combine(_folder, "notes_old.txt");
		File.WriteAllText(oldMatch, "x");
		File.WriteAllText(newMatch, "x");
		File.WriteAllText(oldOther, "x");
		File.SetLastWriteTime(oldMatch, DateTime.Now.AddDays(-2));
		File.SetLastWriteTime(oldOther, DateTime.Now.AddDays(-2));

		int removed = FileHygiene.SweepStale(_folder, "AutoDIS_*.pdf", TimeSpan.FromDays(1));

		Assert.Equal(1, removed);
		Assert.False(File.Exists(oldMatch));
		Assert.True(File.Exists(newMatch));
		Assert.True(File.Exists(oldOther));
	}

	[Fact]
	public void DeleteQuietly_DoesNotThrowForMissingFile()
	{
		FileHygiene.DeleteQuietly(Path.Combine(_folder, "missing.pdf"));
	}
}
