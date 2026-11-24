using System;
using System.IO;
using System.Threading;
using Xunit;

namespace AutoDIS.Tests;

public class FastCopyTests : IDisposable
{
	private readonly string _folder = Path.Combine(Path.GetTempPath(), "autodis_fastcopy_" + Guid.NewGuid().ToString("N"));

	public FastCopyTests() => Directory.CreateDirectory(_folder);

	public void Dispose() => Directory.Delete(_folder, recursive: true);

	[Fact]
	public void CopiesTheFileAndReportsProgress()
	{
		string source = Path.Combine(_folder, "a.bin");
		byte[] data = new byte[3 * 1024 * 1024 + 17];
		new Random(7).NextBytes(data);
		File.WriteAllBytes(source, data);
		string destination = Path.Combine(_folder, "b.bin");
		long lastTransferred = 0;
		long reportedTotal = 0;

		MbrohDrawingIndexCreator.FastCopy.Copy(source, destination, CancellationToken.None, (transferred, total) =>
		{
			lastTransferred = transferred;
			reportedTotal = total;
		});

		Assert.Equal(data, File.ReadAllBytes(destination));
		Assert.Equal(data.Length, reportedTotal);
		Assert.Equal(data.Length, lastTransferred);
	}

	[Fact]
	public void OverwritesAnExistingFile()
	{
		string source = Path.Combine(_folder, "a.txt");
		string destination = Path.Combine(_folder, "b.txt");
		File.WriteAllText(source, "new");
		File.WriteAllText(destination, "old and longer");

		MbrohDrawingIndexCreator.FastCopy.Copy(source, destination, CancellationToken.None, null);

		Assert.Equal("new", File.ReadAllText(destination));
	}

	[Fact]
	public void AMissingSourceIsAnIoError()
	{
		Assert.Throws<IOException>(() => MbrohDrawingIndexCreator.FastCopy.Copy(Path.Combine(_folder, "missing.bin"), Path.Combine(_folder, "x.bin"), CancellationToken.None, null));
	}

	[Fact]
	public void CancellingStopsTheCopyAndRemovesThePartialFile()
	{
		string source = Path.Combine(_folder, "big.bin");
		using (FileStream stream = File.Create(source))
		{
			stream.SetLength(200L * 1024 * 1024);
		}
		string destination = Path.Combine(_folder, "big_copy.bin");
		using CancellationTokenSource cancellation = new CancellationTokenSource();

		Assert.Throws<OperationCanceledException>(() => MbrohDrawingIndexCreator.FastCopy.Copy(source, destination, cancellation.Token, (transferred, total) => cancellation.Cancel()));

		Assert.False(File.Exists(destination));
	}
}
