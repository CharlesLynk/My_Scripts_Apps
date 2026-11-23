using System;
using System.IO;
using Xunit;

namespace MbrohDrawingIndexCreator.Tests;

public sealed class DiagnosticsTests : IDisposable
{
	private readonly bool _savedEnabled;
	private readonly string _savedRoot;
	private readonly string _root = Path.Combine(Path.GetTempPath(), "AutoDIS.Tests", "diag-" + Guid.NewGuid().ToString("N"));

	public DiagnosticsTests()
	{
		_savedEnabled = Diagnostics.Enabled;
		_savedRoot = Diagnostics.RootFolder;
		Diagnostics.RootFolder = _root;
	}

	public void Dispose()
	{
		Diagnostics.EndRun();
		Diagnostics.Enabled = _savedEnabled;
		Diagnostics.RootFolder = _savedRoot;
		try
		{
			Directory.Delete(_root, recursive: true);
		}
		catch (IOException)
		{
		}
	}

	[Fact]
	public void WhenDisabled_TimerIsNull_AndNothingIsWritten()
	{
		Diagnostics.Enabled = false;

		var timer = Diagnostics.Start();
		Diagnostics.BeginRun("disabled-run");
		Diagnostics.Split("step", timer);
		Diagnostics.EndRun();

		Assert.Null(timer);
		Assert.False(Directory.Exists(_root));
	}

	[Fact]
	public void WhenEnabled_WritesLogUnderRootFolder_WithSplitLabel()
	{
		Diagnostics.Enabled = true;

		var timer = Diagnostics.Start();
		Diagnostics.BeginRun("enabled-run");
		Diagnostics.Split("open sheet", timer, "sheet-1.pdf");
		Diagnostics.EndRun();

		string? logPath = Diagnostics.CurrentLogPath;
		Assert.NotNull(logPath);
		Assert.StartsWith(_root, logPath!);
		string text = File.ReadAllText(logPath!);
		Assert.Contains("open sheet", text);
		Assert.Contains("sheet-1.pdf", text);
	}

	[Fact]
	public void WhenEnabled_KeepsOnlyTwentyNewestLogs()
	{
		Diagnostics.Enabled = true;

		for (int i = 0; i < 25; i++)
		{
			Diagnostics.BeginRun("run" + i);
			Diagnostics.EndRun();
		}

		Assert.True(Directory.GetFiles(_root, "*.log").Length <= 20);
	}
}
