using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Xunit;

namespace MbrohDrawingIndexCreator.Tests;

public sealed class RenderPoolTests
{
	private static RenderRequest Request(int page = 0) => new RenderRequest(@"C:\sets\a.pdf", page, 1913.4, 45.5, 2433.1, 295.3, 300.0);

	private sealed class FakeConnection : IRenderConnection
	{
		internal int Calls;
		internal bool Disposed;
		internal Func<RenderRequest, GrayImage?> Behaviour = r => new GrayImage(4, 2);
		public bool IsAlive { get; set; } = true;
		public GrayImage? Render(RenderRequest request)
		{
			Interlocked.Increment(ref Calls);
			return Behaviour(request);
		}
		public void Dispose() => Disposed = true;
	}

	[Fact]
	public void ARequestRoundTripsThroughTheProtocol()
	{
		using MemoryStream stream = new MemoryStream();
		RenderRequest sent = Request(41);
		RenderProtocol.WriteRequest(stream, sent);
		stream.Position = 0;

		Assert.Equal(sent, RenderProtocol.ReadRequest(stream));
	}

	[Fact]
	public void QuitAndAClosedStreamEndTheRequestLoop()
	{
		using MemoryStream quit = new MemoryStream();
		RenderProtocol.WriteQuit(quit);
		quit.Position = 0;
		Assert.Null(RenderProtocol.ReadRequest(quit));
		Assert.Null(RenderProtocol.ReadRequest(new MemoryStream()));
	}

	[Fact]
	public void AnImageRoundTripsThroughTheProtocol()
	{
		GrayImage image = new GrayImage(5, 3);
		for (int i = 0; i < image.Pixels.Length; i++)
		{
			image.Pixels[i] = (byte)(i * 17);
		}
		using MemoryStream stream = new MemoryStream();
		RenderProtocol.WriteImage(stream, image);
		stream.Position = 0;

		(GrayImage? received, string? error) = RenderProtocol.ReadResponse(stream);

		Assert.Null(error);
		Assert.Equal(5, received!.Width);
		Assert.Equal(3, received.Height);
		Assert.Equal(image.Pixels, received.Pixels);
	}

	[Fact]
	public void AFailureMessageRoundTripsThroughTheProtocol()
	{
		using MemoryStream stream = new MemoryStream();
		RenderProtocol.WriteFailure(stream, "page index out of range");
		stream.Position = 0;

		(GrayImage? image, string? error) = RenderProtocol.ReadResponse(stream);

		Assert.Null(image);
		Assert.Equal("page index out of range", error);
	}

	[Fact]
	public void ATruncatedOrNonsensicalResponseIsRejected()
	{
		using MemoryStream full = new MemoryStream();
		RenderProtocol.WriteImage(full, new GrayImage(10, 10));
		byte[] bytes = full.ToArray();
		Assert.ThrowsAny<IOException>(() => RenderProtocol.ReadResponse(new MemoryStream(bytes, 0, bytes.Length - 20)));
		Assert.Throws<InvalidDataException>(() => RenderProtocol.ReadResponse(new MemoryStream(new byte[] { 9, 0, 0, 0 })));
		using MemoryStream huge = new MemoryStream();
		using (BinaryWriter writer = new BinaryWriter(huge, System.Text.Encoding.UTF8, true))
		{
			writer.Write((byte)0);
			writer.Write(900000);
			writer.Write(900000);
		}
		huge.Position = 0;
		Assert.Throws<InvalidDataException>(() => RenderProtocol.ReadResponse(huge));
	}

	[Fact]
	public void NoHelperStartsUntilAPageIsWanted()
	{
		int created = 0;
		using RenderPool pool = new RenderPool(4, () => { created++; return new FakeConnection(); });

		Assert.Equal(0, created);
		Assert.Equal(0, pool.ProcessesStarted);
		Assert.NotNull(pool.TryRender(Request()));
		Assert.Equal(1, created);
	}

	[Fact]
	public void AnIdleHelperIsReusedBeforeAnotherIsStarted()
	{
		int created = 0;
		using RenderPool pool = new RenderPool(4, () => { created++; return new FakeConnection(); });

		for (int i = 0; i < 6; i++)
		{
			Assert.NotNull(pool.TryRender(Request(i)));
		}

		Assert.Equal(1, created);
		Assert.Equal(6, pool.Rendered);
	}

	[Fact]
	public void NeverMoreHelpersThanTheLimitRunAtOnce()
	{
		int active = 0;
		int peak = 0;
		List<FakeConnection> all = new List<FakeConnection>();
		using RenderPool pool = new RenderPool(3, () =>
		{
			FakeConnection c = new FakeConnection();
			c.Behaviour = r =>
			{
				int now = Interlocked.Increment(ref active);
				int seen;
				while (now > (seen = Volatile.Read(ref peak)) && Interlocked.CompareExchange(ref peak, now, seen) != seen)
				{
				}
				Thread.Sleep(20);
				Interlocked.Decrement(ref active);
				return new GrayImage(2, 2);
			};
			lock (all) { all.Add(c); }
			return c;
		});

		Parallel.For(0, 24, new ParallelOptions { MaxDegreeOfParallelism = 8 }, i => Assert.NotNull(pool.TryRender(Request(i))));

		Assert.InRange(peak, 2, 3);
		Assert.InRange(all.Count, 2, 3);
		Assert.Equal(24, pool.Rendered);
	}

	[Fact]
	public void ADeadHelperIsReplacedAndTheRequestFallsBackToTheCaller()
	{
		List<FakeConnection> made = new List<FakeConnection>();
		using RenderPool pool = new RenderPool(2, () =>
		{
			FakeConnection c = new FakeConnection();
			if (made.Count == 0)
			{
				c.Behaviour = r => { throw new IOException("pipe closed"); };
			}
			made.Add(c);
			return c;
		});

		Assert.Null(pool.TryRender(Request(0)));
		Assert.True(made[0].Disposed);
		Assert.NotNull(pool.TryRender(Request(1)));
		Assert.Equal(2, made.Count);
		Assert.Equal(1, pool.Rendered);
		Assert.Equal(1, pool.Declined);
	}

	[Fact]
	public void AHelperThatSaysItIsDeadAfterAnAnswerIsNotReused()
	{
		List<FakeConnection> made = new List<FakeConnection>();
		using RenderPool pool = new RenderPool(2, () =>
		{
			FakeConnection c = new FakeConnection();
			if (made.Count == 0)
			{
				c.Behaviour = r => { c.IsAlive = false; return null; };
			}
			made.Add(c);
			return c;
		});

		Assert.Null(pool.TryRender(Request(0)));
		Assert.NotNull(pool.TryRender(Request(1)));

		Assert.Equal(2, made.Count);
		Assert.True(made[0].Disposed);
	}

	[Fact]
	public void ARenderFailureOfAHealthyHelperDoesNotCountAgainstThePool()
	{
		FakeConnection connection = new FakeConnection();
		connection.Behaviour = r => r.PageIndex == 0 ? null : new GrayImage(2, 2);
		using RenderPool pool = new RenderPool(1, () => connection);

		for (int i = 0; i < 10; i++)
		{
			Assert.Null(pool.TryRender(Request(0)));
		}

		Assert.NotNull(pool.TryRender(Request(1)));
		Assert.False(pool.IsDisabled);
		Assert.False(connection.Disposed);
	}

	[Fact]
	public void APoolThatKeepsFailingSwitchesItselfOff()
	{
		int attempts = 0;
		List<string> messages = new List<string>();
		using RenderPool pool = new RenderPool(2, () => { attempts++; throw new InvalidOperationException("no exe"); }, messages.Add);

		for (int i = 0; i < 10; i++)
		{
			Assert.Null(pool.TryRender(Request(i)));
		}

		Assert.True(pool.IsDisabled);
		Assert.InRange(attempts, 1, 4);
		Assert.Contains(messages, m => m.Contains("repeatedly"));
	}

	[Fact]
	public void DisposeClosesEveryHelperAndLaterRequestsAreDeclined()
	{
		List<FakeConnection> made = new List<FakeConnection>();
		RenderPool pool = new RenderPool(3, () => { FakeConnection c = new FakeConnection(); made.Add(c); return c; });
		Parallel.For(0, 6, i => pool.TryRender(Request(i)));

		pool.Dispose();

		Assert.All(made, c => Assert.True(c.Disposed));
		Assert.Null(pool.TryRender(Request()));
	}
}
