using System;
using System.IO;
using System.Text;

namespace MbrohDrawingIndexCreator
{
	// One page region to draw: which file and page, where on the page (page units, origin bottom-left), and at what resolution.
	internal readonly record struct RenderRequest(string Path, int PageIndex, double X1, double Y1, double X2, double Y2, double Dpi);

	// The messages between the app and a render helper process, over the helper's standard input and output.
	// Request:  'R', path, page index, region (4 doubles), dpi  |  'Q' (quit).
	// Response: 0, width, height, width*height gray bytes  |  1, message.
	internal static class RenderProtocol
	{
		private const byte RequestRender = (byte)'R';
		private const byte RequestQuit = (byte)'Q';
		private const byte ResponseOk = 0;
		private const byte ResponseFailed = 1;
		private const int MaximumSide = 30000;
		private const long MaximumPixels = 200L * 1000 * 1000;

		internal static void WriteRequest(Stream stream, RenderRequest request)
		{
			using BinaryWriter writer = new BinaryWriter(stream, Encoding.UTF8, leaveOpen: true);
			writer.Write(RequestRender);
			writer.Write(request.Path);
			writer.Write(request.PageIndex);
			writer.Write(request.X1);
			writer.Write(request.Y1);
			writer.Write(request.X2);
			writer.Write(request.Y2);
			writer.Write(request.Dpi);
			writer.Flush();
			stream.Flush();
		}

		internal static void WriteQuit(Stream stream)
		{
			using BinaryWriter writer = new BinaryWriter(stream, Encoding.UTF8, leaveOpen: true);
			writer.Write(RequestQuit);
			writer.Flush();
			stream.Flush();
		}

		// Null when the other side asked to quit or closed the stream.
		internal static RenderRequest? ReadRequest(Stream stream)
		{
			using BinaryReader reader = new BinaryReader(stream, Encoding.UTF8, leaveOpen: true);
			int kind = stream.ReadByte();
			if (kind < 0 || kind == RequestQuit)
			{
				return null;
			}
			if (kind != RequestRender)
			{
				throw new InvalidDataException($"Unexpected render request type {kind}.");
			}
			return new RenderRequest(reader.ReadString(), reader.ReadInt32(), reader.ReadDouble(), reader.ReadDouble(), reader.ReadDouble(), reader.ReadDouble(), reader.ReadDouble());
		}

		internal static void WriteImage(Stream stream, GrayImage image)
		{
			using BinaryWriter writer = new BinaryWriter(stream, Encoding.UTF8, leaveOpen: true);
			writer.Write(ResponseOk);
			writer.Write(image.Width);
			writer.Write(image.Height);
			writer.Write(image.Pixels, 0, image.Width * image.Height);
			writer.Flush();
			stream.Flush();
		}

		internal static void WriteFailure(Stream stream, string message)
		{
			using BinaryWriter writer = new BinaryWriter(stream, Encoding.UTF8, leaveOpen: true);
			writer.Write(ResponseFailed);
			writer.Write(message ?? string.Empty);
			writer.Flush();
			stream.Flush();
		}

		// Throws EndOfStreamException / IOException when the helper has gone away, InvalidDataException when it sends nonsense.
		internal static (GrayImage? image, string? error) ReadResponse(Stream stream)
		{
			using BinaryReader reader = new BinaryReader(stream, Encoding.UTF8, leaveOpen: true);
			byte status = reader.ReadByte();
			if (status == ResponseFailed)
			{
				return (null, reader.ReadString());
			}
			if (status != ResponseOk)
			{
				throw new InvalidDataException($"Unexpected render response type {status}.");
			}
			int width = reader.ReadInt32();
			int height = reader.ReadInt32();
			if (width <= 0 || height <= 0 || width > MaximumSide || height > MaximumSide || (long)width * height > MaximumPixels)
			{
				throw new InvalidDataException($"Implausible image size {width} x {height}.");
			}
			byte[] pixels = new byte[width * height];
			stream.ReadExactly(pixels, 0, pixels.Length);
			return (new GrayImage(width, height, pixels), null);
		}
	}
}
