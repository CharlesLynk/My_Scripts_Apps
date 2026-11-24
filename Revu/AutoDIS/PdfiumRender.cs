using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using PDFiumCore;

namespace MbrohDrawingIndexCreator
{
	// Draws a region of a page with PDFium (Google's open-source PDF engine), the only page renderer the app uses: about 30-80 ms per CAD
	// sheet, and it decodes the JBIG2 scans that Windows' own renderer draws as blank pages. PDFium is not thread-safe, so everything here runs under one lock; the render helper processes each have their own copy.
	internal static class PdfiumRender
	{
		private static readonly object Gate = new object();
		private static readonly Dictionary<string, FpdfDocumentT> Documents = new Dictionary<string, FpdfDocumentT>(StringComparer.OrdinalIgnoreCase);
		private static bool _initialised;
		private static bool _unavailable;

		// Null when PDFium cannot be used (not available, the file or page cannot be opened); the caller then uses another renderer.
		// Region in page units of the displayed page, origin bottom-left (the same space as the template's field boxes).
		internal static GrayImage? TryRegion(string path, int pageIndex, double x1, double y1, double x2, double y2, double dpi, out string? error)
		{
			error = null;
			lock (Gate)
			{
				if (_unavailable)
				{
					error = "PDFium is not available";
					return null;
				}
				try
				{
					if (!_initialised)
					{
						fpdfview.FPDF_InitLibrary();
						_initialised = true;
					}
					if (!Documents.TryGetValue(path, out FpdfDocumentT? document))
					{
						// Folder mode opens one small file per sheet; keep only a few open.
						if (Documents.Count >= 3)
						{
							CloseAllLocked();
						}
						document = fpdfview.FPDF_LoadDocument(path, null);
						if (document == null)
						{
							error = "PDFium could not open the file";
							return null;
						}
						Documents[path] = document;
					}
					if (pageIndex < 0 || pageIndex >= fpdfview.FPDF_GetPageCount(document))
					{
						error = "page index out of range";
						return null;
					}
					return Draw(document, pageIndex, x1, y1, x2, y2, dpi, out error);
				}
				catch (Exception ex) when (ex is DllNotFoundException || ex is BadImageFormatException || ex is EntryPointNotFoundException)
				{
					_unavailable = true;
					error = "PDFium could not be loaded: " + ex.Message;
					return null;
				}
			}
		}

		private static GrayImage? Draw(FpdfDocumentT document, int pageIndex, double x1, double y1, double x2, double y2, double dpi, out string? error)
		{
			error = null;
			int width = Math.Max(1, (int)Math.Ceiling((x2 - x1) * dpi / 72.0));
			int height = Math.Max(1, (int)Math.Ceiling((y2 - y1) * dpi / 72.0));
			if ((long)width * height > 200L * 1000 * 1000)
			{
				error = "the region is too large to draw at this resolution";
				return null;
			}
			FpdfPageT? page = fpdfview.FPDF_LoadPage(document, pageIndex);
			if (page == null)
			{
				error = "PDFium could not load the page";
				return null;
			}
			FpdfBitmapT? bitmap = null;
			try
			{
				double pageHeight = fpdfview.FPDF_GetPageHeight(page);
				bitmap = fpdfview.FPDFBitmapCreate(width, height, 0);
				if (bitmap == null)
				{
					error = "PDFium could not make the picture";
					return null;
				}
				fpdfview.FPDFBitmapFillRect(bitmap, 0, 0, width, height, 0xFFFFFFFF);
				// PDFium first maps the page to a top-left origin at one pixel per point; this matrix then scales to the resolution and
				// moves the wanted region to the corner of the picture.
				float scale = (float)(dpi / 72.0);
				FS_MATRIX_ matrix = new FS_MATRIX_ { A = scale, B = 0, C = 0, D = scale, E = (float)(-x1 * scale), F = (float)(-(pageHeight - y2) * scale) };
				FS_RECTF_ clip = new FS_RECTF_ { Left = 0, Top = 0, Right = width, Bottom = height };
				fpdfview.FPDF_RenderPageBitmapWithMatrix(bitmap, page, matrix, clip, 0x01);
				IntPtr buffer = fpdfview.FPDFBitmapGetBuffer(bitmap);
				int stride = fpdfview.FPDFBitmapGetStride(bitmap);
				GrayImage image = new GrayImage(width, height);
				byte[] row = new byte[stride];
				for (int y = 0; y < height; y++)
				{
					Marshal.Copy(IntPtr.Add(buffer, y * stride), row, 0, stride);
					int offset = y * width;
					for (int x = 0; x < width; x++)
					{
						int i = x * 4;
						image.Pixels[offset + x] = (byte)((row[i] * 29 + row[i + 1] * 150 + row[i + 2] * 77) >> 8);
					}
				}
				return image;
			}
			finally
			{
				if (bitmap != null)
				{
					fpdfview.FPDFBitmapDestroy(bitmap);
				}
				fpdfview.FPDF_ClosePage(page);
			}
		}

		private static void CloseAllLocked()
		{
			foreach (FpdfDocumentT document in Documents.Values)
			{
				fpdfview.FPDF_CloseDocument(document);
			}
			Documents.Clear();
		}
	}
}
