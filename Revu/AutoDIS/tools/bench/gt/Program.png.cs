using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using MbrohDrawingIndexCreator;
string pdf = Path.Combine(Path.GetTempPath(), "bachdata", "bach.pdf");
string outDir = args[0];
Directory.CreateDirectory(outDir);
var pages = args.Skip(1).Select(int.Parse).ToList();
int perImage = 14;
for (int chunk = 0; chunk * perImage < pages.Count; chunk++) {
  var set = pages.Skip(chunk * perImage).Take(perImage).ToList();
  using var bmp = new Bitmap(1000, 160 * set.Count, PixelFormat.Format24bppRgb);
  using var g = Graphics.FromImage(bmp); g.Clear(Color.White);
  for (int i = 0; i < set.Count; i++) {
    int pn = set[i];
    var d = PdfiumRender.TryRegion(pdf, pn - 1, 2023.8, 203, 2413.11, 280, 150, out _);
    var id = PdfiumRender.TryRegion(pdf, pn - 1, 1925, 60, 2012, 90, 150, out _);
    int y = i * 160;
    g.DrawString("p" + pn, new Font("Arial", 12, FontStyle.Bold), Brushes.Red, 2, y + 2);
    if (id != null) using (var b = Bmp(id)) g.DrawImage(b, 2, y + 28);
    if (d != null) using (var b = Bmp(d)) g.DrawImage(b, 170, y + 10);
    g.DrawLine(Pens.Gray, 0, y + 159, 1000, y + 159);
  }
  bmp.Save(Path.Combine(outDir, $"sheet{chunk + 1}.png"), ImageFormat.Png);
}
static Bitmap Bmp(GrayImage im) {
  var b = new Bitmap(im.Width, im.Height, PixelFormat.Format24bppRgb);
  for (int y = 0; y < im.Height; y++) for (int x = 0; x < im.Width; x++) { int v = im.Pixels[y * im.Width + x]; b.SetPixel(x, y, Color.FromArgb(v, v, v)); }
  return b;
}
