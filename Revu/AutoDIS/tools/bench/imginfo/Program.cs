using UglyToad.PdfPig;
using var doc = PdfDocument.Open(args[0]);
for (int p = int.Parse(args[1]); p <= int.Parse(args[2]); p++) { var page = doc.GetPage(p); Console.WriteLine($"p{p} {page.Width:0.0}x{page.Height:0.0} rot {page.Rotation.Value}"); }
