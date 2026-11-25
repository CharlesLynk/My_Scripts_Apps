using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using PDFiumCore;
byte[] bytes = File.ReadAllBytes(@"C:\Users\clynk\AppData\Local\Temp\bachdata\bach.pdf");
string path = @"C:\Users\clynk\AppData\Local\Temp\bachdata\bach.pdf";
double X1=1937, X2=2415, Y1=65, Y2=276; // union of the three template fields (page units, origin bottom-left)
var pages = Enumerable.Range(0, 381).Where(i => i % 9 == 2).ToList();
var sw = new Stopwatch();
using var doc = UglyToad.PdfPig.PdfDocument.Open(bytes);
long tOpen=0, tWords=0; var pig = new int[381];
foreach (var p in pages) {
  sw.Restart(); var page = doc.GetPage(p + 1); tOpen += sw.ElapsedMilliseconds;
  sw.Restart(); var n = page.GetWords().Count(w => { var b = w.BoundingBox; double cx=(b.Left+b.Right)/2, cy=(b.Bottom+b.Top)/2; return cx>=X1&&cx<=X2&&cy>=Y1&&cy<=Y2; }); tWords += sw.ElapsedMilliseconds; pig[p] = n;
}
Console.WriteLine($"PdfPig: {pages.Count} pages, GetPage {tOpen} ms, GetWords+filter {tWords} ms");
fpdfview.FPDF_InitLibrary();
sw.Restart(); var d = fpdfview.FPDF_LoadDocument(path, null); long tLoad = sw.ElapsedMilliseconds;
long tPage=0, tText=0; int agree=0, disagree=0; var dis = new System.Collections.Generic.List<string>();
foreach (var p in pages) {
  sw.Restart(); var pg = fpdfview.FPDF_LoadPage(d, p); var tp = fpdf_text.FPDFTextLoadPage(pg); tPage += sw.ElapsedMilliseconds;
  sw.Restart();
  double h = fpdfview.FPDF_GetPageHeight(pg);
  var buf = new ushort[4096];
  int len; unsafe { fixed (ushort* b = buf) { len = fpdf_text.FPDFTextGetBoundedText(tp, X1, Y2, X2, Y1, ref buf[0], buf.Length); } }
  tText += sw.ElapsedMilliseconds;
  string s = new string(buf.Take(Math.Max(0,len-1)).Select(c => (char)c).ToArray()).Trim();
  bool pdfiumHas = s.Length > 0, pigHas = pig[p] > 0;
  if (pdfiumHas == pigHas) agree++; else { disagree++; dis.Add($"p{p+1} pig={pig[p]} pdfium=[{s.Replace('\r',' ').Replace('\n',' ')}]"); }
  fpdf_text.FPDFTextClosePage(tp); fpdfview.FPDF_ClosePage(pg);
}
Console.WriteLine($"PDFium: load doc {tLoad} ms, LoadPage+TextPage {tPage} ms, bounded text {tText} ms; agree on has-text {agree}, disagree {disagree}");
foreach (var x in dis.Take(10)) Console.WriteLine(x);

