using System.Text.RegularExpressions;
string log = args[0]; string truthFile = args[1];
static string Norm(string s) => Regex.Replace(Regex.Replace(Regex.Replace(s ?? "", "[\u2013\u2014]", "-"), "[\"\u201c\u201d\u2018\u2019']", ""), @"\s+", " ").Trim().ToUpperInvariant();
static int Lev(string a, string b) { var d = new int[a.Length + 1, b.Length + 1]; for (int i = 0; i <= a.Length; i++) d[i, 0] = i; for (int j = 0; j <= b.Length; j++) d[0, j] = j; for (int i = 1; i <= a.Length; i++) for (int j = 1; j <= b.Length; j++) d[i, j] = Math.Min(Math.Min(d[i - 1, j] + 1, d[i, j - 1] + 1), d[i - 1, j - 1] + (a[i - 1] == b[j - 1] ? 0 : 1)); return d[a.Length, b.Length]; }
var truth = File.ReadAllLines(truthFile).Where(l => l.Contains('|')).Select(l => l.Split('|')).ToDictionary(p => int.Parse(p[0]), p => p);
var lines = File.ReadAllLines(log);
var starts = Enumerable.Range(0, lines.Length).Where(i => lines[i].Contains("OCR ready")).ToList();
bool detail = args.Length > 2;
for (int k = 0; k < starts.Count; k++) {
  int e = k + 1 < starts.Count ? starts[k + 1] : lines.Length;
  var hdr = Regex.Match(lines[starts[k]], @"preset (\w+).*render DPI (\d+)");
  var rows = new Dictionary<int, string[]>();
  for (int i = starts[k]; i < e; i++) { var m = Regex.Match(lines[i], @"OCR page (\d+) of the combined set: id=\[(.*?)\] \(c\d+\) discipline=\[(.*?)\] \(c\d+\) title=\[(.*?)\] \(c\d+\)"); if (m.Success) rows[int.Parse(m.Groups[1].Value)] = new[] { m.Groups[2].Value, m.Groups[3].Value, m.Groups[4].Value }; }
  int ie = 0, de = 0, te = 0, tc = 0, n = 0;
  foreach (var (pg, t) in truth) { if (!rows.TryGetValue(pg, out var v)) continue; n++;
    if (Norm(v[0]) != Norm(t[1])) { ie++; if (detail && k == int.Parse(args[2])) Console.WriteLine($"  ID p{pg}: [{v[0]}] vs [{t[1]}]"); }
    if (Norm(v[1]) != Norm(t[2])) { de++; if (detail && k == int.Parse(args[2])) Console.WriteLine($"  DISC p{pg}: [{v[1]}] vs [{t[2]}]"); }
    int d = Lev(Norm(v[2]), Norm(t[3])); if (d != 0) { te++; tc += d; if (detail && k == int.Parse(args[2])) Console.WriteLine($"  TITLE p{pg}: [{v[2]}] vs [{t[3]}]"); } }
  Console.WriteLine($"run {k,2} {hdr.Groups[1].Value,-9}{hdr.Groups[2].Value,4}dpi  of {n}: idWrong={ie,2} discWrong={de,2} titleWrong={te,2} titleCharErrors={tc}");
}
