using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using MbrohDrawingIndexCreator;
string file=args[0]; bool verbose=args.Length>1&&args[1]=="v";
var rows=new List<IndexRow>(); var pages=new List<int>();
List<ReadingCandidate> Parse(string s)=>s.Length==0?new(){new ReadingCandidate("",0)}:s.Split('\u0001').Select(x=>{ var p=x.Split('\u0002'); return new ReadingCandidate(p[0].Trim(),double.Parse(p.Length>1?p[1]:"0"));}).ToList();
var vocab=OcrVocabulary.CreateDefault();
var parsed=new List<(int pg,string kind,string[] f)>();
foreach(var line in File.ReadLines(file,Encoding.UTF8)){ var f=line.Split('\t'); parsed.Add((int.Parse(f[0]),f[1],f)); }
// page 1 is the cover sheet
bool expectHyphen=true;
foreach(var (pg,kind,f) in parsed.Skip(1)){
  var r=new IndexRow{Label="p"+pg};
  if(kind=="TEXT"){ r.Id=f[2].Trim(); r.Discipline=f[3].Trim(); r.Title=f[4].Trim(); vocab.AddDiscipline(r.Discipline); vocab.AddTitle(r.Title); }
  else { r.IdOcr=r.DisciplineOcr=r.TitleOcr=true; r.IdCandidates=Parse(f[2]); r.DisciplineCandidates=Parse(f[3]); r.TitleCandidates=Parse(f[4]);
    var bi=r.IdCandidates.OrderByDescending(c=>c.Confidence).First(); r.IdConfidence=bi.Confidence; r.Id=OcrTextCleanup.CleanSheetId(bi.Text,expectHyphen);
    var bd=r.DisciplineCandidates.OrderByDescending(c=>c.Confidence).First(); r.DisciplineConfidence=bd.Confidence; r.Discipline=OcrTextCleanup.CleanDiscipline(bd.Text);
    var bt=r.TitleCandidates.OrderByDescending(c=>c.Confidence).First(); r.TitleConfidence=bt.Confidence; r.Title=vocab.SnapTitle(OcrTextCleanup.CleanTitle(bt.Text)); }
  rows.Add(r); pages.Add(pg);
}
var counts=new Dictionary<string,int>(); foreach(var r in rows.Where(x=>x.DisciplineOcr&&x.DisciplineConfidence>=80)){ var c=vocab.RefineDiscipline(OcrTextCleanup.CleanDiscipline(r.Discipline)); counts[c]=counts.GetValueOrDefault(c)+1; } vocab.LearnDisciplines(counts);
var sum=IndexAnalyzer.Analyze(rows,vocab);
Console.WriteLine($"rows {rows.Count} repaired {sum.IdsRepaired} titles {sum.TitlesRefined} disc {sum.DisciplinesFilled} flagged {sum.FlaggedRows}");
using var w=new StreamWriter(Path.ChangeExtension(file,".result.txt"),false,new UTF8Encoding(false));
for(int i=0;i<rows.Count;i++){ var r=rows[i]; w.WriteLine($"{pages[i]}\t{(r.IdUncertain?"?":" ")}{r.Id}\t{(r.DisciplineUncertain?"?":" ")}{r.Discipline}\t{(r.TitleUncertain?"?":" ")}{r.Title}\t{string.Join("; ",r.Reasons)}"); }


if(args.Length>2){
  static string Norm(string s)=>System.Text.RegularExpressions.Regex.Replace(System.Text.RegularExpressions.Regex.Replace(System.Text.RegularExpressions.Regex.Replace(s??"","[\u2013\u2014]","-"),"[\"\u201c\u201d\u2018\u2019']",""),@"\s+"," ").Trim().ToUpperInvariant();
  static int Lev(string a,string b){ var d=new int[a.Length+1,b.Length+1]; for(int i=0;i<=a.Length;i++)d[i,0]=i; for(int j=0;j<=b.Length;j++)d[0,j]=j; for(int i=1;i<=a.Length;i++)for(int j=1;j<=b.Length;j++)d[i,j]=Math.Min(Math.Min(d[i-1,j]+1,d[i,j-1]+1),d[i-1,j-1]+(a[i-1]==b[j-1]?0:1)); return d[a.Length,b.Length]; }
  var truth=File.ReadAllLines(args[2]).Where(l=>l.Contains('|')).Select(l=>l.Split('|')).ToDictionary(x=>int.Parse(x[0]),x=>x);
  int ie=0,de=0,te=0,tc=0,n=0;
  for(int i=0;i<rows.Count;i++){ if(pages[i]>337||!truth.TryGetValue(pages[i],out var t)) continue; var r=rows[i]; n++;
    if(Norm(r.Id)!=Norm(t[1])){ie++; if(verbose) Console.WriteLine($"  ID p{pages[i]}: [{r.Id}] vs [{t[1]}]");}
    if(Norm(r.Discipline)!=Norm(t[2])){de++; if(verbose) Console.WriteLine($"  DISC p{pages[i]}: [{r.Discipline}] vs [{t[2]}]");}
    int d=Lev(Norm(r.Title),Norm(t[3])); if(d!=0){te++;tc+=d; if(verbose) Console.WriteLine($"  TITLE p{pages[i]}: [{r.Title}] vs [{t[3]}]");} }
  Console.WriteLine($"TRUTH of {n}: idWrong={ie} discWrong={de} titleWrong={te} titleCharErrors={tc}");
}


