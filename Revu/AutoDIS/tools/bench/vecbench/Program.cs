using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Tesseract;
using UglyToad.PdfPig;
using MbrohDrawingIndexCreator;
string pdf=Path.Combine(Path.GetTempPath(),"bachdata","bach.pdf");
string tess=Path.Combine(AppContext.BaseDirectory,"tessdata");
string output=Path.Combine(Path.GetTempPath(),"vecbench");
Directory.CreateDirectory(output);
var presetName=args[0]; int workers=int.Parse(args[1]); double DPI=double.Parse(Environment.GetEnvironmentVariable("EV_DPI")??"300"); if(Environment.GetEnvironmentVariable("EV_TARGET") is string tg) OcrFieldRecognizer.TargetLetterHeight=int.Parse(tg); string tag=Environment.GetEnvironmentVariable("EV_TAG")??"x"; double OVER=double.Parse(Environment.GetEnvironmentVariable("EV_OVER")??"0"); double MAXDPI=double.Parse(Environment.GetEnvironmentVariable("EV_MAX")??"900"); long dpiSum=0; int dpiN=0;
var preset=OcrPresets.Get(Enum.Parse<OcrPreset>(presetName));
var RI=(x1:1937.43,y1:65.46,x2:2000.94,y2:83.16); var RD=(x1:2023.8,y1:258.61,x2:2413.11,y2:275.31); var RT=(x1:2023.8,y1:208.17,x2:2413.11,y2:258.61);
(double x1,double y1,double x2,double y2) Pad((double x1,double y1,double x2,double y2) r,double px,double py)=>(r.x1-px,r.y1-py,r.x2+px,r.y2+py);
var pI=Pad(RI,Math.Max(4,(RI.x2-RI.x1)/6),Math.Max(4,0.6*(RI.y2-RI.y1))); var pD=Pad(RD,4,Math.Max(4,0.6*(RD.y2-RD.y1))); double Anchor((double x1,double y1,double x2,double y2) f,(double x1,double y1,double x2,double y2) cut)=>Math.Clamp((cut.y2-(f.y1+f.y2)/2)/(cut.y2-cut.y1),0,1); var pT=Pad(RT,6,2);
var u=(x1:Math.Min(pI.x1,Math.Min(pD.x1,pT.x1)),y1:Math.Min(pI.y1,Math.Min(pD.y1,pT.y1)),x2:Math.Max(pI.x2,Math.Max(pD.x2,pT.x2)),y2:Math.Max(pI.y2,Math.Max(pD.y2,pT.y2)));
if(args.Length>2&&args[2]=="large"){
  var bytes0=File.ReadAllBytes(pdf); using var d0=PdfDocument.Open(bytes0); using var big=new TesseractEngine(tess,"eng_best",EngineMode.LstmOnly); using var fast=new TesseractEngine(tess,"eng",EngineMode.LstmOnly);
  foreach(var a in args.Skip(3)){ int pn=int.Parse(a); var pg=d0.GetPage(pn);
    if(!PdfPigVectorSource.TryCollect(pg,u.x1,u.y1,u.x2,u.y2,out var shapes,out var why)){Console.WriteLine($"p{pn} declined {why}");continue;}
    if(PdfPigVectorSource.TryCollect(pg,new List<(double,double,double,double)>{(pI.x1,pI.y1,pI.x2,pI.y2)},out var shapes1,out var why1)){ double pen1=VectorRaster.ChoosePenPoints(shapes1,pI.x1,pI.y1,pI.x2,pI.y2); var im1=VectorRaster.Render(shapes1,pI.x1,pI.y1,pI.x2,pI.y2,800,pen1); File.WriteAllBytes($@"{Environment.GetEnvironmentVariable("TEMP")}\vecbench\only{pn}.bmp",im1.ToBmp()); foreach(var (name,eng) in new[]{("fast",fast),("big",big)}){ var r=OcrFieldRecognizer.ReadAll(eng,im1,OcrFieldKind.SheetId,preset.OptionsFor(OcrFieldKind.SheetId,true),out var all,Anchor(RI,pI)); Console.WriteLine($"p{pn} ONLY-ID-rect shapes={shapes1.Count} pen={pen1:0.00} {name}: "+string.Join(" | ",all.Select(x=>x.Text.Replace("\n"," ")+"@"+x.Confidence.ToString("0")))); } } else Console.WriteLine($"p{pn} only-id declined {why1}");
    foreach(var ov in new[]{2.0}){
      double pen=VectorRaster.ChoosePenPoints(shapes,pI.x1,pI.y1,pI.x2,pI.y2); var probe=VectorRaster.Render(shapes,pI.x1,pI.y1,pI.x2,pI.y2,150,pen); int h=probe.TypicalLetterHeight(); double dd=Math.Clamp(150.0*44*ov/h,100,1200);
      var img=VectorRaster.Render(shapes,pI.x1,pI.y1,pI.x2,pI.y2,dd,pen);
      System.IO.File.WriteAllBytes($@"{Environment.GetEnvironmentVariable("TEMP")}\vecbench\id{pn}_{ov}.bmp",img.ToBmp());
      foreach(var (name,eng) in new[]{("fast",fast),("big",big)}){ var r=OcrFieldRecognizer.ReadAll(eng,img,OcrFieldKind.SheetId,preset.OptionsFor(OcrFieldKind.SheetId,true),out var all,Anchor(RI,pI)); Console.WriteLine($"p{pn} over{ov} h={h} dpi={dd:0} {name}: "+string.Join(" | ",all.Select(x=>x.Text.Replace("\n"," ")+"@"+x.Confidence.ToString("0")))); }
    }
  }
  return;
}byte[] bytes=File.ReadAllBytes(pdf); int n; using(var d0=PdfDocument.Open(bytes)) n=d0.NumberOfPages;
var results=new string[n+1]; int vecOk=0,declined=0,unusable=0; var reasons=new System.Collections.Concurrent.ConcurrentDictionary<string,int>(); int next=-1; long tVec=0,tOcr=0,tOpen=0;
GrayImage Crop(GrayImage img,(double x1,double y1,double x2,double y2) r){ double sx=img.Width/(u.x2-u.x1), sy=img.Height/(u.y2-u.y1); return img.Crop((int)Math.Floor((r.x1-u.x1)*sx),(int)Math.Floor((u.y2-r.y2)*sy),(int)Math.Ceiling((r.x2-r.x1)*sx),(int)Math.Ceiling((r.y2-r.y1)*sy)); }
var sw=Stopwatch.StartNew();
Parallel.For(0,workers,w=>{
  using var doc=PdfDocument.Open(bytes); using var e=new TesseractEngine(tess,"eng",EngineMode.LstmOnly);
  while(true){ int pn=Interlocked.Increment(ref next)+1; if(pn>n) break;
    var s1=Stopwatch.StartNew(); var pg=doc.GetPage(pn); Interlocked.Add(ref tOpen,s1.ElapsedMilliseconds);
    bool hasText=pg.GetWords().Any(wd=>wd.BoundingBox.Left<RT.x2&&wd.BoundingBox.Right>RT.x1&&wd.BoundingBox.Bottom<RT.y2&&wd.BoundingBox.Top>RT.y1);
    if(hasText){ string W((double x1,double y1,double x2,double y2) rc)=>string.Join(" ",pg.GetWords().Where(wd=>wd.BoundingBox.Left<rc.x2&&wd.BoundingBox.Right>rc.x1&&wd.BoundingBox.Bottom<rc.y2&&wd.BoundingBox.Top>rc.y1).OrderByDescending(wd=>Math.Round(wd.BoundingBox.Bottom/4)).ThenBy(wd=>wd.BoundingBox.Left).Select(wd=>wd.Text)); results[pn]="TEXT\t"+W(RI)+"\t"+W(RD)+"\t"+W(RT); continue; }
    s1.Restart(); GrayImage? img=null; string? why=null;
    if(PdfPigVectorSource.TryCollect(pg,u.x1,u.y1,u.x2,u.y2,out var shapes,out why)){ double pen=VectorRaster.ChoosePenPoints(shapes,u.x1,u.y1,u.x2,u.y2); img=VectorRaster.Render(shapes,u.x1,u.y1,u.x2,u.y2,DPI,pen); if(VectorRaster.InkFraction(img)<0.002){img=null;why="no ink";} }
    Interlocked.Add(ref tVec,s1.ElapsedMilliseconds);
    if(img==null){ Interlocked.Increment(ref declined); reasons.AddOrUpdate(why??"?",1,(k,v)=>v+1); results[pn]="OCR\t\t\t"; continue; }
    s1.Restart();
    GrayImage Field((double x1,double y1,double x2,double y2) pr){ if(OVER<=0) return Crop(img,pr); double pen2=VectorRaster.ChoosePenPoints(shapes,pr.x1,pr.y1,pr.x2,pr.y2); var probe=VectorRaster.Render(shapes,pr.x1,pr.y1,pr.x2,pr.y2,150,pen2); int h=probe.TypicalLetterHeight(); double d=h<=0?DPI:Math.Clamp(150.0*(OcrFieldRecognizer.TargetLetterHeight*OVER)/h,100,MAXDPI); Interlocked.Add(ref dpiSum,(long)d); Interlocked.Increment(ref dpiN); return VectorRaster.Render(shapes,pr.x1,pr.y1,pr.x2,pr.y2,d,pen2); }
    var ri=OcrFieldRecognizer.ReadAll(e,Field(pI),OcrFieldKind.SheetId,preset.OptionsFor(OcrFieldKind.SheetId),out var ai,Anchor(RI,pI));
    var rd=OcrFieldRecognizer.ReadAll(e,Field(pD),OcrFieldKind.Discipline,preset.OptionsFor(OcrFieldKind.Discipline),out var ad,Anchor(RD,pD));
    var rt=OcrFieldRecognizer.ReadAll(e,Field(pT),OcrFieldKind.Title,preset.OptionsFor(OcrFieldKind.Title),out var at);
    Interlocked.Add(ref tOcr,s1.ElapsedMilliseconds);
    bool bad=!OcrFieldRecognizer.IsPlausible(ri.Text,OcrFieldKind.SheetId)||!OcrFieldRecognizer.IsPlausible(rd.Text,OcrFieldKind.Discipline)||!OcrFieldRecognizer.IsPlausible(rt.Text,OcrFieldKind.Title);
    if(bad) { Console.WriteLine($"UNUSABLE p{pn}: id=[{ri.Text}]@{ri.Confidence:0} disc=[{rd.Text}]@{rd.Confidence:0} title=[{rt.Text.Replace("\n"," ")}]@{rt.Confidence:0}"); Interlocked.Increment(ref unusable);} else Interlocked.Increment(ref vecOk);
    string J(List<OcrFieldResult> l)=>string.Join("\u0001",l.Select(x=>x.Text.Replace("\n"," ").Replace("\t"," ")+"\u0002"+x.Confidence.ToString("0")));
    results[pn]="OCR\t"+J(ai)+"\t"+J(ad)+"\t"+J(at);
  }
});
Console.WriteLine($"avg field dpi {(dpiN==0?0:dpiSum/dpiN)}"); Console.WriteLine($"PROD {presetName} w={workers}: wall {sw.Elapsed.TotalSeconds:0.0}s open {tOpen/1000.0:0.0}s vector(sum) {tVec/1000.0:0.0}s ocr(sum) {tOcr/1000.0:0.0}s | vector usable {vecOk}, vector unusable (would fall back) {unusable}, declined {declined} [{string.Join("; ",reasons.Select(k=>k.Key+" x"+k.Value))}] | GC pause {GC.GetTotalPauseDuration().TotalSeconds:0.0}s");
File.WriteAllLines(Path.Combine(output,$"rows_ev_{presetName}_{tag}.tsv"),Enumerable.Range(1,n).Select(i=>$"{i}\t{results[i]}"));





