using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Tesseract;
using MbrohDrawingIndexCreator;
string pdf=@"C:\Users\clynk\AppData\Local\Temp\bachdata\bach.pdf";
string tess=@"C:\Users\clynk\Documents\My_Scripts_Apps\Revu\AutoDIS\tessdata";
double dpi=double.Parse(args[0]); if(args.Length>1) OcrFieldRecognizer.TargetLetterHeight=int.Parse(args[1]);
var preset=OcrPresets.Get(OcrPreset.Balanced);
var RI=(x1:1937.43,y1:65.46,x2:2000.94,y2:83.16); var RD=(x1:2023.8,y1:258.61,x2:2413.11,y2:275.31); var RT=(x1:2023.8,y1:208.17,x2:2413.11,y2:258.61);
(double x1,double y1,double x2,double y2) Pad((double x1,double y1,double x2,double y2) r,double px,double py)=>(r.x1-px,r.y1-py,r.x2+px,r.y2+py);
var pI=Pad(RI,Math.Max(4,(RI.x2-RI.x1)/6),Math.Max(4,0.6*(RI.y2-RI.y1))); var pD=Pad(RD,4,Math.Max(4,0.6*(RD.y2-RD.y1))); double Anchor((double x1,double y1,double x2,double y2) f,(double x1,double y1,double x2,double y2) cut)=>Math.Clamp((cut.y2-(f.y1+f.y2)/2)/(cut.y2-cut.y1),0,1); var pT=Pad(RT,6,2);
var u=(x1:Math.Min(pI.x1,Math.Min(pD.x1,pT.x1)),y1:Math.Min(pI.y1,Math.Min(pD.y1,pT.y1)),x2:Math.Max(pI.x2,Math.Max(pD.x2,pT.x2)),y2:Math.Max(pI.y2,Math.Max(pD.y2,pT.y2)));
GrayImage Crop(GrayImage img,(double x1,double y1,double x2,double y2) r){ double sx=img.Width/(u.x2-u.x1), sy=img.Height/(u.y2-u.y1); return img.Crop((int)Math.Floor((r.x1-u.x1)*sx),(int)Math.Floor((u.y2-r.y2)*sy),(int)Math.Ceiling((r.x2-r.x1)*sx),(int)Math.Ceiling((r.y2-r.y1)*sy)); }
using var e=new TesseractEngine(tess,"eng",EngineMode.LstmOnly);
var truth=File.ReadAllLines(Environment.GetEnvironmentVariable("TEMP")+@"\gt\truth.txt").Where(l=>l.Contains('|')).Select(l=>l.Split('|')).ToDictionary(p=>int.Parse(p[0]),p=>p);
if(args.Length>2&&args[2]=="idtest"){
  using var big=new TesseractEngine(tess,"eng_best",EngineMode.LstmOnly); var ultra=OcrPresets.Get(OcrPreset.Ultra);
  foreach(var a in args.Skip(3)){ int pg=int.Parse(a); var im=PdfiumRender.TryRegion(pdf,pg-1,u.x1,u.y1,u.x2,u.y2,dpi,out _)!;
    foreach(var (n,eng) in new[]{("fast",e),("big",big)}){ var r=OcrFieldRecognizer.ReadAll(eng,Crop(im,pI),OcrFieldKind.SheetId,ultra.OptionsFor(OcrFieldKind.SheetId,true),out var all,Anchor(RI,pI)); Console.WriteLine($"p{pg} dpi{dpi} {n}: "+string.Join(" | ",all.Select(x=>x.Text.Replace("\n"," ")+"@"+x.Confidence.ToString("0")))); }
    File.WriteAllBytes($@"{Environment.GetEnvironmentVariable("TEMP")}\scanbench\pid{pg}.bmp",Crop(im,pI).ToBmp()); }
  return;
}if(args.Length>2){ foreach(var pg in args.Skip(2).Select(int.Parse)){ var im=PdfiumRender.TryRegion(pdf,pg-1,u.x1,u.y1,u.x2,u.y2,dpi,out _)!; var d0=Crop(im,pD); File.WriteAllBytes($@"{Environment.GetEnvironmentVariable("TEMP")}\scanbench\d{pg}_raw.bmp",d0.ToBmp()); var l=d0.RemoveLongLines(0.25,0.85,true); File.WriteAllBytes($@"{Environment.GetEnvironmentVariable("TEMP")}\scanbench\d{pg}_lines.bmp",l.ToBmp()); var c=l.SelectNearestLine(Anchor(RD,pD)).CropToInk(); var t0=Crop(im,pT); File.WriteAllBytes($@"{Environment.GetEnvironmentVariable("TEMP")}\scanbench\t{pg}_raw.bmp",t0.ToBmp()); var t1=t0.RemoveLongLines(0.25,0.85,false).RemoveFrameBands().KeepTextLines(false).RemoveEdgeBars().CropToInk(); File.WriteAllBytes($@"{Environment.GetEnvironmentVariable("TEMP")}\scanbench\t{pg}_clean.bmp",t1.ToBmp()); File.WriteAllBytes($@"{Environment.GetEnvironmentVariable("TEMP")}\scanbench\d{pg}_line.bmp",c.ToBmp()); } return; }
foreach(var pg in truth.Keys.Where(k=>k>=338).OrderBy(k=>k)){
  var img=PdfiumRender.TryRegion(pdf,pg-1,u.x1,u.y1,u.x2,u.y2,dpi,out var err); if(img==null){Console.WriteLine($"p{pg} render failed {err}");continue;}
  var ri=OcrFieldRecognizer.ReadAll(e,Crop(img,pI),OcrFieldKind.SheetId,preset.OptionsFor(OcrFieldKind.SheetId),out _,Anchor(RI,pI));
  var rd=OcrFieldRecognizer.ReadAll(e,Crop(img,pD),OcrFieldKind.Discipline,preset.OptionsFor(OcrFieldKind.Discipline),out _,Anchor(RD,pD));
  var rt=OcrFieldRecognizer.ReadAll(e,Crop(img,pT),OcrFieldKind.Title,preset.OptionsFor(OcrFieldKind.Title),out _);
  var t=truth[pg]; Console.WriteLine($"p{pg} id[{ri.Text}] vs [{t[1]}] | disc[{rd.Text.Replace('\n',' ')}] vs [{t[2]}] | title[{rt.Text.Replace('\n',' ')}] vs [{t[3]}]");
}




