using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;
using UglyToad.PdfPig.Graphics;
using UglyToad.PdfPig.Core;
using MbrohDrawingIndexCreator;

static class VecRaster
{
	public static double FixedWidth=double.Parse(Environment.GetEnvironmentVariable("FIXEDW")??"0"); public static double WidthFactor=double.Parse(Environment.GetEnvironmentVariable("WIDTHF")??"1");
	// Region in PDF user space (origin bottom-left); output pixels at dpi, origin top-left.
	public static GrayImage Render(IReadOnlyList<PdfPath> paths, (double x1,double y1,double x2,double y2) r, double dpi, int pad=0)
	{
		double s=dpi/72.0;
		int w=(int)Math.Ceiling((r.x2-r.x1)*s), h=(int)Math.Ceiling((r.y2-r.y1)*s);
		var cov=new float[w*h]; // ink coverage 0..1
		foreach(var p in paths)
		{
			var bb=p.GetBoundingRectangle(); if(bb==null) continue; var b=bb.Value;
			double lw=Math.Max(p.IsStroked?(FixedWidth>0?FixedWidth:p.LineWidth*WidthFactor):0,0);
			double m=lw/2+1;
			if(b.Right<r.x1-m||b.Left>r.x2+m||b.Top<r.y1-m||b.Bottom>r.y2+m) continue;
			double half=Math.Max(lw*s/2,0.5);
			foreach(var sp in p)
			{
				var pts=new List<(double x,double y)>();
				PdfPoint cur=default; PdfPoint start=default; bool have=false;
				void Flush(bool closed){ if(pts.Count>1){ if(p.IsStroked||!p.IsFilled){ for(int i=0;i+1<pts.Count;i++) Seg(cov,w,h,pts[i],pts[i+1],half); if(closed) Seg(cov,w,h,pts[^1],pts[0],half);} if(p.IsFilled) Fill(cov,w,h,pts);} pts=new List<(double,double)>(); }
				(double x,double y) T(PdfPoint q)=>((q.X-r.x1)*s,(r.y2-q.Y)*s);
				foreach(var c in sp.Commands)
				{
					switch(c)
					{
						case PdfSubpath.Move mv: Flush(false); pts.Add(T(mv.Location)); cur=mv.Location; start=cur; break;
						case PdfSubpath.Line ln: if(pts.Count==0) pts.Add(T(ln.From)); pts.Add(T(ln.To)); cur=ln.To; break;
						case PdfSubpath.BezierCurve bz: { if(pts.Count==0) pts.Add(T(bz.StartPoint)); foreach(var seg in bz.ToLines(8)) pts.Add(T(seg.To)); cur=bz.EndPoint; break; } case PdfSubpath.Close: Flush(true); break;
					}
				}
				Flush(false);
			}
		}
		var img=new GrayImage(w,h);
		for(int i=0;i<w*h;i++) img.Pixels[i]=(byte)(255-Math.Round(Math.Min(1f,cov[i])*255));
		return img;
	}
	static void Seg(float[] cov,int w,int h,(double x,double y) a,(double x,double y) b,double half)
	{
		double x0=Math.Min(a.x,b.x)-half-1,x1=Math.Max(a.x,b.x)+half+1,y0=Math.Min(a.y,b.y)-half-1,y1=Math.Max(a.y,b.y)+half+1;
		int ix0=Math.Max(0,(int)Math.Floor(x0)),ix1=Math.Min(w-1,(int)Math.Ceiling(x1)),iy0=Math.Max(0,(int)Math.Floor(y0)),iy1=Math.Min(h-1,(int)Math.Ceiling(y1));
		double dx=b.x-a.x,dy=b.y-a.y,len2=dx*dx+dy*dy;
		for(int y=iy0;y<=iy1;y++) for(int x=ix0;x<=ix1;x++)
		{
			double px=x+0.5-a.x,py=y+0.5-a.y; double t=len2<1e-9?0:Math.Clamp((px*dx+py*dy)/len2,0,1);
			double ex=px-t*dx,ey=py-t*dy; double d=Math.Sqrt(ex*ex+ey*ey);
			float c=(float)Math.Clamp(half+0.5-d,0,1); if(c>cov[y*w+x]) cov[y*w+x]=c;
		}
	}
	static void Fill(float[] cov,int w,int h,List<(double x,double y)> pts)
	{
		if(pts.Count<3) return; double miny=pts.Min(p=>p.y),maxy=pts.Max(p=>p.y);
		for(int y=Math.Max(0,(int)Math.Floor(miny));y<=Math.Min(h-1,(int)Math.Ceiling(maxy));y++)
		{
			double sy=y+0.5; var xs=new List<double>();
			for(int i=0;i<pts.Count;i++){ var a=pts[i]; var b=pts[(i+1)%pts.Count]; if((a.y<=sy&&b.y>sy)||(b.y<=sy&&a.y>sy)){ xs.Add(a.x+(sy-a.y)/(b.y-a.y)*(b.x-a.x)); } }
			xs.Sort(); for(int k=0;k+1<xs.Count;k+=2){ int xa=Math.Max(0,(int)Math.Round(xs[k])),xb=Math.Min(w-1,(int)Math.Round(xs[k+1])); for(int x=xa;x<=xb;x++) cov[y*w+x]=1f; }
		}
	}
}





