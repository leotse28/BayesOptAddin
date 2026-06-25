using System;
using System.Collections.Generic;
using System.Linq;
namespace BayesOptAddin
{
    public class VarDef{public string Cell;public double LB,UB;}
    public class ObjDef{public string Cell;public bool Minimize;public double Weight=1.0;}
    public class ConDef{public string Cell;public string Op;public double Limit;}
    public class BayesPoint
    {
        public double[]X,F;public double ScalarF;public bool Feasible;public int BatchIndex;
        public BayesPoint(int nv,int no){X=new double[nv];F=new double[no];}
    }
    public class BayesEngine
    {
        public int MaxIter=50,InitPoints=10,AcqRestarts=5;public double Xi=0.01;public bool StopFlag=false;
        public List<VarDef>Vars=new List<VarDef>();public List<ObjDef>Objs=new List<ObjDef>();public List<ConDef>Cons=new List<ConDef>();
        public Action<BayesPoint>Evaluate;public Action<int,int,string>OnProgress;
        private static readonly Random rng=new Random();
        private double _ls=1.0,_s2=1.0,_noise=1e-6;
        private List<double[]>_X=new List<double[]>();private List<double>_Y=new List<double>();
        private double[,]_L;private double[]_alpha;
        public BayesPoint Run()
        {
            StopFlag=false;int nv=Vars.Count,no=Objs.Count;_X.Clear();_Y.Clear();int tot=InitPoints+MaxIter;
            for(int i=0;i<InitPoints&&!StopFlag;i++){OnProgress?.Invoke(i+1,tot,"随机初始化采样...");var pt=RandPt(nv,no);Evaluate(pt);_X.Add((double[])pt.X.Clone());_Y.Add(ScalarF(pt));}
            for(int iter=0;iter<MaxIter&&!StopFlag;iter++)
            {
                FitGP();double[]nx=MaxEI(nv);var pt=new BayesPoint(nv,no);Array.Copy(nx,pt.X,nv);Evaluate(pt);double fy=ScalarF(pt);_X.Add((double[])pt.X.Clone());_Y.Add(fy);
                OnProgress?.Invoke(InitPoints+iter+1,tot,string.Format("最优={0:F6}",_Y.Min()));
            }
            int bi=-1;double bv=double.MaxValue;for(int i=0;i<_Y.Count;i++)if(_Y[i]<bv){bv=_Y[i];bi=i;}
            if(bi<0)return null;
            var best=new BayesPoint(nv,no);Array.Copy(_X[bi],best.X,nv);best.ScalarF=bv;best.Feasible=true;Evaluate(best);return best;
        }
        private BayesPoint RandPt(int nv,int no){var p=new BayesPoint(nv,no);for(int j=0;j<nv;j++)p.X[j]=Vars[j].LB+rng.NextDouble()*(Vars[j].UB-Vars[j].LB);return p;}
        private double ScalarF(BayesPoint pt){double s=0,ws=Objs.Sum(o=>o.Weight);for(int j=0;j<Objs.Count;j++){double w=Objs[j].Weight/ws;double v=Objs[j].Minimize?pt.F[j]:-pt.F[j];s+=w*v;}if(!pt.Feasible)s+=1e10;return s;}
        private void FitGP(){int n=_X.Count;_s2=Math.Max(1e-4,Var(_Y));_ls=Math.Max(0.1,MDist());var K=KMat(n);for(int i=0;i<n;i++)K[i,i]+=_noise;_L=Chol(K,n);_alpha=SolveL(_L,_Y.ToArray(),n);}
        private double Ker(double[]a,double[]b){double sq=0;for(int j=0;j<a.Length;j++){double d=(a[j]-b[j])/_ls;sq+=d*d;}return _s2*Math.Exp(-0.5*sq);}
        private double[,]KMat(int n){var K=new double[n,n];for(int i=0;i<n;i++)for(int j=i;j<n;j++){K[i,j]=Ker(_X[i],_X[j]);K[j,i]=K[i,j];}return K;}
        private void Pred(double[]x,out double mu,out double sig){int n=_X.Count;var ks=new double[n];for(int i=0;i<n;i++)ks[i]=Ker(x,_X[i]);mu=0;for(int i=0;i<n;i++)mu+=_alpha[i]*ks[i];var v=SolveLvec(_L,ks,n);double var=Ker(x,x);for(int i=0;i<n;i++)var-=v[i]*ks[i];sig=Math.Sqrt(Math.Max(0,var));}
        private double EI(double[]x){double mu,sig;Pred(x,out mu,out sig);if(sig<1e-10)return 0;double yb=_Y.Min(),z=(yb-mu-Xi)/sig;return(yb-mu-Xi)*Phi(z)+sig*Phi0(z);}
        private double[]MaxEI(int nv){double[]bx=null;double bei=double.NegativeInfinity;int cands=AcqRestarts*200;for(int c=0;c<cands;c++){var x=RandX(nv);double ei=EI(x);if(ei>bei){bei=ei;bx=(double[])x.Clone();}}for(int r=0;r<50&&bx!=null;r++){var x=new double[nv];for(int j=0;j<nv;j++){double range=(Vars[j].UB-Vars[j].LB)*0.05;x[j]=Math.Max(Vars[j].LB,Math.Min(Vars[j].UB,bx[j]+(rng.NextDouble()*2-1)*range));}double ei=EI(x);if(ei>bei){bei=ei;bx=(double[])x.Clone();}}return bx??RandX(nv);}
        private double[]RandX(int nv){var x=new double[nv];for(int j=0;j<nv;j++)x[j]=Vars[j].LB+rng.NextDouble()*(Vars[j].UB-Vars[j].LB);return x;}
        private double Var(List<double>y){double m=y.Average();return y.Select(v=>(v-m)*(v-m)).Average();}
        private double MDist(){if(_X.Count<2)return 1.0;double s=0;int cnt=0;for(int i=0;i<Math.Min(_X.Count,20);i++)for(int j=i+1;j<Math.Min(_X.Count,20);j++){double d=0;for(int k=0;k<_X[i].Length;k++)d+=(_X[i][k]-_X[j][k])*(_X[i][k]-_X[j][k]);s+=Math.Sqrt(d);cnt++;}return cnt>0?s/cnt:1.0;}
        private double[,]Chol(double[,]A,int n){var L=new double[n,n];for(int i=0;i<n;i++){for(int j=0;j<=i;j++){double s=A[i,j];for(int k=0;k<j;k++)s-=L[i,k]*L[j,k];L[i,j]=(i==j)?Math.Sqrt(Math.Max(s,1e-12)):s/L[j,j];}}return L;}
        private double[]SolveL(double[,]L,double[]y,int n){var a=new double[n];for(int i=0;i<n;i++){double s=y[i];for(int k=0;k<i;k++)s-=L[i,k]*a[k];a[i]=s/L[i,i];}for(int i=n-1;i>=0;i--){double s=a[i];for(int k=i+1;k<n;k++)s-=L[k,i]*a[k];a[i]=s/L[i,i];}return a;}
        private double[]SolveLvec(double[,]L,double[]ks,int n){var v=new double[n];for(int i=0;i<n;i++){double s=ks[i];for(int k=0;k<i;k++)s-=L[i,k]*v[k];v[i]=s/L[i,i];}return v;}
        private static double Phi(double x)=>0.5*(1+Erf(x/Math.Sqrt(2)));
        private static double Phi0(double x)=>Math.Exp(-0.5*x*x)/Math.Sqrt(2*Math.PI);
        private static double Erf(double x){double t=1/(1+0.3275911*Math.Abs(x));double y=1-(((((1.061405429*t-1.453152027)*t)+1.421413741)*t-0.284496736)*t+0.254829592)*t*Math.Exp(-x*x);return x>=0?y:-y;}
    }
}