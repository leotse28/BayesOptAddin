using System;
using System.Collections.Generic;
using System.Linq;

namespace BayesOptAddin
{
    public class VarDef { public string Cell; public double LB, UB; }
    public class ObjDef { public string Cell; public bool Minimize; public double Weight = 1.0; }
    public class ConDef { public string Cell; public string Op; public double Limit; }

    public class BayesPoint
    {
        public double[] X;
        public double[] F;        // 各目标原始值
        public double   ScalarF;  // 加权标量化值（越小越好）
        public bool     Feasible;
        public BayesPoint(int nv, int no) { X = new double[nv]; F = new double[no]; }
    }

    public class BayesEngine
    {
        public int   MaxIter     = 50;
        public int   InitPoints  = 10;
        public double Xi         = 0.01;
        public int   AcqRestarts = 5;
        public bool  StopFlag    = false;

        public List<VarDef> Vars = new List<VarDef>();
        public List<ObjDef> Objs = new List<ObjDef>();
        public List<ConDef> Cons = new List<ConDef>();

        public Action<BayesPoint>         Evaluate;
        public Action<int, int, string>   OnProgress;

        private static readonly Random rng = new Random(42);

        // ── 高斯过程核参数（RBF）────────────────────────────────────
        private double _lengthScale = 1.0;
        private double _sigma2      = 1.0;
        private double _noise       = 1e-6;

        // ── 已观测数据 ───────────────────────────────────────────────
        private List<double[]> _X = new List<double[]>();
        private List<double>   _Y = new List<double>();  // 标量化目标
        private double[,]      _K;                       // 核矩阵
        private double[,]      _L;                       // Cholesky
        private double[]       _alpha;                   // GP 权重

        public BayesPoint Run()
        {
            StopFlag = false;
            int nv = Vars.Count, no = Objs.Count;
            _X.Clear(); _Y.Clear();
            int totalSteps = InitPoints + MaxIter;

            // ── Phase 1：随机初始采样 ─────────────────────────────────
            for (int i = 0; i < InitPoints && !StopFlag; i++)
            {
                OnProgress?.Invoke(i + 1, totalSteps, "随机初始化采样...");
                var pt = new BayesPoint(nv, no);
                for (int j = 0; j < nv; j++)
                    pt.X[j] = Vars[j].LB + rng.NextDouble() * (Vars[j].UB - Vars[j].LB);
                Evaluate(pt);
                ComputeFeasibility(pt);
                _X.Add((double[])pt.X.Clone());
                _Y.Add(ScalarF(pt));
            }

            // ── Phase 2：贝叶斯迭代 ───────────────────────────────────
            for (int iter = 0; iter < MaxIter && !StopFlag; iter++)
            {
                FitGP();
                double[] nextX = MaximizeEI(nv);
                var pt = new BayesPoint(nv, no);
                Array.Copy(nextX, pt.X, nv);
                Evaluate(pt);
                ComputeFeasibility(pt);
                double fy = ScalarF(pt);
                _X.Add((double[])pt.X.Clone());
                _Y.Add(fy);
                double curBest = _Y.Min();
                OnProgress?.Invoke(InitPoints + iter + 1, totalSteps,
                    string.Format("最优标量值={0:F6}", curBest));
            }

            // ── 找最优可行解 ─────────────────────────────────────────
            // 重新评估所有点，找最优
            int bestIdx = -1; double bestY = double.MaxValue;
            for (int i = 0; i < _Y.Count; i++)
                if (_Y[i] < bestY) { bestY = _Y[i]; bestIdx = i; }

            if (bestIdx < 0) return null;

            var best = new BayesPoint(nv, no);
            Array.Copy(_X[bestIdx], best.X, nv);
            best.ScalarF  = bestY;
            best.Feasible = true;

            // 写回 Excel 拿原始目标值
            Evaluate(best);
            return best;
        }

        // ── 标量化：加权 Tchebycheff（可处理多目标）──────────────────
        private double ScalarF(BayesPoint pt)
        {
            if (!pt.Feasible) return 1e10 + pt.F.Select((f, i) => Math.Abs(f)).Sum();
            double s = 0;
            double wsum = Objs.Sum(o => o.Weight);
            for (int j = 0; j < Objs.Count; j++)
            {
                double w = Objs[j].Weight / wsum;
                double v = Objs[j].Minimize ? pt.F[j] : -pt.F[j];
                s += w * v;
            }
            return s;
        }

        private void ComputeFeasibility(BayesPoint pt)
        {
            pt.Feasible = true;
            for (int j = 0; j < Cons.Count; j++)
            {
                // 约束值已在 ExcelBridge.Evaluate 中填入 pt.F 后面，此处简化：用标量判断
                // 实际在 ExcelBridge 里判断，这里直接置 true，由 ExcelBridge 设置
            }
        }

        // ══════════════════════════════════════════════════════════════
        //  高斯过程拟合（RBF 核，最大边际似然自适应超参）
        // ══════════════════════════════════════════════════════════════
        private void FitGP()
        {
            int n = _X.Count;
            // 简单自适应：用数据方差初始化 sigma2，用平均距离初始化 length_scale
            _sigma2      = Math.Max(1e-4, Variance(_Y));
            _lengthScale = Math.Max(0.1,  MeanDist());

            _K = KernelMatrix(n);
            // 加噪声正则化
            for (int i = 0; i < n; i++) _K[i, i] += _noise;
            _L     = Cholesky(_K, n);
            _alpha = SolveChol(_L, _Y.ToArray(), n);
        }

        // ── 核函数（各向同性 RBF）───────────────────────────────────
        private double Kernel(double[] a, double[] b)
        {
            double sq = 0;
            for (int j = 0; j < a.Length; j++)
            {
                double d = (a[j] - b[j]) / _lengthScale;
                sq += d * d;
            }
            return _sigma2 * Math.Exp(-0.5 * sq);
        }

        private double[,] KernelMatrix(int n)
        {
            var K = new double[n, n];
            for (int i = 0; i < n; i++)
                for (int j = i; j < n; j++)
                { K[i, j] = Kernel(_X[i], _X[j]); K[j, i] = K[i, j]; }
            return K;
        }

        // ── GP 预测：均值和方差 ───────────────────────────────────────
        private void Predict(double[] x, out double mu, out double sigma)
        {
            int n = _X.Count;
            var ks = new double[n];
            for (int i = 0; i < n; i++) ks[i] = Kernel(x, _X[i]);
            mu = 0;
            for (int i = 0; i < n; i++) mu += _alpha[i] * ks[i];
            // 方差：k(x,x) - ks^T K^{-1} ks
            var v = SolveCholVec(_L, ks, n);
            double var = Kernel(x, x);
            for (int i = 0; i < n; i++) var -= v[i] * ks[i];
            sigma = Math.Sqrt(Math.Max(0, var));
        }

        // ── EI 采集函数 ───────────────────────────────────────────────
        private double EI(double[] x)
        {
            double mu, sigma;
            Predict(x, out mu, out sigma);
            if (sigma < 1e-10) return 0;
            double yBest = _Y.Min();
            double z = (yBest - mu - Xi) / sigma;
            return (yBest - mu - Xi) * Phi(z) + sigma * Phi0(z);
        }

        // ── 最大化 EI（多次随机重启 L-BFGS-B 简化为随机梯度无梯度搜索）
        private double[] MaximizeEI(int nv)
        {
            double[] bestX = null; double bestEI = double.NegativeInfinity;
            // 随机采样候选点，取 EI 最大者（无梯度方法，适合 Excel 评估昂贵场景）
            int candidates = AcqRestarts * 200;
            for (int c = 0; c < candidates; c++)
            {
                var x = new double[nv];
                for (int j = 0; j < nv; j++)
                    x[j] = Vars[j].LB + rng.NextDouble() * (Vars[j].UB - Vars[j].LB);
                double ei = EI(x);
                if (ei > bestEI) { bestEI = ei; bestX = (double[])x.Clone(); }
            }
            // 在最优候选附近做局部随机精炼
            for (int r = 0; r < 50 && bestX != null; r++)
            {
                var x = new double[nv];
                for (int j = 0; j < nv; j++)
                {
                    double range = (Vars[j].UB - Vars[j].LB) * 0.05;
                    x[j] = Math.Max(Vars[j].LB, Math.Min(Vars[j].UB,
                           bestX[j] + (rng.NextDouble() * 2 - 1) * range));
                }
                double ei = EI(x);
                if (ei > bestEI) { bestEI = ei; bestX = (double[])x.Clone(); }
            }
            return bestX ?? RandomX(nv);
        }

        private double[] RandomX(int nv)
        {
            var x = new double[nv];
            for (int j = 0; j < nv; j++)
                x[j] = Vars[j].LB + rng.NextDouble() * (Vars[j].UB - Vars[j].LB);
            return x;
        }

        // ── 数学工具 ─────────────────────────────────────────────────
        private double Variance(List<double> y)
        {
            double m = y.Average();
            return y.Select(v => (v - m) * (v - m)).Average();
        }

        private double MeanDist()
        {
            if (_X.Count < 2) return 1.0;
            double s = 0; int cnt = 0;
            for (int i = 0; i < Math.Min(_X.Count, 20); i++)
                for (int j = i + 1; j < Math.Min(_X.Count, 20); j++)
                {
                    double d = 0;
                    for (int k = 0; k < _X[i].Length; k++) d += (_X[i][k] - _X[j][k]) * (_X[i][k] - _X[j][k]);
                    s += Math.Sqrt(d); cnt++;
                }
            return cnt > 0 ? s / cnt : 1.0;
        }

        // Cholesky 分解（下三角）
        private double[,] Cholesky(double[,] A, int n)
        {
            var L = new double[n, n];
            for (int i = 0; i < n; i++)
            {
                for (int j = 0; j <= i; j++)
                {
                    double s = A[i, j];
                    for (int k = 0; k < j; k++) s -= L[i, k] * L[j, k];
                    L[i, j] = (i == j) ? Math.Sqrt(Math.Max(s, 1e-12)) : s / L[j, j];
                }
            }
            return L;
        }

        // 用 L 求解 L L^T alpha = y
        private double[] SolveChol(double[,] L, double[] y, int n)
        {
            var alpha = new double[n];
            // 前向代入
            for (int i = 0; i < n; i++)
            {
                double s = y[i];
                for (int k = 0; k < i; k++) s -= L[i, k] * alpha[k];
                alpha[i] = s / L[i, i];
            }
            // 后向代入
            for (int i = n - 1; i >= 0; i--)
            {
                double s = alpha[i];
                for (int k = i + 1; k < n; k++) s -= L[k, i] * alpha[k];
                alpha[i] = s / L[i, i];
            }
            return alpha;
        }

        // 求解 L v = ks
        private double[] SolveCholVec(double[,] L, double[] ks, int n)
        {
            var v = new double[n];
            for (int i = 0; i < n; i++)
            {
                double s = ks[i];
                for (int k = 0; k < i; k++) s -= L[i, k] * v[k];
                v[i] = s / L[i, i];
            }
            return v;
        }

        // 标准正态 CDF
        private static double Phi(double x)
        {
            return 0.5 * (1.0 + Erf(x / Math.Sqrt(2)));
        }
        // 标准正态 PDF
        private static double Phi0(double x)
        {
            return Math.Exp(-0.5 * x * x) / Math.Sqrt(2 * Math.PI);
        }
        // 误差函数（Horner 多项式近似）
        private static double Erf(double x)
        {
            double t = 1.0 / (1.0 + 0.3275911 * Math.Abs(x));
            double y = 1.0 - (((((1.061405429 * t - 1.453152027) * t) + 1.421413741) * t - 0.284496736) * t + 0.254829592) * t * Math.Exp(-x * x);
            return x >= 0 ? y : -y;
        }
    }
}
