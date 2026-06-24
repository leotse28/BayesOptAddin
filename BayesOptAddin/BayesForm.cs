using System;
using System.Collections.Generic;
using System.Drawing;
using System.Threading;
using System.Windows.Forms;
using ExcelDna.Integration;

namespace BayesOptAddin
{
    public class BayesForm : Form
    {
        private readonly List<(Panel row, TextBox addr, TextBox lb, TextBox ub)>  _vars = new List<(Panel, TextBox, TextBox, TextBox)>();
        private readonly List<(Panel row, TextBox addr, ComboBox dir)>             _objs = new List<(Panel, TextBox, ComboBox)>();
        private readonly List<(Panel row, TextBox addr, ComboBox op, TextBox lim)> _cons = new List<(Panel, TextBox, ComboBox, TextBox)>();

        private Panel pnlVars, pnlObjs, pnlCons;
        private NumericUpDown nudIter, nudInit, nudXi, nudRestarts;
        private Button    btnRun, btnStop;
        private ProgressBar progressBar;
        private Label     lblStatus;
        private volatile bool _stopFlag;

        // ── 配色（橙金主题，区别于 NSGA-II 蓝色）────────────────────
        static readonly Color C_MAIN    = Color.FromArgb(217, 119,   6);  // 橙金
        static readonly Color C_MAIN_LT = Color.FromArgb(254, 243, 199);
        static readonly Color C_MAIN_DK = Color.FromArgb(120,  53,  15);
        static readonly Color C_RED     = Color.FromArgb(220,  38,  38);
        static readonly Color C_DEL_BG  = Color.FromArgb(255, 241, 242);
        static readonly Color C_DEL_FG  = Color.FromArgb(190,  18,  60);
        static readonly Color C_BG      = Color.FromArgb(255, 251, 235);
        static readonly Color C_HDR     = Color.FromArgb(254, 243, 199);
        static readonly Color C_WHITE   = Color.White;
        static readonly Color C_GRAY    = Color.FromArgb(120, 113,  90);

        static readonly Font F_TITLE = new Font("微软雅黑 UI", 13f, FontStyle.Bold);
        static readonly Font F_GRP   = new Font("微软雅黑 UI", 10f, FontStyle.Bold);
        static readonly Font F_BODY  = new Font("微软雅黑 UI", 10f);
        static readonly Font F_SMALL = new Font("微软雅黑 UI",  9f);
        static readonly Font F_BTN   = new Font("微软雅黑 UI", 10f, FontStyle.Bold);
        static readonly Font F_HDR   = new Font("微软雅黑 UI",  9f, FontStyle.Bold);

        const int ROW_H = 36;
        const int HDR_H = 26;

        public BayesForm()
        {
            Text          = "贝叶斯优化  Bayesian Optimization  (Office 2016)";
            Size          = new Size(740, 840);
            MinimumSize   = new Size(680, 740);
            StartPosition = FormStartPosition.CenterScreen;
            BackColor     = C_BG;
            Font          = F_BODY;
            BuildUI();
            AddVarRow(); AddVarRow();
            AddObjRow();
        }

        // ══════════════════════════════════════════════════════════════
        private void BuildUI()
        {
            var root = new TableLayoutPanel
            {
                Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 8,
                Padding = new Padding(12, 8, 12, 8), BackColor = C_BG,
            };
            root.RowStyles.Add(new RowStyle(SizeType.Absolute,  52)); // 0 标题
            root.RowStyles.Add(new RowStyle(SizeType.Percent,   28)); // 1 变量区
            root.RowStyles.Add(new RowStyle(SizeType.Absolute,  34)); // 2 变量工具栏
            root.RowStyles.Add(new RowStyle(SizeType.Percent,   20)); // 3 目标区
            root.RowStyles.Add(new RowStyle(SizeType.Absolute,  34)); // 4 目标工具栏
            root.RowStyles.Add(new RowStyle(SizeType.Percent,   16)); // 5 约束区
            root.RowStyles.Add(new RowStyle(SizeType.Absolute,  34)); // 6 约束工具栏
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 196)); // 7 参数+底部
            Controls.Add(root);

            // 标题
            var titlePnl = new Panel { Dock = DockStyle.Fill, BackColor = C_MAIN };
            titlePnl.Controls.Add(new Label
            {
                Text = "  Bayesian Optimization  贝叶斯优化",
                Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft,
                Font = F_TITLE, ForeColor = C_WHITE, BackColor = Color.Transparent,
            });
            root.Controls.Add(titlePnl, 0, 0);

            // 设计变量
            var grpVar = MakeGrp("  [设计变量]");
            var tblVar = SectionLayout(HDR_H, out pnlVars);
            tblVar.Controls.Add(MakeHdr(
                new[]{ "选取",  "单元格地址",  "下界 LB", "上界 UB" },
                new[]{   66,    184,             78,         78      }), 0, 0);
            grpVar.Controls.Add(tblVar);
            root.Controls.Add(grpVar, 0, 1);
            root.Controls.Add(Toolbar(
                () => AddVarRow(), "  + 添加变量行",
                () => DelVarRow(), "  - 删除末行",
                "点 [选取] -> 工作表点击单元格 -> 地址自动填入"), 0, 2);

            // 目标函数（贝叶斯一次只优化一个或多个加权目标）
            var grpObj = MakeGrp("  [目标函数]");
            var tblObj = SectionLayout(HDR_H, out pnlObjs);
            tblObj.Controls.Add(MakeHdr(
                new[]{ "选取",  "单元格地址（目标公式格）",  "方向",  "权重 W" },
                new[]{   66,    242,                           80,       72     }), 0, 0);
            grpObj.Controls.Add(tblObj);
            root.Controls.Add(grpObj, 0, 3);
            root.Controls.Add(Toolbar(
                () => AddObjRow(), "  + 添加目标行",
                () => DelObjRow(), "  - 删除末行",
                "多目标时按权重加权标量化  |  方向：min / max"), 0, 4);

            // 约束条件
            var grpCon = MakeGrp("  [约束条件（可选）]");
            var tblCon = SectionLayout(HDR_H, out pnlCons);
            tblCon.Controls.Add(MakeHdr(
                new[]{ "选取",  "约束单元格地址",  "符号",  "限制值" },
                new[]{   66,    186,                80,       80     }), 0, 0);
            grpCon.Controls.Add(tblCon);
            root.Controls.Add(grpCon, 0, 5);
            root.Controls.Add(Toolbar(
                () => AddConRow(), "  + 添加约束行",
                () => DelConRow(), "  - 删除末行",
                "约束公式结果 op 限制值（如 <=0）"), 0, 6);

            // 参数 + 底部
            var btmPanel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 2,
            };
            btmPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 126));
            btmPanel.RowStyles.Add(new RowStyle(SizeType.Absolute,  70));

            var paramGrp = MakeGrp("  [贝叶斯算法参数]");
            var pt = new TableLayoutPanel
            {
                Dock = DockStyle.Fill, ColumnCount = 4, Padding = new Padding(8, 4, 8, 2),
            };
            pt.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 34));
            pt.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 16));
            pt.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 34));
            pt.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 16));

            nudIter     = MakeNud(5,    500,  50,   false);
            nudInit     = MakeNud(3,    200,  10,   false);
            nudXi       = MakeNud(0.0m, 1.0m, 0.01m, true);
            nudRestarts = MakeNud(1,     20,   5,   false);

            pt.Controls.Add(PL("优化迭代次数 :"),      0, 0); pt.Controls.Add(nudIter,     1, 0);
            pt.Controls.Add(PL("初始随机采样点 :"),    2, 0); pt.Controls.Add(nudInit,     3, 0);
            pt.Controls.Add(PL("EI 探索参数 xi :"),   0, 1); pt.Controls.Add(nudXi,       1, 1);
            pt.Controls.Add(PL("采集函数重启次数 :"), 2, 1); pt.Controls.Add(nudRestarts, 3, 1);

            var tip = new Label
            {
                Text = "  推荐：迭代 >= 30，初始采样 >= 5。xi 越大越倾向探索（默认 0.01）。",
                Dock = DockStyle.Fill, AutoSize = false, ForeColor = C_GRAY,
                Font = F_SMALL, TextAlign = ContentAlignment.MiddleLeft,
            };
            pt.Controls.Add(tip, 0, 2); pt.SetColumnSpan(tip, 4);

            var tip2 = new Label
            {
                Text = "  原理：高斯过程(GP)+期望改进(EI)采集函数，每次评估后更新代理模型。",
                Dock = DockStyle.Fill, AutoSize = false, ForeColor = C_GRAY,
                Font = F_SMALL, TextAlign = ContentAlignment.MiddleLeft,
            };
            pt.Controls.Add(tip2, 0, 3); pt.SetColumnSpan(tip2, 4);
            pt.RowStyles.Add(new RowStyle(SizeType.Absolute, 20));
            pt.RowStyles.Add(new RowStyle(SizeType.Absolute, 20));
            pt.RowStyles.Add(new RowStyle(SizeType.Absolute, 20));
            pt.RowStyles.Add(new RowStyle(SizeType.Absolute, 20));

            paramGrp.Controls.Add(pt);
            btmPanel.Controls.Add(paramGrp, 0, 0);

            var runPnl = new TableLayoutPanel
            {
                Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 3,
                Padding = new Padding(0, 4, 0, 0),
            };
            runPnl.RowStyles.Add(new RowStyle(SizeType.Absolute, 42));
            runPnl.RowStyles.Add(new RowStyle(SizeType.Absolute, 20));
            runPnl.RowStyles.Add(new RowStyle(SizeType.Absolute, 18));

            var bRow = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill, FlowDirection = FlowDirection.LeftToRight,
                Padding = new Padding(0, 2, 0, 2), WrapContents = false,
            };
            btnRun  = MB("  [Ru] 运行优化", C_MAIN, C_WHITE, 140);
            btnStop = MB("  [X]  停止",     C_RED,  C_WHITE, 104);
            btnStop.Enabled = false;
            btnRun.Click  += OnRun;
            btnStop.Click += (s, e) => { _stopFlag = true; lblStatus.Text = "正在停止..."; };
            bRow.Controls.AddRange(new Control[] { btnRun, btnStop });
            runPnl.Controls.Add(bRow, 0, 0);

            progressBar = new ProgressBar { Dock = DockStyle.Fill, Minimum = 0, Maximum = 100 };
            runPnl.Controls.Add(progressBar, 0, 1);

            lblStatus = new Label
            {
                Text = "就绪，填写参数后点击 [运行优化]",
                Dock = DockStyle.Fill, AutoSize = false,
                TextAlign = ContentAlignment.MiddleLeft, ForeColor = C_GRAY, Font = F_SMALL,
            };
            runPnl.Controls.Add(lblStatus, 0, 2);
            btmPanel.Controls.Add(runPnl, 0, 1);
            root.Controls.Add(btmPanel, 0, 7);
        }

        // ── 布局工具 ─────────────────────────────────────────────────
        private TableLayoutPanel SectionLayout(int hdrHeight, out Panel dataPanel)
        {
            var tbl = new TableLayoutPanel
            {
                Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 2,
                Padding = new Padding(0),
            };
            tbl.RowStyles.Add(new RowStyle(SizeType.Absolute, hdrHeight));
            tbl.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            dataPanel = new Panel { Dock = DockStyle.Fill, AutoScroll = true };
            tbl.Controls.Add(dataPanel, 0, 1);
            return tbl;
        }

        private Panel MakeHdr(string[] titles, int[] widths)
        {
            var pnl = new Panel { Dock = DockStyle.Fill, BackColor = C_HDR };
            int x = 6;
            for (int i = 0; i < titles.Length; i++)
            {
                pnl.Controls.Add(new Label
                {
                    Text = titles[i], Left = x, Top = 4,
                    Width = widths[i], Height = HDR_H - 6,
                    Font = F_HDR, ForeColor = C_MAIN_DK,
                    TextAlign = ContentAlignment.MiddleLeft,
                });
                x += widths[i] + 4;
            }
            return pnl;
        }

        private FlowLayoutPanel Toolbar(Action addAct, string addTxt,
                                         Action delAct, string delTxt, string hint)
        {
            var row = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill, FlowDirection = FlowDirection.LeftToRight,
                Padding = new Padding(0, 3, 0, 3), WrapContents = false,
            };
            var bAdd = new Button { Text = addTxt, Width = 110, Height = 26, FlatStyle = FlatStyle.Flat, Font = F_SMALL, BackColor = C_MAIN_LT, ForeColor = C_MAIN, Margin = new Padding(0, 0, 6, 0) };
            bAdd.FlatAppearance.BorderColor = C_MAIN;
            bAdd.Click += (s, e) => addAct();
            var bDel = new Button { Text = delTxt, Width = 102, Height = 26, FlatStyle = FlatStyle.Flat, Font = F_SMALL, BackColor = C_DEL_BG, ForeColor = C_DEL_FG, Margin = new Padding(0, 0, 8, 0) };
            bDel.FlatAppearance.BorderColor = C_DEL_FG;
            bDel.Click += (s, e) => delAct();
            row.Controls.Add(bAdd);
            row.Controls.Add(bDel);
            row.Controls.Add(new Label { Text = hint, AutoSize = true, ForeColor = C_GRAY, Font = F_SMALL, TextAlign = ContentAlignment.MiddleLeft, Padding = new Padding(0, 4, 0, 0) });
            return row;
        }

        // ══════════════════════════════════════════════════════════════
        //  添加 / 删除行
        // ══════════════════════════════════════════════════════════════
        private void AddVarRow()
        {
            var rowPnl = new Panel { Left=0, Top=_vars.Count*ROW_H, Width=680, Height=ROW_H-2 };
            int x = 6;
            var btn  = PickBtn(rowPnl, x); x += 70;
            var addr = TB(rowPnl, x, 182); x += 186;
            var lb   = TB(rowPnl, x, 76);  lb.Text="0"; x += 80;
            var ub   = TB(rowPnl, x, 76);  ub.Text="1";
            WirePick(btn, addr);
            pnlVars.Controls.Add(rowPnl);
            _vars.Add((rowPnl, addr, lb, ub));
            pnlVars.AutoScrollMinSize = new Size(0, (_vars.Count+1)*ROW_H);
        }
        private void DelVarRow()
        {
            if (_vars.Count == 0) return;
            int last = _vars.Count - 1;
            pnlVars.Controls.Remove(_vars[last].row);
            _vars[last].row.Dispose(); _vars.RemoveAt(last);
            pnlVars.AutoScrollMinSize = new Size(0, (_vars.Count+1)*ROW_H);
        }

        private void AddObjRow()
        {
            var rowPnl = new Panel { Left=0, Top=_objs.Count*ROW_H, Width=680, Height=ROW_H-2 };
            int x = 6;
            var btn  = PickBtn(rowPnl, x); x += 70;
            var addr = TB(rowPnl, x, 240); x += 244;
            var dir  = new ComboBox { Left=x, Top=4, Width=78, Height=26, DropDownStyle=ComboBoxStyle.DropDownList, Font=F_BODY };
            dir.Items.AddRange(new object[]{"min","max"}); dir.SelectedIndex=0; x += 82;
            rowPnl.Controls.Add(dir);
            var wt = TB(rowPnl, x, 70); wt.Text="1.0";
            WirePick(btn, addr);
            pnlObjs.Controls.Add(rowPnl);
            _objs.Add((rowPnl, addr, dir));
            // 把权重框暂存在 Tag
            rowPnl.Tag = wt;
            pnlObjs.AutoScrollMinSize = new Size(0, (_objs.Count+1)*ROW_H);
        }
        private void DelObjRow()
        {
            if (_objs.Count == 0) return;
            int last = _objs.Count - 1;
            pnlObjs.Controls.Remove(_objs[last].row);
            _objs[last].row.Dispose(); _objs.RemoveAt(last);
            pnlObjs.AutoScrollMinSize = new Size(0, (_objs.Count+1)*ROW_H);
        }

        private void AddConRow()
        {
            var rowPnl = new Panel { Left=0, Top=_cons.Count*ROW_H, Width=680, Height=ROW_H-2 };
            int x = 6;
            var btn  = PickBtn(rowPnl, x); x += 70;
            var addr = TB(rowPnl, x, 184); x += 188;
            var op   = new ComboBox { Left=x, Top=4, Width=78, Height=26, DropDownStyle=ComboBoxStyle.DropDownList, Font=F_BODY };
            op.Items.AddRange(new object[]{"<=",">="}); op.SelectedIndex=0; x += 82;
            rowPnl.Controls.Add(op);
            var lim = TB(rowPnl, x, 78); lim.Text="0";
            WirePick(btn, addr);
            pnlCons.Controls.Add(rowPnl);
            _cons.Add((rowPnl, addr, op, lim));
            pnlCons.AutoScrollMinSize = new Size(0, (_cons.Count+1)*ROW_H);
        }
        private void DelConRow()
        {
            if (_cons.Count == 0) return;
            int last = _cons.Count - 1;
            pnlCons.Controls.Remove(_cons[last].row);
            _cons[last].row.Dispose(); _cons.RemoveAt(last);
            pnlCons.AutoScrollMinSize = new Size(0, (_cons.Count+1)*ROW_H);
        }

        // ══════════════════════════════════════════════════════════════
        //  选取按钮（后台线程 + QueueAsMacro）
        // ══════════════════════════════════════════════════════════════
        private Button PickBtn(Panel parent, int x)
        {
            var b = new Button
            {
                Text = "  [选取]", Left = x, Top = 4, Width = 64, Height = 26,
                FlatStyle = FlatStyle.Flat, Font = F_SMALL,
                BackColor = C_MAIN_LT, ForeColor = C_MAIN,
            };
            b.FlatAppearance.BorderColor = C_MAIN;
            parent.Controls.Add(b);
            return b;
        }

        private void WirePick(Button btn, TextBox addrBox)
        {
            btn.Click += (s, e) =>
            {
                btn.Enabled = false; btn.Text = "  ...";
                var bg = new Thread(() =>
                {
                    string picked = null; bool done = false; var gate = new object();
                    ExcelAsyncUtil.QueueAsMacro(() =>
                    {
                        try
                        {
                            dynamic app = ExcelDnaUtil.Application;
                            object  res = app.InputBox(
                                "请在工作表中点击目标单元格，然后点确定：",
                                "选取单元格", Type.Missing, Type.Missing,
                                Type.Missing, Type.Missing, Type.Missing, 8);
                            if (!(res is bool))
                            {
                                dynamic rng = res;
                                picked = ((string)rng.Address[false, false, 1, true]).Replace("$","");
                            }
                        }
                        catch { }
                        finally { lock (gate) { done = true; Monitor.PulseAll(gate); } }
                    });
                    lock (gate) { while (!done) Monitor.Wait(gate, 20000); }
                    addrBox.Invoke((Action)(() =>
                    {
                        if (picked != null) addrBox.Text = picked;
                        btn.Enabled = true; btn.Text = "  [选取]";
                    }));
                });
                bg.IsBackground = true; bg.Start();
            };
        }

        // ══════════════════════════════════════════════════════════════
        //  运行
        // ══════════════════════════════════════════════════════════════
        private void OnRun(object sender, EventArgs e)
        {
            BayesEngine eng;
            try { eng = BuildEngine(); }
            catch (Exception ex) { MessageBox.Show("参数错误：\n" + ex.Message, "贝叶斯优化"); return; }

            btnRun.Enabled = false; btnStop.Enabled = true;
            _stopFlag = false; progressBar.Value = 0; lblStatus.Text = "初始随机采样...";

            var t = new Thread(() =>
            {
                try
                {
                    eng.StopFlag = false;
                    eng.OnProgress = (iter, total, msg) =>
                    {
                        this.Invoke((Action)(() =>
                        {
                            progressBar.Value = (int)(100.0 * iter / total);
                            lblStatus.Text = string.Format("第 {0}/{1} 次迭代  {2}", iter, total, msg);
                        }));
                        if (_stopFlag) eng.StopFlag = true;
                    };
                    eng.Evaluate = ind =>
                    {
                        bool done = false; Exception err = null;
                        ExcelAsyncUtil.QueueAsMacro(() =>
                        {
                            try { ExcelBridge.Evaluate(ind, eng); }
                            catch (Exception ex) { err = ex; }
                            finally { lock (ind) { done = true; Monitor.PulseAll(ind); } }
                        });
                        lock (ind) { while (!done) Monitor.Wait(ind, 10000); }
                        if (err != null) throw err;
                    };

                    BayesPoint best = eng.Run();

                    ExcelAsyncUtil.QueueAsMacro(() =>
                    {
                        try { ExcelBridge.WriteResult(best, eng); }
                        catch (Exception ex) { MessageBox.Show("写结果出错：\n" + ex.Message, "贝叶斯优化"); }
                    });

                    this.Invoke((Action)(() =>
                    {
                        progressBar.Value = 100;
                        string fStr = best == null ? "N/A" : Math.Round(best.ScalarF, 6).ToString();
                        lblStatus.Text = string.Format("完成！最优标量目标值 = {0}  结果已写入 [贝叶斯结果] Sheet", fStr);
                        btnRun.Enabled = true; btnStop.Enabled = false;
                    }));
                }
                catch (Exception ex)
                {
                    this.Invoke((Action)(() =>
                    {
                        lblStatus.Text = "错误: " + ex.Message;
                        btnRun.Enabled = true; btnStop.Enabled = false;
                        MessageBox.Show("运行错误：\n" + ex.Message, "贝叶斯优化");
                    }));
                }
            });
            t.IsBackground = true; t.Start();
        }

        private BayesEngine BuildEngine()
        {
            var eng = new BayesEngine
            {
                MaxIter    = (int)nudIter.Value,
                InitPoints = (int)nudInit.Value,
                Xi         = (double)nudXi.Value,
                AcqRestarts= (int)nudRestarts.Value,
            };
            foreach (var (r, addr, lb, ub) in _vars)
            {
                string c=addr.Text.Trim(), l=lb.Text.Trim(), u=ub.Text.Trim();
                if (string.IsNullOrEmpty(c) || string.IsNullOrEmpty(l)) continue;
                eng.Vars.Add(new VarDef { Cell=c, LB=double.Parse(l), UB=string.IsNullOrEmpty(u)?1.0:double.Parse(u) });
            }
            foreach (var (r, addr, dir) in _objs)
            {
                string c = addr.Text.Trim();
                if (string.IsNullOrEmpty(c)) continue;
                double w = 1.0;
                if (r.Tag is TextBox wt && !string.IsNullOrEmpty(wt.Text.Trim()))
                    double.TryParse(wt.Text.Trim(), out w);
                eng.Objs.Add(new ObjDef { Cell=c, Minimize=dir.Text!="max", Weight=w });
            }
            foreach (var (r, addr, op, lim) in _cons)
            {
                string c = addr.Text.Trim();
                if (string.IsNullOrEmpty(c)) continue;
                eng.Cons.Add(new ConDef { Cell=c, Op=op.Text, Limit=string.IsNullOrEmpty(lim.Text)?0.0:double.Parse(lim.Text) });
            }
            if (eng.Vars.Count == 0) throw new Exception("请至少添加一个设计变量！");
            if (eng.Objs.Count == 0) throw new Exception("请至少添加一个目标函数！");
            return eng;
        }

        // ── 工具方法 ─────────────────────────────────────────────────
        private TextBox TB(Panel p, int x, int w)
        {
            var t = new TextBox { Left=x, Top=5, Width=w, Height=26, Font=F_BODY, BackColor=C_WHITE, BorderStyle=BorderStyle.FixedSingle };
            p.Controls.Add(t); return t;
        }
        private GroupBox MakeGrp(string txt)
            => new GroupBox { Text=txt, Dock=DockStyle.Fill, Font=F_GRP, ForeColor=C_MAIN };
        private NumericUpDown MakeNud(decimal min, decimal max, decimal val, bool dec)
            => new NumericUpDown { Minimum=min, Maximum=max, Value=val, Width=90, DecimalPlaces=dec?2:0, Increment=dec?0.01m:5, Font=F_BODY };
        private Label PL(string t)
            => new Label { Text=t, Dock=DockStyle.Fill, TextAlign=ContentAlignment.MiddleRight, Font=F_BODY, ForeColor=C_MAIN_DK };
        private Button MB(string text, Color back, Color fore, int width)
            => new Button { Text=text, Width=width, Height=36, BackColor=back, ForeColor=fore, FlatStyle=FlatStyle.Flat, Margin=new Padding(0,0,10,0), Font=F_BTN };
    }
}
