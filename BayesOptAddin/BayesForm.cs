using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using ExcelDna.Integration;

namespace BayesOptAddin
{
    // ── 极简 JSON 读写（无外部依赖）─────────────────────────────────
    public static class SimpleJson
    {
        public static string Escape(string s) => s.Replace("\\","\\\\").Replace("\"","\\\"");
        public static string Unescape(string s) => s.Replace("\\\"","\"").Replace("\\\\","\\");
        public static string Obj(params string[] kvPairs)
        {
            var sb=new StringBuilder("{");
            for(int i=0;i<kvPairs.Length;i+=2){if(i>0)sb.Append(",");sb.Append("\"").Append(Escape(kvPairs[i])).Append("\":\"").Append(Escape(kvPairs[i+1])).Append("\"");}
            sb.Append("}"); return sb.ToString();
        }
        public static string Get(string json, string key)
        {
            string search="\""+key+"\":\"";int idx=json.IndexOf(search);if(idx<0)return "";
            int start=idx+search.Length,end=json.IndexOf("\"",start);if(end<0)return "";
            return Unescape(json.Substring(start,end-start));
        }
        public static List<string> SplitArray(string json)
        {
            var list=new List<string>();int depth=0,start=-1;
            for(int i=0;i<json.Length;i++){if(json[i]=='{'){if(depth==0)start=i;depth++;}else if(json[i]=='}'){depth--;if(depth==0&&start>=0){list.Add(json.Substring(start,i-start+1));start=-1;}}}
            return list;
        }
    }

    public class BayesForm : Form
    {
        private readonly List<(Panel row, TextBox addr, TextBox lb, TextBox ub)>  _vars = new List<(Panel, TextBox, TextBox, TextBox)>();
        private readonly List<(Panel row, TextBox addr, ComboBox dir)>             _objs = new List<(Panel, TextBox, ComboBox)>();
        private readonly List<(Panel row, TextBox addr, ComboBox op, TextBox lim)> _cons = new List<(Panel, TextBox, ComboBox, TextBox)>();

        private Panel pnlVars, pnlObjs, pnlCons;
        private NumericUpDown nudIter, nudInit, nudXi, nudRestarts;
        private NumericUpDown nudBatch;   // 批量次数
        private Button    btnRun, btnStop, btnSave, btnLoad;
        private ProgressBar progressBar;
        private Label     lblStatus;
        private volatile bool _stopFlag;

        static readonly Color C_MAIN    = Color.FromArgb(217, 119,   6);
        static readonly Color C_MAIN_LT = Color.FromArgb(254, 243, 199);
        static readonly Color C_MAIN_DK = Color.FromArgb(120,  53,  15);
        static readonly Color C_RED     = Color.FromArgb(220,  38,  38);
        static readonly Color C_DEL_BG  = Color.FromArgb(255, 241, 242);
        static readonly Color C_DEL_FG  = Color.FromArgb(190,  18,  60);
        static readonly Color C_BG      = Color.FromArgb(255, 251, 235);
        static readonly Color C_HDR     = Color.FromArgb(254, 243, 199);
        static readonly Color C_WHITE   = Color.White;
        static readonly Color C_GRAY    = Color.FromArgb(120, 113,  90);
        static readonly Color C_GREEN   = Color.FromArgb(22, 163, 74);
        static readonly Color C_GREEN_LT= Color.FromArgb(220, 252, 231);
        static readonly Color C_BLUE    = Color.FromArgb(37,  99, 235);
        static readonly Color C_BLUE_LT = Color.FromArgb(219, 234, 254);

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
            Size          = new Size(740, 900);
            MinimumSize   = new Size(680, 800);
            StartPosition = FormStartPosition.CenterScreen;
            BackColor     = C_BG;
            Font          = F_BODY;
            BuildUI();
            AddVarRow(); AddVarRow();
            AddObjRow();
        }

        private void BuildUI()
        {
            var root = new TableLayoutPanel
            {
                Dock=DockStyle.Fill, ColumnCount=1, RowCount=9,
                Padding=new Padding(12,8,12,8), BackColor=C_BG,
            };
            root.RowStyles.Add(new RowStyle(SizeType.Absolute,  52)); // 0 标题
            root.RowStyles.Add(new RowStyle(SizeType.Absolute,  38)); // 1 配置工具栏
            root.RowStyles.Add(new RowStyle(SizeType.Percent,   26)); // 2 变量区
            root.RowStyles.Add(new RowStyle(SizeType.Absolute,  34)); // 3 变量工具栏
            root.RowStyles.Add(new RowStyle(SizeType.Percent,   18)); // 4 目标区
            root.RowStyles.Add(new RowStyle(SizeType.Absolute,  34)); // 5 目标工具栏
            root.RowStyles.Add(new RowStyle(SizeType.Percent,   14)); // 6 约束区
            root.RowStyles.Add(new RowStyle(SizeType.Absolute,  34)); // 7 约束工具栏
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 220)); // 8 参数+底部
            Controls.Add(root);

            // 标题
            var titlePnl = new Panel{Dock=DockStyle.Fill,BackColor=C_MAIN};
            titlePnl.Controls.Add(new Label{Text="  Bayesian Optimization  贝叶斯优化",Dock=DockStyle.Fill,TextAlign=ContentAlignment.MiddleLeft,Font=F_TITLE,ForeColor=C_WHITE,BackColor=Color.Transparent});
            root.Controls.Add(titlePnl,0,0);

            // 配置工具栏
            root.Controls.Add(ConfigToolbar(),0,1);

            // 设计变量
            var grpVar=MakeGrp("  [设计变量]");
            var tblVar=SectionLayout(HDR_H,out pnlVars);
            tblVar.Controls.Add(MakeHdr(new[]{"选取","单元格地址","下界 LB","上界 UB"},new[]{66,184,78,78}),0,0);
            grpVar.Controls.Add(tblVar);
            root.Controls.Add(grpVar,0,2);
            root.Controls.Add(Toolbar(()=>AddVarRow(),"  + 添加变量行",()=>DelVarRow(),"  - 删除末行","点 [选取] -> 工作表点击单元格 -> 地址自动填入"),0,3);

            // 目标函数
            var grpObj=MakeGrp("  [目标函数]");
            var tblObj=SectionLayout(HDR_H,out pnlObjs);
            tblObj.Controls.Add(MakeHdr(new[]{"选取","单元格地址（目标公式格）","方向","权重 W"},new[]{66,242,80,72}),0,0);
            grpObj.Controls.Add(tblObj);
            root.Controls.Add(grpObj,0,4);
            root.Controls.Add(Toolbar(()=>AddObjRow(),"  + 添加目标行",()=>DelObjRow(),"  - 删除末行","多目标时按权重加权标量化  |  方向：min / max"),0,5);

            // 约束条件
            var grpCon=MakeGrp("  [约束条件（可选）]");
            var tblCon=SectionLayout(HDR_H,out pnlCons);
            tblCon.Controls.Add(MakeHdr(new[]{"选取","约束单元格地址","符号","限制值"},new[]{66,186,80,80}),0,0);
            grpCon.Controls.Add(tblCon);
            root.Controls.Add(grpCon,0,6);
            root.Controls.Add(Toolbar(()=>AddConRow(),"  + 添加约束行",()=>DelConRow(),"  - 删除末行","约束公式结果 op 限制值（如 <=0）"),0,7);

            // 参数+底部
            var btmPanel=new TableLayoutPanel{Dock=DockStyle.Fill,ColumnCount=1,RowCount=2};
            btmPanel.RowStyles.Add(new RowStyle(SizeType.Absolute,148));
            btmPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 72));

            var paramGrp=MakeGrp("  [贝叶斯算法参数]");
            var pt=new TableLayoutPanel{Dock=DockStyle.Fill,ColumnCount=4,Padding=new Padding(8,4,8,2)};
            pt.ColumnStyles.Add(new ColumnStyle(SizeType.Percent,34));
            pt.ColumnStyles.Add(new ColumnStyle(SizeType.Percent,16));
            pt.ColumnStyles.Add(new ColumnStyle(SizeType.Percent,34));
            pt.ColumnStyles.Add(new ColumnStyle(SizeType.Percent,16));

            nudIter     = MakeNud(5,   500, 50,  false);
            nudInit     = MakeNud(3,   200, 10,  false);
            nudXi       = MakeNud(0.0m,1.0m,0.01m,true);
            nudRestarts = MakeNud(1,    20,  5,  false);
            nudBatch    = MakeNud(1,   100,  1,  false);  // 批量次数

            pt.Controls.Add(PL("优化迭代次数 :"),      0,0); pt.Controls.Add(nudIter,    1,0);
            pt.Controls.Add(PL("初始随机采样点 :"),    2,0); pt.Controls.Add(nudInit,    3,0);
            pt.Controls.Add(PL("EI 探索参数 xi :"),   0,1); pt.Controls.Add(nudXi,      1,1);
            pt.Controls.Add(PL("采集函数重启次数 :"), 2,1); pt.Controls.Add(nudRestarts,3,1);
            pt.Controls.Add(PL("批量优化次数 :"),     0,2); pt.Controls.Add(nudBatch,   1,2);
            var batchHint=new Label{Text="批量 > 1 时每次结果追加到 Sheet，最后汇总最优解",Dock=DockStyle.Fill,AutoSize=false,ForeColor=C_GRAY,Font=F_SMALL,TextAlign=ContentAlignment.MiddleLeft};
            pt.Controls.Add(batchHint,2,2); pt.SetColumnSpan(batchHint,2);

            var tip=new Label{Text="  推荐：迭代 >= 30，初始采样 >= 5。xi 越大越倾向探索。批量 >= 3 可提高鲁棒性。",Dock=DockStyle.Fill,AutoSize=false,ForeColor=C_GRAY,Font=F_SMALL,TextAlign=ContentAlignment.MiddleLeft};
            pt.Controls.Add(tip,0,3); pt.SetColumnSpan(tip,4);
            for(int r=0;r<4;r++) pt.RowStyles.Add(new RowStyle(SizeType.Absolute,26));

            paramGrp.Controls.Add(pt);
            btmPanel.Controls.Add(paramGrp,0,0);

            var runPnl=new TableLayoutPanel{Dock=DockStyle.Fill,ColumnCount=1,RowCount=3,Padding=new Padding(0,4,0,0)};
            runPnl.RowStyles.Add(new RowStyle(SizeType.Absolute,42));
            runPnl.RowStyles.Add(new RowStyle(SizeType.Absolute,20));
            runPnl.RowStyles.Add(new RowStyle(SizeType.Absolute,18));
            var bRow=new FlowLayoutPanel{Dock=DockStyle.Fill,FlowDirection=FlowDirection.LeftToRight,Padding=new Padding(0,2,0,2),WrapContents=false};
            btnRun  = MB("  [Ru] 运行优化",C_MAIN,C_WHITE,140);
            btnStop = MB("  [X]  停止",    C_RED, C_WHITE,104);
            btnStop.Enabled=false;
            btnRun.Click  +=OnRun;
            btnStop.Click +=(s,e)=>{_stopFlag=true;lblStatus.Text="正在停止...";};
            bRow.Controls.AddRange(new Control[]{btnRun,btnStop});
            runPnl.Controls.Add(bRow,0,0);
            progressBar=new ProgressBar{Dock=DockStyle.Fill,Minimum=0,Maximum=100};
            runPnl.Controls.Add(progressBar,0,1);
            lblStatus=new Label{Text="就绪，填写参数后点击 [运行优化]",Dock=DockStyle.Fill,AutoSize=false,TextAlign=ContentAlignment.MiddleLeft,ForeColor=C_GRAY,Font=F_SMALL};
            runPnl.Controls.Add(lblStatus,0,2);
            btmPanel.Controls.Add(runPnl,0,1);
            root.Controls.Add(btmPanel,0,8);
        }

        // ── 配置工具栏 ─────────────────────────────────────────────
        private FlowLayoutPanel ConfigToolbar()
        {
            var row=new FlowLayoutPanel{Dock=DockStyle.Fill,FlowDirection=FlowDirection.LeftToRight,Padding=new Padding(0,4,0,4),WrapContents=false};
            btnSave=new Button{Text="  [S] 保存配置",Width=114,Height=28,FlatStyle=FlatStyle.Flat,Font=F_SMALL,BackColor=C_GREEN_LT,ForeColor=C_GREEN,Margin=new Padding(0,0,6,0)};
            btnSave.FlatAppearance.BorderColor=C_GREEN; btnSave.Click+=OnSave;
            btnLoad=new Button{Text="  [L] 加载配置",Width=114,Height=28,FlatStyle=FlatStyle.Flat,Font=F_SMALL,BackColor=C_BLUE_LT,ForeColor=C_BLUE,Margin=new Padding(0,0,8,0)};
            btnLoad.FlatAppearance.BorderColor=C_BLUE; btnLoad.Click+=OnLoad;
            var hint=new Label{Text="配置保存为 .bayescfg 文件，下次打开可直接加载恢复",AutoSize=true,ForeColor=C_GRAY,Font=F_SMALL,TextAlign=ContentAlignment.MiddleLeft,Padding=new Padding(0,5,0,0)};
            row.Controls.AddRange(new Control[]{btnSave,btnLoad,hint});
            return row;
        }

        // ══════════════════════════════════════════════════════════════
        //  保存 / 加载配置
        // ══════════════════════════════════════════════════════════════
        private void OnSave(object sender, EventArgs e)
        {
            using(var dlg=new SaveFileDialog()){
                dlg.Title="保存贝叶斯优化配置";dlg.Filter="贝叶斯配置 (*.bayescfg)|*.bayescfg|所有文件|*.*";dlg.DefaultExt="bayescfg";
                if(dlg.ShowDialog()!=DialogResult.OK)return;
                try{File.WriteAllText(dlg.FileName,BuildConfigJson(),Encoding.UTF8);lblStatus.Text="配置已保存："+Path.GetFileName(dlg.FileName);}
                catch(Exception ex){MessageBox.Show("保存失败：\n"+ex.Message,"贝叶斯优化");}
            }
        }
        private void OnLoad(object sender, EventArgs e)
        {
            using(var dlg=new OpenFileDialog()){
                dlg.Title="加载贝叶斯优化配置";dlg.Filter="贝叶斯配置 (*.bayescfg)|*.bayescfg|所有文件|*.*";
                if(dlg.ShowDialog()!=DialogResult.OK)return;
                try{ApplyConfigJson(File.ReadAllText(dlg.FileName,Encoding.UTF8));lblStatus.Text="配置已加载："+Path.GetFileName(dlg.FileName);}
                catch(Exception ex){MessageBox.Show("加载失败：\n"+ex.Message,"贝叶斯优化");}
            }
        }

        private string BuildConfigJson()
        {
            var sb=new StringBuilder();
            sb.Append("{\"params\":");
            sb.Append(SimpleJson.Obj("iter",nudIter.Value.ToString(),"init",nudInit.Value.ToString(),"xi",nudXi.Value.ToString(),"restarts",nudRestarts.Value.ToString(),"batch",nudBatch.Value.ToString()));
            sb.Append(",\"vars\":[");
            for(int i=0;i<_vars.Count;i++){if(i>0)sb.Append(",");sb.Append(SimpleJson.Obj("addr",_vars[i].addr.Text,"lb",_vars[i].lb.Text,"ub",_vars[i].ub.Text));}
            sb.Append("],\"objs\":[");
            for(int i=0;i<_objs.Count;i++){if(i>0)sb.Append(",");string wt=((_objs[i].row.Tag as TextBox)?.Text)??"1.0";sb.Append(SimpleJson.Obj("addr",_objs[i].addr.Text,"dir",_objs[i].dir.Text,"wt",wt));}
            sb.Append("],\"cons\":[");
            for(int i=0;i<_cons.Count;i++){if(i>0)sb.Append(",");sb.Append(SimpleJson.Obj("addr",_cons[i].addr.Text,"op",_cons[i].op.Text,"lim",_cons[i].lim.Text));}
            sb.Append("]}");
            return sb.ToString();
        }

        private void ApplyConfigJson(string json)
        {
            string ps=ExtractSection(json,"params");
            if(!string.IsNullOrEmpty(ps)){TrySet(nudIter,SimpleJson.Get(ps,"iter"),false);TrySet(nudInit,SimpleJson.Get(ps,"init"),false);TrySet(nudXi,SimpleJson.Get(ps,"xi"),true);TrySet(nudRestarts,SimpleJson.Get(ps,"restarts"),false);TrySet(nudBatch,SimpleJson.Get(ps,"batch"),false);}
            while(_vars.Count>0)DelVarRow();while(_objs.Count>0)DelObjRow();while(_cons.Count>0)DelConRow();
            foreach(var item in SimpleJson.SplitArray(ExtractSection(json,"vars"))){AddVarRow();var v=_vars[_vars.Count-1];v.addr.Text=SimpleJson.Get(item,"addr");v.lb.Text=SimpleJson.Get(item,"lb");v.ub.Text=SimpleJson.Get(item,"ub");}
            foreach(var item in SimpleJson.SplitArray(ExtractSection(json,"objs"))){AddObjRow();var o=_objs[_objs.Count-1];o.addr.Text=SimpleJson.Get(item,"addr");o.dir.SelectedIndex=SimpleJson.Get(item,"dir")=="max"?1:0;if(o.row.Tag is TextBox wt)wt.Text=SimpleJson.Get(item,"wt");}
            foreach(var item in SimpleJson.SplitArray(ExtractSection(json,"cons"))){AddConRow();var c=_cons[_cons.Count-1];c.addr.Text=SimpleJson.Get(item,"addr");c.op.SelectedIndex=SimpleJson.Get(item,"op")==">="?1:0;c.lim.Text=SimpleJson.Get(item,"lim");}
        }

        private string ExtractSection(string json,string key)
        {
            string search="\""+key+"\":";int idx=json.IndexOf(search);if(idx<0)return "";
            int start=idx+search.Length;char open=json[start];char close=open=='{'?'}':']';int depth=0,end=start;
            for(int i=start;i<json.Length;i++){if(json[i]==open)depth++;else if(json[i]==close){depth--;if(depth==0){end=i;break;}}}
            return json.Substring(start,end-start+1);
        }
        private void TrySet(NumericUpDown nud,string val,bool dec){if(string.IsNullOrEmpty(val))return;try{decimal v=decimal.Parse(val);if(v>=nud.Minimum&&v<=nud.Maximum)nud.Value=v;}catch{}}

        // ══════════════════════════════════════════════════════════════
        //  布局工具
        // ══════════════════════════════════════════════════════════════
        private TableLayoutPanel SectionLayout(int hdrHeight,out Panel dataPanel)
        {
            var tbl=new TableLayoutPanel{Dock=DockStyle.Fill,ColumnCount=1,RowCount=2,Padding=new Padding(0)};
            tbl.RowStyles.Add(new RowStyle(SizeType.Absolute,hdrHeight));
            tbl.RowStyles.Add(new RowStyle(SizeType.Percent,100));
            dataPanel=new Panel{Dock=DockStyle.Fill,AutoScroll=true};
            tbl.Controls.Add(dataPanel,0,1);
            return tbl;
        }
        private Panel MakeHdr(string[]titles,int[]widths)
        {
            var pnl=new Panel{Dock=DockStyle.Fill,BackColor=C_HDR};int x=6;
            for(int i=0;i<titles.Length;i++){pnl.Controls.Add(new Label{Text=titles[i],Left=x,Top=4,Width=widths[i],Height=HDR_H-6,Font=F_HDR,ForeColor=C_MAIN_DK,TextAlign=ContentAlignment.MiddleLeft});x+=widths[i]+4;}
            return pnl;
        }
        private FlowLayoutPanel Toolbar(Action addAct,string addTxt,Action delAct,string delTxt,string hint)
        {
            var row=new FlowLayoutPanel{Dock=DockStyle.Fill,FlowDirection=FlowDirection.LeftToRight,Padding=new Padding(0,3,0,3),WrapContents=false};
            var bAdd=new Button{Text=addTxt,Width=110,Height=26,FlatStyle=FlatStyle.Flat,Font=F_SMALL,BackColor=C_MAIN_LT,ForeColor=C_MAIN,Margin=new Padding(0,0,6,0)};
            bAdd.FlatAppearance.BorderColor=C_MAIN;bAdd.Click+=(s,e)=>addAct();
            var bDel=new Button{Text=delTxt,Width=102,Height=26,FlatStyle=FlatStyle.Flat,Font=F_SMALL,BackColor=C_DEL_BG,ForeColor=C_DEL_FG,Margin=new Padding(0,0,8,0)};
            bDel.FlatAppearance.BorderColor=C_DEL_FG;bDel.Click+=(s,e)=>delAct();
            row.Controls.Add(bAdd);row.Controls.Add(bDel);
            row.Controls.Add(new Label{Text=hint,AutoSize=true,ForeColor=C_GRAY,Font=F_SMALL,TextAlign=ContentAlignment.MiddleLeft,Padding=new Padding(0,4,0,0)});
            return row;
        }

        // ══════════════════════════════════════════════════════════════
        //  添加 / 删除行
        // ══════════════════════════════════════════════════════════════
        private void AddVarRow()
        {
            var rowPnl=new Panel{Left=0,Top=_vars.Count*ROW_H,Width=680,Height=ROW_H-2};
            int x=6;var btn=PickBtn(rowPnl,x);x+=70;var addr=TB(rowPnl,x,182);x+=186;var lb=TB(rowPnl,x,76);lb.Text="0";x+=80;var ub=TB(rowPnl,x,76);ub.Text="1";
            WirePick(btn,addr);pnlVars.Controls.Add(rowPnl);_vars.Add((rowPnl,addr,lb,ub));
            pnlVars.AutoScrollMinSize=new Size(0,(_vars.Count+1)*ROW_H);
        }
        private void DelVarRow(){if(_vars.Count==0)return;int last=_vars.Count-1;pnlVars.Controls.Remove(_vars[last].row);_vars[last].row.Dispose();_vars.RemoveAt(last);pnlVars.AutoScrollMinSize=new Size(0,(_vars.Count+1)*ROW_H);}
        private void AddObjRow()
        {
            var rowPnl=new Panel{Left=0,Top=_objs.Count*ROW_H,Width=680,Height=ROW_H-2};
            int x=6;var btn=PickBtn(rowPnl,x);x+=70;var addr=TB(rowPnl,x,240);x+=244;
            var dir=new ComboBox{Left=x,Top=4,Width=78,Height=26,DropDownStyle=ComboBoxStyle.DropDownList,Font=F_BODY};dir.Items.AddRange(new object[]{"min","max"});dir.SelectedIndex=0;x+=82;
            rowPnl.Controls.Add(dir);var wt=TB(rowPnl,x,70);wt.Text="1.0";
            WirePick(btn,addr);pnlObjs.Controls.Add(rowPnl);_objs.Add((rowPnl,addr,dir));rowPnl.Tag=wt;
            pnlObjs.AutoScrollMinSize=new Size(0,(_objs.Count+1)*ROW_H);
        }
        private void DelObjRow(){if(_objs.Count==0)return;int last=_objs.Count-1;pnlObjs.Controls.Remove(_objs[last].row);_objs[last].row.Dispose();_objs.RemoveAt(last);pnlObjs.AutoScrollMinSize=new Size(0,(_objs.Count+1)*ROW_H);}
        private void AddConRow()
        {
            var rowPnl=new Panel{Left=0,Top=_cons.Count*ROW_H,Width=680,Height=ROW_H-2};
            int x=6;var btn=PickBtn(rowPnl,x);x+=70;var addr=TB(rowPnl,x,184);x+=188;
            var op=new ComboBox{Left=x,Top=4,Width=78,Height=26,DropDownStyle=ComboBoxStyle.DropDownList,Font=F_BODY};op.Items.AddRange(new object[]{"<=",">="});op.SelectedIndex=0;x+=82;
            rowPnl.Controls.Add(op);var lim=TB(rowPnl,x,78);lim.Text="0";
            WirePick(btn,addr);pnlCons.Controls.Add(rowPnl);_cons.Add((rowPnl,addr,op,lim));
            pnlCons.AutoScrollMinSize=new Size(0,(_cons.Count+1)*ROW_H);
        }
        private void DelConRow(){if(_cons.Count==0)return;int last=_cons.Count-1;pnlCons.Controls.Remove(_cons[last].row);_cons[last].row.Dispose();_cons.RemoveAt(last);pnlCons.AutoScrollMinSize=new Size(0,(_cons.Count+1)*ROW_H);}

        // ══════════════════════════════════════════════════════════════
        //  选取
        // ══════════════════════════════════════════════════════════════
        private Button PickBtn(Panel parent,int x)
        {
            var b=new Button{Text="  [选取]",Left=x,Top=4,Width=64,Height=26,FlatStyle=FlatStyle.Flat,Font=F_SMALL,BackColor=C_MAIN_LT,ForeColor=C_MAIN};
            b.FlatAppearance.BorderColor=C_MAIN;parent.Controls.Add(b);return b;
        }
        private void WirePick(Button btn,TextBox addrBox)
        {
            btn.Click+=(s,e)=>
            {
                btn.Enabled=false;btn.Text="  ...";
                var bg=new Thread(()=>{
                    string picked=null;bool done=false;var gate=new object();
                    ExcelAsyncUtil.QueueAsMacro(()=>{
                        try{dynamic app=ExcelDnaUtil.Application;object res=app.InputBox("请在工作表中点击目标单元格，然后点确定：","选取单元格",Type.Missing,Type.Missing,Type.Missing,Type.Missing,Type.Missing,8);if(!(res is bool)){dynamic rng=res;picked=((string)rng.Address[false,false,1,true]).Replace("$","");}}
                        catch{}finally{lock(gate){done=true;Monitor.PulseAll(gate);}}
                    });
                    lock(gate){while(!done)Monitor.Wait(gate,20000);}
                    addrBox.Invoke((Action)(()=>{if(picked!=null)addrBox.Text=picked;btn.Enabled=true;btn.Text="  [选取]";}));
                });
                bg.IsBackground=true;bg.Start();
            };
        }

        // ══════════════════════════════════════════════════════════════
        //  运行（支持批量）
        // ══════════════════════════════════════════════════════════════
        private void OnRun(object sender, EventArgs e)
        {
            BayesEngine eng;
            try{eng=BuildEngine();}
            catch(Exception ex){MessageBox.Show("参数错误：\n"+ex.Message,"贝叶斯优化");return;}
            int batchCount=(int)nudBatch.Value;
            btnRun.Enabled=false;btnStop.Enabled=true;
            _stopFlag=false;progressBar.Value=0;lblStatus.Text="初始化...";

            var t=new Thread(()=>{
                var allBests=new List<BayesPoint>();
                try
                {
                    for(int b=0;b<batchCount&&!_stopFlag;b++)
                    {
                        int batchIdx=b;
                        // 重建引擎（每次批量独立采样）
                        BayesEngine beng=BuildEngine();
                        beng.StopFlag=false;
                        int totalSteps=beng.InitPoints+beng.MaxIter;
                        beng.OnProgress=(iter,total,msg)=>{
                            this.Invoke((Action)(()=>{
                                int globalPct=(int)((b*100.0+100.0*iter/total)/batchCount);
                                progressBar.Value=Math.Min(100,globalPct);
                                lblStatus.Text=string.Format("批次 {0}/{1} | 迭代 {2}/{3}  {4}",b+1,batchCount,iter,total,msg);
                            }));
                            if(_stopFlag)beng.StopFlag=true;
                        };
                        beng.Evaluate=ind=>{
                            bool done=false;Exception err=null;
                            ExcelAsyncUtil.QueueAsMacro(()=>{try{ExcelBridge.Evaluate(ind,beng);}catch(Exception ex){err=ex;}finally{lock(ind){done=true;Monitor.PulseAll(ind);}}});
                            lock(ind){while(!done)Monitor.Wait(ind,10000);}
                            if(err!=null)throw err;
                        };
                        BayesPoint best=beng.Run();
                        if(best!=null)
                        {
                            best.BatchIndex=b+1;
                            allBests.Add(best);
                            // 追加写入每批结果
                            ExcelAsyncUtil.QueueAsMacro(()=>{try{ExcelBridge.AppendBatchResult(best,beng,b+1,batchCount);}catch(Exception ex){MessageBox.Show("写结果出错：\n"+ex.Message,"贝叶斯优化");}});
                        }
                    }
                    // 汇总最优
                    if(allBests.Count>0)
                    {
                        BayesPoint globalBest=null;
                        foreach(var p in allBests)
                            if(globalBest==null||p.ScalarF<globalBest.ScalarF)globalBest=p;
                        ExcelAsyncUtil.QueueAsMacro(()=>{try{ExcelBridge.WriteSummary(allBests,globalBest,BuildEngine());}catch{}});
                    }
                    this.Invoke((Action)(()=>{
                        progressBar.Value=100;
                        lblStatus.Text=string.Format("完成！共 {0} 批次，全局最优标量值 = {1:F6}  结果见 [贝叶斯结果] Sheet",allBests.Count,allBests.Count>0?(object)allBests[0].ScalarF:"N/A");
                        btnRun.Enabled=true;btnStop.Enabled=false;
                    }));
                }
                catch(Exception ex)
                {
                    this.Invoke((Action)(()=>{lblStatus.Text="错误: "+ex.Message;btnRun.Enabled=true;btnStop.Enabled=false;MessageBox.Show("运行错误：\n"+ex.Message,"贝叶斯优化");}));
                }
            });
            t.IsBackground=true;t.Start();
        }

        private BayesEngine BuildEngine()
        {
            var eng=new BayesEngine{MaxIter=(int)nudIter.Value,InitPoints=(int)nudInit.Value,Xi=(double)nudXi.Value,AcqRestarts=(int)nudRestarts.Value};
            foreach(var(r,addr,lb,ub)in _vars){string c=addr.Text.Trim(),l=lb.Text.Trim(),u=ub.Text.Trim();if(string.IsNullOrEmpty(c)||string.IsNullOrEmpty(l))continue;eng.Vars.Add(new VarDef{Cell=c,LB=double.Parse(l),UB=string.IsNullOrEmpty(u)?1.0:double.Parse(u)});}
            foreach(var(r,addr,dir)in _objs){string c=addr.Text.Trim();if(string.IsNullOrEmpty(c))continue;double w=1.0;if(r.Tag is TextBox wt&&!string.IsNullOrEmpty(wt.Text.Trim()))double.TryParse(wt.Text.Trim(),out w);eng.Objs.Add(new ObjDef{Cell=c,Minimize=dir.Text!="max",Weight=w});}
            foreach(var(r,addr,op,lim)in _cons){string c=addr.Text.Trim();if(string.IsNullOrEmpty(c))continue;eng.Cons.Add(new ConDef{Cell=c,Op=op.Text,Limit=string.IsNullOrEmpty(lim.Text)?0.0:double.Parse(lim.Text)});}
            if(eng.Vars.Count==0)throw new Exception("请至少添加一个设计变量！");
            if(eng.Objs.Count==0)throw new Exception("请至少添加一个目标函数！");
            return eng;
        }

        private TextBox TB(Panel p,int x,int w){var t=new TextBox{Left=x,Top=5,Width=w,Height=26,Font=F_BODY,BackColor=C_WHITE,BorderStyle=BorderStyle.FixedSingle};p.Controls.Add(t);return t;}
        private GroupBox MakeGrp(string txt)=>new GroupBox{Text=txt,Dock=DockStyle.Fill,Font=F_GRP,ForeColor=C_MAIN};
        private NumericUpDown MakeNud(decimal min,decimal max,decimal val,bool dec)=>new NumericUpDown{Minimum=min,Maximum=max,Value=val,Width=90,DecimalPlaces=dec?2:0,Increment=dec?0.01m:1,Font=F_BODY};
        private Label PL(string t)=>new Label{Text=t,Dock=DockStyle.Fill,TextAlign=ContentAlignment.MiddleRight,Font=F_BODY,ForeColor=C_MAIN_DK};
        private Button MB(string text,Color back,Color fore,int width)=>new Button{Text=text,Width=width,Height=36,BackColor=back,ForeColor=fore,FlatStyle=FlatStyle.Flat,Margin=new Padding(0,0,10,0),Font=F_BTN};
    }
}
