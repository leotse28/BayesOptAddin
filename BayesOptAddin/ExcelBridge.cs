using System;
using System.Collections.Generic;
using ExcelDna.Integration;

namespace BayesOptAddin
{
    public static class ExcelBridge
    {
        private static dynamic App => ExcelDnaUtil.Application;
        const string SHEET_NAME = "贝叶斯结果";

        public static void Evaluate(BayesPoint pt, BayesEngine eng)
        {
            dynamic app=App;
            for(int j=0;j<eng.Vars.Count;j++) app.Range[eng.Vars[j].Cell].Value2=pt.X[j];
            app.Calculate();
            for(int j=0;j<eng.Objs.Count;j++) pt.F[j]=Convert.ToDouble(app.Range[eng.Objs[j].Cell].Value2);
            pt.Feasible=true;
            for(int j=0;j<eng.Cons.Count;j++)
            {
                double v=Convert.ToDouble(app.Range[eng.Cons[j].Cell].Value2);
                double vio=eng.Cons[j].Op=="<="?v-eng.Cons[j].Limit:eng.Cons[j].Limit-v;
                if(vio>0){pt.Feasible=false;break;}
            }
        }

        // 每批次结果追加（第1批时初始化 Sheet 和标题）
        public static void AppendBatchResult(BayesPoint best, BayesEngine eng, int batchIdx, int totalBatch)
        {
            dynamic app=App,wb=app.ActiveWorkbook;
            dynamic ws=null;
            foreach(dynamic sh in wb.Sheets) if((string)sh.Name==SHEET_NAME){ws=sh;break;}
            if(ws==null||(batchIdx==1))
            {
                // 第一批：清空或新建
                if(ws!=null){app.DisplayAlerts=false;ws.Delete();app.DisplayAlerts=true;}
                ws=wb.Sheets.Add(After:wb.Sheets[wb.Sheets.Count]);ws.Name=SHEET_NAME;
                WriteHeader(ws, eng);
            }
            // 找下一行
            int lastRow=(int)ws.Cells[ws.Rows.Count,1].End(-4162).Row; // xlUp
            if(lastRow<2)lastRow=1;
            int writeRow=lastRow+1;
            WriteRow(ws, writeRow, best, eng, batchIdx);
            ws.Columns.AutoFit();
            ws.Activate();
        }

        // 汇总：在末尾追加空行+全局最优标注
        public static void WriteSummary(List<BayesPoint> allBests, BayesPoint globalBest, BayesEngine eng)
        {
            dynamic app=App,wb=app.ActiveWorkbook;
            dynamic ws=null;
            foreach(dynamic sh in wb.Sheets) if((string)sh.Name==SHEET_NAME){ws=sh;break;}
            if(ws==null) return;
            int lastRow=(int)ws.Cells[ws.Rows.Count,1].End(-4162).Row;
            int writeRow=lastRow+2;
            H(ws.Cells[writeRow,1],"全局最优解"); ws.SetColumnSpan(ws.Cells[writeRow,1],ws.Cells[writeRow,1+1+eng.Vars.Count+eng.Objs.Count+1]);
            WriteRow(ws, writeRow+1, globalBest, eng, -1);
            ws.Columns.AutoFit();ws.Activate();
        }

        private static void WriteHeader(dynamic ws, BayesEngine eng)
        {
            int col=1;
            H(ws.Cells[1,col],"批次");col++;
            for(int j=0;j<eng.Vars.Count;j++){H(ws.Cells[1,col],string.Format("变量{0}[{1}]LB={2}UB={3}",j+1,eng.Vars[j].Cell,eng.Vars[j].LB,eng.Vars[j].UB));col++;}
            for(int j=0;j<eng.Objs.Count;j++){H(ws.Cells[1,col],string.Format("目标{0}[{1}]{2}",j+1,eng.Objs[j].Cell,eng.Objs[j].Minimize?"min":"max"));col++;}
            H(ws.Cells[1,col],"加权标量值");col++;
            H(ws.Cells[1,col],"可行性");
        }

        private static void WriteRow(dynamic ws, int row, BayesPoint pt, BayesEngine eng, int batchIdx)
        {
            int col=1;
            ws.Cells[row,col].Value2=(batchIdx>0?(object)batchIdx:"全局最优");col++;
            for(int j=0;j<eng.Vars.Count;j++){ws.Cells[row,col].Value2=Math.Round(pt.X[j],8);col++;}
            for(int j=0;j<eng.Objs.Count;j++){ws.Cells[row,col].Value2=Math.Round(pt.F[j],8);col++;}
            ws.Cells[row,col].Value2=Math.Round(pt.ScalarF,8);col++;
            ws.Cells[row,col].Value2=pt.Feasible?"可行":"不可行";
            if(batchIdx%2==0)
            {
                dynamic rng=(object)ws.Range[ws.Cells[row,1],ws.Cells[row,col]];
                ((dynamic)rng).Interior.Color=0xFFF3CD;
            }
        }

        private static void H(dynamic c,string t){c.Value2=t;c.Interior.Color=0xD97706;c.Font.Color=0xFFFFFF;c.Font.Bold=true;c.HorizontalAlignment=-4108;}
    }
}
