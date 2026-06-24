using System;
using System.Collections.Generic;
using ExcelDna.Integration;

namespace BayesOptAddin
{
    public static class ExcelBridge
    {
        private static dynamic App => ExcelDnaUtil.Application;

        public static void Evaluate(BayesPoint pt, BayesEngine eng)
        {
            dynamic app = App;
            // 写入变量
            for (int j = 0; j < eng.Vars.Count; j++)
                app.Range[eng.Vars[j].Cell].Value2 = pt.X[j];
            app.Calculate();
            // 读取目标
            for (int j = 0; j < eng.Objs.Count; j++)
                pt.F[j] = Convert.ToDouble(app.Range[eng.Objs[j].Cell].Value2);
            // 判断约束
            pt.Feasible = true;
            for (int j = 0; j < eng.Cons.Count; j++)
            {
                double v   = Convert.ToDouble(app.Range[eng.Cons[j].Cell].Value2);
                double vio = eng.Cons[j].Op == "<=" ? v - eng.Cons[j].Limit : eng.Cons[j].Limit - v;
                if (vio > 0) { pt.Feasible = false; break; }
            }
        }

        public static void WriteResult(BayesPoint best, BayesEngine eng)
        {
            if (best == null) return;
            dynamic app = App, wb = app.ActiveWorkbook;
            // 删除旧结果
            foreach (dynamic sh in wb.Sheets)
                if ((string)sh.Name == "贝叶斯结果") { app.DisplayAlerts = false; sh.Delete(); app.DisplayAlerts = true; break; }
            dynamic ws = wb.Sheets.Add(After: wb.Sheets[wb.Sheets.Count]);
            ws.Name = "贝叶斯结果";

            int nv = eng.Vars.Count, no = eng.Objs.Count, col = 1;
            // 标题行
            H(ws.Cells[1, col], "序号"); col++;
            for (int j = 0; j < nv; j++) { H(ws.Cells[1, col], string.Format("变量{0}[{1}]LB={2}UB={3}", j+1, eng.Vars[j].Cell, eng.Vars[j].LB, eng.Vars[j].UB)); col++; }
            for (int j = 0; j < no; j++) { H(ws.Cells[1, col], string.Format("目标{0}[{1}]{2}", j+1, eng.Objs[j].Cell, eng.Objs[j].Minimize ? "min" : "max")); col++; }
            H(ws.Cells[1, col], "加权标量值"); col++;
            H(ws.Cells[1, col], "可行性");

            // 最优解写入第 2 行
            col = 1;
            ws.Cells[2, col].Value2 = 1; col++;
            for (int j = 0; j < nv; j++) { ws.Cells[2, col].Value2 = Math.Round(best.X[j], 8); col++; }
            for (int j = 0; j < no; j++) { ws.Cells[2, col].Value2 = Math.Round(best.F[j], 8); col++; }
            ws.Cells[2, col].Value2 = Math.Round(best.ScalarF, 8); col++;
            ws.Cells[2, col].Value2 = best.Feasible ? "可行" : "不可行";

            ws.Columns.AutoFit();
            ws.Activate();
        }

        private static void H(dynamic c, string t)
        {
            c.Value2 = t;
            c.Interior.Color = 0xD97706; // 橙金
            c.Font.Color = 0xFFFFFF;
            c.Font.Bold = true;
            c.HorizontalAlignment = -4108;
        }
    }
}
