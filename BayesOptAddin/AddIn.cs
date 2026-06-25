using System.Runtime.InteropServices;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
namespace BayesOptAddin
{
    [ComVisible(true)]
    public class BayesRibbon:ExcelRibbon
    {
        public override string GetCustomUI(string RibbonID)=>
            @"<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui'>
  <ribbon><tabs><tab id='tabBayes' label='贝叶斯优化'>
    <group id='grpMain' label='Bayesian Opt'>
      <button id='btnShow' label='打开优化面板' screentip='打开贝叶斯优化面板'
              size='large' imageMso='AnalysisToolPak' onAction='ShowPanel'/>
    </group>
  </tab></tabs></ribbon></customUI>";
        public void ShowPanel(IRibbonControl control){PanelManager.Show();}
    }
    public class AddInEntry:IExcelAddIn
    {
        public void AutoOpen(){}
        public void AutoClose(){PanelManager.Close();}
    }
}