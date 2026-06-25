using System.Windows.Forms;
using ExcelDna.Integration;
namespace BayesOptAddin
{
    public static class PanelManager
    {
        private static BayesForm _form;
        public static void Show()
        {
            if(_form==null||_form.IsDisposed){_form=new BayesForm();_form.Show(new ExcelWin(ExcelDnaUtil.WindowHandle));}
            else{if(!_form.Visible)_form.Show();_form.BringToFront();}
        }
        public static void Close(){try{_form?.Close();}catch{}_form=null;}
    }
    public class ExcelWin:IWin32Window
    {
        private readonly System.IntPtr _h;
        public ExcelWin(System.IntPtr h){_h=h;}
        public System.IntPtr Handle=>_h;
    }
}