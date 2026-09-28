using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
namespace PackingListManager {
    static class ExportDialogsSmoke {
        [STAThread] static void Main(string[] args) {
            Application.EnableVisualStyles();Directory.CreateDirectory(args[0]);
            using(var form=PackingDialogs.ExportOptions()) {
                form.StartPosition=FormStartPosition.Manual;form.Location=new Point(-30000,-30000);form.Show();Application.DoEvents();
                var pdf=(RadioButton)form.Controls.Find("exportPdf",true)[0];var word=(RadioButton)form.Controls.Find("exportWord",true)[0];
                if(!pdf.Checked||word.Checked)throw new Exception("PDF must be the default");
                Save(form,Path.Combine(args[0],"export-options.png"));
                word.Checked=true;if(pdf.Checked||!word.Checked)throw new Exception("Word selection failed");form.Close();
            }
            using(var form=PackingDialogs.ExportOptions()) {if(!((RadioButton)form.Controls.Find("exportPdf",true)[0]).Checked)throw new Exception("New export dialog must reset to PDF");}
            using(var form=PackingDialogs.RepeatWarning(17)) {
                form.StartPosition=FormStartPosition.Manual;form.Location=new Point(-30000,-30000);form.Show();Application.DoEvents();
                var no=(Button)form.Controls.Find("repeatNo",true)[0];var yes=(Button)form.Controls.Find("repeatYes",true)[0];
                if(form.AcceptButton!=no||form.CancelButton!=no||no.Width<=yes.Width||no.Height<=yes.Height||form.ActiveControl!=no)throw new Exception("No must be larger and the keyboard default");
                Save(form,Path.Combine(args[0],"repeat-warning.png"));
                form.AcceptButton.PerformClick();if(form.DialogResult!=DialogResult.No)throw new Exception("Default action did not cancel");form.Close();
            }
            File.WriteAllText(Path.Combine(args[0],"dialogs-ok.txt"),"PDF defaults on each open; Word can be selected; repeat warning defaults Enter/Escape to No; No is larger; default action returns No.");
        }
        static void Save(Form form,string path){using(var bitmap=new Bitmap(form.Width,form.Height)){form.DrawToBitmap(bitmap,new Rectangle(Point.Empty,form.Size));bitmap.Save(path);}}
    }
}
