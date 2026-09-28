using System;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Windows.Forms;
namespace PackingListManager {
    static class ExportDialogsSmoke {
        [STAThread] static void Main(string[] args) {
            Application.EnableVisualStyles();Directory.CreateDirectory(args[0]);
            using(var form=PackingDialogs.RepeatWarning(282)) {
                form.StartPosition=FormStartPosition.Manual;form.Location=new Point(-30000,-30000);form.Show();Application.DoEvents();
                var no=(Button)form.Controls.Find("repeatNo",true)[0];var proceed=(LinkLabel)form.Controls.Find("repeatContinue",true)[0];
                if(form.AcceptButton!=no||form.CancelButton!=no||form.ActiveControl!=no)throw new Exception("No must remain the keyboard default");
                if(no.Left*2+no.Width!=form.ClientSize.Width||no.Width<300||no.Height<50||proceed.Top<no.Bottom||proceed.Width>no.Width||proceed.Left!=no.Left)throw new Exception("No and continue link are not centered/aligned");
                if(proceed.Text!="Continue to Generate the Repeated Orders"||proceed.BorderStyle!=BorderStyle.None||proceed.Font.Size>=no.Font.Size||proceed.LinkBehavior!=LinkBehavior.NeverUnderline)throw new Exception("Continue control must be small borderless text");
                Save(form,Path.Combine(args[0],"repeat-warning.png"));
                form.AcceptButton.PerformClick();if(form.DialogResult!=DialogResult.No)throw new Exception("Default action did not cancel");form.Close();
            }
            using(var form=PackingDialogs.RepeatWarning(282)) {
                form.StartPosition=FormStartPosition.Manual;form.Location=new Point(-30000,-30000);form.Show();Application.DoEvents();
                var proceed=(LinkLabel)form.Controls.Find("repeatContinue",true)[0];
                typeof(LinkLabel).GetMethod("OnLinkClicked",BindingFlags.NonPublic|BindingFlags.Instance).Invoke(proceed,new object[]{new LinkLabelLinkClickedEventArgs(proceed.Links[0])});
                if(form.DialogResult!=DialogResult.Yes)throw new Exception("Explicit continue click must approve");
            }
            File.WriteAllText(Path.Combine(args[0],"dialogs-ok.txt"),"Large centered No is the Enter/Escape default. The smaller gray borderless continue link fits beneath its width. Default action cancels; deliberate continuation approves.");
        }
        static void Save(Form form,string path){using(var bitmap=new Bitmap(form.Width,form.Height)){form.DrawToBitmap(bitmap,new Rectangle(Point.Empty,form.Size));bitmap.Save(path);}}
    }
}
