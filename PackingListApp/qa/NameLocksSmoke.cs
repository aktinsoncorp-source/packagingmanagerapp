using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;
namespace PackingListManager {
    static class NameLocksSmoke {
        [STAThread] static void Main(string[] args) {
            Application.EnableVisualStyles();Directory.CreateDirectory(args[0]);
            var catalog=new Dictionary<string,object>{{"revision",0},{"rows",new object[0]}};
            using(var form=new NameDefinitionsForm(catalog,request=>Task.FromResult(catalog))) {
                form.StartPosition=FormStartPosition.Manual;form.Location=new Point(-30000,-30000);form.Show();Application.DoEvents();
                var locked=(Button)form.Controls.Find("lockNames",true)[0];var unlocked=(Button)form.Controls.Find("unlockNames",true)[0];
                var grid=(DataGridView)form.Controls.Find("namesGrid",true)[0];
                if(locked.Top!=unlocked.Top||locked.Bottom!=unlocked.Bottom||locked.Size!=unlocked.Size)throw new Exception("Padlocks are not aligned");
                if(!grid.ReadOnly||locked.Enabled||!unlocked.Enabled)throw new Exception("Must start locked");
                Save(form,Path.Combine(args[0],"names-locked.png"));
                unlocked.PerformClick();Application.DoEvents();
                if(grid.ReadOnly||!locked.Enabled||unlocked.Enabled)throw new Exception("Unlock did not enable editing");
                Save(form,Path.Combine(args[0],"names-unlocked.png"));
                locked.PerformClick();Application.DoEvents();if(!grid.ReadOnly)throw new Exception("Relocking did not work");
                using(var icons=new Bitmap(96,48))using(var g=Graphics.FromImage(icons)) {
                    g.Clear(Color.White);g.InterpolationMode=System.Drawing.Drawing2D.InterpolationMode.NearestNeighbor;
                    g.DrawImage(locked.Image,new Rectangle(0,0,48,48));g.DrawImage(unlocked.Image,new Rectangle(48,0,48,48));icons.Save(Path.Combine(args[0],"locks.png"));
                }
                form.Close();
            }
            File.WriteAllText(Path.Combine(args[0],"locks-ok.txt"),"Both outline locks have identical control size and vertical alignment. Default locked, unlock and relock behavior passed.");
        }
        static void Save(Form f,string path){using(var b=new Bitmap(f.Width,f.Height)){f.DrawToBitmap(b,new Rectangle(Point.Empty,f.Size));b.Save(path);}}
    }
}
