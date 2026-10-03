using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
namespace PackingListManager {
 static class NameSortSmoke {
  static List<string> checks=new List<string>();
  static void Check(bool value,string label){if(!value)throw new Exception(label);checks.Add(label);}
  static string Order(DataGridView grid){return String.Join(",",grid.Rows.Cast<DataGridViewRow>().Where(r=>!r.IsNewRow).Select(r=>Convert.ToString(((Dictionary<string,object>)r.Tag)["id"])));}
  [STAThread] static void Main(string[] args){
   Application.EnableVisualStyles();Directory.CreateDirectory(args[0]);
   var rows=new List<Dictionary<string,object>>();
   string[][] data={new[]{"A30","ALPHA","30ML","ZULU","1L"},new[]{"A100","ALPHA","100ML","ZULU","30ML"},new[]{"A1L","ALPHA","1L","ALPHA INV","100ML"},new[]{"B30","BETA","30ML","ZULU","100ML"},new[]{"B100","BETA","100ML","ALPHA INV","1L"},new[]{"B1L","BETA","1L","ALPHA INV","30ML"}};
   foreach(var values in data)rows.Add(new Dictionary<string,object>{{"id",values[0]},{"original_product",values[1]},{"original_variation",values[2]},{"alternative_product",values[3]},{"alternative_variation",values[4]},{"locked",false},{"aliases",new object[0]}});
   Func<Dictionary<string,object>,Task<Dictionary<string,object>>> call=request=>Task.FromResult((string)request["action"]=="names_prepare"?new Dictionary<string,object>{{"rows",request["rows"]},{"merges",new object[0]}}:new Dictionary<string,object>{{"revision",2},{"rows",request["rows"]}});
   using(var form=new NameDefinitionsForm(new Dictionary<string,object>{{"revision",1},{"rows",rows}},call)){
    form.SuppressDialogs=true;form.StartPosition=FormStartPosition.Manual;form.Location=new Point(-30000,-30000);
    form.Shown+=async(s,e)=>{try{
     var grid=(DataGridView)form.Controls.Find("namesGrid",true)[0];
     Check(Order(grid)=="A30,A100,A1L,B30,B100,B1L","Initial product A-Z and numeric sizes within products");
     Check(grid.ReadOnly&&grid.Rows[0].Cells[0].ReadOnly,"Sorting available while names stay locked");
     Check(form.SortNames("original_variation"),"Reverse original size sort");
     Check(Order(grid)=="A1L,A100,A30,B1L,B100,B30","Sizes reverse within each product, products stay A-Z");
     Check(form.SortNames("alternative_product"),"Choose alternative product independently");
     Check(Order(grid)=="A1L,B1L,B100,A100,A30,B30","Alternative groups retain original variation descending");
     Check(form.SortNames("alternative_variation"),"Choose alternative variation independently");
     Check(Order(grid)=="B1L,A1L,B100,A100,B30,A30","Alternative groups use alternative sizes ascending");
     Check(form.SortNames("alternative_product"),"Reverse alternative products only");
     Check(Order(grid)=="A100,B30,A30,B1L,A1L,B100","Z-A retains ascending alternative sizes");
     Check(form.SortNames("alternative_variation"),"Reverse alternative sizes only");
     Check(Order(grid)=="A30,B30,A100,B100,A1L,B1L","Z-A groups and descending alternative sizes");
     Check(grid.Columns["alternative_product"].HeaderText.Contains("Z-A")&&grid.Columns["alternative_product"].HeaderText.Contains("click A-Z"),"Name header describes current and next direction");
     Check(grid.Columns["alternative_variation"].HeaderText.Contains("Largest to Smallest")&&grid.Columns["alternative_variation"].HeaderCell.ToolTipText.Contains("Click for smallest to largest"),"Variation header and tooltip describe direction toggle");
     Check(grid.Rows.Cast<DataGridViewRow>().Where(r=>!r.IsNewRow).All(r=>r.Cells[0].Style.BackColor.IsEmpty&&r.Cells[0].ReadOnly),"Sorting preserves row locks and saved highlights");
     Application.DoEvents();using(var image=new Bitmap(form.Width,form.Height)){
      form.DrawToBitmap(image,new Rectangle(Point.Empty,form.Size));image.Save(Path.Combine(args[0],"independent-sorting.png"));
      var screen=grid.PointToScreen(Point.Empty);var origin=new Point(screen.X-form.Left,screen.Y-form.Top);
      foreach(string field in new[]{"alternative_product","alternative_variation"}){
       var header=grid.GetCellDisplayRectangle(grid.Columns[field].Index,-1,true);int black=0;
       for(int x=header.Right-19;x<header.Right-8;x++)for(int y=header.Top+35;y<header.Top+51;y++){var pixel=image.GetPixel(origin.X+x,origin.Y+y);if(pixel.R<20&&pixel.G<20&&pixel.B<20)black++;}
       Check(black>12,"Black arrow rendered for "+field);
      }
     }
     Check(form.SetEditLocked(false)&&form.SetAllRowLocks(false),"Unlock completed rows for edit checks");
     var changed=grid.Rows[0];changed.Cells["alternative_product"].Value="ZULU CHANGED";
     Check(changed.Cells[0].Style.BackColor==Color.FromArgb(255,229,153),"Changed row yellow before sort");
     Check(form.SortNames("original_product")&&changed.Cells[0].Style.BackColor==Color.FromArgb(255,229,153),"Sorting preserves the same dirty row");
     Check(await form.SaveNames(),"Save sorted definitions");
     Check(Order(grid)=="A30,A1L,A100,B100,B30,B1L","Save retains chosen product field and variation direction");
     Check(grid.Rows.Cast<DataGridViewRow>().Where(r=>!r.IsNewRow).All(r=>r.Cells[0].Style.BackColor.IsEmpty),"Save clears dirty colors after sorting");
     form.SetEditLocked(true);
     File.WriteAllText(Path.Combine(args[0],"result.txt"),"PASS: "+checks.Count+" checks\n"+String.Join("\n",checks));
    }catch(Exception ex){File.WriteAllText(Path.Combine(args[0],"error.txt"),ex.ToString());Environment.ExitCode=1;}finally{form.Dispose();}};
    Application.Run(form);
   }
  }
 }
}
