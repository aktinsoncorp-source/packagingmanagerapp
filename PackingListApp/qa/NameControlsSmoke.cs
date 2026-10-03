using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
namespace PackingListManager {
 static class NameControlsSmoke {
  static void Check(bool condition,string message){if(!condition)throw new Exception(message);}
  [STAThread] static void Main(string[] args){
   Application.EnableVisualStyles();Directory.CreateDirectory(args[0]);
   var rows=new List<Dictionary<string,object>>();
   foreach(string size in new[]{"1L","30mL","100mL","1kg","100g"})rows.Add(new Dictionary<string,object>{{"id",size},{"original_product","SOAP"},{"original_variation",size},{"alternative_product","INVENTORY SOAP"},{"alternative_variation",size},{"locked",false},{"aliases",new object[0]}});
   var catalog=new Dictionary<string,object>{{"revision",1},{"rows",rows}};
   Func<Dictionary<string,object>,Task<Dictionary<string,object>>> call=request=>{
    if((string)request["action"]=="names_prepare")return Task.FromResult(new Dictionary<string,object>{{"rows",request["rows"]},{"merges",new object[0]}});
    return Task.FromResult(new Dictionary<string,object>{{"revision",2},{"rows",request["rows"]}});
   };
   using(var form=new NameDefinitionsForm(catalog,call)){
    form.SuppressDialogs=true;form.StartPosition=FormStartPosition.Manual;form.Location=new Point(-30000,-30000);
    form.Shown+=async(s,e)=>{try {
     var grid=(DataGridView)form.Controls.Find("namesGrid",true)[0];
     Check(grid.ReadOnly,"Starts globally locked");Check(form.SetEditLocked(false),"Unlock global");Check(form.SetAllRowLocks(false)&&await form.SaveNames(),"Explicit unlock retained during session");
     Check(NameDefinitionsForm.CompareVariation("30mL","100mL")<0,"30mL before100mL");Check(NameDefinitionsForm.CompareVariation("100mL","1L")<0,"100mL before1L");Check(NameDefinitionsForm.CompareVariation("100g","1kg")<0,"100g before1kg");Check(NameDefinitionsForm.CompareVariation("0.5L","100ML")>0,"Decimal units");
     form.SortNames("alternative_variation");
     Check(Convert.ToString(grid.Rows[0].Cells["alternative_variation"].Value).ToUpperInvariant()=="100G","Ascending size groups");
     form.SortNames("alternative_variation");
     Check(Convert.ToString(grid.Rows[0].Cells["alternative_variation"].Value)=="1L","Descending size");
     var row=grid.Rows[0];Check(form.SetRowLocked(row,true),"Lock row");Check(row.Cells["original_product"].ReadOnly&&row.Cells["alternative_variation"].ReadOnly,"Whole row protected");
     Check(form.SetRowLocked(row,false),"Unlock row");row.Cells["alternative_product"].Value="CHANGED INVENTORY NAME";
     Check(row.Cells["original_product"].Style.BackColor==Color.FromArgb(255,229,153),"Changed row highlighted");
     Check(NameDefinitionsForm.CorrectAlternative(" @ SOAP ")=="SOAP","Automatic leading correction");
     Check(form.SetAllRowLocks(true),"All lock");Check(grid.Rows.Cast<DataGridViewRow>().Where(r=>!r.IsNewRow).All(r=>r.Cells[0].ReadOnly),"Every row locked");
     Check(form.SetAllRowLocks(false),"All unlock");Check(await form.SaveNames(),"Save names");
     Check(grid.Rows[0].Cells[0].Style.BackColor.IsEmpty,"Save clears changed color");
     row=grid.Rows[0];form.SetRowLocked(row,true);grid.Rows[1].Cells["alternative_product"].Value="UNSAVED INVENTORY NAME";
     Application.DoEvents();using(var b=new Bitmap(form.Width,form.Height)){form.DrawToBitmap(b,new Rectangle(Point.Empty,form.Size));b.Save(Path.Combine(args[0],"names-controls.png"));}
     Check(await form.SaveNames(),"Save locks");Check(form.SetEditLocked(true),"Relock global");
     File.WriteAllText(Path.Combine(args[0],"result.txt"),"PASS: sorting both directions, normalized units, row locks, all locks, row highlighting, save reset, correction and global locking.");
    }catch(Exception ex){File.WriteAllText(Path.Combine(args[0],"error.txt"),ex.ToString());Environment.ExitCode=1;}finally{form.Dispose();}};
    Application.Run(form);
   }
  }
 }
}
