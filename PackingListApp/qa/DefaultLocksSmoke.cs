using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;
namespace PackingListManager {
 static class DefaultLocksSmoke {
  [STAThread] static void Main(string[] args){
   Application.EnableVisualStyles();Directory.CreateDirectory(args[0]);
   var rows=new List<Dictionary<string,object>>();
   string[] products={"STANDARD","","STANDARD",""," ","STANDARD"};
   string[] variations={"100ML","","","100ML"," ","100ML"};
   for(int i=0;i<6;i++)rows.Add(new Dictionary<string,object>{{"id",i.ToString()},{"original_product","LISTING "+i},{"original_variation","100ML"},{"alternative_product",products[i]},{"alternative_variation",variations[i]},{"locked",i==1},{"aliases",new object[0]}});
   var catalog=new Dictionary<string,object>{{"revision",0},{"rows",rows}};
   using(var form=new NameDefinitionsForm(catalog,r=>Task.FromResult(catalog))){
    form.StartPosition=FormStartPosition.Manual;form.Location=new System.Drawing.Point(-30000,-30000);form.Show();Application.DoEvents();
    if(!form.SetEditLocked(false))throw new Exception("Global unlock failed");
    var grid=(DataGridView)form.Controls.Find("namesGrid",true)[0];
    for(int i=0;i<6;i++){
     bool expected=i==0||i==5;
     if(grid.Rows[i].Cells[0].ReadOnly!=expected||grid.Rows[i].Cells[3].ReadOnly!=expected)throw new Exception("Incorrect initial lock for row "+i);
    }
    if(Convert.ToBoolean(rows[0]["locked"])||!Convert.ToBoolean(rows[1]["locked"]))throw new Exception("Opening changed source catalog objects");
    form.SetEditLocked(true);form.Close();
   }
   File.WriteAllText(Path.Combine(args[0],"result.txt"),"PASS: complete alternative pairs start locked; absent/partial/whitespace alternatives start unlocked; catalog data untouched on open.");
  }
 }
}
