using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using System.Windows.Forms;

namespace PackingListManager {
    static class LockedExportSmoke {
        static string engine,work,catalog;
        static JavaScriptSerializer json=new JavaScriptSerializer {MaxJsonLength=10000000};
        static List<string> checks=new List<string>();
        [STAThread] static void Main(string[] args) {
            engine=Path.GetFullPath(args[0]);work=Path.GetFullPath(args[1]);Directory.CreateDirectory(work);catalog=Path.Combine(work,"ProductNames.json");
            Application.EnableVisualStyles();
            var host=new Form {ShowInTaskbar=false,StartPosition=FormStartPosition.Manual,Location=new System.Drawing.Point(-30000,-30000)};
            host.Shown+=async(s,e)=>{
                try {await Test();File.WriteAllText(Path.Combine(work,"result.json"),json.Serialize(new{ok=true,checks=checks}));}
                catch(Exception ex){File.WriteAllText(Path.Combine(work,"result.json"),json.Serialize(new{ok=false,error=ex.ToString(),checks=checks}));}
                finally{host.Close();}
            };
            Application.Run(host);
        }
        static void Check(bool condition,string label){if(!condition)throw new Exception(label);checks.Add(label);}
        static async Task<Dictionary<string,object>> Call(Dictionary<string,object> request) {
            request["catalog_path"]=catalog;
            string id=Guid.NewGuid().ToString("N"),input=Path.Combine(work,id+"-request.json"),output=Path.Combine(work,id+"-response.json");
            File.WriteAllText(input,json.Serialize(request));
            await Task.Run(()=>{var psi=new ProcessStartInfo(engine,"--request \""+input+"\" --response \""+output+"\""){UseShellExecute=false,CreateNoWindow=true};using(var p=Process.Start(psi)){if(!p.WaitForExit(30000))throw new Exception("Engine timeout");}});
            var result=json.Deserialize<Dictionary<string,object>>(File.ReadAllText(output));
            if(!(bool)result["ok"])throw new Exception(Convert.ToString(result["error"]));return result;
        }
        static bool Enabled(NameDefinitionsForm form,string name){return form.Controls.Find(name,true)[0].Enabled;}
        static List<Dictionary<string,object>> Rows(object value){return ((System.Collections.IEnumerable)value).Cast<object>().Cast<Dictionary<string,object>>().ToList();}
        static async Task Test() {
            var data=await Call(new Dictionary<string,object>{{"action","names_save"},{"revision",0},{"rows",new[]{
                new {original_product="A LOCKED",original_variation="100ML",alternative_product="STANDARD A",alternative_variation="100ML",locked=true},
                new {original_product="B EDITABLE",original_variation="1L",alternative_product="STANDARD B",alternative_variation="1L",locked=false}}}});
            using(var form=new NameDefinitionsForm(data,Call)) {
                form.SuppressDialogs=true;form.StartPosition=FormStartPosition.Manual;form.Location=new System.Drawing.Point(-30000,-30000);form.Show();
                Check(form.SetEditLocked(false),"Unlock overall editing; individual locks remain");
                var grid=(DataGridView)form.Controls.Find("namesGrid",true)[0];
                Check(grid.Rows[0].Cells[0].ReadOnly&&grid.Rows[1].Cells[0].ReadOnly,"Complete alternatives default locked");
                Check(form.SetRowLocked(grid.Rows[1],false),"Explicitly unlock one completed definition");
                Check(await form.ExportToExcel(),"Mixed locked/unlocked definitions export: "+form.LastError);
                var sheet=form.excel.Workbook.Worksheets["Names"];
                Check((bool)sheet.ProtectContents,"Export retains worksheet protection");
                Check(Convert.ToString(sheet.Range["A2"].Value2)=="A LOCKED","Locked row included in export");
                bool blocked=false;try{sheet.Range["C2"].Value2="FORBIDDEN";}catch{blocked=true;}
                Check(blocked,"Locked name remains uneditable in Excel");
                sheet.Range["C3"].Value2="UPDATED STANDARD B";
                Check(await form.ImportFromExcel(),"Unlocked edit imports with locked rows retained: "+form.LastError);
                Check(File.ReadAllText(catalog).Contains("UPDATED STANDARD B"),"Unlocked edit saved");
                Check(form.SetAllRowLocks(true),"Lock every definition");
                Check(await form.ExportToExcel(),"All-locked definitions export: "+form.LastError);
                Check((bool)form.excel.Workbook.Worksheets["Names"].ProtectContents,"All-locked workbook stays protected");
                Check(Convert.ToInt32(form.excel.Workbook.Worksheets["Names"].UsedRange.Rows.Count)==3,"All locked definitions present");
                Check(await form.ImportFromExcel(),"Unchanged all-locked workbook imports");
                Check(form.SetAllRowLocks(false)&&await form.ExportToExcel(),"Unlocked export still works");
                Check(!(bool)form.excel.Workbook.Worksheets["Names"].ProtectContents,"Fully unlocked workbook permits editing");
                Check(await form.ImportFromExcel(),"Unlocked workbook imports");
                form.Close();
            }
        }
    }
}
