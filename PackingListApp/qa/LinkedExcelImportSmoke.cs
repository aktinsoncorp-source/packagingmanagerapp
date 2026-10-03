using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using System.Windows.Forms;

namespace PackingListManager {
    static class LinkedExcelImportSmoke {
        static string engine,work,catalog;
        static JavaScriptSerializer json=new JavaScriptSerializer {MaxJsonLength=10000000};
        static List<string> checks=new List<string>();
        [STAThread] static void Main(string[] args) {
            engine=Path.GetFullPath(args[0]);work=Path.GetFullPath(args[1]);Directory.CreateDirectory(work);
            Application.EnableVisualStyles();
            var host=new Form {ShowInTaskbar=false,StartPosition=FormStartPosition.Manual,Location=new System.Drawing.Point(-30000,-30000)};
            host.Shown+=async(s,e)=>{
                try {await Test(false);await Test(true);File.WriteAllText(Path.Combine(work,"result.json"),json.Serialize(new{ok=true,checks=checks}));}
                catch(Exception ex){File.WriteAllText(Path.Combine(work,"result.json"),json.Serialize(new{ok=false,error=ex.ToString(),checks=checks}));}
                finally{host.Close();}
            };
            Application.Run(host);
        }
        static void Check(bool condition,string label){if(!condition)throw new Exception(label);checks.Add(label);}
        static List<Dictionary<string,object>> Rows(object value){return ((System.Collections.IEnumerable)value).Cast<object>().Cast<Dictionary<string,object>>().ToList();}
        static async Task<Dictionary<string,object>> Call(Dictionary<string,object> request) {
            request["catalog_path"]=catalog;
            string id=Guid.NewGuid().ToString("N"),input=Path.Combine(work,id+"-request.json"),output=Path.Combine(work,id+"-response.json");
            File.WriteAllText(input,json.Serialize(request));
            await Task.Run(()=>{var psi=new ProcessStartInfo(engine,"--request \""+input+"\" --response \""+output+"\""){UseShellExecute=false,CreateNoWindow=true};using(var p=Process.Start(psi)){if(!p.WaitForExit(30000))throw new Exception("Engine timeout");}});
            var result=json.Deserialize<Dictionary<string,object>>(File.ReadAllText(output));
            if(!(bool)result["ok"])throw new Exception(Convert.ToString(result["error"]));return result;
        }
        static async Task Test(bool approve) {
            catalog=Path.Combine(work,approve?"approved.json":"separate.json");
            var data=await Call(new Dictionary<string,object>{{"action","names_save"},{"revision",0},{"rows",new object[]{
                new {id="base",original_product="A BASE",original_variation="SMALL",alternative_product="STANDARD",alternative_variation="100ML",locked=false,aliases=new[]{new {original_product="B LINK",original_variation="SMALL"}}},
                new {id="match",original_product="C MATCH",original_variation="OTHER",alternative_product="OTHER",alternative_variation="1L",locked=false},
                new {id="locked",original_product="D LOCKED",original_variation="BIG",alternative_product="FINISHED",alternative_variation="1KG",locked=true}}}});
            var locked=Rows(data["rows"]).Single(r=>Convert.ToString(r["id"])=="locked");
            using(var form=new NameDefinitionsForm(data,Call)) {
                form.SuppressDialogs=true;form.StartPosition=FormStartPosition.Manual;form.Location=new System.Drawing.Point(-30000,-30000);form.Show();
                Check(form.SetEditLocked(false),"Overall unlock (merge "+approve+")");
                var grid=(DataGridView)form.Controls.Find("namesGrid",true)[0];
                Check(form.SetRowLocked(grid.Rows[0],false)&&form.SetRowLocked(grid.Rows[1],false),"Unlock edited definitions, retain finished row lock");
                int prompts=0;
                form.ConfirmMerge=group=>{prompts++;Check(Convert.ToString(group["base_product"])=="B LINK","App merge uses first matching Excel original as base");return approve;};
                Check(await form.ExportToExcel(),"Export linked originals: "+form.LastError);
                var sheet=form.excel.Workbook.Worksheets["Names"];
                Check((bool)sheet.ProtectContents,"Mixed workbook protection retained");
                sheet.Range["C2"].Value2="FIRST NAME";sheet.Range["D2"].Value2="30ML";
                sheet.Range["C3"].Value2="SECOND NAME";sheet.Range["D3"].Value2="100ML";
                sheet.Range["C4"].Value2="SECOND NAME";sheet.Range["D4"].Value2="100ML";
                Check(await form.ImportFromExcel(),"Different linked values import without warning: "+form.LastError);
                Check(prompts==1,"One merge confirmation handled in app");
                var saved=json.Deserialize<Dictionary<string,object>>(File.ReadAllText(catalog));var rows=Rows(saved["rows"]);
                Check(rows.Count==(approve?3:4),"Merge choice respected");
                var original=rows.Single(r=>Convert.ToString(r["id"])=="base");
                Check(Convert.ToString(original["alternative_product"])=="FIRST NAME"&&Convert.ToString(original["alternative_variation"])=="30ML","Base keeps its own complete alternatives and ID");
                var linked=rows.Single(r=>Convert.ToString(r["original_product"])=="B LINK");
                Check(Convert.ToString(linked["alternative_product"])=="SECOND NAME"&&Convert.ToString(linked["alternative_variation"])=="100ML","Separated original keeps its own alternatives");
                Check(Rows(linked["aliases"]).Count==(approve?1:0),"Linked originals join only when approved");
                Check(json.Serialize(rows.Single(r=>Convert.ToString(r["id"])=="locked"))==json.Serialize(locked),"Finished locked definition unchanged");
                Check(rows.Sum(r=>1+Rows(r["aliases"]).Count)==4,"Every original pair preserved");
                Check(form.excel==null,"Successful import closes temporary session");
                form.SetEditLocked(true);form.Close();
            }
        }
    }
}
