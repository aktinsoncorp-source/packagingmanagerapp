using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using System.Windows.Forms;

namespace PackingListManager {
    static class NamesUiSmoke {
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
        static async Task Test() {
            var data=await Call(new Dictionary<string,object>{{"action","names_save"},{"revision",0},{"rows",new[]{
                new {original_product="Listing A",original_variation="Lemon",alternative_product="",alternative_variation=""},
                new {original_product="Listing B",original_variation="Rose",alternative_product="",alternative_variation=""}}}});
            using(var form=new NameDefinitionsForm(data,Call)) {
                form.SuppressDialogs=true;form.StartPosition=FormStartPosition.Manual;form.Location=new System.Drawing.Point(-30000,-30000);form.Show();
                Check(Enabled(form,"exportNames")&&!Enabled(form,"importNames"),"Only Export initially enabled");
                Check(await form.ExportToExcel(),"Export opens temporary Excel workbook");
                Check(!Enabled(form,"exportNames")&&Enabled(form,"importNames"),"Only Import enabled during Excel edit");
                string temp=form.excel.FilePath;
                form.excel.Workbook.Worksheets["Names"].Range["A2"].Value2="Updated original";
                form.excel.Workbook.Worksheets["Names"].Range["B2"].Value2="Updated variation";
                form.excel.Workbook.Worksheets["Names"].Range["C2"].Value2="Standard soap";
                form.excel.Workbook.Worksheets["Names"].Range["D2"].Value2="Lemon (1 L)";
                form.excel.Workbook.Worksheets["Names"].Range["A4"].Value2="Added original";
                form.excel.Workbook.Worksheets["Names"].Range["C4"].Value2="Added alternative";
                dynamic other=form.excel.Application.Workbooks.Add();
                Check(await form.ImportFromExcel(),"Import saves unsaved Excel edits and persists them");
                Check(!File.Exists(temp),"Temporary workbook deleted after import");
                Check(Enabled(form,"exportNames")&&!Enabled(form,"importNames"),"Buttons reset after successful import");
                Check(Convert.ToInt32(other.Worksheets.Count)>0,"Other Excel workbook remains open");
                dynamic otherApp=other.Application;other.Close(false);otherApp.Quit();
                var saved=File.ReadAllText(catalog);
                Check(saved.Contains("Standard soap")&&saved.Contains("Updated original")&&saved.Contains("Updated variation")&&saved.Contains("Added original"),"All four columns and new rows imported");

                Check(await form.ExportToExcel(),"Second export starts new session");
                temp=form.excel.FilePath;form.excel.Workbook.Close(false);
                Check(!await form.ImportFromExcel(),"Closed workbook is rejected");
                Check(form.LastError==NameExcelSession.MissingMessage,"Exact requested closed-workbook error displayed: "+form.LastError);
                Check(Enabled(form,"exportNames")&&!Enabled(form,"importNames")&&!File.Exists(temp),"Closed-workbook error resets and cleans up");

                Check(await form.ExportToExcel(),"Third export starts new session");
                form.excel.Workbook.Saved=true;form.excel.Application.Quit();
                Check(!await form.ImportFromExcel()&&form.LastError==NameExcelSession.MissingMessage,"Exiting Excel is detected");

                Check(await form.ExportToExcel(),"Export after Excel exit works");
                saved=File.ReadAllText(catalog);form.excel.Workbook.Worksheets["Names"].Range["C2"].Value2=" Invalid";
                Check(!await form.ImportFromExcel(),"Invalid alternative blocks import");
                Check(File.ReadAllText(catalog)==saved&&form.excel.IsOpen()&&Enabled(form,"importNames"),"Invalid import retains workbook and original catalog for correction");
                form.excel.Workbook.Worksheets["Names"].Range["C2"].Value2="Corrected name";
                Check(await form.ImportFromExcel(),"Corrected workbook imports successfully");
                form.Close();
            }
        }
    }
}
