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
        static List<Dictionary<string,object>> Rows(object value){return ((System.Collections.IEnumerable)value).Cast<object>().Cast<Dictionary<string,object>>().ToList();}
        static async Task TestMerging() {
            catalog=Path.Combine(work,"MergedProductNames.json");
            var data=await Call(new Dictionary<string,object>{{"action","names_save"},{"revision",0},{"rows",new[]{
                new {original_product="First original",original_variation="One",alternative_product="Same name",alternative_variation="Size"},
                new {original_product="Second original",original_variation="Two",alternative_product="SAME NAME ",alternative_variation="size "}}}});
            using(var form=new NameDefinitionsForm(data,Call)) {
                form.SuppressDialogs=true;form.StartPosition=FormStartPosition.Manual;form.Location=new System.Drawing.Point(-30000,-30000);form.Show();
                var lockedGrid=(DataGridView)form.Controls.Find("namesGrid",true)[0];
                Check(lockedGrid.ReadOnly&&!lockedGrid.AllowUserToAddRows&&!Enabled(form,"exportNames")&&!Enabled(form,"importNames"),"Name editor opens locked");
                Check(!await form.SaveNames()&&!await form.ExportToExcel()&&!await form.ImportFromExcel(),"Locked editor refuses every mutation entry point");
                Check(form.SetEditLocked(false)&&!lockedGrid.ReadOnly&&lockedGrid.AllowUserToAddRows,"Unlock enables name editing");
                var grid=(DataGridView)form.Controls.Find("namesGrid",true)[0];
                Check(form.SetAllRowLocks(false),"Explicitly unlock completed definitions before merging");
                int prompts=0;string baseName="";
                form.ConfirmMerge=g=>{prompts++;baseName=Convert.ToString(g["base_product"]);return false;};
                Check(await form.SaveNames()&&prompts==1&&grid.Rows.Count==3,"Declining merge keeps two definitions");
                Check(form.SetAllRowLocks(false),"Unlock saved definitions before confirming merge");
                grid.ClearSelection();grid.CurrentCell=grid.Rows[1].Cells[0];grid.Rows[1].Selected=true;grid.Rows[0].Selected=true;
                string firstSelected=Convert.ToString(((Dictionary<string,object>)grid.Rows[1].Tag)["id"]);
                form.ConfirmMerge=g=>{prompts++;baseName=Convert.ToString(g["base_product"]);return true;};
                Check(await form.SaveNames()&&prompts==2&&grid.Rows.Count==2,"Confirming merge leaves one definition");
                Check(baseName=="SECOND ORIGINAL"&&Convert.ToString(((Dictionary<string,object>)grid.Rows[0].Tag)["id"])==firstSelected,"First selected row is kept as base");
                Check(Convert.ToString(grid.Rows[0].Cells[4].Value)=="View (2)","Merged definition shows linked original count");
                Check(grid.Rows[0].Cells[0].ReadOnly&&form.SetRowLocked(grid.Rows[0],false),"Merged definition is locked and can be unlocked again");
                grid.CurrentCell=grid.Rows[0].Cells[2];grid.BeginEdit(true);
                var editor=grid.EditingControl as TextBox;
                Check(editor!=null&&!editor.Multiline&&editor.CharacterCasing==CharacterCasing.Upper,"Grid editor enforces single line and uppercase");
                editor.Text="new standardized name   ";grid.EndEdit();
                Check(Convert.ToString(grid.Rows[0].Cells[2].Value)=="NEW STANDARDIZED NAME","Grid trims trailing spaces and normalizes case");
                Check(await form.SaveNames(),"Merged definition edits save");
                Check(form.SetRowLocked(grid.Rows[0],false),"Unlock saved definition for validation checks");
                grid.Rows[0].Cells[2].Value="First\nSecond";
                string before=File.ReadAllText(catalog);
                Check(!await form.SaveNames()&&File.ReadAllText(catalog)==before,"Programmatic multiline paste cannot save");
                grid.Rows[0].Cells[2].Value="NEW STANDARDIZED NAME";
                Check(await form.SaveNames(),"Correction after multiline rejection saves");
                Check(form.SetRowLocked(grid.Rows[0],false),"Unlock merged definition for linked Excel edits");
                Check(await form.ExportToExcel(),"Merged definition exports to Excel");
                Check(Convert.ToInt32(form.excel.Workbook.Worksheets["Names"].UsedRange.Rows.Count)==3,"Excel exposes both linked originals for editing");
                form.excel.Workbook.Worksheets["Names"].Range["A3"].Value2="Renamed linked original ";
                form.excel.Workbook.Worksheets["Names"].Range["C3"].Value2="Updated shared name ";
                Check(await form.ImportFromExcel(),"Linked original and shared alternative edits import");
                Check(grid.Rows.Count==2&&File.ReadAllText(catalog).Contains("RENAMED LINKED ORIGINAL")&&File.ReadAllText(catalog).Contains("UPDATED SHARED NAME"),"Excel import preserves one merged definition");
                Check(form.SetRowLocked(grid.Rows[0],false),"Unlock imported definition before adding matching originals");
                Check(await form.ExportToExcel(),"Merged definition re-exports");
                form.excel.Workbook.Worksheets["Names"].Range["A4"].Value2="Third original";
                form.excel.Workbook.Worksheets["Names"].Range["B4"].Value2="Three";
                form.excel.Workbook.Worksheets["Names"].Range["C4"].Value2="updated shared name";
                form.excel.Workbook.Worksheets["Names"].Range["D4"].Value2="size";
                int oldPrompts=prompts;
                Check(await form.ImportFromExcel()&&prompts==oldPrompts+1&&grid.Rows.Count==2,"Excel import asks before merging a new matching definition");
                Check(Convert.ToString(grid.Rows[0].Cells[4].Value)=="View (3)","All three original pairs remain linked");
                var search=form.Controls.OfType<TableLayoutPanel>().First().Controls.OfType<FlowLayoutPanel>().First().Controls.OfType<TextBox>().First();
                search.Text="Renamed linked original";
                Check(grid.Rows[0].Visible,"Search finds linked originals as well as the base");
                search.Clear();
                using(var bitmap=new System.Drawing.Bitmap(form.Width,form.Height)){form.DrawToBitmap(bitmap,new System.Drawing.Rectangle(System.Drawing.Point.Empty,form.Size));bitmap.Save(Path.Combine(work,"merged-names.png"));}
                Check(form.SetRowLocked(grid.Rows[0],false),"Explicitly unlock saved merged definition for deletion");
                grid.ClearSelection();grid.Rows[0].Selected=true;
                form.Controls.OfType<TableLayoutPanel>().First().Controls.OfType<FlowLayoutPanel>().SelectMany(p=>p.Controls.OfType<Button>()).Single(b=>b.Text=="Delete selected").PerformClick();
                Check(await form.SaveNames()&&Rows(json.Deserialize<Dictionary<string,object>>(File.ReadAllText(catalog))["rows"]).Count==0,"Explicitly unlocked definition and linked originals can be deleted after auto-locking");
                Check(form.SetEditLocked(true)&&lockedGrid.ReadOnly,"Saved name editor can be relocked");
                form.Close();
            }
        }
        static async Task Test() {
            var data=await Call(new Dictionary<string,object>{{"action","names_save"},{"revision",0},{"rows",new[]{
                new {original_product="Listing A",original_variation="Lemon",alternative_product="",alternative_variation=""},
                new {original_product="Listing B",original_variation="Rose",alternative_product="",alternative_variation=""}}}});
            using(var form=new NameDefinitionsForm(data,Call)) {
                form.SuppressDialogs=true;form.StartPosition=FormStartPosition.Manual;form.Location=new System.Drawing.Point(-30000,-30000);form.Show();
                var lockedGrid=(DataGridView)form.Controls.Find("namesGrid",true)[0];
                Check(lockedGrid.ReadOnly&&!lockedGrid.AllowUserToAddRows&&!Enabled(form,"exportNames")&&!Enabled(form,"importNames"),"Name editor opens locked");
                Check(!await form.SaveNames()&&!await form.ExportToExcel()&&!await form.ImportFromExcel(),"Locked editor refuses every mutation entry point");
                Check(form.SetEditLocked(false)&&!lockedGrid.ReadOnly&&lockedGrid.AllowUserToAddRows,"Unlock enables name editing");
                Check(Enabled(form,"exportNames")&&!Enabled(form,"importNames"),"Only Export initially enabled");
                Check(await form.ExportToExcel(),"Export opens temporary Excel workbook");
                Check(!Enabled(form,"exportNames")&&Enabled(form,"importNames"),"Only Import enabled during Excel edit");
                Check(!form.SetEditLocked(true)&&!Enabled(form,"lockNames"),"Pending Excel import cannot be locked halfway through");
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
                Check(saved.Contains("STANDARD SOAP")&&saved.Contains("UPDATED ORIGINAL")&&saved.Contains("UPDATED VARIATION")&&saved.Contains("ADDED ORIGINAL"),"All four columns and new rows imported");

                Check(await form.ExportToExcel(),"Second export starts new session");
                temp=form.excel.FilePath;form.excel.Workbook.Close(false);
                Check(!await form.ImportFromExcel(),"Closed workbook is rejected");
                Check(form.LastError==NameExcelSession.MissingMessage,"Exact requested closed-workbook error displayed: "+form.LastError);
                Check(Enabled(form,"exportNames")&&!Enabled(form,"importNames")&&!File.Exists(temp),"Closed-workbook error resets and cleans up");

                Check(await form.ExportToExcel(),"Third export starts new session");
                form.excel.Workbook.Saved=true;form.excel.Application.Quit();
                Check(!await form.ImportFromExcel()&&form.LastError==NameExcelSession.MissingMessage,"Exiting Excel is detected");

                Check(form.SetAllRowLocks(false),"Unlock imported rows before correction workflow");
                Check(await form.ExportToExcel(),"Export after Excel exit works");
                saved=File.ReadAllText(catalog);form.excel.Workbook.Worksheets["Names"].Range["C2"].Value2="Invalid:Name";
                Check(!await form.ImportFromExcel(),"Invalid alternative blocks import");
                Check(File.ReadAllText(catalog)==saved&&form.excel.IsOpen()&&Enabled(form,"importNames"),"Invalid import retains workbook and original catalog for correction");
                form.excel.Workbook.Worksheets["Names"].Range["C2"].Value2="Corrected name";
                Check(await form.ImportFromExcel(),"Corrected workbook imports successfully");
                Check(form.SetEditLocked(true)&&lockedGrid.ReadOnly,"Saved name editor can be relocked");
                form.Close();
            }
            await TestMerging();
        }
    }
}
