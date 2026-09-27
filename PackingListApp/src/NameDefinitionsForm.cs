using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PackingListManager {
    public class NameDefinitionsForm : Form {
        readonly Func<Dictionary<string,object>,Task<Dictionary<string,object>>> call;
        readonly DataGridView grid = new DataGridView();
        readonly TextBox search = new TextBox();
        readonly Label message = new Label();
        readonly Label count = new Label();
        readonly FlowLayoutPanel buttons = new FlowLayoutPanel();
        readonly string[] fields = {"original_product","original_variation","alternative_product","alternative_variation"};
        int revision;
        bool dirty, loading, working;
        Button saveButton,addButton,deleteButton,exportButton,importButton;
        readonly List<DataGridViewRow> selectionOrder=new List<DataGridViewRow>();
        internal Func<Dictionary<string,object>,bool> ConfirmMerge=null;
        internal NameExcelSession excel;
        internal bool SuppressDialogs=false;
        internal string LastError;

        public NameDefinitionsForm(Dictionary<string,object> catalog, Func<Dictionary<string,object>,Task<Dictionary<string,object>>> send) {
            call=send;Text="Define Name";Font=new Font("Segoe UI",10F);Size=new Size(1320,760);MinimumSize=new Size(1000,560);
            StartPosition=FormStartPosition.CenterParent;BackColor=Color.FromArgb(246,248,251);
            var layout=new TableLayoutPanel {Dock=DockStyle.Fill,Padding=new Padding(20),ColumnCount=1,RowCount=5};
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute,62));layout.RowStyles.Add(new RowStyle(SizeType.Absolute,42));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent,100));layout.RowStyles.Add(new RowStyle(SizeType.Absolute,48));layout.RowStyles.Add(new RowStyle(SizeType.Absolute,46));
            var note=new Label {Dock=DockStyle.Fill,Text="Names are UPPERCASE and single-line. Blank alternatives keep the original field. Matching alternatives can be merged when saving.\nSelect the base row first. Use View to see its linked originals. For Excel: Export, edit, keep Excel open, then Import."};
            layout.Controls.Add(note,0,0);
            var find=new FlowLayoutPanel {Dock=DockStyle.Fill,WrapContents=false};
            find.Controls.Add(new Label {Text="Search",AutoSize=true,Margin=new Padding(0,6,10,0)});
            search.Width=390;find.Controls.Add(search);count.AutoSize=true;count.Margin=new Padding(16,6,0,0);find.Controls.Add(count);layout.Controls.Add(find,0,1);
            grid.Name="namesGrid";grid.Dock=DockStyle.Fill;grid.BackgroundColor=Color.White;grid.RowHeadersWidth=32;
            grid.AllowUserToAddRows=true;grid.AllowUserToDeleteRows=false;grid.MultiSelect=true;grid.SelectionMode=DataGridViewSelectionMode.FullRowSelect;
            grid.AutoSizeColumnsMode=DataGridViewAutoSizeColumnsMode.Fill;grid.AutoSizeRowsMode=DataGridViewAutoSizeRowsMode.AllCells;
            grid.DefaultCellStyle.WrapMode=DataGridViewTriState.True;grid.DefaultCellStyle.Padding=new Padding(5);grid.DefaultCellStyle.SelectionForeColor=Color.Black;
            grid.DefaultCellStyle.SelectionBackColor=Color.FromArgb(220,234,252);grid.RowTemplate.MinimumHeight=48;
            grid.EnableHeadersVisualStyles=false;grid.ColumnHeadersHeight=40;grid.ColumnHeadersDefaultCellStyle.BackColor=Color.FromArgb(232,238,247);
            grid.ColumnHeadersDefaultCellStyle.Font=new Font(Font,FontStyle.Bold);
            var titles=new[]{"Original product name","Original variation","Alternative product name","Alternative variation"};
            var weights=new[]{31F,16F,26F,16F};
            for(int i=0;i<4;i++) {
                var column=new DataGridViewTextBoxColumn {Name=fields[i],HeaderText=titles[i],FillWeight=weights[i],SortMode=DataGridViewColumnSortMode.Automatic};
                if(i>=2){column.DefaultCellStyle.BackColor=Color.FromArgb(255,242,216);column.HeaderCell.Style.BackColor=Color.FromArgb(248,203,173);column.DefaultCellStyle.Font=new Font(Font,FontStyle.Bold);}
                grid.Columns.Add(column);
            }
            grid.Columns.Add(new DataGridViewButtonColumn {Name="linkedOriginals",HeaderText="Original pairs",FillWeight=11F,ReadOnly=true,SortMode=DataGridViewColumnSortMode.NotSortable});
            grid.CellContentClick+=(s,e)=>{if(e.RowIndex>=0&&e.ColumnIndex==4&&!grid.Rows[e.RowIndex].IsNewRow)ShowOriginals(grid.Rows[e.RowIndex]);};
            grid.EditingControlShowing+=(s,e)=>{var text=e.Control as TextBox;if(text!=null){text.CharacterCasing=CharacterCasing.Upper;text.Multiline=false;text.AcceptsReturn=false;}};
            grid.CellValidating+=(s,e)=>{
                if(e.RowIndex<0||e.ColumnIndex>=4||grid.Rows[e.RowIndex].IsNewRow)return;
                if(!SingleLine(Convert.ToString(e.FormattedValue)??"")){e.Cancel=true;grid.Rows[e.RowIndex].ErrorText="Use one line only; paragraphs and line breaks are not allowed.";message.Text=grid.Rows[e.RowIndex].ErrorText;}
            };
            grid.CellEndEdit+=(s,e)=>{
                if(e.RowIndex<0||e.ColumnIndex>=4)return;
                var row=grid.Rows[e.RowIndex];row.ErrorText="";
                var value=Convert.ToString(row.Cells[e.ColumnIndex].Value)??"";
                row.Cells[e.ColumnIndex].Value=value.TrimEnd(' ').ToUpperInvariant();
            };
            grid.SelectionChanged+=(s,e)=>{
                if(loading)return;
                selectionOrder.RemoveAll(r=>r.DataGridView!=grid||!r.Selected||r.IsNewRow);
                if(grid.CurrentRow!=null&&grid.CurrentRow.Selected&&!grid.CurrentRow.IsNewRow&&!selectionOrder.Contains(grid.CurrentRow))selectionOrder.Add(grid.CurrentRow);
                foreach(var row in grid.SelectedRows.Cast<DataGridViewRow>().Where(r=>!r.IsNewRow).OrderBy(r=>r.Index))if(!selectionOrder.Contains(row))selectionOrder.Add(row);
            };
            grid.CellValueChanged+=(s,e)=>{if(!loading&&e.RowIndex>=0){dirty=true;message.Text="Unsaved changes. Click Save names to apply them.";}};
            grid.UserAddedRow+=(s,e)=>{if(!loading)dirty=true;};
            grid.DataError+=(s,e)=>{e.ThrowException=false;message.Text="Please enter a single line of text in each name field.";};
            search.TextChanged+=(s,e)=>Filter();layout.Controls.Add(grid,0,2);
            buttons.Dock=DockStyle.Fill;buttons.WrapContents=false;buttons.Padding=new Padding(0,8,0,0);
            saveButton=Button("Save names",140);saveButton.BackColor=Color.FromArgb(25,87,157);saveButton.ForeColor=Color.White;saveButton.Click+=async(s,e)=>await SaveNames();
            addButton=Button("Add row",100);addButton.Click+=(s,e)=>{search.Clear();grid.CurrentCell=grid.Rows[grid.NewRowIndex].Cells[0];grid.BeginEdit(true);};
            deleteButton=Button("Delete selected",150);deleteButton.Click+=(s,e)=>{
                var selected=grid.SelectedRows.Cast<DataGridViewRow>().Where(r=>!r.IsNewRow).ToArray();
                foreach(var row in selected)grid.Rows.Remove(row);
                if(selected.Length>0){dirty=true;message.Text="Rows removed from this edit. Save names to delete them; a later PDF scan can add them back with blank alternatives.";Filter();}
            };
            buttons.Controls.Add(new Label {Text="Excel",AutoSize=true,Margin=new Padding(12,6,6,0)});
            exportButton=Button("Export",85);exportButton.Name="exportNames";exportButton.Click+=async(s,e)=>await ExportToExcel();
            importButton=Button("Import",85);importButton.Name="importNames";importButton.Click+=async(s,e)=>await ImportFromExcel();
            var close=Button("Close",95);close.Click+=(s,e)=>Close();layout.Controls.Add(buttons,0,3);
            message.Dock=DockStyle.Fill;message.Font=new Font("Segoe UI",9F);message.Text="Start alternatives with a letter or number. Common symbols are allowed afterwards. Trailing spaces are removed.";layout.Controls.Add(message,0,4);Controls.Add(layout);
            revision=Convert.ToInt32(catalog["revision"]);Fill(Rows(catalog["rows"]));RefreshExcelState();
            Disposed+=(s,e)=>{if(excel!=null){excel.Dispose();excel=null;}};
            FormClosing+=async(s,e)=>{
                if(working){e.Cancel=true;return;}
                grid.EndEdit();
                if(excel!=null){
                    if(MessageBox.Show(this,"An Excel edit is still pending. Close Define Name and discard the unimported Excel changes?","Excel edit pending",MessageBoxButtons.YesNo,MessageBoxIcon.Question,MessageBoxDefaultButton.Button2)!=DialogResult.Yes){e.Cancel=true;return;}
                    ResetExcel();
                }
                if(!dirty)return;
                var answer=MessageBox.Show(this,"Save your name changes before closing?","Unsaved names",MessageBoxButtons.YesNoCancel,MessageBoxIcon.Question);
                if(answer==DialogResult.Cancel){e.Cancel=true;return;}
                if(answer==DialogResult.Yes){e.Cancel=true;if(await SaveNames())Close();}
            };
        }

        Button Button(string title,int width) {
            var result=new Button {Text=title,Width=width,Height=32,BackColor=Color.White,FlatStyle=FlatStyle.Flat,Margin=new Padding(0,0,8,0)};
            buttons.Controls.Add(result);return result;
        }
        static List<Dictionary<string,object>> Rows(object value) {
            return ((System.Collections.IEnumerable)value).Cast<object>().Cast<Dictionary<string,object>>().ToList();
        }
        void Fill(List<Dictionary<string,object>> rows) {
            loading=true;
            try {
                selectionOrder.Clear();grid.Rows.Clear();
                foreach(var row in rows){int i=grid.Rows.Add(fields.Select(f=>row[f]).Concat(new object[]{"View ("+(Aliases(row).Count+1)+")"}).ToArray());grid.Rows[i].Tag=row;}
                grid.ClearSelection();
            }
            finally {loading=false;}
            Filter();
        }
        List<Dictionary<string,object>> CollectRows() {
            if(!grid.EndEdit())throw new Exception("Use one line only; paragraphs and line breaks are not allowed.");
            var result=new List<Dictionary<string,object>>();
            foreach(DataGridViewRow row in grid.Rows) {
                if(row.IsNewRow)continue;
                var values=fields.Select(f=>Convert.ToString(row.Cells[f].Value)??"").ToArray();
                if(values.All(v=>v=="")&&row.Tag==null)continue;
                var old=row.Tag as Dictionary<string,object>;
                var entry=old==null?new Dictionary<string,object>{{"id",Guid.NewGuid().ToString("N")}}:new Dictionary<string,object>(old);
                for(int i=0;i<4;i++) {
                    if(!SingleLine(values[i]))throw new Exception("Use one line only; paragraphs and line breaks are not allowed.");
                    entry[fields[i]]=values[i].TrimEnd(' ').ToUpperInvariant();
                }
                row.Tag=entry;result.Add(entry);
            }
            return result;
        }
        void Filter() {
            if(loading)return;
            string query=search.Text.Trim();grid.CurrentCell=null;int total=0,visible=0;
            foreach(DataGridViewRow row in grid.Rows) {
                if(row.IsNewRow)continue;total++;
                var entry=row.Tag as Dictionary<string,object>;
                row.Visible=query.Length==0||row.Cells.Cast<DataGridViewCell>().Take(4).Any(c=>(Convert.ToString(c.Value)??"").IndexOf(query,StringComparison.OrdinalIgnoreCase)>=0)||
                    (entry!=null&&Aliases(entry).Any(a=>fields.Take(2).Any(f=>Convert.ToString(a[f]).IndexOf(query,StringComparison.OrdinalIgnoreCase)>=0)));
                if(row.Visible)visible++;
            }
            count.Text=visible+" of "+total+" definitions";
        }
        async Task<bool> Run(Func<Task> action) {
            working=true;buttons.Enabled=false;grid.Enabled=false;search.Enabled=false;UseWaitCursor=true;
            try {await action();return true;}
            catch(Exception ex){LastError=ex.Message;message.Text=ex.Message;if(!SuppressDialogs)MessageBox.Show(this,ex.Message,"Check names",MessageBoxButtons.OK,MessageBoxIcon.Warning);return false;}
            finally{working=false;buttons.Enabled=true;search.Enabled=true;UseWaitCursor=false;RefreshExcelState();}
        }
        void RefreshExcelState() {
            bool pending=excel!=null;
            exportButton.Enabled=!working&&!pending;importButton.Enabled=!working&&pending;
            grid.Enabled=!working&&!pending;saveButton.Enabled=!working&&!pending;addButton.Enabled=!working&&!pending;deleteButton.Enabled=!working&&!pending;
        }
        void ResetExcel(){if(excel!=null)excel.Dispose();excel=null;RefreshExcelState();}
        internal async Task<bool> ExportToExcel() {
            if(excel!=null)return false;
            if(!await SaveNames())return false;
            return await Run(async()=>{
                excel=new NameExcelSession();
                try {
                    await call(new Dictionary<string,object>{{"action","names_export"},{"path",excel.FilePath}});
                    excel.Open(!SuppressDialogs);
                    message.Text="Edit the Excel workbook. Keep it open and click Import here when finished; your edits will be saved automatically.";
                } catch {ResetExcel();throw;}
            });
        }
        internal async Task<bool> ImportFromExcel() {
            return await Run(async()=>{
                if(excel==null||!excel.IsOpen()||!System.IO.File.Exists(excel.FilePath)) {
                    ResetExcel();throw new Exception(NameExcelSession.MissingMessage);
                }
                excel.Save();
                var preview=await call(new Dictionary<string,object>{{"action","names_import"},{"path",excel.FilePath},{"rows",CollectRows()}});
                var saved=await SavePrepared(Rows(preview["rows"]),preview.ContainsKey("priority")?preview["priority"]:new string[0]);
                revision=Convert.ToInt32(saved["revision"]);Fill(Rows(saved["rows"]));dirty=false;
                ResetExcel();message.Text="Excel names imported and saved: "+preview["added"]+" added, "+preview["updated"]+" updated. Temporary workbook closed and removed.";
            });
        }
        static bool SingleLine(string value){return !value.Any(c=>Char.IsControl(c)||c=='\u2028'||c=='\u2029');}
        static List<Dictionary<string,object>> Aliases(Dictionary<string,object> row){return row.ContainsKey("aliases")?Rows(row["aliases"]):new List<Dictionary<string,object>>();}
        void ShowOriginals(DataGridViewRow row) {
            grid.EndEdit();var entry=row.Tag as Dictionary<string,object>;
            var originals=entry==null?new List<Dictionary<string,object>>():Aliases(entry);
            originals.Insert(0,new Dictionary<string,object>{{"original_product",Convert.ToString(row.Cells[0].Value)??""},{"original_variation",Convert.ToString(row.Cells[1].Value)??""}});
            using(var view=new Form {Text="Linked original names - "+originals.Count+" pairs",Size=new Size(980,520),MinimumSize=new Size(650,350),StartPosition=FormStartPosition.CenterParent,Font=Font}) {
                var list=new DataGridView {Dock=DockStyle.Fill,ReadOnly=true,AllowUserToAddRows=false,AllowUserToDeleteRows=false,RowHeadersVisible=false,AutoSizeColumnsMode=DataGridViewAutoSizeColumnsMode.Fill,AutoSizeRowsMode=DataGridViewAutoSizeRowsMode.AllCells,BackgroundColor=Color.White};
                list.DefaultCellStyle.WrapMode=DataGridViewTriState.True;list.DefaultCellStyle.Padding=new Padding(6);
                list.Columns.Add(new DataGridViewTextBoxColumn {HeaderText="Original product name",FillWeight=70});list.Columns.Add(new DataGridViewTextBoxColumn {HeaderText="Original variation",FillWeight=30});
                foreach(var original in originals)list.Rows.Add(original["original_product"],original["original_variation"]);
                var help=new Label {Dock=DockStyle.Top,Height=60,Padding=new Padding(10),Text="All these original pairs use this definition. The first pair is the base.\nExcel export includes every pair so you can edit their original names and variations."};
                view.Controls.Add(list);view.Controls.Add(help);view.ShowDialog(this);
            }
        }
        async Task<Dictionary<string,object>> SavePrepared(List<Dictionary<string,object>> pending,object priority) {
            var preview=await call(new Dictionary<string,object>{{"action","names_prepare"},{"rows",pending},{"priority",priority}});
            var approvals=new List<object>();
            foreach(var group in Rows(preview["merges"])) {
                bool yes;
                if(ConfirmMerge!=null)yes=ConfirmMerge(group);
                else if(SuppressDialogs)yes=false;
                else {
                    string text="These definitions have the same alternative product name AND variation:\n\n"+group["alternative_product"]+"\n"+group["alternative_variation"]+"\n\nMerge them into one definition and remember all "+Rows(group["originals"]).Count+" original pairs?\n\nBase original (first selected, otherwise first listed):\n"+group["base_product"]+" / "+group["base_variation"]+"\n\nYes = merge. No = keep separate definitions.";
                    yes=MessageBox.Show(this,text,"Merge matching names?",MessageBoxButtons.YesNo,MessageBoxIcon.Question,MessageBoxDefaultButton.Button2)==DialogResult.Yes;
                }
                if(yes)approvals.Add(group["ids"]);
            }
            return await call(new Dictionary<string,object>{{"action","names_save"},{"rows",preview["rows"]},{"revision",revision},{"merge_groups",approvals}});
        }
        internal async Task<bool> SaveNames() {
            return await Run(async()=>{
                var pending=CollectRows();
                var priority=selectionOrder.Where(r=>r.Tag is Dictionary<string,object>).Select(r=>Convert.ToString(((Dictionary<string,object>)r.Tag)["id"])).ToArray();
                var result=await SavePrepared(pending,priority);
                revision=Convert.ToInt32(result["revision"]);Fill(Rows(result["rows"]));dirty=false;message.Text="Names saved. They will be used in the next generated documents.";
            });
        }
    }
}
