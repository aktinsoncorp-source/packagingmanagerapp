using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Globalization;
using System.Text.RegularExpressions;
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
        readonly string[] fieldTitles = {"Original product name","Original variation","Alternative product name","Alternative variation"};
        readonly Font sortFont = new Font("Segoe UI",8.5F);
        string productSortField="original_product",variationSortField="original_variation";
        bool productDescending,variationDescending;
        int revision;
        bool dirty, loading, working;
        bool editLocked=true;
        Button lockButton,unlockButton;
        readonly ToolTip lockTips=new ToolTip();
        Button saveButton,addButton,deleteButton,exportButton,importButton;
        readonly Bitmap rowClosed=LockIcon(false),rowOpen=LockIcon(true);
        readonly HashSet<DataGridViewRow> lockedRows=new HashSet<DataGridViewRow>();
        readonly Dictionary<DataGridViewRow,string> savedRows=new Dictionary<DataGridViewRow,string>();
        readonly List<DataGridViewRow> selectionOrder=new List<DataGridViewRow>();
        internal Func<Dictionary<string,object>,bool> ConfirmMerge=null;
        internal NameExcelSession excel;
        internal bool SuppressDialogs=false;
        internal string LastError;

        public NameDefinitionsForm(Dictionary<string,object> catalog, Func<Dictionary<string,object>,Task<Dictionary<string,object>>> send) {
            call=send;Text="Define Name";Font=new Font("Segoe UI",10F);Size=new Size(1320,760);MinimumSize=new Size(1000,560);
            StartPosition=FormStartPosition.CenterParent;BackColor=Color.FromArgb(246,248,251);
            var layout=new TableLayoutPanel {Dock=DockStyle.Fill,Padding=new Padding(20),ColumnCount=1,RowCount=5};
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute,78));layout.RowStyles.Add(new RowStyle(SizeType.Absolute,42));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent,100));layout.RowStyles.Add(new RowStyle(SizeType.Absolute,48));layout.RowStyles.Add(new RowStyle(SizeType.Absolute,46));
            var note=new Label {Dock=DockStyle.Fill,Text="Choose either product heading for A-Z / Z-A, then either variation heading for sizes within each product group.\nName and variation choices are independent. Black arrows show direction; hover over a heading for its next action.\nYellow marks unsaved changes. Unlock editing to use row locks. Blank alternatives keep their original values."};
            layout.Controls.Add(note,0,0);
            var find=new FlowLayoutPanel {Dock=DockStyle.Fill,WrapContents=false};
            find.Controls.Add(new Label {Text="Search",AutoSize=true,Margin=new Padding(0,6,10,0)});
            search.Width=390;find.Controls.Add(search);count.AutoSize=true;count.Margin=new Padding(16,6,0,0);find.Controls.Add(count);
            lockButton=new Button {Name="lockNames",Text="",Width=38,Height=32,Image=LockIcon(false),AccessibleName="Lock name editing",FlatStyle=FlatStyle.Flat,Margin=new Padding(18,0,4,0)};
            unlockButton=new Button {Name="unlockNames",Text="",Width=38,Height=32,Image=LockIcon(true),AccessibleName="Unlock name editing",FlatStyle=FlatStyle.Flat,Margin=new Padding(0,0,0,0)};
            lockTips.SetToolTip(lockButton,"Lock name editing");lockTips.SetToolTip(unlockButton,"Unlock name editing");
            lockButton.Click+=(s,e)=>SetEditLocked(true);unlockButton.Click+=(s,e)=>SetEditLocked(false);
            find.Controls.Add(lockButton);find.Controls.Add(unlockButton);layout.Controls.Add(find,0,1);
            grid.Name="namesGrid";grid.Dock=DockStyle.Fill;grid.BackgroundColor=Color.White;grid.RowHeadersWidth=52;
            grid.AllowUserToAddRows=true;grid.AllowUserToDeleteRows=false;grid.MultiSelect=true;grid.SelectionMode=DataGridViewSelectionMode.FullRowSelect;
            grid.AutoSizeColumnsMode=DataGridViewAutoSizeColumnsMode.Fill;grid.AutoSizeRowsMode=DataGridViewAutoSizeRowsMode.AllCells;
            grid.DefaultCellStyle.WrapMode=DataGridViewTriState.True;grid.DefaultCellStyle.Padding=new Padding(5);grid.DefaultCellStyle.SelectionForeColor=Color.Black;
            grid.DefaultCellStyle.SelectionBackColor=Color.FromArgb(220,234,252);grid.RowTemplate.MinimumHeight=48;
            grid.EnableHeadersVisualStyles=false;grid.ColumnHeadersHeightSizeMode=DataGridViewColumnHeadersHeightSizeMode.DisableResizing;grid.ColumnHeadersHeight=62;grid.ColumnHeadersDefaultCellStyle.BackColor=Color.FromArgb(232,238,247);
            grid.ColumnHeadersDefaultCellStyle.ForeColor=Color.Black;grid.ColumnHeadersDefaultCellStyle.SelectionForeColor=Color.Black;
            grid.ColumnHeadersDefaultCellStyle.Font=new Font(Font,FontStyle.Bold);
            var weights=new[]{31F,16F,26F,16F};
            for(int i=0;i<4;i++) {
                var column=new DataGridViewTextBoxColumn {Name=fields[i],HeaderText=fieldTitles[i],FillWeight=weights[i],MinimumWidth=i%2==0?225:180,SortMode=DataGridViewColumnSortMode.Programmatic};
                if(i>=2){column.DefaultCellStyle.BackColor=Color.FromArgb(255,242,216);column.HeaderCell.Style.BackColor=Color.FromArgb(248,203,173);column.DefaultCellStyle.Font=new Font(Font,FontStyle.Bold);}
                grid.Columns.Add(column);
            }
            grid.Columns.Add(new DataGridViewButtonColumn {Name="linkedOriginals",HeaderText="Original pairs",FillWeight=11F,ReadOnly=true,SortMode=DataGridViewColumnSortMode.NotSortable});
            var rowLockColumn=new DataGridViewImageColumn {Name="rowLock",HeaderText="",Width=72,AutoSizeMode=DataGridViewAutoSizeColumnMode.None,ReadOnly=true,ImageLayout=DataGridViewImageCellLayout.Normal,SortMode=DataGridViewColumnSortMode.NotSortable};
            rowLockColumn.DefaultCellStyle.NullValue=null;grid.Columns.Add(rowLockColumn);rowLockColumn.DisplayIndex=0;
            rowLockColumn.HeaderCell.ToolTipText="Lock all rows (left padlock) / unlock all rows (right padlock)";
            grid.CellPainting+=(s,e)=>{if(e.RowIndex==-1&&e.ColumnIndex>=0&&e.ColumnIndex<4){PaintSortHeader(e);return;}if(e.RowIndex>=0&&grid.Rows[e.RowIndex].IsNewRow&&e.ColumnIndex==grid.Columns["rowLock"].Index){e.PaintBackground(e.ClipBounds,false);e.Handled=true;return;}if(e.RowIndex==-1&&e.ColumnIndex==grid.Columns["rowLock"].Index){e.PaintBackground(e.ClipBounds,false);e.Graphics.DrawImage(rowClosed,e.CellBounds.Left+8,e.CellBounds.Top+8);e.Graphics.DrawImage(rowOpen,e.CellBounds.Left+39,e.CellBounds.Top+8);e.Handled=true;}};
            grid.ColumnHeaderMouseClick+=(s,e)=>{if(e.ColumnIndex>=0&&e.ColumnIndex<4)SortNames(fields[e.ColumnIndex]);};
            grid.CellMouseClick+=(s,e)=>{if(e.ColumnIndex!=grid.Columns["rowLock"].Index)return;if(e.RowIndex==-1)SetAllRowLocks(e.X<36);else if(!grid.Rows[e.RowIndex].IsNewRow)SetRowLocked(grid.Rows[e.RowIndex],!lockedRows.Contains(grid.Rows[e.RowIndex]));};
            grid.CellBeginEdit+=(s,e)=>{if(editLocked||lockedRows.Contains(grid.Rows[e.RowIndex]))e.Cancel=true;};
            grid.RowPostPaint+=(s,e)=>{if(!grid.Rows[e.RowIndex].IsNewRow)TextRenderer.DrawText(e.Graphics,(e.RowIndex+1).ToString(),grid.Font,new Rectangle(e.RowBounds.Left,e.RowBounds.Top,grid.RowHeadersWidth-5,e.RowBounds.Height),Color.DimGray,TextFormatFlags.Right|TextFormatFlags.VerticalCenter);};
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
                row.Cells[e.ColumnIndex].Value=e.ColumnIndex>=2?CorrectAlternative(value):value.TrimEnd(' ').ToUpperInvariant();
            };
            grid.SelectionChanged+=(s,e)=>{
                if(loading)return;
                selectionOrder.RemoveAll(r=>r.DataGridView!=grid||!r.Selected||r.IsNewRow);
                if(grid.CurrentRow!=null&&grid.CurrentRow.Selected&&!grid.CurrentRow.IsNewRow&&!selectionOrder.Contains(grid.CurrentRow))selectionOrder.Add(grid.CurrentRow);
                foreach(var row in grid.SelectedRows.Cast<DataGridViewRow>().Where(r=>!r.IsNewRow).OrderBy(r=>r.Index))if(!selectionOrder.Contains(row))selectionOrder.Add(row);
            };
            grid.CellValueChanged+=(s,e)=>{if(!loading&&e.RowIndex>=0){dirty=true;MarkChanged(grid.Rows[e.RowIndex]);message.Text="Unsaved changes are highlighted. Click Save names to apply them.";}};
            grid.UserAddedRow+=(s,e)=>{if(!loading)dirty=true;};
            grid.DataError+=(s,e)=>{e.ThrowException=false;message.Text="Please enter a single line of text in each name field.";};
            search.TextChanged+=(s,e)=>Filter();layout.Controls.Add(grid,0,2);
            buttons.Dock=DockStyle.Fill;buttons.WrapContents=false;buttons.Padding=new Padding(0,8,0,0);
            saveButton=Button("Save names",140);saveButton.BackColor=Color.FromArgb(25,87,157);saveButton.ForeColor=Color.White;saveButton.Click+=async(s,e)=>await SaveNames();
            addButton=Button("Add row",100);addButton.Click+=(s,e)=>{search.Clear();grid.CurrentCell=grid.Rows[grid.NewRowIndex].Cells[0];grid.BeginEdit(true);};
            deleteButton=Button("Delete selected",150);deleteButton.Click+=(s,e)=>{
                var selected=grid.SelectedRows.Cast<DataGridViewRow>().Where(r=>!r.IsNewRow&&!lockedRows.Contains(r)).ToArray();
                foreach(var row in selected)grid.Rows.Remove(row);
                if(selected.Length>0){dirty=true;message.Text="Rows removed from this edit. Save names to delete them; a later PDF scan can add them back with blank alternatives.";Filter();}
            };
            buttons.Controls.Add(new Label {Text="Excel",AutoSize=true,Margin=new Padding(12,6,6,0)});
            exportButton=Button("Export",85);exportButton.Name="exportNames";exportButton.Click+=async(s,e)=>await ExportToExcel();
            importButton=Button("Import",85);importButton.Name="importNames";importButton.Click+=async(s,e)=>await ImportFromExcel();
            var close=Button("Close",95);close.Click+=(s,e)=>Close();layout.Controls.Add(buttons,0,3);
            message.Dock=DockStyle.Fill;message.Font=new Font("Segoe UI",9F);message.Text="Names are locked. Click the open-padlock icon to edit or use Excel.";layout.Controls.Add(message,0,4);Controls.Add(layout);
            revision=Convert.ToInt32(catalog["revision"]);Fill(Rows(catalog["rows"]),true);RefreshExcelState();
            Disposed+=(s,e)=>{if(excel!=null){excel.Dispose();excel=null;}lockButton.Image.Dispose();unlockButton.Image.Dispose();rowClosed.Dispose();rowOpen.Dispose();lockTips.Dispose();sortFont.Dispose();};
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
        void Fill(List<Dictionary<string,object>> rows,bool opening=false) {
            loading=true;
            try {
                selectionOrder.Clear();savedRows.Clear();lockedRows.Clear();grid.Rows.Clear();
                foreach(var source in rows){
                    var row=new Dictionary<string,object>(source);
                    if(opening)row["locked"]=!String.IsNullOrWhiteSpace(Convert.ToString(row["alternative_product"]))&&!String.IsNullOrWhiteSpace(Convert.ToString(row["alternative_variation"]));
                    int i=grid.Rows.Add(fields.Select(f=>row[f]).Concat(new object[]{"View ("+(Aliases(row).Count+1)+")"}).ToArray());
                    grid.Rows[i].Tag=row;if(row.ContainsKey("locked")&&Convert.ToBoolean(row["locked"]))lockedRows.Add(grid.Rows[i]);
                    ApplyRowLock(grid.Rows[i]);savedRows[grid.Rows[i]]=RowSignature(grid.Rows[i]);
                }
                ApplySorting();grid.ClearSelection();
            }
            finally {loading=false;}
            Filter();
        }
        internal bool SortNames(string field) {
            if(working||excel!=null||!fields.Contains(field)||!grid.EndEdit())return false;
            if(field.EndsWith("variation")) {
                variationDescending=field==variationSortField?!variationDescending:false;variationSortField=field;
            } else {
                productDescending=field==productSortField?!productDescending:false;productSortField=field;
            }
            bool wasLoading=loading;loading=true;
            try {ApplySorting();}finally{loading=wasLoading;}
            return true;
        }
        void ApplySorting() {
            grid.Sort(new DefinitionComparer(productSortField,productDescending,variationSortField,variationDescending));
            foreach(string field in fields) {
                int index=Array.IndexOf(fields,field);bool variation=field.EndsWith("variation");
                bool active=field==(variation?variationSortField:productSortField);
                bool descending=active&&(variation?variationDescending:productDescending);
                grid.Columns[field].HeaderText=fieldTitles[index]+"\n"+SortCaption(field);
                string current=variation?(descending?"largest to smallest":"smallest to largest"):(descending?"Z-A":"A-Z");
                string next=variation?(descending?"smallest to largest":"largest to smallest"):(descending?"A-Z":"Z-A");
                grid.Columns[field].HeaderCell.ToolTipText=active?"Sorted "+current+" by "+fieldTitles[index].ToLowerInvariant()+". Click for "+next+". The other sort stays unchanged.":"Click to sort "+current+" by "+fieldTitles[index].ToLowerInvariant()+". The other sort stays unchanged.";
            }
            grid.Invalidate();
        }
        string SortCaption(string field) {
            bool variation=field.EndsWith("variation"),active=field==(field.EndsWith("variation")?variationSortField:productSortField);
            bool descending=active&&(variation?variationDescending:productDescending);
            if(variation)return active?(descending?"Largest to Smallest":"Smallest to Largest"):"Sort: Small to Large";
            return active?(descending?"Z-A  |  click A-Z":"A-Z  |  click Z-A"):"Sort: A-Z";
        }
        void PaintSortHeader(DataGridViewCellPaintingEventArgs e) {
            e.Paint(e.ClipBounds,DataGridViewPaintParts.Background|DataGridViewPaintParts.Border);
            string field=fields[e.ColumnIndex];bool variation=field.EndsWith("variation"),active=field==(variation?variationSortField:productSortField);
            bool descending=active&&(variation?variationDescending:productDescending);
            var title=new Rectangle(e.CellBounds.Left+5,e.CellBounds.Top+5,e.CellBounds.Width-10,25);
            var caption=new Rectangle(e.CellBounds.Left+5,e.CellBounds.Top+32,e.CellBounds.Width-30,22);
            TextRenderer.DrawText(e.Graphics,fieldTitles[e.ColumnIndex],e.CellStyle.Font,title,Color.Black,TextFormatFlags.Left|TextFormatFlags.VerticalCenter|TextFormatFlags.EndEllipsis);
            TextRenderer.DrawText(e.Graphics,SortCaption(field),sortFont,caption,Color.Black,TextFormatFlags.Left|TextFormatFlags.VerticalCenter|TextFormatFlags.EndEllipsis);
            int x=e.CellBounds.Right-14,y=e.CellBounds.Top+43;
            using(var pen=new Pen(Color.Black,2F)) {
                e.Graphics.DrawLine(pen,x,y-6,x,y+6);
                int tip=y+(descending?-6:6),tail=tip+(descending?4:-4);
                e.Graphics.DrawLine(pen,x,tip,x-4,tail);e.Graphics.DrawLine(pen,x,tip,x+4,tail);
                if(active)e.Graphics.DrawLine(pen,caption.Left,caption.Bottom-1,caption.Left+Math.Min(caption.Width-5,TextRenderer.MeasureText(SortCaption(field),sortFont).Width-5),caption.Bottom-1);
            }
            e.Handled=true;
        }
        sealed class DefinitionComparer : System.Collections.IComparer {
            readonly string product,variation;readonly int productDirection,variationDirection;
            internal DefinitionComparer(string product,bool productDescending,string variation,bool variationDescending){this.product=product;this.variation=variation;productDirection=productDescending?-1:1;variationDirection=variationDescending?-1:1;}
            public int Compare(object left,object right) {
                var a=(DataGridViewRow)left;var b=(DataGridViewRow)right;
                int result=StringComparer.OrdinalIgnoreCase.Compare(Convert.ToString(a.Cells[product].Value),Convert.ToString(b.Cells[product].Value));
                if(result!=0)return Math.Sign(result)*productDirection;
                result=CompareVariation(Convert.ToString(a.Cells[variation].Value),Convert.ToString(b.Cells[variation].Value));
                if(result!=0)return Math.Sign(result)*variationDirection;
                result=StringComparer.OrdinalIgnoreCase.Compare(Convert.ToString(a.Cells["original_product"].Value),Convert.ToString(b.Cells["original_product"].Value));
                if(result!=0)return result;
                return CompareVariation(Convert.ToString(a.Cells["original_variation"].Value),Convert.ToString(b.Cells["original_variation"].Value));
            }
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
                    entry[fields[i]]=i>=2?CorrectAlternative(values[i]):values[i].TrimEnd(' ').ToUpperInvariant();
                }
                entry["locked"]=lockedRows.Contains(row);row.Tag=entry;result.Add(entry);
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
        internal bool SetEditLocked(bool value) {
            if(working||excel!=null)return false;
            if(value&&dirty){message.Text="Save your name changes before locking.";return false;}
            editLocked=value;RefreshExcelState();
            message.Text=value?"Names are locked. Click the open-padlock icon to edit or use Excel.":"Editing unlocked. Unlock rows to edit them. Save names locks all rows again.";
            return true;
        }
        static Bitmap LockIcon(bool open) {
            var image=new Bitmap(24,24);
            using(var g=Graphics.FromImage(image))using(var pen=new Pen(Color.Black,2.5F))using(var body=new System.Drawing.Drawing2D.GraphicsPath()) {
                g.SmoothingMode=System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                pen.StartCap=pen.EndCap=System.Drawing.Drawing2D.LineCap.Round;
                pen.LineJoin=System.Drawing.Drawing2D.LineJoin.Round;
                // Matching outline padlocks: same body and baseline, no circle or keyhole.
                g.DrawArc(pen,7,2,10,10,180,180);g.DrawLine(pen,7,7,7,12);
                if(!open)g.DrawLine(pen,17,7,17,13);
                body.AddArc(4,12,3,3,180,90);body.AddArc(17,12,3,3,270,90);
                body.AddArc(17,19,3,3,0,90);body.AddArc(4,19,3,3,90,90);body.CloseFigure();
                g.DrawPath(pen,body);
            }
            return image;
        }
        void RefreshExcelState() {
            bool pending=excel!=null,canEdit=!working&&!editLocked;
            exportButton.Enabled=canEdit&&!pending;importButton.Enabled=canEdit&&pending;
            grid.Enabled=!working&&!pending;grid.ReadOnly=editLocked||pending;grid.AllowUserToAddRows=canEdit&&!pending;
            saveButton.Enabled=canEdit&&!pending;addButton.Enabled=canEdit&&!pending;deleteButton.Enabled=canEdit&&!pending;
            lockButton.Enabled=!working&&!pending&&!editLocked;unlockButton.Enabled=!working&&!pending&&editLocked;
            lockButton.BackColor=editLocked?Color.FromArgb(248,203,173):Color.White;
            unlockButton.BackColor=editLocked?Color.White:Color.FromArgb(220,234,252);
            foreach(DataGridViewRow row in grid.Rows)if(!row.IsNewRow)ApplyRowLock(row);
        }
        void ResetExcel(){if(excel!=null)excel.Dispose();excel=null;RefreshExcelState();}
        internal async Task<bool> ExportToExcel() {
            if(editLocked||excel!=null)return false;
            // Save before exporting without relocking the rows selected for Excel editing.
            if(!await SaveNames(false))return false;
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
            if(editLocked)return false;
            return await Run(async()=>{
                if(excel==null||!excel.IsOpen()||!System.IO.File.Exists(excel.FilePath)) {
                    ResetExcel();throw new Exception(NameExcelSession.MissingMessage);
                }
                excel.Save();
                var preview=await call(new Dictionary<string,object>{{"action","names_import"},{"path",excel.FilePath},{"rows",CollectRows()}});
                var saved=await SavePrepared(Rows(preview["rows"]),preview.ContainsKey("priority")?preview["priority"]:new string[0]);
                revision=Convert.ToInt32(saved["revision"]);Fill(Rows(saved["rows"]));dirty=false;
                ResetExcel();message.Text="Excel names imported and saved: "+preview["added"]+" added, "+preview["updated"]+" updated. All rows locked.";
                if(preview.ContainsKey("separated")&&Convert.ToInt32(preview["separated"])>0)message.Text+=" Different linked alternatives separated into "+preview["separated"]+" additional definitions.";
            });
        }
        void ApplyRowLock(DataGridViewRow row) {
            bool locked=lockedRows.Contains(row);
            foreach(string field in fields)row.Cells[field].ReadOnly=locked;
            row.Cells["rowLock"].Value=locked?rowClosed:rowOpen;
            row.Cells["rowLock"].ToolTipText=locked?"Unlock this definition":"Lock this definition";
        }
        internal bool SetRowLocked(DataGridViewRow row,bool value) {
            if(editLocked||working||excel!=null||row.IsNewRow||!grid.EndEdit())return false;
            string baseline;string current=RowSignature(row);
            if(value&&savedRows.TryGetValue(row,out baseline)&&baseline.EndsWith("|True")&&baseline.Substring(0,baseline.LastIndexOf('|'))!=current.Substring(0,current.LastIndexOf('|'))){message.Text="Save this row before locking it again.";return false;}
            if(value)lockedRows.Add(row);else lockedRows.Remove(row);
            ApplyRowLock(row);dirty=true;MarkChanged(row);
            message.Text="Row lock changed. Click Save names to keep this setting.";return true;
        }
        internal bool SetAllRowLocks(bool value) {
            if(editLocked||working||excel!=null||!grid.EndEdit())return false;
            bool complete=true;foreach(DataGridViewRow row in grid.Rows)if(!row.IsNewRow&&!SetRowLocked(row,value))complete=false;
            if(!complete){message.Text="Save changed rows before locking all definitions.";return false;}
            message.Text=value?"All rows locked. Save names to keep these settings.":"All rows unlocked. Save names to keep these settings.";return true;
        }

        internal static string CorrectAlternative(string value) {
            value=value.TrimEnd(' ').ToUpperInvariant();int start=0;
            while(start<value.Length&&!Char.IsLetterOrDigit(value[start]))start++;
            return value.Substring(start);
        }
        internal static int CompareVariation(string a,string b) {
            a=a??"";b=b??"";
            var pattern=@"(?<![\d.])(\d+(?:[.,]\d+)?)\s*(ML|KG|L|G)\b";
            var x=Regex.Match(a,pattern,RegexOptions.IgnoreCase);var y=Regex.Match(b,pattern,RegexOptions.IgnoreCase);
            if(x.Success&&y.Success){
                string xu=x.Groups[2].Value.ToUpperInvariant(),yu=y.Groups[2].Value.ToUpperInvariant();
                int family=String.CompareOrdinal(xu.EndsWith("L")?"V":"M",yu.EndsWith("L")?"V":"M");
                if(family!=0)return family;
                double xv,yv;
                if(!Double.TryParse(x.Groups[1].Value.Replace(',','.'),NumberStyles.Float,CultureInfo.InvariantCulture,out xv)||!Double.TryParse(y.Groups[1].Value.Replace(',','.'),NumberStyles.Float,CultureInfo.InvariantCulture,out yv))return StringComparer.OrdinalIgnoreCase.Compare(a,b);
                xv*=xu=="L"||xu=="KG"?1000:1;yv*=yu=="L"||yu=="KG"?1000:1;
                int size=xv.CompareTo(yv);if(size!=0)return size;
            } else if(x.Success!=y.Success)return x.Success?-1:1;
            return StringComparer.OrdinalIgnoreCase.Compare(a,b);
        }
        string RowSignature(DataGridViewRow row){return String.Join("\u001f",fields.Select(f=>Convert.ToString(row.Cells[f].Value)??""))+"|"+lockedRows.Contains(row);}
        void MarkChanged(DataGridViewRow row) {
            if(row.IsNewRow)return;
            string saved;bool changed=!savedRows.TryGetValue(row,out saved)||saved!=RowSignature(row);
            foreach(DataGridViewCell cell in row.Cells){cell.Style.BackColor=changed?Color.FromArgb(255,229,153):Color.Empty;cell.Style.SelectionBackColor=changed?Color.FromArgb(255,217,102):Color.Empty;}
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
        async Task<Dictionary<string,object>> SavePrepared(List<Dictionary<string,object>> pending,object priority,bool lockAllRows=true) {
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
            // Retain explicit unlocks for saved rows removed by deletion or a confirmed merge.
            var unlocked=savedRows.Keys.Where(r=>!lockedRows.Contains(r)&&r.Tag is Dictionary<string,object>).Select(r=>Convert.ToString(((Dictionary<string,object>)r.Tag)["id"])).ToArray();
            return await call(new Dictionary<string,object>{{"action","names_save"},{"rows",preview["rows"]},{"revision",revision},{"merge_groups",approvals},{"lock_all",lockAllRows},{"unlocked_ids",unlocked}});
        }
        internal async Task<bool> SaveNames(bool lockAllRows=true) {
            if(editLocked)return false;
            return await Run(async()=>{
                var pending=CollectRows();
                var priority=selectionOrder.Where(r=>r.Tag is Dictionary<string,object>).Select(r=>Convert.ToString(((Dictionary<string,object>)r.Tag)["id"])).ToArray();
                var result=await SavePrepared(pending,priority,lockAllRows);
                revision=Convert.ToInt32(result["revision"]);Fill(Rows(result["rows"]));dirty=false;
                message.Text=lockAllRows?"Names saved. All rows locked. Unlock a row to edit it again.":"Names saved. They will be used in the next generated documents.";
            });
        }
    }
}
