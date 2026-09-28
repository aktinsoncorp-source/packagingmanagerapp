using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using System.Windows.Forms;

[assembly: System.Reflection.AssemblyVersion("1.8.0.0")]
[assembly: System.Reflection.AssemblyFileVersion("1.8.0.0")]

namespace PackingListManager {
    static class Program {
        [STAThread]
        static void Main(string[] args) {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            bool created;
            using (var mutex = new Mutex(true, "Local\\PackingListManagerSingleComputer", out created)) {
                if (!created) {
                    MessageBox.Show("Packing List Manager is already open. Use the existing window.", "Already open");
                    return;
                }
                Application.ThreadException += (s,e) => MessageBox.Show(e.Exception.Message, "Packing List Manager", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Application.Run(new MainForm(args));
            }
        }
    }

    public class MainForm : Form {
        readonly JavaScriptSerializer json = new JavaScriptSerializer { MaxJsonLength = 100000000 };
        readonly string appDir = AppDomain.CurrentDomain.BaseDirectory;
        string rootFolder;
        readonly string settingsPath;
        bool busy, loadingFiles;
        string lastError;
        readonly TextBox folderText = new TextBox();
        readonly DateTimePicker date = new DateTimePicker();
        readonly DataGridView grid = new DataGridView();
        readonly CheckBox selectAllUse = new CheckBox();
        bool selectingAll;
        readonly Label summary = new Label();
        readonly Label selectionTotals = new Label();
        readonly Label status = new Label();
        readonly ProgressBar progress = new ProgressBar();
        readonly Button chooseButton = new Button();
        readonly Button scanButton = new Button();
        readonly Button addButton = new Button();
        readonly Button generateButton = new Button();
        readonly Button outputButton = new Button();
        readonly Button reviewButton = new Button();
        readonly Button sourceButton = new Button();
        readonly Button historyButton = new Button();
        readonly Button helpButton = new Button();
        readonly Button namesButton = new Button();
        readonly HashSet<string> revisions = new HashSet<string>();
        List<Dictionary<string,object>> entries = new List<Dictionary<string,object>>();
        List<Dictionary<string,object>> outputs = new List<Dictionary<string,object>>();
        readonly string[] startupArgs;

        public MainForm(string[] args) {
            startupArgs = args;
            settingsPath = Path.Combine(appDir,"Settings.json");
            Text = "Packing List Manager 1.8";
            Icon = System.Drawing.Icon.ExtractAssociatedIcon(Application.ExecutablePath);
            Font = new Font("Segoe UI", 10F);
            BackColor = Color.FromArgb(246,248,251);
            ForeColor = Color.FromArgb(31,42,55);
            MinimumSize = new Size(980,730);
            Size = new Size(1160,860);
            StartPosition = FormStartPosition.CenterScreen;
            AutoScaleMode = AutoScaleMode.Dpi;
            BuildUI();
            rootFolder = Directory.GetParent(appDir.TrimEnd(Path.DirectorySeparatorChar)).FullName;
            if (File.Exists(settingsPath)) {
                try {
                    var saved = json.Deserialize<Dictionary<string,object>>(File.ReadAllText(settingsPath));
                    if (Directory.Exists(S(saved,"folder"))) rootFolder = S(saved,"folder");
                } catch { }
            }
            folderText.Text = rootFolder;
            Shown += async (s,e) => {
                if (startupArgs.Length >= 2 && startupArgs[0] == "--self-test") {
                    try {
                        SelectFolder(startupArgs[1]);
                        await Scan();
                        bool courierControlsValid=CheckCourierControls();
                        bool courierChangesWork=!startupArgs.Contains("--check-couriers")||await CheckCourierChanges();
                        if(!courierControlsValid||!courierChangesWork)throw new Exception("Courier controls verification failed: controls="+courierControlsValid+", changes="+courierChangesWork+", rows="+grid.Rows.Count+", error="+lastError+", entries="+json.Serialize(entries));
                        int selectable=grid.Rows.Cast<DataGridViewRow>().Count(r=>!r.Cells[0].ReadOnly);
                        int autoSelected=grid.Rows.Cast<DataGridViewRow>().Count(r=>Convert.ToBoolean(r.Cells[0].Value));
                        bool manualSelectionWorks=false;
                        var probe=grid.Rows.Cast<DataGridViewRow>().FirstOrDefault(r=>!r.Cells[0].ReadOnly);
                        if(probe!=null){var previous=probe.Cells[0].Value;probe.Cells[0].Value=true;RefreshSelection();manualSelectionWorks=generateButton.Enabled;probe.Cells[0].Value=previous;RefreshSelection();}
                        bool selectAllWorks=CheckSelectAllBehavior();
                        bool selectionTotalsWork=CheckSelectionTotals();
                        if(!selectionTotalsWork)throw new Exception("Selection totals verification failed.");
                        if(!selectAllWorks)throw new Exception("Select-all verification failed.");
                        if(startupArgs.Contains("--select-for-preview"))ToggleAllUse();
                        if (startupArgs.Length >= 3) {
                            using (var bitmap = new Bitmap(Width,Height)) {
                                DrawToBitmap(bitmap,new Rectangle(Point.Empty,Size));
                                bitmap.Save(startupArgs[2]);
                            }
                        }
                        if(startupArgs.Length>=4) {
                            using(var preview=CreateHistoryForm()) {
                                var tree=(TreeView)preview.Controls.Find("completedTree",true)[0];
                                if(tree.Nodes.Count>0&&tree.Nodes[0].Nodes.Count>0) {
                                    var document=tree.Nodes[0].Nodes[0];document.Expand();
                                    if(document.Nodes.Count>0)document.Nodes[0].Expand();
                                }
                                preview.StartPosition=FormStartPosition.Manual;preview.Location=new Point(-30000,-30000);
                                preview.Show();preview.PerformLayout();Application.DoEvents();
                                using(var bitmap=new Bitmap(preview.Width,preview.Height)) {
                                    preview.DrawToBitmap(bitmap,new Rectangle(Point.Empty,preview.Size));bitmap.Save(startupArgs[3]);
                                }
                                File.WriteAllText(Path.Combine(appDir,"HistorySelfTest.json"),json.Serialize(new {
                                    groups=tree.Nodes.Cast<TreeNode>().Select(n=>new {name=n.Text,documents=n.Nodes.Cast<TreeNode>().Select(d=>new {
                                        filename=S(((CompletedSelection)d.Tag).Document,"filename"),sequence=N(((CompletedSelection)d.Tag).Document,"sequence"),
                                        sources=d.Nodes.Cast<TreeNode>().Where(p=>(p.Tag as CompletedSelection)!=null&&!String.IsNullOrEmpty(((CompletedSelection)p.Tag).Pdf)).Select(p=>new {path=((CompletedSelection)p.Tag).Pdf,orders=p.Nodes.Count}).ToArray()
                                    }).ToArray()}).ToArray()
                                }));
                            }
                        }
                        if(startupArgs.Length>=5) {
                            var catalog=await Call(new Dictionary<string,object>{{"action","names_list"}},"Checking Define Name...");
                            using(var preview=new NameDefinitionsForm(catalog,request=>Call(request,"Checking names..."))) {
                                preview.StartPosition=FormStartPosition.Manual;preview.Location=new Point(-30000,-30000);
                                preview.Show();preview.PerformLayout();Application.DoEvents();
                                using(var bitmap=new Bitmap(preview.Width,preview.Height)) {
                                    preview.DrawToBitmap(bitmap,new Rectangle(Point.Empty,preview.Size));bitmap.Save(startupArgs[4]);
                                }
                            }
                        }
                        File.WriteAllText(Path.Combine(appDir,"SelfTest.json"),json.Serialize(new { ok=lastError==null, rows=grid.Rows.Count, selectable=selectable, autoSelected=autoSelected, manualSelectionWorks=manualSelectionWorks, selectAllWorks=selectAllWorks, selectionTotalsWork=selectionTotalsWork, courierControlsValid=courierControlsValid, courierChangesWork=courierChangesWork, summary=summary.Text, status=status.Text, error=lastError }));
                    } catch(Exception ex) {
                        File.WriteAllText(Path.Combine(appDir,"SelfTest.json"),json.Serialize(new {ok=false,error=ex.ToString()}));
                    }
                    Close();
                } else {
                    await Scan();
                }
            };
            FormClosing += (s,e) => {
                if (busy) {
                    e.Cancel = true;
                    MessageBox.Show(this,"Please wait for the current scan or batch to finish.","Working",MessageBoxButtons.OK,MessageBoxIcon.Information);
                }
            };
        }

        static string S(Dictionary<string,object> d,string key) { return d.ContainsKey(key) && d[key]!=null ? Convert.ToString(d[key],CultureInfo.InvariantCulture) : ""; }
        static int N(Dictionary<string,object> d,string key) { return d.ContainsKey(key) && d[key]!=null ? Convert.ToInt32(d[key]) : 0; }
        static bool B(Dictionary<string,object> d,string key) { return d.ContainsKey(key) && d[key]!=null && Convert.ToBoolean(d[key]); }
        bool CanSelectFile(Dictionary<string,object> en) { return en!=null && S(en,"status")!="ERROR" && S(en,"status")!="COURIER" && (!B(en,"requires_revision") || revisions.Contains(S(en,"path"))); }
        static List<Dictionary<string,object>> Rows(object value) {
            var result=new List<Dictionary<string,object>>();
            if (value is IEnumerable) foreach(var v in (IEnumerable)value) if(v is Dictionary<string,object>) result.Add((Dictionary<string,object>)v);
            return result;
        }
        static IEnumerable<object> Values(object value) { return value is IEnumerable ? ((IEnumerable)value).Cast<object>() : Enumerable.Empty<object>(); }
        string DateFolder { get { return Path.Combine(rootFolder,date.Value.ToString("yyyyMMdd")); } }
        string OutputFolder { get { return Path.Combine(DateFolder,"000 Output"); } }

        void BuildUI() {
            var layout=new TableLayoutPanel { Dock=DockStyle.Fill,Padding=new Padding(24),ColumnCount=1,RowCount=8 };
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute,66));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute,64));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute,56));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute,38));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent,100));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute,42));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute,62));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute,30));
            Controls.Add(layout);

            var header=new TableLayoutPanel {Dock=DockStyle.Fill,ColumnCount=2};
            header.ColumnStyles.Add(new ColumnStyle(SizeType.Percent,100));
            header.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute,100));
            var title=new Label {Text="Packing List Manager",Font=new Font("Segoe UI",23F,FontStyle.Bold),Dock=DockStyle.Fill,AutoSize=false};
            header.Controls.Add(title,0,0);
            StyleButton(helpButton,"Help",false); helpButton.Dock=DockStyle.Top;
            helpButton.Click+=(s,e)=>OpenFile(Path.Combine(appDir,"Employee Guide.txt"));
            header.Controls.Add(helpButton,1,0);
            layout.Controls.Add(header,0,0);

            var folders=new TableLayoutPanel {Dock=DockStyle.Fill,ColumnCount=2,RowCount=2};
            folders.ColumnStyles.Add(new ColumnStyle(SizeType.Percent,100)); folders.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute,150));
            folders.RowStyles.Add(new RowStyle(SizeType.Absolute,24));folders.RowStyles.Add(new RowStyle(SizeType.Absolute,36));
            folders.Controls.Add(new Label {Text="Source Folder",Dock=DockStyle.Fill,Font=new Font(Font,FontStyle.Bold)},0,0);
            folderText.ReadOnly=true;folderText.Dock=DockStyle.Fill;folderText.BackColor=Color.White;folderText.Margin=new Padding(0,2,10,2);
            folders.Controls.Add(folderText,0,1);
            StyleButton(chooseButton,"Choose folder...",false);chooseButton.Dock=DockStyle.Fill;
            chooseButton.Click+=async(s,e)=>{
                using(var dlg=new FolderBrowserDialog {Description="Choose your main packing folder or a dated YYYYMMDD folder.",SelectedPath=rootFolder,ShowNewFolderButton=true}) {
                    if(dlg.ShowDialog(this)==DialogResult.OK){SelectFolder(dlg.SelectedPath);SaveSettings();await Scan();}
                }
            };
            folders.Controls.Add(chooseButton,1,1);layout.Controls.Add(folders,0,1);

            var controls=new FlowLayoutPanel {Dock=DockStyle.Fill,WrapContents=false,Padding=new Padding(0,8,0,0)};
            controls.Controls.Add(new Label {Text="Date",AutoSize=true,Margin=new Padding(0,7,8,0)});
            date.Format=DateTimePickerFormat.Custom;date.CustomFormat="yyyy-MM-dd";date.Width=140;date.Margin=new Padding(0,3,18,0);
            date.ValueChanged+=async(s,e)=>{if(IsHandleCreated&&!busy)await Scan();};controls.Controls.Add(date);
            StyleButton(scanButton,"Scan for new PDFs",false);scanButton.Width=168;scanButton.Click+=async(s,e)=>await Scan();controls.Controls.Add(scanButton);
            StyleButton(addButton,"Add PDFs...",false);addButton.Width=124;addButton.Click+=async(s,e)=>await AddPdfs();controls.Controls.Add(addButton);
            layout.Controls.Add(controls,0,2);

            summary.Text="Review files";summary.Dock=DockStyle.Fill;summary.Font=new Font(Font,FontStyle.Bold);summary.TextAlign=ContentAlignment.MiddleLeft;
            layout.Controls.Add(summary,0,3);
            grid.Dock=DockStyle.Fill;grid.BackgroundColor=Color.White;grid.BorderStyle=BorderStyle.FixedSingle;
            grid.AllowUserToAddRows=false;grid.AllowUserToDeleteRows=false;grid.RowHeadersVisible=false;grid.MultiSelect=false;
            grid.SelectionMode=DataGridViewSelectionMode.FullRowSelect;grid.AutoSizeColumnsMode=DataGridViewAutoSizeColumnsMode.None;
            grid.EnableHeadersVisualStyles=false;grid.ColumnHeadersDefaultCellStyle.BackColor=Color.FromArgb(232,238,247);grid.ColumnHeadersDefaultCellStyle.ForeColor=ForeColor;grid.ColumnHeadersDefaultCellStyle.Font=new Font(Font,FontStyle.Bold);
            grid.ColumnHeadersHeight=38;grid.RowTemplate.Height=34;
            grid.DefaultCellStyle.SelectionBackColor=Color.FromArgb(220,234,252);grid.DefaultCellStyle.SelectionForeColor=ForeColor;
            grid.AlternatingRowsDefaultCellStyle.BackColor=Color.FromArgb(248,250,252);
            grid.Columns.Add(new DataGridViewCheckBoxColumn {Name="include",HeaderText="",FillWeight=40,MinimumWidth=70});
            grid.Columns.Add(new DataGridViewTextBoxColumn {Name="platform",HeaderText="Platform",FillWeight=65,ReadOnly=true});
            grid.Columns.Add(new DataGridViewTextBoxColumn {Name="file",HeaderText="PDF file",FillWeight=220,ReadOnly=true});
            grid.Columns.Add(new DataGridViewTextBoxColumn {Name="orders",HeaderText="Orders",FillWeight=48,ReadOnly=true});
            grid.Columns.Add(new DataGridViewTextBoxColumn {Name="pieces",HeaderText="Pieces",FillWeight=48,ReadOnly=true});
            grid.Columns.Add(new DataGridViewTextBoxColumn {Name="status",HeaderText="Status",FillWeight=180,ReadOnly=true});
            var courierColumn=new DataGridViewComboBoxColumn {Name="courier",HeaderText="Courier",FillWeight=76,MinimumWidth=88,FlatStyle=FlatStyle.Flat,DisplayStyle=DataGridViewComboBoxDisplayStyle.DropDownButton};
            courierColumn.Items.AddRange("Auto","SPX","FLASH");grid.Columns.Add(courierColumn);courierColumn.DisplayIndex=3;
            grid.Columns[0].AutoSizeMode=DataGridViewAutoSizeColumnMode.None;grid.Columns[0].Width=70;
            grid.Columns["platform"].Width=95;grid.Columns["courier"].Width=100;
            grid.Columns["orders"].Width=65;grid.Columns["pieces"].Width=65;
            grid.Columns["file"].AutoSizeMode=DataGridViewAutoSizeColumnMode.Fill;
            grid.Columns["status"].AutoSizeMode=DataGridViewAutoSizeColumnMode.Fill;
            grid.Columns["file"].FillWeight=55;grid.Columns["status"].FillWeight=45;
            selectAllUse.Text="Use";selectAllUse.AutoSize=true;selectAllUse.AutoCheck=false;selectAllUse.ThreeState=true;
            selectAllUse.Font=new Font("Segoe UI",9F,FontStyle.Bold);selectAllUse.BackColor=grid.ColumnHeadersDefaultCellStyle.BackColor;
            selectAllUse.AccessibleName="Select all compatible PDFs";selectAllUse.Cursor=Cursors.Hand;
            selectAllUse.Click+=(s,e)=>ToggleAllUse();grid.Controls.Add(selectAllUse);
            grid.ColumnWidthChanged+=(s,e)=>PositionUseHeader();grid.SizeChanged+=(s,e)=>PositionUseHeader();
            grid.Scroll+=(s,e)=>PositionUseHeader();grid.ColumnHeadersHeightChanged+=(s,e)=>PositionUseHeader();
            grid.SelectionChanged+=(s,e)=>ShowSelected();
            grid.CurrentCellDirtyStateChanged+=(s,e)=>{if(grid.IsCurrentCellDirty)grid.CommitEdit(DataGridViewDataErrorContexts.Commit);};
            grid.CellValueChanged+=async(s,e)=>{
                if(loadingFiles)return;
                if(e.RowIndex>=0&&e.ColumnIndex==grid.Columns["courier"].Index)await ChangeCourier(grid.Rows[e.RowIndex]);
                else RefreshSelection();
            };
            grid.DataError+=(s,e)=>{e.ThrowException=false;};
            grid.CellBeginEdit+=(s,e)=>{if(e.ColumnIndex==0){var en=(Dictionary<string,object>)grid.Rows[e.RowIndex].Tag;if(!CanSelectFile(en))e.Cancel=true;}else if(e.ColumnIndex==grid.Columns["courier"].Index&&!B((Dictionary<string,object>)grid.Rows[e.RowIndex].Tag,"is_shopee"))e.Cancel=true;};
            var fileArea=new TableLayoutPanel {Dock=DockStyle.Fill,ColumnCount=1,RowCount=2,Margin=Padding.Empty};
            fileArea.RowStyles.Add(new RowStyle(SizeType.Percent,100));fileArea.RowStyles.Add(new RowStyle(SizeType.Absolute,76));
            fileArea.Controls.Add(grid,0,0);
            selectionTotals.Name="selectionTotals";selectionTotals.Dock=DockStyle.Fill;selectionTotals.Padding=new Padding(5,8,5,0);
            selectionTotals.Font=new Font("Segoe UI",9F);selectionTotals.BackColor=Color.FromArgb(232,238,247);
            fileArea.Controls.Add(selectionTotals,0,1);layout.Controls.Add(fileArea,0,4);

            var secondary=new FlowLayoutPanel {Dock=DockStyle.Fill,WrapContents=false,Padding=new Padding(0,5,0,0)};
            StyleButton(sourceButton,"Open selected PDF",false);sourceButton.Width=180;sourceButton.Click+=(s,e)=>{var en=Selected();if(en!=null)OpenFile(Path.Combine(DateFolder,S(en,"path").Replace('/',Path.DirectorySeparatorChar)));};secondary.Controls.Add(sourceButton);
            StyleButton(reviewButton,"Review changed file...",false);reviewButton.Width=198;reviewButton.Click+=async(s,e)=>await Review();secondary.Controls.Add(reviewButton);
            StyleButton(historyButton,"Completed documents...",false);historyButton.Width=210;historyButton.Click+=(s,e)=>ShowHistory();secondary.Controls.Add(historyButton);
            StyleButton(namesButton,"Define Name...",false);namesButton.Width=150;namesButton.Click+=async(s,e)=>await ShowNames();secondary.Controls.Add(namesButton);
            layout.Controls.Add(secondary,0,5);

            var final=new TableLayoutPanel {Dock=DockStyle.Fill,ColumnCount=3,Padding=new Padding(0,8,0,5)};
            final.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute,260));final.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute,190));final.ColumnStyles.Add(new ColumnStyle(SizeType.Percent,100));
            StyleButton(generateButton,"Generate Packing Lists",true);generateButton.Dock=DockStyle.Fill;generateButton.Click+=async(s,e)=>await Generate();final.Controls.Add(generateButton,0,0);
            StyleButton(outputButton,"Open Output Folder",false);outputButton.Dock=DockStyle.Fill;outputButton.Click+=(s,e)=>{if(Directory.Exists(OutputFolder))OpenFile(OutputFolder);else MessageBox.Show(this,"The Output folder will appear after your first batch.","Output folder");};final.Controls.Add(outputButton,1,0);
            var hint=new Label {Text="A4  |  One table per order\nSPX / FLASH orange  ·  Lazada + TikTok blue",Dock=DockStyle.Fill,TextAlign=ContentAlignment.MiddleRight,Font=new Font("Segoe UI",9F),ForeColor=Color.FromArgb(82,96,112)};
            final.Controls.Add(hint,2,0);layout.Controls.Add(final,0,6);
            var footer=new TableLayoutPanel {Dock=DockStyle.Fill,ColumnCount=2};footer.ColumnStyles.Add(new ColumnStyle(SizeType.Percent,100));footer.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute,150));
            status.Text="Ready";status.Dock=DockStyle.Fill;status.Font=new Font("Segoe UI",9F);status.TextAlign=ContentAlignment.MiddleLeft;footer.Controls.Add(status,0,0);
            progress.Dock=DockStyle.Fill;progress.Style=ProgressBarStyle.Marquee;progress.Visible=false;progress.MarqueeAnimationSpeed=24;footer.Controls.Add(progress,1,0);layout.Controls.Add(footer,0,7);
        }

        bool CheckCourierControls() {
            return grid.Rows.Cast<DataGridViewRow>().All(row=>{
                var entry=row.Tag as Dictionary<string,object>;var cell=row.Cells["courier"];
                return B(entry,"is_shopee")?cell is DataGridViewComboBoxCell&&!cell.ReadOnly&&Convert.ToString(cell.Value)==S(entry,"courier"):
                    cell is DataGridViewTextBoxCell&&cell.ReadOnly;
            });
        }
        async Task<bool> CheckCourierChanges() {
            var source=entries.FirstOrDefault(e=>S(e,"status")=="COURIER");
            if(source==null)return false;
            string path=S(source,"path");
            foreach(string mode in new[]{"FLASH","SPX","Auto"}) {
                var row=grid.Rows.Cast<DataGridViewRow>().First(r=>S((Dictionary<string,object>)r.Tag,"path")==path);
                loadingFiles=true;row.Cells["courier"].Value=mode;loadingFiles=false;
                await ChangeCourier(row);
                var changed=entries.First(e=>S(e,"path")==path);
                if(S(changed,"courier")!=mode||S(changed,"platform")!=(mode=="Auto"?"SHOPEE":mode))return false;
                await Scan(true);
                if(S(entries.First(e=>S(e,"path")==path),"courier")!=mode||!CheckCourierControls())return false;
            }
            return true;
        }
        void PositionUseHeader() {
            if(grid.Columns.Count==0)return;
            var bounds=grid.GetCellDisplayRectangle(0,-1,true);
            selectAllUse.Location=new Point(bounds.Left+6,bounds.Top+Math.Max(0,(bounds.Height-selectAllUse.Height)/2));
            selectAllUse.Visible=bounds.Width>=selectAllUse.Width+6;selectAllUse.BringToFront();
        }
        void ToggleAllUse() {
            if(busy)return;
            bool use=selectAllUse.CheckState!=CheckState.Checked;
            selectingAll=true;
            try {
                grid.EndEdit();
                foreach(DataGridViewRow row in grid.Rows)
                    row.Cells[0].Value=use&&CanSelectFile(row.Tag as Dictionary<string,object>);
            } finally {selectingAll=false;}
            RefreshSelection();
        }
        void RefreshUseHeader() {
            var eligible=grid.Rows.Cast<DataGridViewRow>().Where(r=>CanSelectFile(r.Tag as Dictionary<string,object>)).ToList();
            int selected=eligible.Count(r=>Convert.ToBoolean(r.Cells[0].Value));
            selectAllUse.Enabled=!busy&&eligible.Count>0;
            selectAllUse.CheckState=selected==0?CheckState.Unchecked:selected==eligible.Count?CheckState.Checked:CheckState.Indeterminate;
            PositionUseHeader();
        }
        bool CheckSelectAllBehavior() {
            var original=grid.Rows.Cast<DataGridViewRow>().Select(r=>r.Cells[0].Value).ToArray();
            var added=new List<DataGridViewRow>();
            try {
                selectingAll=true;
                foreach(string state in new[]{"ERROR","REVIEW"}) {
                    int i=grid.Rows.Add(false,"","QA excluded PDF",0,0,state,null);
                    var row=grid.Rows[i];added.Add(row);
                    row.Tag=new Dictionary<string,object>{{"status",state},{"requires_revision",state=="REVIEW"},{"path","QA "+state}};
                    row.Cells[0].ReadOnly=true;
                }
                selectingAll=false;RefreshSelection();
                var eligible=grid.Rows.Cast<DataGridViewRow>().Where(r=>CanSelectFile(r.Tag as Dictionary<string,object>)).ToList();
                selectAllUse.CheckState=CheckState.Unchecked;ToggleAllUse();
                bool ok=eligible.All(r=>Convert.ToBoolean(r.Cells[0].Value))&&added.All(r=>!Convert.ToBoolean(r.Cells[0].Value));
                if(eligible.Count>0) {
                    ok=ok&&selectAllUse.CheckState==CheckState.Checked;
                    eligible[0].Cells[0].Value=false;RefreshSelection();
                    ok=ok&&selectAllUse.CheckState==(eligible.Count>1?CheckState.Indeterminate:CheckState.Unchecked);
                    ToggleAllUse();
                }
                ToggleAllUse();
                return ok&&grid.Rows.Cast<DataGridViewRow>().All(r=>!Convert.ToBoolean(r.Cells[0].Value));
            } finally {
                selectingAll=true;
                foreach(var row in added)grid.Rows.Remove(row);
                for(int i=0;i<original.Length;i++)grid.Rows[i].Cells[0].Value=original[i];
                selectingAll=false;RefreshSelection();
            }
        }

        void StyleButton(Button button,string text,bool primary) {
            button.Text=text;button.Height=34;button.FlatStyle=FlatStyle.Flat;button.FlatAppearance.BorderSize=1;
            button.FlatAppearance.BorderColor=primary?Color.FromArgb(25,87,157):Color.FromArgb(198,208,222);
            button.BackColor=primary?Color.FromArgb(25,87,157):Color.White;button.ForeColor=primary?Color.White:ForeColor;
            button.Font=new Font("Segoe UI",10F,primary?FontStyle.Bold:FontStyle.Regular);button.Cursor=Cursors.Hand;button.Margin=new Padding(0,0,10,0);
        }
        void SelectFolder(string path) {
            DateTime parsed;
            var di=new DirectoryInfo(path);
            bool isDate=DateTime.TryParseExact(di.Name,"yyyyMMdd",CultureInfo.InvariantCulture,DateTimeStyles.None,out parsed);
            rootFolder=isDate?di.Parent.FullName:di.FullName;
            folderText.Text=rootFolder;
            if(isDate){bool old=busy;busy=true;date.Value=parsed;busy=old;}
        }
        void SaveSettings() {
            if(String.IsNullOrEmpty(rootFolder))return;
            try {File.WriteAllText(settingsPath,json.Serialize(new {folder=rootFolder}));}catch{}
        }
        void Busy(bool value,string message) {
            busy=value;status.Text=message;progress.Visible=value;
            selectAllUse.Enabled=!value;
            chooseButton.Enabled=!value;scanButton.Enabled=!value;addButton.Enabled=!value;date.Enabled=!value;
            grid.Enabled=!value;generateButton.Enabled=!value;reviewButton.Enabled=!value;sourceButton.Enabled=!value;historyButton.Enabled=!value;
            namesButton.Enabled=!value;
            if(!value)RefreshSelection();
        }
        async Task<Dictionary<string,object>> Call(Dictionary<string,object> request,string message) {
            Busy(true,message);
            string work=Path.Combine(Path.GetTempPath(),"PackingList-"+Guid.NewGuid().ToString("N"));Directory.CreateDirectory(work);
            try {
                request["folder"]=DateFolder;
                if(!request.ContainsKey("catalog_path"))request["catalog_path"]=Path.Combine(appDir,"ProductNames.json");
                string input=Path.Combine(work,"request.json"),output=Path.Combine(work,"response.json");
                File.WriteAllText(input,json.Serialize(request),new UTF8Encoding(false));
                string engine=Path.Combine(appDir,"App Files","PackingEngine.exe");
                if(!File.Exists(engine))throw new Exception("App Files\\PackingEngine.exe is missing. Keep the application folder together.");
                var psi=new ProcessStartInfo(engine,"--request \""+input+"\" --response \""+output+"\"") {UseShellExecute=false,CreateNoWindow=true,WorkingDirectory=appDir,RedirectStandardError=true};
                string error=await Task.Run(()=>{
                    using(var process=Process.Start(psi)) {string err=process.StandardError.ReadToEnd();process.WaitForExit();return err;}
                });
                if(!File.Exists(output))throw new Exception("The processing engine stopped unexpectedly. "+error);
                var result=json.Deserialize<Dictionary<string,object>>(File.ReadAllText(output,Encoding.UTF8));
                if(!result.ContainsKey("ok")||!(bool)result["ok"])throw new Exception(S(result,"error"));
                return result;
            } finally {Busy(false,"Ready");try{Directory.Delete(work,true);}catch{}}
        }

        public async Task Scan(bool preserveSelections=false) {
            if(busy)return;
            string selectedPath=Selected()==null?"":S(Selected(),"path");
            var previous=new Dictionary<string,bool>();
            if(preserveSelections)foreach(DataGridViewRow row in grid.Rows)if(CanSelectFile(row.Tag as Dictionary<string,object>))previous[S((Dictionary<string,object>)row.Tag,"path")]=Convert.ToBoolean(row.Cells[0].Value);
            try {
                var result=await Call(new Dictionary<string,object>{{"action","scan"}},"Reading PDFs and checking completed batches...");
                entries=Rows(result["entries"]);outputs=Rows(result["outputs"]);revisions.Clear();loadingFiles=true;grid.Rows.Clear();
                foreach(var entry in entries) {
                    string state=S(entry,"status"),label=S(entry,"message");
                    int i=grid.Rows.Add(state=="READY"||state=="MISSING",S(entry,"platform"),S(entry,"path"),N(entry,"orders"),N(entry,"pieces"),label,B(entry,"is_shopee")?S(entry,"courier"):null);
                    var row=grid.Rows[i];row.Tag=entry;row.Cells[0].ReadOnly=!CanSelectFile(entry);
                    if(B(entry,"is_shopee"))row.Cells["courier"].ToolTipText="Auto uses the nearest SPX/FLASH folder first, then the filename. Choose SPX or FLASH to override.";
                    else {row.Cells["courier"]=new DataGridViewTextBoxCell {Value="—",Style=new DataGridViewCellStyle {ForeColor=Color.Gray,BackColor=Color.FromArgb(242,244,247)}};row.Cells["courier"].ReadOnly=true;}
                    if(previous.ContainsKey(S(entry,"path"))&&CanSelectFile(entry))row.Cells[0].Value=previous[S(entry,"path")];
                    row.Cells[5].ToolTipText=label;
                    row.Cells[5].Style.ForeColor=state=="READY"||state=="MISSING"?Color.FromArgb(20,112,67):(state=="REVIEW"||state=="ERROR"||state=="COURIER"?Color.FromArgb(166,77,10):Color.FromArgb(95,106,119));
                }
                loadingFiles=false;
                if(preserveSelections)foreach(DataGridViewRow row in grid.Rows)if(S((Dictionary<string,object>)row.Tag,"path")==selectedPath){grid.CurrentCell=row.Cells["file"];break;}
                int fresh=entries.Sum(e=>N(e,"new_orders")),issues=entries.Count(e=>S(e,"status")=="REVIEW"||S(e,"status")=="ERROR"||S(e,"status")=="COURIER");
                int missing=entries.Sum(e=>N(e,"missing_orders"));
                summary.Text="Review files — "+fresh+" new orders"+(missing>0?"  |  "+missing+" orders missing output documents":"")+(issues>0?"  |  "+issues+" file(s) need attention":"");
                status.Text=entries.Count==0?"No readable, supported packing PDFs for this date. Choose a dated folder or use Add PDFs.":"Tick Use for any PDF to generate again. New files and files missing output are selected automatically.";
                ShowSelected();RefreshSelection();
            } catch(Exception ex){loadingFiles=false;ShowError(ex);grid.Rows.Clear();entries.Clear();outputs.Clear();RefreshSelection();}
        }
        async Task ChangeCourier(DataGridViewRow row) {
            var entry=row.Tag as Dictionary<string,object>;
            if(busy||entry==null||!B(entry,"is_shopee"))return;
            string mode=Convert.ToString(row.Cells["courier"].Value);
            if(mode==S(entry,"courier"))return;
            try {
                await Call(new Dictionary<string,object>{{"action","courier"},{"carrier",mode},{"selected",new Dictionary<string,object>{{"path",S(entry,"path")},{"sha256",S(entry,"sha256")}}}},"Saving courier selection...");
                await Scan(true);
            } catch(Exception ex){loadingFiles=true;row.Cells["courier"].Value=S(entry,"courier");loadingFiles=false;ShowError(ex);}
        }
        Dictionary<string,object> Selected() {return grid.CurrentRow==null?null:grid.CurrentRow.Tag as Dictionary<string,object>;}
        void ShowSelected() {
            var en=Selected();
            reviewButton.Enabled=!busy&&en!=null&&B(en,"requires_revision");
            sourceButton.Enabled=!busy&&en!=null;
        }
        void RefreshSelection() {
            if(busy||selectingAll)return;
            generateButton.Enabled=grid.Rows.Cast<DataGridViewRow>().Any(r=>Convert.ToBoolean(r.Cells[0].Value));
            generateButton.BackColor=generateButton.Enabled?Color.FromArgb(25,87,157):Color.FromArgb(223,231,241);
            ShowSelected();historyButton.Enabled=outputs.Count>0;
            RefreshUseHeader();UpdateSelectionTotals();
        }
        internal Dictionary<string,int[]> SelectionCounts() {
            var counts=new Dictionary<string,int[]>();var seen=new HashSet<string>();
            var rows=grid.Rows.Cast<DataGridViewRow>().Where(r=>Convert.ToBoolean(r.Cells[0].Value)&&r.Tag is Dictionary<string,object>)
                .OrderBy(r=>!revisions.Contains(S((Dictionary<string,object>)r.Tag,"path")));
            foreach(var row in rows) {
                var entry=(Dictionary<string,object>)row.Tag;string platform=S(entry,"platform");
                if(!counts.ContainsKey(platform))counts[platform]=new int[3];counts[platform][0]++;
                if(entry.ContainsKey("order_totals"))foreach(var order in Rows(entry["order_totals"]))if(seen.Add(S(order,"key"))){counts[platform][1]++;counts[platform][2]+=N(order,"pieces");}
            }
            return counts;
        }
        void UpdateSelectionTotals() {
            var counts=SelectionCounts();
            selectionTotals.Text="Selected: "+counts.Values.Sum(c=>c[0])+" PDFs  |  "+counts.Values.Sum(c=>c[1])+" orders  |  "+counts.Values.Sum(c=>c[2])+" pieces";
            if(counts.Count>0)selectionTotals.Text+="\n"+String.Join("     ",new[]{"SPX","FLASH","LAZADA","TIKTOK"}.Where(counts.ContainsKey).Select(p=>p+": "+counts[p][0]+" PDFs / "+counts[p][1]+" orders / "+counts[p][2]+" pieces"));
            selectionTotals.Text+="\nMatching orders are counted once.";
        }
        internal bool CheckSelectionTotals() {
            var saved=grid.Rows.Cast<DataGridViewRow>().ToDictionary(r=>r,r=>Convert.ToBoolean(r.Cells[0].Value));
            try {
                foreach(DataGridViewRow r in grid.Rows)r.Cells[0].Value=false;
                if(SelectionCounts().Count!=0)return false;
                foreach(var row in grid.Rows.Cast<DataGridViewRow>().Where(r=>!r.Cells[0].ReadOnly)) {
                    row.Cells[0].Value=true;var entry=(Dictionary<string,object>)row.Tag;
                    var counts=SelectionCounts();var values=counts[S(entry,"platform")];
                    if(counts.Count!=1||values[0]!=1||values[1]!=N(entry,"orders")||values[2]!=N(entry,"pieces"))return false;
                    row.Cells[0].Value=false;
                }
                return true;
            } finally {foreach(var pair in saved)pair.Key.Cells[0].Value=pair.Value;RefreshSelection();}
        }
        void ShowError(Exception ex) {lastError=ex.Message;status.Text="Needs attention: "+ex.Message;if(startupArgs.Length==0||startupArgs[0]!="--self-test")MessageBox.Show(this,ex.Message,"Needs attention",MessageBoxButtons.OK,MessageBoxIcon.Warning);}
        void OpenFile(string path) {try{Process.Start(new ProcessStartInfo(path){UseShellExecute=true});}catch(Exception ex){ShowError(ex);}}

        async Task AddPdfs() {
            using(var dlg=new OpenFileDialog {Title="Select packing-list PDFs",Filter="PDF files (*.pdf)|*.pdf",Multiselect=true}) {
                if(dlg.ShowDialog(this)!=DialogResult.OK)return;
                try {
                    var req=new Dictionary<string,object>{{"action","import"},{"files",dlg.FileNames}};
                    var result=await Call(req,"Reading and adding the selected PDFs...");
                    var problems=Rows(result["errors"]);
                    string msg=Values(result["imported"]).Count()+" PDF(s) added. "+Values(result["skipped"]).Count()+" identical file(s) already present.";
                    if(problems.Count>0)msg+="\r\n\r\n"+String.Join("\r\n",problems.Select(p=>S(p,"file")+": "+S(p,"message")));
                    MessageBox.Show(this,msg,"Add PDFs",MessageBoxButtons.OK,problems.Count>0?MessageBoxIcon.Warning:MessageBoxIcon.Information);
                    await Scan();
                } catch(Exception ex){ShowError(ex);}
            }
        }

        async Task Review() {
            var en=Selected();if(en==null||!B(en,"requires_revision"))return;
            using(var review=new Form {Text="Review changed packing file",Size=new Size(900,700),StartPosition=FormStartPosition.CenterParent,Font=Font,MinimumSize=new Size(700,500)}) {
                var panel=new TableLayoutPanel {Dock=DockStyle.Fill,RowCount=3,Padding=new Padding(18)};
                panel.RowStyles.Add(new RowStyle(SizeType.Absolute,60));panel.RowStyles.Add(new RowStyle(SizeType.Percent,100));panel.RowStyles.Add(new RowStyle(SizeType.Absolute,55));
                panel.Controls.Add(new Label {Dock=DockStyle.Fill,Text="Compare the earlier packing record with the current PDF. A revision includes this file's orders again. Check whether they have already been packed before using the revision."},0,0);
                panel.Controls.Add(new TextBox {Text=S(en,"details"),Multiline=true,ReadOnly=true,ScrollBars=ScrollBars.Both,WordWrap=false,Dock=DockStyle.Fill,BackColor=Color.White},0,1);
                var buttons=new FlowLayoutPanel {Dock=DockStyle.Fill,FlowDirection=FlowDirection.RightToLeft,Padding=new Padding(0,10,0,0)};
                var cancel=new Button {Text="Cancel",Width=90,Height=34,DialogResult=DialogResult.Cancel};
                var revise=new Button {Text="Create revision",Width=155,Height=34,DialogResult=DialogResult.Yes};
                var keep=new Button {Text="Keep earlier output",Width=175,Height=34,DialogResult=DialogResult.No};
                buttons.Controls.Add(cancel);buttons.Controls.Add(revise);buttons.Controls.Add(keep);panel.Controls.Add(buttons,0,2);review.Controls.Add(panel);review.CancelButton=cancel;
                var decision=review.ShowDialog(this);
                if(decision==DialogResult.Yes) {
                    revisions.Add(S(en,"path"));grid.CurrentRow.Cells[0].ReadOnly=false;grid.CurrentRow.Cells[0].Value=true;grid.CurrentRow.Cells[5].Value="Reviewed — revision selected";RefreshSelection();
                } else if(decision==DialogResult.No) {
                    try {await Call(new Dictionary<string,object>{{"action","acknowledge"},{"selected",new Dictionary<string,object>{{"path",S(en,"path")},{"sha256",S(en,"sha256")}}}},"Saving the changed-file review...");await Scan();}catch(Exception ex){ShowError(ex);}
                }
            }
        }

        async Task Generate() {
            if(busy)return;
            var selected=new List<Dictionary<string,object>>();
            foreach(DataGridViewRow row in grid.Rows)if(Convert.ToBoolean(row.Cells[0].Value)) {
                var en=(Dictionary<string,object>)row.Tag;
                selected.Add(new Dictionary<string,object>{{"path",S(en,"path")},{"sha256",S(en,"sha256")},{"revision",revisions.Contains(S(en,"path"))},{"courier",S(en,"courier")},{"platform",S(en,"platform")}});
            }
            if(selected.Count==0)return;
            try {
                string format;
                using(var options=PackingDialogs.ExportOptions()) {
                    if(options.ShowDialog(this)!=DialogResult.OK)return;
                    format=((RadioButton)options.Controls.Find("exportPdf",true)[0]).Checked?"pdf":"docx";
                }
                var request=new Dictionary<string,object>{{"action","generate"},{"selected",selected},{"output_format",format}};
                var result=await Call(request,"Preparing "+(format=="pdf"?"PDF":"Word")+" packing lists and checking A4 pagination. Please wait...");
                while(B(result,"confirmation_required")) {
                    using(var warning=PackingDialogs.RepeatWarning(N(result,"repeat_count"))) {
                        if(warning.ShowDialog(this)!=DialogResult.Yes){status.Text="Generation cancelled. No documents created.";return;}
                    }
                    request["repeat_confirmation"]=result["repeat_token"];
                    result=await Call(request,"Generating the confirmed packing lists. Please wait...");
                }
                var made=Rows(result["outputs"]);
                await Scan();
                using(var done=new Form {Text="Packing lists ready",Size=new Size(720,340),StartPosition=FormStartPosition.CenterParent,Font=Font}) {
                    var layout=new TableLayoutPanel {Dock=DockStyle.Fill,Padding=new Padding(20),RowCount=3};
                    layout.RowStyles.Add(new RowStyle(SizeType.Absolute,50));layout.RowStyles.Add(new RowStyle(SizeType.Percent,100));layout.RowStyles.Add(new RowStyle(SizeType.Absolute,50));
                    layout.Controls.Add(new Label {Text="Packing lists are ready. Open a document and print on A4.",Dock=DockStyle.Fill,Font=new Font(Font,FontStyle.Bold)},0,0);
                    var list=new ListBox {Dock=DockStyle.Fill};foreach(var o in made)list.Items.Add(S(o,"filename")+"  ("+S(o,"pages")+" pages)");if(list.Items.Count>0)list.SelectedIndex=0;layout.Controls.Add(list,0,1);
                    var buttons=new FlowLayoutPanel {Dock=DockStyle.Fill,Padding=new Padding(0,10,0,0)};
                    var open=new Button {Text="Open document",Width=150,Height=32};open.Click+=(s,e)=>{if(list.SelectedIndex>=0)OpenFile(S(made[list.SelectedIndex],"path"));};buttons.Controls.Add(open);
                    var folder=new Button {Text="Open Output Folder",Width=180,Height=32};folder.Click+=(s,e)=>OpenFile(OutputFolder);buttons.Controls.Add(folder);
                    var close=new Button {Text="Done",Width=90,Height=32,DialogResult=DialogResult.OK};buttons.Controls.Add(close);layout.Controls.Add(buttons,0,2);done.Controls.Add(layout);done.AcceptButton=close;done.ShowDialog(this);
                }
            } catch(Exception ex){ShowError(ex);}
        }

        class CompletedSelection {
            public Dictionary<string,object> Document;
            public string Pdf;
        }
        static int PlatformOrder(string group) {return group=="SPX"?0:group=="FLASH"?1:group=="LAZADA_TIKTOK"?2:3;}
        static string PlatformLabel(string group) {return group=="LAZADA_TIKTOK"?"Lazada + TikTok":group;}

        void PopulateHistoryTree(TreeView tree,Label caption) {
            tree.BeginUpdate();tree.Nodes.Clear();
            var visible=outputs.Where(o=>File.Exists(Path.Combine(OutputFolder,S(o,"filename"))))
                .OrderBy(o=>PlatformOrder(S(o,"group"))).ThenBy(o=>N(o,"sequence")).ToList();
            foreach(var group in visible.GroupBy(o=>S(o,"group"))) {
                var platform=new TreeNode(PlatformLabel(group.Key));tree.Nodes.Add(platform);
                foreach(var output in group) {
                    string note=S(output,"status")=="active"?"":"  ["+S(output,"status")+"]";
                    var docNode=new TreeNode(S(output,"filename")+"  |  "+N(output,"order_count")+" orders  |  "+N(output,"pieces")+" pieces"+note);
                    docNode.Tag=new CompletedSelection {Document=output};docNode.ToolTipText=docNode.Text;platform.Nodes.Add(docNode);
                    var sources=output.ContainsKey("included_sources")?Rows(output["included_sources"]):new List<Dictionary<string,object>>();
                    foreach(var source in sources) {
                        var orders=Rows(source["orders"]);
                        var tag=new CompletedSelection {Document=output,Pdf=S(source,"path")};
                        var pdfNode=new TreeNode(S(source,"path")+"  |  "+orders.Count+" orders  |  "+N(source,"pieces")+" pieces") {Tag=tag};
                        pdfNode.ToolTipText=pdfNode.Text;docNode.Nodes.Add(pdfNode);
                        foreach(var order in orders) {
                            var items=Rows(order["items"]);
                            var orderNode=new TreeNode((S(order,"platform")=="LAZADA"?"Order Number: ":"Order ID: ")+S(order,"order_id")) {Tag=tag};
                            pdfNode.Nodes.Add(orderNode);
                            foreach(var item in items) {
                                string text=N(item,"pieces")+" pcs | "+S(item,"product")+" | "+S(item,"variation");
                                orderNode.Nodes.Add(new TreeNode(text){Tag=tag,ToolTipText=text});
                            }
                            if(items.Count==0)orderNode.Nodes.Add(new TreeNode("Item details were not recorded for this earlier document."){Tag=tag});
                        }
                    }
                    if(sources.Count==0)docNode.Nodes.Add(new TreeNode("Source PDFs were not recorded for this earlier document."){Tag=docNode.Tag});
                }
                platform.Expand();
            }
            caption.Text=visible.Count==0?"No completed PDF or Word files are currently in this date's Output folder.":
                "Expand a document to see its source PDFs. Expand a PDF and order to see the included products and pieces.";
            tree.EndUpdate();
        }

        Form CreateHistoryForm() {
            var form=new Form {Text="Completed packing documents",Size=new Size(1100,680),MinimumSize=new Size(850,450),StartPosition=FormStartPosition.CenterParent,Font=Font};
            var layout=new TableLayoutPanel {Dock=DockStyle.Fill,Padding=new Padding(18),RowCount=3};
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute,50));layout.RowStyles.Add(new RowStyle(SizeType.Percent,100));layout.RowStyles.Add(new RowStyle(SizeType.Absolute,50));
            var caption=new Label {Dock=DockStyle.Fill};layout.Controls.Add(caption,0,0);
            var tree=new TreeView {Name="completedTree",Dock=DockStyle.Fill,HideSelection=false,ShowNodeToolTips=true,ItemHeight=28,Indent=22,BackColor=Color.White,BorderStyle=BorderStyle.FixedSingle};
            layout.Controls.Add(tree,0,1);
            var buttons=new FlowLayoutPanel {Dock=DockStyle.Fill,Padding=new Padding(0,10,0,0),WrapContents=false};
            var open=new Button {Text="Open document",Width=180,Height=32,Enabled=false};
            var pdf=new Button {Text="Open source PDF",Width=155,Height=32,Enabled=false};
            var refresh=new Button {Text="Refresh list",Width=120,Height=32};
            var log=new Button {Text="Open processing log",Width=175,Height=32};
            var close=new Button {Text="Close",Width=85,Height=32,DialogResult=DialogResult.OK};
            tree.AfterSelect+=(s,e)=>{var chosen=e.Node.Tag as CompletedSelection;open.Enabled=chosen!=null;pdf.Enabled=chosen!=null&&!String.IsNullOrEmpty(chosen.Pdf);};
            open.Click+=(s,e)=>{
                var chosen=tree.SelectedNode==null?null:tree.SelectedNode.Tag as CompletedSelection;if(chosen==null)return;
                var path=Path.Combine(OutputFolder,S(chosen.Document,"filename"));
                if(!File.Exists(path)){PopulateHistoryTree(tree,caption);open.Enabled=false;pdf.Enabled=false;return;}
                if(S(chosen.Document,"status")!="active"){MessageBox.Show(form,"This document is "+S(chosen.Document,"status")+". Use its replacement or generate another copy from the main screen.","Earlier document");return;}
                OpenFile(path);
            };
            pdf.Click+=(s,e)=>{var chosen=tree.SelectedNode==null?null:tree.SelectedNode.Tag as CompletedSelection;if(chosen==null||String.IsNullOrEmpty(chosen.Pdf))return;var path=Path.Combine(DateFolder,chosen.Pdf.Replace('/',Path.DirectorySeparatorChar));if(File.Exists(path))OpenFile(path);else MessageBox.Show(form,"This source PDF is no longer in its recorded folder. The included order data shown here was saved with the generated document.","Source PDF missing");};
            refresh.Click+=(s,e)=>{PopulateHistoryTree(tree,caption);open.Enabled=false;pdf.Enabled=false;};
            log.Click+=(s,e)=>OpenFile(Path.Combine(OutputFolder,"Processed_Files.txt"));
            buttons.Controls.Add(open);buttons.Controls.Add(pdf);buttons.Controls.Add(refresh);buttons.Controls.Add(log);buttons.Controls.Add(close);
            layout.Controls.Add(buttons,0,2);form.Controls.Add(layout);form.CancelButton=close;
            PopulateHistoryTree(tree,caption);return form;
        }

        async Task ShowNames() {
            try {
                var catalog=await Call(new Dictionary<string,object>{{"action","names_list"}},"Loading product names...");
                using(var form=new NameDefinitionsForm(catalog,request=>Call(request,"Updating product names...")))form.ShowDialog(this);
            } catch(Exception ex){ShowError(ex);}
        }
        void ShowHistory() {using(var form=CreateHistoryForm())form.ShowDialog(this);}
    }
}

