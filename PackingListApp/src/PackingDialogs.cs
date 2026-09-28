using System;
using System.Drawing;
using System.Windows.Forms;

namespace PackingListManager {
    internal static class PackingDialogs {
        internal static Form ExportOptions() {
            var form=new Form {Text="Export packing lists",ClientSize=new Size(450,225),FormBorderStyle=FormBorderStyle.FixedDialog,MaximizeBox=false,MinimizeBox=false,StartPosition=FormStartPosition.CenterParent,Font=new Font("Segoe UI",10F)};
            form.Controls.Add(new Label {Text="Choose the file format",Location=new Point(22,20),AutoSize=true,Font=new Font(form.Font,FontStyle.Bold)});
            var pdf=new RadioButton {Name="exportPdf",Text="PDF",Checked=true,Location=new Point(26,59),AutoSize=true,TabIndex=0};
            var word=new RadioButton {Name="exportWord",Text="Word (.docx)",Location=new Point(26,94),AutoSize=true,TabIndex=1};
            form.Controls.Add(pdf);form.Controls.Add(word);
            form.Controls.Add(new Label {Text="The selected format is saved in Completed documents.",Location=new Point(22,130),Size=new Size(410,25)});
            var generate=new Button {Name="exportGenerate",Text="Generate",DialogResult=DialogResult.OK,Location=new Point(205,172),Size=new Size(130,36),BackColor=Color.FromArgb(25,87,157),ForeColor=Color.White,FlatStyle=FlatStyle.Flat,TabIndex=2};
            var cancel=new Button {Name="exportCancel",Text="Cancel",DialogResult=DialogResult.Cancel,Location=new Point(345,172),Size=new Size(85,36),TabIndex=3};
            form.Controls.Add(generate);form.Controls.Add(cancel);form.AcceptButton=generate;form.CancelButton=cancel;
            return form;
        }
        internal static Form RepeatWarning(int count) {
            var form=new Form {Text="Repeated orders",ClientSize=new Size(580,245),FormBorderStyle=FormBorderStyle.FixedDialog,MaximizeBox=false,MinimizeBox=false,StartPosition=FormStartPosition.CenterParent,Font=new Font("Segoe UI",10F)};
            form.Controls.Add(new Label {Text="Warning: There are orders in this packing list that has repeated. Do you wish to continue and generate the repeating orders again?",Location=new Point(24,22),Size=new Size(532,76),Font=new Font(form.Font,FontStyle.Bold)});
            form.Controls.Add(new Label {Text=count+" repeated order(s) found in earlier generation history or other source PDFs. Matching orders within this generation will appear once.",Location=new Point(24,105),Size=new Size(532,62)});
            var no=new Button {Name="repeatNo",Text="No",DialogResult=DialogResult.No,Location=new Point(274,187),Size=new Size(170,40),BackColor=Color.FromArgb(25,87,157),ForeColor=Color.White,FlatStyle=FlatStyle.Flat,TabIndex=0};
            var yes=new Button {Name="repeatYes",Text="Yes",DialogResult=DialogResult.Yes,Location=new Point(461,193),Size=new Size(90,28),BackColor=Color.White,FlatStyle=FlatStyle.Flat,TabIndex=1};
            form.Controls.Add(no);form.Controls.Add(yes);form.AcceptButton=no;form.CancelButton=no;
            form.Shown+=(s,e)=>{form.ActiveControl=no;no.Focus();};
            return form;
        }
    }
}
