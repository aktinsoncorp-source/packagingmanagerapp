using System;
using System.Drawing;
using System.Windows.Forms;

namespace PackingListManager {
    internal static class PackingDialogs {
        internal static Form RepeatWarning(int count) {
            var form=new Form {Text="Repeated orders",ClientSize=new Size(580,292),FormBorderStyle=FormBorderStyle.FixedDialog,MaximizeBox=false,MinimizeBox=false,StartPosition=FormStartPosition.CenterParent,Font=new Font("Segoe UI",10F)};
            form.Controls.Add(new Label {Text="Warning: There are orders in this packing list that has repeated. Do you wish to continue and generate the repeating orders again?",Location=new Point(24,22),Size=new Size(532,76),Font=new Font(form.Font,FontStyle.Bold)});
            form.Controls.Add(new Label {Text=count+" repeated order(s) found in earlier generation history or other source PDFs. Matching orders within this generation will appear once.",Location=new Point(24,105),Size=new Size(532,62)});
            var no=new Button {Name="repeatNo",Text="No",DialogResult=DialogResult.No,Location=new Point(125,180),Size=new Size(330,54),Font=new Font("Segoe UI",12F,FontStyle.Bold),BackColor=Color.FromArgb(25,87,157),ForeColor=Color.White,FlatStyle=FlatStyle.Flat,TabIndex=0};
            var proceed=new LinkLabel {Name="repeatContinue",Text="Continue to Generate the Repeated Orders",Location=new Point(125,240),Size=new Size(330,34),Font=new Font("Segoe UI",9F),TextAlign=ContentAlignment.MiddleCenter,LinkBehavior=LinkBehavior.NeverUnderline,LinkColor=Color.FromArgb(140,140,140),ActiveLinkColor=Color.FromArgb(110,110,110),VisitedLinkColor=Color.FromArgb(140,140,140),BorderStyle=BorderStyle.None,TabIndex=1};
            proceed.LinkClicked+=(s,e)=>{form.DialogResult=DialogResult.Yes;form.Close();};
            form.Controls.Add(no);form.Controls.Add(proceed);form.AcceptButton=no;form.CancelButton=no;
            form.Shown+=(s,e)=>{form.ActiveControl=no;no.Focus();};
            return form;
        }
    }
}
