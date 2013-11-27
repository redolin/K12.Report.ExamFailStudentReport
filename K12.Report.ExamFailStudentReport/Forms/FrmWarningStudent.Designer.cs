namespace K12.Report.ExamFailStudentReport.Forms
{
    partial class FrmWarningStudent
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.labelX1 = new DevComponents.DotNetBar.LabelX();
            this.labelX2 = new DevComponents.DotNetBar.LabelX();
            this.btnClose = new DevComponents.DotNetBar.ButtonX();
            this.btnContinue = new DevComponents.DotNetBar.ButtonX();
            this.btnWait = new DevComponents.DotNetBar.ButtonX();
            this.lvWarningStudent = new DevComponents.DotNetBar.Controls.ListViewEx();
            this.ColStudentNumber = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.ColClassName = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.ColStudentName = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.SuspendLayout();
            // 
            // labelX1
            // 
            this.labelX1.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.labelX1.BackgroundStyle.Class = "";
            this.labelX1.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.labelX1.Location = new System.Drawing.Point(13, 13);
            this.labelX1.Name = "labelX1";
            this.labelX1.Size = new System.Drawing.Size(355, 23);
            this.labelX1.TabIndex = 0;
            this.labelX1.Text = "以下學生有多個成績及格標準，可能造成統計數值有誤，";
            // 
            // labelX2
            // 
            this.labelX2.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.labelX2.BackgroundStyle.Class = "";
            this.labelX2.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.labelX2.Location = new System.Drawing.Point(13, 43);
            this.labelX2.Name = "labelX2";
            this.labelX2.Size = new System.Drawing.Size(184, 23);
            this.labelX2.TabIndex = 1;
            this.labelX2.Text = "請確認學生身分類別：";
            // 
            // btnClose
            // 
            this.btnClose.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnClose.BackColor = System.Drawing.Color.Transparent;
            this.btnClose.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnClose.Location = new System.Drawing.Point(317, 268);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(75, 23);
            this.btnClose.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.btnClose.TabIndex = 3;
            this.btnClose.Text = "離開";
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // btnContinue
            // 
            this.btnContinue.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnContinue.BackColor = System.Drawing.Color.Transparent;
            this.btnContinue.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnContinue.Location = new System.Drawing.Point(236, 268);
            this.btnContinue.Name = "btnContinue";
            this.btnContinue.Size = new System.Drawing.Size(75, 23);
            this.btnContinue.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.btnContinue.TabIndex = 4;
            this.btnContinue.Text = "繼續產生";
            this.btnContinue.Click += new System.EventHandler(this.btnContinue_Click);
            // 
            // btnWait
            // 
            this.btnWait.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnWait.BackColor = System.Drawing.Color.Transparent;
            this.btnWait.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnWait.Location = new System.Drawing.Point(13, 268);
            this.btnWait.Name = "btnWait";
            this.btnWait.Size = new System.Drawing.Size(88, 23);
            this.btnWait.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.btnWait.TabIndex = 5;
            this.btnWait.Text = "加到待處理";
            this.btnWait.Click += new System.EventHandler(this.btnWait_Click);
            // 
            // lvWarningStudent
            // 
            // 
            // 
            // 
            this.lvWarningStudent.Border.Class = "ListViewBorder";
            this.lvWarningStudent.Border.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.lvWarningStudent.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.ColStudentNumber,
            this.ColClassName,
            this.ColStudentName});
            this.lvWarningStudent.Location = new System.Drawing.Point(13, 73);
            this.lvWarningStudent.Name = "lvWarningStudent";
            this.lvWarningStudent.Size = new System.Drawing.Size(379, 184);
            this.lvWarningStudent.TabIndex = 6;
            this.lvWarningStudent.UseCompatibleStateImageBehavior = false;
            this.lvWarningStudent.View = System.Windows.Forms.View.Details;
            // 
            // ColStudentNumber
            // 
            this.ColStudentNumber.Text = "學號";
            this.ColStudentNumber.Width = 88;
            // 
            // ColClassName
            // 
            this.ColClassName.Text = "班級";
            this.ColClassName.Width = 87;
            // 
            // ColStudentName
            // 
            this.ColStudentName.Text = "姓名";
            this.ColStudentName.Width = 117;
            // 
            // FrmWarningStudent
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(404, 298);
            this.Controls.Add(this.lvWarningStudent);
            this.Controls.Add(this.btnWait);
            this.Controls.Add(this.btnContinue);
            this.Controls.Add(this.btnClose);
            this.Controls.Add(this.labelX2);
            this.Controls.Add(this.labelX1);
            this.DoubleBuffered = true;
            this.Name = "FrmWarningStudent";
            this.Text = "提示";
            this.Load += new System.EventHandler(this.FrmWarningStudent_Load);
            this.ResumeLayout(false);

        }

        #endregion

        private DevComponents.DotNetBar.LabelX labelX1;
        private DevComponents.DotNetBar.LabelX labelX2;
        private DevComponents.DotNetBar.ButtonX btnClose;
        private DevComponents.DotNetBar.ButtonX btnContinue;
        private DevComponents.DotNetBar.ButtonX btnWait;
        private DevComponents.DotNetBar.Controls.ListViewEx lvWarningStudent;
        private System.Windows.Forms.ColumnHeader ColStudentNumber;
        private System.Windows.Forms.ColumnHeader ColClassName;
        private System.Windows.Forms.ColumnHeader ColStudentName;
    }
}