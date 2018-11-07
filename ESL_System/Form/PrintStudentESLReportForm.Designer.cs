namespace ESL_System.Form
{
    partial class PrintStudentESLReportForm
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
            this.linklabel3 = new System.Windows.Forms.LinkLabel();
            this.linklabel1 = new System.Windows.Forms.LinkLabel();
            this.linklabel2 = new System.Windows.Forms.LinkLabel();
            this.btnPrint = new DevComponents.DotNetBar.ButtonX();
            this.btnClose = new DevComponents.DotNetBar.ButtonX();
            this.dtEnd = new DevComponents.Editors.DateTimeAdv.DateTimeInput();
            this.dtBegin = new DevComponents.Editors.DateTimeAdv.DateTimeInput();
            this.labelX14 = new DevComponents.DotNetBar.LabelX();
            this.labelX13 = new DevComponents.DotNetBar.LabelX();
            this.labelX12 = new DevComponents.DotNetBar.LabelX();
            ((System.ComponentModel.ISupportInitialize)(this.dtEnd)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dtBegin)).BeginInit();
            this.SuspendLayout();
            // 
            // linklabel3
            // 
            this.linklabel3.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.linklabel3.AutoSize = true;
            this.linklabel3.BackColor = System.Drawing.Color.Transparent;
            this.linklabel3.Location = new System.Drawing.Point(42, 92);
            this.linklabel3.Name = "linklabel3";
            this.linklabel3.Size = new System.Drawing.Size(112, 17);
            this.linklabel3.TabIndex = 31;
            this.linklabel3.TabStop = true;
            this.linklabel3.Text = "下載合併欄位總表";
            this.linklabel3.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linklabel3_LinkClicked);
            // 
            // linklabel1
            // 
            this.linklabel1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.linklabel1.AutoSize = true;
            this.linklabel1.BackColor = System.Drawing.Color.Transparent;
            this.linklabel1.Location = new System.Drawing.Point(287, 16);
            this.linklabel1.Name = "linklabel1";
            this.linklabel1.Size = new System.Drawing.Size(86, 17);
            this.linklabel1.TabIndex = 29;
            this.linklabel1.TabStop = true;
            this.linklabel1.Text = "檢視套印樣板";
            this.linklabel1.Visible = false;
            this.linklabel1.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linklabel1_LinkClicked);
            // 
            // linklabel2
            // 
            this.linklabel2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.linklabel2.AutoSize = true;
            this.linklabel2.BackColor = System.Drawing.Color.Transparent;
            this.linklabel2.Location = new System.Drawing.Point(379, 16);
            this.linklabel2.Name = "linklabel2";
            this.linklabel2.Size = new System.Drawing.Size(86, 17);
            this.linklabel2.TabIndex = 30;
            this.linklabel2.TabStop = true;
            this.linklabel2.Text = "變更套印樣板";
            this.linklabel2.Visible = false;
            this.linklabel2.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linklabel2_LinkClicked);
            // 
            // btnPrint
            // 
            this.btnPrint.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnPrint.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnPrint.BackColor = System.Drawing.Color.Transparent;
            this.btnPrint.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnPrint.Location = new System.Drawing.Point(298, 86);
            this.btnPrint.Name = "btnPrint";
            this.btnPrint.Size = new System.Drawing.Size(75, 23);
            this.btnPrint.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.btnPrint.TabIndex = 32;
            this.btnPrint.Text = "列印";
            this.btnPrint.Click += new System.EventHandler(this.btnPrint_Click);
            // 
            // btnClose
            // 
            this.btnClose.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnClose.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnClose.BackColor = System.Drawing.Color.Transparent;
            this.btnClose.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnClose.Location = new System.Drawing.Point(389, 86);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(75, 23);
            this.btnClose.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.btnClose.TabIndex = 33;
            this.btnClose.Text = "離開";
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // dtEnd
            // 
            this.dtEnd.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.dtEnd.BackgroundStyle.Class = "DateTimeInputBackground";
            this.dtEnd.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.dtEnd.ButtonDropDown.Shortcut = DevComponents.DotNetBar.eShortcut.AltDown;
            this.dtEnd.ButtonDropDown.Visible = true;
            this.dtEnd.IsPopupCalendarOpen = false;
            this.dtEnd.Location = new System.Drawing.Point(290, 42);
            // 
            // 
            // 
            this.dtEnd.MonthCalendar.AnnuallyMarkedDates = new System.DateTime[0];
            // 
            // 
            // 
            this.dtEnd.MonthCalendar.BackgroundStyle.BackColor = System.Drawing.SystemColors.Window;
            this.dtEnd.MonthCalendar.BackgroundStyle.Class = "";
            this.dtEnd.MonthCalendar.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.dtEnd.MonthCalendar.ClearButtonVisible = true;
            // 
            // 
            // 
            this.dtEnd.MonthCalendar.CommandsBackgroundStyle.BackColor2SchemePart = DevComponents.DotNetBar.eColorSchemePart.BarBackground2;
            this.dtEnd.MonthCalendar.CommandsBackgroundStyle.BackColorGradientAngle = 90;
            this.dtEnd.MonthCalendar.CommandsBackgroundStyle.BackColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.BarBackground;
            this.dtEnd.MonthCalendar.CommandsBackgroundStyle.BorderTop = DevComponents.DotNetBar.eStyleBorderType.Solid;
            this.dtEnd.MonthCalendar.CommandsBackgroundStyle.BorderTopColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.BarDockedBorder;
            this.dtEnd.MonthCalendar.CommandsBackgroundStyle.BorderTopWidth = 1;
            this.dtEnd.MonthCalendar.CommandsBackgroundStyle.Class = "";
            this.dtEnd.MonthCalendar.CommandsBackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.dtEnd.MonthCalendar.DayNames = new string[] {
        "日",
        "一",
        "二",
        "三",
        "四",
        "五",
        "六"};
            this.dtEnd.MonthCalendar.DisplayMonth = new System.DateTime(2012, 11, 1, 0, 0, 0, 0);
            this.dtEnd.MonthCalendar.MarkedDates = new System.DateTime[0];
            this.dtEnd.MonthCalendar.MaxDate = new System.DateTime(3000, 12, 31, 0, 0, 0, 0);
            this.dtEnd.MonthCalendar.MinDate = new System.DateTime(1900, 1, 1, 0, 0, 0, 0);
            this.dtEnd.MonthCalendar.MonthlyMarkedDates = new System.DateTime[0];
            // 
            // 
            // 
            this.dtEnd.MonthCalendar.NavigationBackgroundStyle.BackColor2SchemePart = DevComponents.DotNetBar.eColorSchemePart.PanelBackground2;
            this.dtEnd.MonthCalendar.NavigationBackgroundStyle.BackColorGradientAngle = 90;
            this.dtEnd.MonthCalendar.NavigationBackgroundStyle.BackColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.PanelBackground;
            this.dtEnd.MonthCalendar.NavigationBackgroundStyle.Class = "";
            this.dtEnd.MonthCalendar.NavigationBackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.dtEnd.MonthCalendar.TodayButtonVisible = true;
            this.dtEnd.MonthCalendar.WeeklyMarkedDays = new System.DayOfWeek[0];
            this.dtEnd.Name = "dtEnd";
            this.dtEnd.Size = new System.Drawing.Size(143, 25);
            this.dtEnd.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.dtEnd.TabIndex = 35;
            // 
            // dtBegin
            // 
            this.dtBegin.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.dtBegin.BackgroundStyle.Class = "DateTimeInputBackground";
            this.dtBegin.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.dtBegin.ButtonDropDown.Shortcut = DevComponents.DotNetBar.eShortcut.AltDown;
            this.dtBegin.ButtonDropDown.Visible = true;
            this.dtBegin.IsPopupCalendarOpen = false;
            this.dtBegin.Location = new System.Drawing.Point(72, 42);
            this.dtBegin.MaxDate = new System.DateTime(3000, 12, 31, 0, 0, 0, 0);
            this.dtBegin.MinDate = new System.DateTime(1900, 1, 1, 0, 0, 0, 0);
            // 
            // 
            // 
            this.dtBegin.MonthCalendar.AnnuallyMarkedDates = new System.DateTime[0];
            // 
            // 
            // 
            this.dtBegin.MonthCalendar.BackgroundStyle.BackColor = System.Drawing.SystemColors.Window;
            this.dtBegin.MonthCalendar.BackgroundStyle.Class = "";
            this.dtBegin.MonthCalendar.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.dtBegin.MonthCalendar.ClearButtonVisible = true;
            // 
            // 
            // 
            this.dtBegin.MonthCalendar.CommandsBackgroundStyle.BackColor2SchemePart = DevComponents.DotNetBar.eColorSchemePart.BarBackground2;
            this.dtBegin.MonthCalendar.CommandsBackgroundStyle.BackColorGradientAngle = 90;
            this.dtBegin.MonthCalendar.CommandsBackgroundStyle.BackColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.BarBackground;
            this.dtBegin.MonthCalendar.CommandsBackgroundStyle.BorderTop = DevComponents.DotNetBar.eStyleBorderType.Solid;
            this.dtBegin.MonthCalendar.CommandsBackgroundStyle.BorderTopColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.BarDockedBorder;
            this.dtBegin.MonthCalendar.CommandsBackgroundStyle.BorderTopWidth = 1;
            this.dtBegin.MonthCalendar.CommandsBackgroundStyle.Class = "";
            this.dtBegin.MonthCalendar.CommandsBackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.dtBegin.MonthCalendar.DayNames = new string[] {
        "日",
        "一",
        "二",
        "三",
        "四",
        "五",
        "六"};
            this.dtBegin.MonthCalendar.DisplayMonth = new System.DateTime(2012, 11, 1, 0, 0, 0, 0);
            this.dtBegin.MonthCalendar.MarkedDates = new System.DateTime[0];
            this.dtBegin.MonthCalendar.MaxDate = new System.DateTime(3000, 12, 31, 0, 0, 0, 0);
            this.dtBegin.MonthCalendar.MinDate = new System.DateTime(1900, 1, 1, 0, 0, 0, 0);
            this.dtBegin.MonthCalendar.MonthlyMarkedDates = new System.DateTime[0];
            // 
            // 
            // 
            this.dtBegin.MonthCalendar.NavigationBackgroundStyle.BackColor2SchemePart = DevComponents.DotNetBar.eColorSchemePart.PanelBackground2;
            this.dtBegin.MonthCalendar.NavigationBackgroundStyle.BackColorGradientAngle = 90;
            this.dtBegin.MonthCalendar.NavigationBackgroundStyle.BackColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.PanelBackground;
            this.dtBegin.MonthCalendar.NavigationBackgroundStyle.Class = "";
            this.dtBegin.MonthCalendar.NavigationBackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.dtBegin.MonthCalendar.TodayButtonVisible = true;
            this.dtBegin.MonthCalendar.WeeklyMarkedDays = new System.DayOfWeek[0];
            this.dtBegin.Name = "dtBegin";
            this.dtBegin.Size = new System.Drawing.Size(143, 25);
            this.dtBegin.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.dtBegin.TabIndex = 34;
            // 
            // labelX14
            // 
            this.labelX14.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.labelX14.AutoSize = true;
            this.labelX14.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.labelX14.BackgroundStyle.Class = "";
            this.labelX14.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.labelX14.Location = new System.Drawing.Point(250, 46);
            this.labelX14.Name = "labelX14";
            this.labelX14.Size = new System.Drawing.Size(34, 21);
            this.labelX14.TabIndex = 38;
            this.labelX14.Text = "結束";
            // 
            // labelX13
            // 
            this.labelX13.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.labelX13.AutoSize = true;
            this.labelX13.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.labelX13.BackgroundStyle.Class = "";
            this.labelX13.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.labelX13.Location = new System.Drawing.Point(32, 46);
            this.labelX13.Name = "labelX13";
            this.labelX13.Size = new System.Drawing.Size(34, 21);
            this.labelX13.TabIndex = 37;
            this.labelX13.Text = "開始";
            // 
            // labelX12
            // 
            this.labelX12.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.labelX12.AutoSize = true;
            this.labelX12.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.labelX12.BackgroundStyle.Class = "";
            this.labelX12.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.labelX12.Location = new System.Drawing.Point(12, 15);
            this.labelX12.Name = "labelX12";
            this.labelX12.Size = new System.Drawing.Size(256, 21);
            this.labelX12.TabIndex = 36;
            this.labelX12.Text = "請選擇日期區間：(依區間統計缺曠、獎懲)";
            // 
            // PrintStudentESLReportForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(479, 118);
            this.Controls.Add(this.dtEnd);
            this.Controls.Add(this.dtBegin);
            this.Controls.Add(this.labelX14);
            this.Controls.Add(this.labelX13);
            this.Controls.Add(this.labelX12);
            this.Controls.Add(this.btnClose);
            this.Controls.Add(this.btnPrint);
            this.Controls.Add(this.linklabel3);
            this.Controls.Add(this.linklabel1);
            this.Controls.Add(this.linklabel2);
            this.DoubleBuffered = true;
            this.MaximumSize = new System.Drawing.Size(495, 157);
            this.MinimumSize = new System.Drawing.Size(495, 157);
            this.Name = "PrintStudentESLReportForm";
            this.Text = "列印學生ESL成績單";
            ((System.ComponentModel.ISupportInitialize)(this.dtEnd)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dtBegin)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.LinkLabel linklabel3;
        private System.Windows.Forms.LinkLabel linklabel1;
        private System.Windows.Forms.LinkLabel linklabel2;
        private DevComponents.DotNetBar.ButtonX btnPrint;
        private DevComponents.DotNetBar.ButtonX btnClose;
        private DevComponents.Editors.DateTimeAdv.DateTimeInput dtEnd;
        private DevComponents.Editors.DateTimeAdv.DateTimeInput dtBegin;
        private DevComponents.DotNetBar.LabelX labelX14;
        private DevComponents.DotNetBar.LabelX labelX13;
        private DevComponents.DotNetBar.LabelX labelX12;
    }
}