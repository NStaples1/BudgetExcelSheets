namespace BudgetExcelSheets
{
    partial class frmMain
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
            this.cmdGo = new DevExpress.XtraEditors.SimpleButton();
            this.dteReportDate = new DevExpress.XtraEditors.DateEdit();
            ((System.ComponentModel.ISupportInitialize)(this.dteReportDate.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dteReportDate.Properties.CalendarTimeProperties)).BeginInit();
            this.SuspendLayout();
            // 
            // cmdGo
            // 
            this.cmdGo.Appearance.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cmdGo.Appearance.Options.UseFont = true;
            this.cmdGo.Location = new System.Drawing.Point(12, 107);
            this.cmdGo.Name = "cmdGo";
            this.cmdGo.Size = new System.Drawing.Size(190, 74);
            this.cmdGo.TabIndex = 0;
            this.cmdGo.Text = "GO";
            this.cmdGo.Click += new System.EventHandler(this.cmdGo_Click);
            // 
            // dteReportDate
            // 
            this.dteReportDate.EditValue = null;
            this.dteReportDate.Location = new System.Drawing.Point(12, 48);
            this.dteReportDate.Name = "dteReportDate";
            this.dteReportDate.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
            new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
            this.dteReportDate.Properties.CalendarTimeProperties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
            new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
            this.dteReportDate.Properties.DisplayFormat.FormatString = "MMM yyyy";
            this.dteReportDate.Properties.DisplayFormat.FormatType = DevExpress.Utils.FormatType.DateTime;
            this.dteReportDate.Properties.EditFormat.FormatString = "MMM yyyy";
            this.dteReportDate.Properties.EditFormat.FormatType = DevExpress.Utils.FormatType.DateTime;
            this.dteReportDate.Properties.MaskSettings.Set("mask", "MMM yyyy");
            this.dteReportDate.Properties.VistaCalendarViewStyle = DevExpress.XtraEditors.VistaCalendarViewStyle.YearView;
            this.dteReportDate.Size = new System.Drawing.Size(190, 20);
            this.dteReportDate.TabIndex = 1;
            // 
            // frmMain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(214, 193);
            this.Controls.Add(this.dteReportDate);
            this.Controls.Add(this.cmdGo);
            this.Name = "frmMain";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "frmMain";
            this.Load += new System.EventHandler(this.frmMain_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dteReportDate.Properties.CalendarTimeProperties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dteReportDate.Properties)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private DevExpress.XtraEditors.SimpleButton cmdGo;
        private DevExpress.XtraEditors.DateEdit dteReportDate;
    }
}