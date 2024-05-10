using DevExpress.XtraEditors;

namespace BOD_Utility
{
    partial class Settings
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
            this.lbl_Attempts = new DevExpress.XtraEditors.LabelControl();
            this.labelControl1 = new DevExpress.XtraEditors.LabelControl();
            this.labelControl2 = new DevExpress.XtraEditors.LabelControl();
            this.labelControl3 = new DevExpress.XtraEditors.LabelControl();
            this.txt_FromAddress = new DevExpress.XtraEditors.TextEdit();
            this.labelControl4 = new DevExpress.XtraEditors.LabelControl();
            this.txt_ToAddress = new DevExpress.XtraEditors.TextEdit();
            this.labelControl5 = new DevExpress.XtraEditors.LabelControl();
            this.txt_Password = new DevExpress.XtraEditors.TextEdit();
            this.labelControl6 = new DevExpress.XtraEditors.LabelControl();
            this.txt_SMTP = new DevExpress.XtraEditors.TextEdit();
            this.labelControl7 = new DevExpress.XtraEditors.LabelControl();
            this.groupControl1 = new DevExpress.XtraEditors.GroupControl();
            this.groupControl2 = new DevExpress.XtraEditors.GroupControl();
            this.radioBtn_No = new System.Windows.Forms.RadioButton();
            this.radioBtn_Yes = new System.Windows.Forms.RadioButton();
            this.groupControl3 = new DevExpress.XtraEditors.GroupControl();
            this.spEdit_Interval = new DevExpress.XtraEditors.SpinEdit();
            this.spEdit_Attempts = new DevExpress.XtraEditors.SpinEdit();
            ((System.ComponentModel.ISupportInitialize)(this.txt_FromAddress.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.txt_ToAddress.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.txt_Password.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.txt_SMTP.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.groupControl1)).BeginInit();
            this.groupControl1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.groupControl2)).BeginInit();
            this.groupControl2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.groupControl3)).BeginInit();
            this.groupControl3.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.spEdit_Interval.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.spEdit_Attempts.Properties)).BeginInit();
            this.SuspendLayout();
            // 
            // lbl_Attempts
            // 
            this.lbl_Attempts.Appearance.Font = new System.Drawing.Font("Segoe UI", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbl_Attempts.Appearance.Options.UseFont = true;
            this.lbl_Attempts.Location = new System.Drawing.Point(22, 54);
            this.lbl_Attempts.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.lbl_Attempts.Name = "lbl_Attempts";
            this.lbl_Attempts.Size = new System.Drawing.Size(74, 23);
            this.lbl_Attempts.TabIndex = 1;
            this.lbl_Attempts.Text = "Attempts:";
            // 
            // labelControl1
            // 
            this.labelControl1.Appearance.Font = new System.Drawing.Font("Segoe UI", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelControl1.Appearance.Options.UseFont = true;
            this.labelControl1.Location = new System.Drawing.Point(293, 56);
            this.labelControl1.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.labelControl1.Name = "labelControl1";
            this.labelControl1.Size = new System.Drawing.Size(61, 23);
            this.labelControl1.TabIndex = 1;
            this.labelControl1.Text = "Interval:";
            // 
            // labelControl2
            // 
            this.labelControl2.Appearance.Font = new System.Drawing.Font("Segoe UI", 7.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelControl2.Appearance.Options.UseFont = true;
            this.labelControl2.Location = new System.Drawing.Point(462, 60);
            this.labelControl2.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.labelControl2.Name = "labelControl2";
            this.labelControl2.Size = new System.Drawing.Size(56, 17);
            this.labelControl2.TabIndex = 1;
            this.labelControl2.Text = "(seconds)";
            // 
            // labelControl3
            // 
            this.labelControl3.Appearance.Font = new System.Drawing.Font("Segoe UI", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelControl3.Appearance.Options.UseFont = true;
            this.labelControl3.Location = new System.Drawing.Point(22, 51);
            this.labelControl3.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.labelControl3.Name = "labelControl3";
            this.labelControl3.Size = new System.Drawing.Size(130, 23);
            this.labelControl3.TabIndex = 1;
            this.labelControl3.Text = "Start NOTIS API :";
            // 
            // txt_FromAddress
            // 
            this.txt_FromAddress.Location = new System.Drawing.Point(173, 59);
            this.txt_FromAddress.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.txt_FromAddress.Name = "txt_FromAddress";
            this.txt_FromAddress.Properties.Appearance.Font = new System.Drawing.Font("Segoe UI", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txt_FromAddress.Properties.Appearance.Options.UseFont = true;
            this.txt_FromAddress.Size = new System.Drawing.Size(383, 30);
            this.txt_FromAddress.TabIndex = 3;
            // 
            // labelControl4
            // 
            this.labelControl4.Appearance.Font = new System.Drawing.Font("Segoe UI", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelControl4.Appearance.Options.UseFont = true;
            this.labelControl4.Location = new System.Drawing.Point(35, 62);
            this.labelControl4.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.labelControl4.Name = "labelControl4";
            this.labelControl4.Size = new System.Drawing.Size(108, 23);
            this.labelControl4.TabIndex = 1;
            this.labelControl4.Text = "From Address:";
            // 
            // txt_ToAddress
            // 
            this.txt_ToAddress.Location = new System.Drawing.Point(173, 108);
            this.txt_ToAddress.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.txt_ToAddress.Name = "txt_ToAddress";
            this.txt_ToAddress.Properties.Appearance.Font = new System.Drawing.Font("Segoe UI", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txt_ToAddress.Properties.Appearance.Options.UseFont = true;
            this.txt_ToAddress.Size = new System.Drawing.Size(383, 30);
            this.txt_ToAddress.TabIndex = 4;
            // 
            // labelControl5
            // 
            this.labelControl5.Appearance.Font = new System.Drawing.Font("Segoe UI", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelControl5.Appearance.Options.UseFont = true;
            this.labelControl5.Location = new System.Drawing.Point(49, 111);
            this.labelControl5.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.labelControl5.Name = "labelControl5";
            this.labelControl5.Size = new System.Drawing.Size(88, 23);
            this.labelControl5.TabIndex = 1;
            this.labelControl5.Text = "To Address:";
            // 
            // txt_Password
            // 
            this.txt_Password.Location = new System.Drawing.Point(173, 156);
            this.txt_Password.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.txt_Password.Name = "txt_Password";
            this.txt_Password.Properties.Appearance.Font = new System.Drawing.Font("Segoe UI", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txt_Password.Properties.Appearance.Options.UseFont = true;
            this.txt_Password.Properties.PasswordChar = '*';
            this.txt_Password.Size = new System.Drawing.Size(383, 30);
            this.txt_Password.TabIndex = 5;
            // 
            // labelControl6
            // 
            this.labelControl6.Appearance.Font = new System.Drawing.Font("Segoe UI", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelControl6.Appearance.Options.UseFont = true;
            this.labelControl6.Location = new System.Drawing.Point(62, 159);
            this.labelControl6.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.labelControl6.Name = "labelControl6";
            this.labelControl6.Size = new System.Drawing.Size(75, 23);
            this.labelControl6.TabIndex = 1;
            this.labelControl6.Text = "Password:";
            // 
            // txt_SMTP
            // 
            this.txt_SMTP.Location = new System.Drawing.Point(173, 204);
            this.txt_SMTP.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.txt_SMTP.Name = "txt_SMTP";
            this.txt_SMTP.Properties.Appearance.Font = new System.Drawing.Font("Segoe UI", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txt_SMTP.Properties.Appearance.Options.UseFont = true;
            this.txt_SMTP.Size = new System.Drawing.Size(383, 30);
            this.txt_SMTP.TabIndex = 6;
            // 
            // labelControl7
            // 
            this.labelControl7.Appearance.Font = new System.Drawing.Font("Segoe UI", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelControl7.Appearance.Options.UseFont = true;
            this.labelControl7.Location = new System.Drawing.Point(90, 207);
            this.labelControl7.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.labelControl7.Name = "labelControl7";
            this.labelControl7.Size = new System.Drawing.Size(47, 23);
            this.labelControl7.TabIndex = 1;
            this.labelControl7.Text = "SMTP:";
            // 
            // groupControl1
            // 
            this.groupControl1.Controls.Add(this.txt_FromAddress);
            this.groupControl1.Controls.Add(this.txt_ToAddress);
            this.groupControl1.Controls.Add(this.labelControl4);
            this.groupControl1.Controls.Add(this.txt_Password);
            this.groupControl1.Controls.Add(this.labelControl5);
            this.groupControl1.Controls.Add(this.labelControl7);
            this.groupControl1.Controls.Add(this.txt_SMTP);
            this.groupControl1.Controls.Add(this.labelControl6);
            this.groupControl1.Location = new System.Drawing.Point(12, 259);
            this.groupControl1.Name = "groupControl1";
            this.groupControl1.Size = new System.Drawing.Size(589, 259);
            this.groupControl1.TabIndex = 7;
            this.groupControl1.Text = "Email Configuration";
            // 
            // groupControl2
            // 
            this.groupControl2.Controls.Add(this.radioBtn_No);
            this.groupControl2.Controls.Add(this.radioBtn_Yes);
            this.groupControl2.Controls.Add(this.labelControl3);
            this.groupControl2.Location = new System.Drawing.Point(12, 141);
            this.groupControl2.Name = "groupControl2";
            this.groupControl2.Size = new System.Drawing.Size(589, 103);
            this.groupControl2.TabIndex = 8;
            this.groupControl2.Text = "NOTIS API Configuration";
            // 
            // radioBtn_No
            // 
            this.radioBtn_No.AutoSize = true;
            this.radioBtn_No.Font = new System.Drawing.Font("Segoe UI", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.radioBtn_No.Location = new System.Drawing.Point(350, 52);
            this.radioBtn_No.Name = "radioBtn_No";
            this.radioBtn_No.Size = new System.Drawing.Size(54, 27);
            this.radioBtn_No.TabIndex = 4;
            this.radioBtn_No.TabStop = true;
            this.radioBtn_No.Text = "No";
            this.radioBtn_No.UseVisualStyleBackColor = true;
            // 
            // radioBtn_Yes
            // 
            this.radioBtn_Yes.AutoSize = true;
            this.radioBtn_Yes.Font = new System.Drawing.Font("Segoe UI", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.radioBtn_Yes.ForeColor = System.Drawing.Color.OldLace;
            this.radioBtn_Yes.Location = new System.Drawing.Point(204, 52);
            this.radioBtn_Yes.Name = "radioBtn_Yes";
            this.radioBtn_Yes.Size = new System.Drawing.Size(55, 27);
            this.radioBtn_Yes.TabIndex = 3;
            this.radioBtn_Yes.TabStop = true;
            this.radioBtn_Yes.Text = "Yes";
            this.radioBtn_Yes.UseVisualStyleBackColor = true;
            // 
            // groupControl3
            // 
            this.groupControl3.Controls.Add(this.spEdit_Interval);
            this.groupControl3.Controls.Add(this.spEdit_Attempts);
            this.groupControl3.Controls.Add(this.lbl_Attempts);
            this.groupControl3.Controls.Add(this.labelControl2);
            this.groupControl3.Controls.Add(this.labelControl1);
            this.groupControl3.Location = new System.Drawing.Point(12, 22);
            this.groupControl3.Name = "groupControl3";
            this.groupControl3.Size = new System.Drawing.Size(589, 100);
            this.groupControl3.TabIndex = 9;
            this.groupControl3.Text = "Connection Configuration";
            // 
            // spEdit_Interval
            // 
            this.spEdit_Interval.EditValue = new decimal(new int[] {
            0,
            0,
            0,
            0});
            this.spEdit_Interval.Location = new System.Drawing.Point(371, 55);
            this.spEdit_Interval.Name = "spEdit_Interval";
            this.spEdit_Interval.Properties.Appearance.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.spEdit_Interval.Properties.Appearance.Options.UseFont = true;
            this.spEdit_Interval.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
            new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
            this.spEdit_Interval.Properties.MaxValue = new decimal(new int[] {
            300,
            0,
            0,
            0});
            this.spEdit_Interval.Size = new System.Drawing.Size(84, 26);
            this.spEdit_Interval.TabIndex = 0;
            // 
            // spEdit_Attempts
            // 
            this.spEdit_Attempts.EditValue = new decimal(new int[] {
            0,
            0,
            0,
            0});
            this.spEdit_Attempts.Location = new System.Drawing.Point(113, 53);
            this.spEdit_Attempts.Name = "spEdit_Attempts";
            this.spEdit_Attempts.Properties.Appearance.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.spEdit_Attempts.Properties.Appearance.Options.UseFont = true;
            this.spEdit_Attempts.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
            new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
            this.spEdit_Attempts.Properties.MaxValue = new decimal(new int[] {
            10,
            0,
            0,
            0});
            this.spEdit_Attempts.Size = new System.Drawing.Size(84, 26);
            this.spEdit_Attempts.TabIndex = 0;
            // 
            // Settings
            // 
            this.Appearance.Options.UseFont = true;
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(625, 547);
            this.Controls.Add(this.groupControl3);
            this.Controls.Add(this.groupControl2);
            this.Controls.Add(this.groupControl1);
            this.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "Settings";
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Settings";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.Settings_FormClosing);
            this.Shown += new System.EventHandler(this.Settings_Shown);
            ((System.ComponentModel.ISupportInitialize)(this.txt_FromAddress.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.txt_ToAddress.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.txt_Password.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.txt_SMTP.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.groupControl1)).EndInit();
            this.groupControl1.ResumeLayout(false);
            this.groupControl1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.groupControl2)).EndInit();
            this.groupControl2.ResumeLayout(false);
            this.groupControl2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.groupControl3)).EndInit();
            this.groupControl3.ResumeLayout(false);
            this.groupControl3.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.spEdit_Interval.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.spEdit_Attempts.Properties)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion
        private DevExpress.XtraEditors.LabelControl lbl_Attempts;
        private DevExpress.XtraEditors.LabelControl labelControl1;
        private DevExpress.XtraEditors.LabelControl labelControl2;
        private DevExpress.XtraEditors.LabelControl labelControl3;
        private DevExpress.XtraEditors.TextEdit txt_FromAddress;
        private DevExpress.XtraEditors.LabelControl labelControl4;
        private DevExpress.XtraEditors.TextEdit txt_ToAddress;
        private DevExpress.XtraEditors.LabelControl labelControl5;
        private DevExpress.XtraEditors.TextEdit txt_Password;
        private DevExpress.XtraEditors.LabelControl labelControl6;
        private DevExpress.XtraEditors.TextEdit txt_SMTP;
        private DevExpress.XtraEditors.LabelControl labelControl7;
        private DevExpress.XtraEditors.GroupControl groupControl1;
        private DevExpress.XtraEditors.GroupControl groupControl2;
        private DevExpress.XtraEditors.GroupControl groupControl3;
        private SpinEdit spEdit_Attempts;
        private SpinEdit spEdit_Interval;
        private System.Windows.Forms.RadioButton radioBtn_No;
        private System.Windows.Forms.RadioButton radioBtn_Yes;
    }
}