namespace RTP3Tools
{
    partial class FormLoadFromXLSParams
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
            this.buttonStartLoading = new System.Windows.Forms.Button();
            this.buttonCancel = new System.Windows.Forms.Button();
            this.spinEditColumnCodeOld = new DevExpress.XtraEditors.SpinEdit();
            this.spinEditColumnCodeNew = new DevExpress.XtraEditors.SpinEdit();
            this.spinEditDataFirstRow = new DevExpress.XtraEditors.SpinEdit();
            this.labelControl1 = new DevExpress.XtraEditors.LabelControl();
            this.labelControl2 = new DevExpress.XtraEditors.LabelControl();
            this.labelControl3 = new DevExpress.XtraEditors.LabelControl();
            this.labelControl4 = new DevExpress.XtraEditors.LabelControl();
            ((System.ComponentModel.ISupportInitialize)(this.spinEditColumnCodeOld.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.spinEditColumnCodeNew.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.spinEditDataFirstRow.Properties)).BeginInit();
            this.SuspendLayout();
            // 
            // buttonStartLoading
            // 
            this.buttonStartLoading.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.buttonStartLoading.Location = new System.Drawing.Point(111, 197);
            this.buttonStartLoading.Name = "buttonStartLoading";
            this.buttonStartLoading.Size = new System.Drawing.Size(75, 23);
            this.buttonStartLoading.TabIndex = 0;
            this.buttonStartLoading.Text = "Загрузить";
            this.buttonStartLoading.UseVisualStyleBackColor = true;
            // 
            // buttonCancel
            // 
            this.buttonCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.buttonCancel.Location = new System.Drawing.Point(192, 197);
            this.buttonCancel.Name = "buttonCancel";
            this.buttonCancel.Size = new System.Drawing.Size(75, 23);
            this.buttonCancel.TabIndex = 1;
            this.buttonCancel.Text = "Отмена";
            this.buttonCancel.UseVisualStyleBackColor = true;
            // 
            // spinEditColumnCodeOld
            // 
            this.spinEditColumnCodeOld.EditValue = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.spinEditColumnCodeOld.Location = new System.Drawing.Point(192, 21);
            this.spinEditColumnCodeOld.Name = "spinEditColumnCodeOld";
            this.spinEditColumnCodeOld.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
            new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
            this.spinEditColumnCodeOld.Properties.IsFloatValue = false;
            this.spinEditColumnCodeOld.Properties.Mask.EditMask = "N00";
            this.spinEditColumnCodeOld.Properties.MaxValue = new decimal(new int[] {
            50,
            0,
            0,
            0});
            this.spinEditColumnCodeOld.Properties.MinValue = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.spinEditColumnCodeOld.Properties.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.DisableTextEditor;
            this.spinEditColumnCodeOld.Size = new System.Drawing.Size(75, 20);
            this.spinEditColumnCodeOld.TabIndex = 3;
            // 
            // spinEditColumnCodeNew
            // 
            this.spinEditColumnCodeNew.EditValue = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.spinEditColumnCodeNew.Location = new System.Drawing.Point(192, 63);
            this.spinEditColumnCodeNew.Name = "spinEditColumnCodeNew";
            this.spinEditColumnCodeNew.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
            new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
            this.spinEditColumnCodeNew.Properties.IsFloatValue = false;
            this.spinEditColumnCodeNew.Properties.Mask.EditMask = "N00";
            this.spinEditColumnCodeNew.Properties.MaxValue = new decimal(new int[] {
            50,
            0,
            0,
            0});
            this.spinEditColumnCodeNew.Properties.MinValue = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.spinEditColumnCodeNew.Properties.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.DisableTextEditor;
            this.spinEditColumnCodeNew.Size = new System.Drawing.Size(75, 20);
            this.spinEditColumnCodeNew.TabIndex = 5;
            // 
            // spinEditDataFirstRow
            // 
            this.spinEditDataFirstRow.EditValue = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.spinEditDataFirstRow.Location = new System.Drawing.Point(192, 104);
            this.spinEditDataFirstRow.Name = "spinEditDataFirstRow";
            this.spinEditDataFirstRow.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
            new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
            this.spinEditDataFirstRow.Properties.IsFloatValue = false;
            this.spinEditDataFirstRow.Properties.Mask.EditMask = "N00";
            this.spinEditDataFirstRow.Properties.MaxValue = new decimal(new int[] {
            50,
            0,
            0,
            0});
            this.spinEditDataFirstRow.Properties.MinValue = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.spinEditDataFirstRow.Properties.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.DisableTextEditor;
            this.spinEditDataFirstRow.Size = new System.Drawing.Size(75, 20);
            this.spinEditDataFirstRow.TabIndex = 7;
            // 
            // labelControl1
            // 
            this.labelControl1.Appearance.Options.UseTextOptions = true;
            this.labelControl1.Appearance.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
            this.labelControl1.AutoSizeMode = DevExpress.XtraEditors.LabelAutoSizeMode.None;
            this.labelControl1.Location = new System.Drawing.Point(12, 12);
            this.labelControl1.Name = "labelControl1";
            this.labelControl1.Size = new System.Drawing.Size(174, 29);
            this.labelControl1.TabIndex = 8;
            this.labelControl1.Text = "Номер столбца со \"старым\" кодом (начиная с 1):";
            // 
            // labelControl2
            // 
            this.labelControl2.Appearance.Options.UseTextOptions = true;
            this.labelControl2.Appearance.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
            this.labelControl2.AutoSizeMode = DevExpress.XtraEditors.LabelAutoSizeMode.None;
            this.labelControl2.Location = new System.Drawing.Point(12, 54);
            this.labelControl2.Name = "labelControl2";
            this.labelControl2.Size = new System.Drawing.Size(174, 29);
            this.labelControl2.TabIndex = 9;
            this.labelControl2.Text = "Номер столбца с \"новым\" кодом (начиная с 1):";
            // 
            // labelControl3
            // 
            this.labelControl3.Appearance.Options.UseTextOptions = true;
            this.labelControl3.Appearance.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
            this.labelControl3.AutoSizeMode = DevExpress.XtraEditors.LabelAutoSizeMode.None;
            this.labelControl3.Location = new System.Drawing.Point(12, 95);
            this.labelControl3.Name = "labelControl3";
            this.labelControl3.Size = new System.Drawing.Size(174, 29);
            this.labelControl3.TabIndex = 10;
            this.labelControl3.Text = "Номер первой строки с данными (начиная с 1):";
            // 
            // labelControl4
            // 
            this.labelControl4.Appearance.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
            this.labelControl4.Appearance.Options.UseFont = true;
            this.labelControl4.Appearance.Options.UseTextOptions = true;
            this.labelControl4.Appearance.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
            this.labelControl4.AutoSizeMode = DevExpress.XtraEditors.LabelAutoSizeMode.None;
            this.labelControl4.Location = new System.Drawing.Point(12, 142);
            this.labelControl4.Name = "labelControl4";
            this.labelControl4.Size = new System.Drawing.Size(255, 29);
            this.labelControl4.TabIndex = 11;
            this.labelControl4.Text = "Важно! Количество строк с данными определяется по первому столбцу файла";
            // 
            // FormLoadFromXLSParams
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(281, 232);
            this.Controls.Add(this.labelControl4);
            this.Controls.Add(this.labelControl3);
            this.Controls.Add(this.labelControl2);
            this.Controls.Add(this.labelControl1);
            this.Controls.Add(this.spinEditDataFirstRow);
            this.Controls.Add(this.spinEditColumnCodeNew);
            this.Controls.Add(this.spinEditColumnCodeOld);
            this.Controls.Add(this.buttonCancel);
            this.Controls.Add(this.buttonStartLoading);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Name = "FormLoadFromXLSParams";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Параметры исходного файла";
            ((System.ComponentModel.ISupportInitialize)(this.spinEditColumnCodeOld.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.spinEditColumnCodeNew.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.spinEditDataFirstRow.Properties)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button buttonStartLoading;
        private System.Windows.Forms.Button buttonCancel;
        private DevExpress.XtraEditors.LabelControl labelControl1;
        private DevExpress.XtraEditors.LabelControl labelControl2;
        private DevExpress.XtraEditors.LabelControl labelControl3;
        private DevExpress.XtraEditors.LabelControl labelControl4;
        public DevExpress.XtraEditors.SpinEdit spinEditColumnCodeOld;
        public DevExpress.XtraEditors.SpinEdit spinEditColumnCodeNew;
        public DevExpress.XtraEditors.SpinEdit spinEditDataFirstRow;
    }
}