namespace exceltabl
{
    partial class Form1
    {
        private System.ComponentModel.IContainer components = null;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            this.dgvResult = new System.Windows.Forms.DataGridView();
            this.btnOpenFile = new System.Windows.Forms.Button();
            this.btnRun = new System.Windows.Forms.Button();
            this.btnSave = new System.Windows.Forms.Button();
            this.lblFilePath = new System.Windows.Forms.Label();
            this.progressBar = new System.Windows.Forms.ProgressBar();
            this.lblStatus = new System.Windows.Forms.Label();
            this.button_max = new System.Windows.Forms.Button();
            this.chkAllVariants = new System.Windows.Forms.CheckBox();
            this.numVariants = new System.Windows.Forms.NumericUpDown();
            ((System.ComponentModel.ISupportInitialize)(this.dgvResult)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numVariants)).BeginInit();
            this.SuspendLayout();
            // 
            // dgvResult
            // 
            this.dgvResult.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvResult.Location = new System.Drawing.Point(16, 15);
            this.dgvResult.Margin = new System.Windows.Forms.Padding(4);
            this.dgvResult.Name = "dgvResult";
            this.dgvResult.RowHeadersWidth = 51;
            this.dgvResult.Size = new System.Drawing.Size(1013, 492);
            this.dgvResult.TabIndex = 0;
            // 
            // btnOpenFile
            // 
            this.btnOpenFile.Location = new System.Drawing.Point(16, 514);
            this.btnOpenFile.Margin = new System.Windows.Forms.Padding(4);
            this.btnOpenFile.Name = "btnOpenFile";
            this.btnOpenFile.Size = new System.Drawing.Size(133, 28);
            this.btnOpenFile.TabIndex = 1;
            this.btnOpenFile.Text = "Открыть файл";
            this.btnOpenFile.UseVisualStyleBackColor = true;
            this.btnOpenFile.Click += new System.EventHandler(this.btnOpenFile_Click);
            // 
            // btnRun
            // 
            this.btnRun.Location = new System.Drawing.Point(157, 514);
            this.btnRun.Margin = new System.Windows.Forms.Padding(4);
            this.btnRun.Name = "btnRun";
            this.btnRun.Size = new System.Drawing.Size(133, 28);
            this.btnRun.TabIndex = 2;
            this.btnRun.Text = "Оптимизировать";
            this.btnRun.UseVisualStyleBackColor = true;
            this.btnRun.Click += new System.EventHandler(this.btnRun_Click);
            // 
            // btnSave
            // 
            this.btnSave.Location = new System.Drawing.Point(896, 514);
            this.btnSave.Margin = new System.Windows.Forms.Padding(4);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(133, 28);
            this.btnSave.TabIndex = 3;
            this.btnSave.Text = "Сохранить";
            this.btnSave.UseVisualStyleBackColor = true;
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // lblFilePath
            // 
            this.lblFilePath.AutoSize = true;
            this.lblFilePath.Location = new System.Drawing.Point(16, 546);
            this.lblFilePath.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lblFilePath.Name = "lblFilePath";
            this.lblFilePath.Size = new System.Drawing.Size(0, 16);
            this.lblFilePath.TabIndex = 4;
            // 
            // progressBar
            // 
            this.progressBar.Location = new System.Drawing.Point(451, 514);
            this.progressBar.Margin = new System.Windows.Forms.Padding(4);
            this.progressBar.Name = "progressBar";
            this.progressBar.Size = new System.Drawing.Size(114, 28);
            this.progressBar.TabIndex = 5;
            this.progressBar.Visible = false;
            // 
            // lblStatus
            // 
            this.lblStatus.AutoSize = true;
            this.lblStatus.Location = new System.Drawing.Point(573, 521);
            this.lblStatus.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lblStatus.Name = "lblStatus";
            this.lblStatus.Size = new System.Drawing.Size(0, 16);
            this.lblStatus.TabIndex = 6;
            // 
            // button_max
            // 
            this.button_max.Location = new System.Drawing.Point(297, 514);
            this.button_max.Name = "button_max";
            this.button_max.Size = new System.Drawing.Size(143, 28);
            this.button_max.TabIndex = 7;
            this.button_max.Text = "Максимальное";
            this.button_max.UseVisualStyleBackColor = true;
            this.button_max.Click += new System.EventHandler(this.button_max_Click);
            // 
            // chkAllVariants
            // 
            this.chkAllVariants.Location = new System.Drawing.Point(680, 515);
            this.chkAllVariants.Name = "chkAllVariants";
            this.chkAllVariants.Size = new System.Drawing.Size(140, 24);
            this.chkAllVariants.Text = "Все варианты";
            this.chkAllVariants.Checked = true;
            this.chkAllVariants.UseVisualStyleBackColor = true;
            this.chkAllVariants.CheckedChanged += new System.EventHandler(this.chkAllVariants_CheckedChanged);
            // 
            // numVariants
            // 
            this.numVariants.Location = new System.Drawing.Point(830, 515);
            this.numVariants.Minimum = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.numVariants.Maximum = new decimal(new int[] {
            10000,
            0,
            0,
            0});
            this.numVariants.Name = "numVariants";
            this.numVariants.Size = new System.Drawing.Size(60, 22);
            this.numVariants.Value = new decimal(new int[] {
            1000,
            0,
            0,
            0});
            this.numVariants.Enabled = false;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1045, 567);
            this.Controls.Add(this.numVariants);
            this.Controls.Add(this.chkAllVariants);
            this.Controls.Add(this.button_max);
            this.Controls.Add(this.lblStatus);
            this.Controls.Add(this.progressBar);
            this.Controls.Add(this.lblFilePath);
            this.Controls.Add(this.btnSave);
            this.Controls.Add(this.btnRun);
            this.Controls.Add(this.btnOpenFile);
            this.Controls.Add(this.dgvResult);
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "Form1";
            this.Text = "Оптимизатор учебной нагрузки";
            this.Load += new System.EventHandler(this.Form1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dgvResult)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numVariants)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();
        }

        private System.Windows.Forms.DataGridView dgvResult;
        private System.Windows.Forms.Button btnOpenFile;
        private System.Windows.Forms.Button btnRun;
        private System.Windows.Forms.Button btnSave;
        private System.Windows.Forms.Label lblFilePath;
        private System.Windows.Forms.ProgressBar progressBar;
        private System.Windows.Forms.Label lblStatus;
        private System.Windows.Forms.Button button_max;
        private System.Windows.Forms.CheckBox chkAllVariants;
        private System.Windows.Forms.NumericUpDown numVariants;
    }
}
