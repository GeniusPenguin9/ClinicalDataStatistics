namespace ClinicalDataStatistics
{
    partial class Form1
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.ChooseRootDirectory = new System.Windows.Forms.Button();
            this.CreateAnalysis = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.CreateExcelAnalysis = new System.Windows.Forms.Button();
            this.CreateStatistics = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // ChooseRootDirectory
            // 
            this.ChooseRootDirectory.Location = new System.Drawing.Point(31, 36);
            this.ChooseRootDirectory.Name = "ChooseRootDirectory";
            this.ChooseRootDirectory.Size = new System.Drawing.Size(231, 57);
            this.ChooseRootDirectory.TabIndex = 0;
            this.ChooseRootDirectory.Text = "ChooseRootDirectory";
            this.ChooseRootDirectory.UseVisualStyleBackColor = true;
            this.ChooseRootDirectory.Click += new System.EventHandler(this.ChooseRootDirectory_Click);
            // 
            // CreateAnalysis
            // 
            this.CreateAnalysis.Location = new System.Drawing.Point(31, 128);
            this.CreateAnalysis.Name = "CreateAnalysis";
            this.CreateAnalysis.Size = new System.Drawing.Size(231, 63);
            this.CreateAnalysis.TabIndex = 1;
            this.CreateAnalysis.Text = "CreateTxtAnalysis";
            this.CreateAnalysis.UseVisualStyleBackColor = true;
            this.CreateAnalysis.Click += new System.EventHandler(this.CreateAnalysis_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(45, 249);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(431, 15);
            this.label1.TabIndex = 2;
            this.label1.Text = "Function Description: calculate Joule of each channel";
            // 
            // CreateExcelAnalysis
            // 
            this.CreateExcelAnalysis.Location = new System.Drawing.Point(288, 35);
            this.CreateExcelAnalysis.Name = "CreateExcelAnalysis";
            this.CreateExcelAnalysis.Size = new System.Drawing.Size(231, 58);
            this.CreateExcelAnalysis.TabIndex = 3;
            this.CreateExcelAnalysis.Text = "CreateExcelAnalysis";
            this.CreateExcelAnalysis.UseVisualStyleBackColor = true;
            this.CreateExcelAnalysis.Click += new System.EventHandler(this.CreateExcelAnalysis_Click);
            // 
            // CreateStatistics
            // 
            this.CreateStatistics.Location = new System.Drawing.Point(288, 128);
            this.CreateStatistics.Name = "CreateStatistics";
            this.CreateStatistics.Size = new System.Drawing.Size(231, 63);
            this.CreateStatistics.TabIndex = 4;
            this.CreateStatistics.Text = "CreateStatistics";
            this.CreateStatistics.UseVisualStyleBackColor = true;
            this.CreateStatistics.Click += new System.EventHandler(this.CreateStatistics_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(538, 303);
            this.Controls.Add(this.CreateStatistics);
            this.Controls.Add(this.CreateExcelAnalysis);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.CreateAnalysis);
            this.Controls.Add(this.ChooseRootDirectory);
            this.Name = "Form1";
            this.Text = "Form1";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button ChooseRootDirectory;
        private System.Windows.Forms.Button CreateAnalysis;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button CreateExcelAnalysis;
        private System.Windows.Forms.Button CreateStatistics;
    }
}

