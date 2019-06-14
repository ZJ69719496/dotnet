namespace CraftImport
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
            this.LoadBtn = new System.Windows.Forms.Button();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.comboBox1 = new System.Windows.Forms.ComboBox();
            this.butSql = new System.Windows.Forms.Button();
            this.chkOperation = new System.Windows.Forms.CheckBox();
            this.chkTool = new System.Windows.Forms.CheckBox();
            this.chkWorkContent = new System.Windows.Forms.CheckBox();
            this.SuspendLayout();
            // 
            // LoadBtn
            // 
            this.LoadBtn.Location = new System.Drawing.Point(14, 17);
            this.LoadBtn.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.LoadBtn.Name = "LoadBtn";
            this.LoadBtn.Size = new System.Drawing.Size(87, 26);
            this.LoadBtn.TabIndex = 0;
            this.LoadBtn.Text = "加载Excel";
            this.LoadBtn.UseVisualStyleBackColor = true;
            this.LoadBtn.Click += new System.EventHandler(this.LoadBtn_Click);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // comboBox1
            // 
            this.comboBox1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBox1.FormattingEnabled = true;
            this.comboBox1.Location = new System.Drawing.Point(112, 18);
            this.comboBox1.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.comboBox1.Name = "comboBox1";
            this.comboBox1.Size = new System.Drawing.Size(194, 25);
            this.comboBox1.TabIndex = 3;
            // 
            // butSql
            // 
            this.butSql.Location = new System.Drawing.Point(320, 17);
            this.butSql.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.butSql.Name = "butSql";
            this.butSql.Size = new System.Drawing.Size(87, 26);
            this.butSql.TabIndex = 4;
            this.butSql.Text = "生成SQL";
            this.butSql.UseVisualStyleBackColor = true;
            this.butSql.Click += new System.EventHandler(this.BtnSql_Click);
            // 
            // checkBox1
            // 
            this.chkOperation.AutoSize = true;
            this.chkOperation.Checked = true;
            this.chkOperation.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkOperation.Location = new System.Drawing.Point(44, 66);
            this.chkOperation.Name = "checkBox1";
            this.chkOperation.Size = new System.Drawing.Size(75, 21);
            this.chkOperation.TabIndex = 5;
            this.chkOperation.Text = "装配零件";
            this.chkOperation.UseVisualStyleBackColor = true;
            // 
            // checkBox2
            // 
            this.chkTool.AutoSize = true;
            this.chkTool.Checked = true;
            this.chkTool.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkTool.Location = new System.Drawing.Point(161, 66);
            this.chkTool.Name = "checkBox2";
            this.chkTool.Size = new System.Drawing.Size(75, 21);
            this.chkTool.TabIndex = 6;
            this.chkTool.Text = "使用工具";
            this.chkTool.UseVisualStyleBackColor = true;
            // 
            // checkBox3
            // 
            this.chkWorkContent.AutoSize = true;
            this.chkWorkContent.Checked = true;
            this.chkWorkContent.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkWorkContent.Location = new System.Drawing.Point(283, 66);
            this.chkWorkContent.Name = "checkBox3";
            this.chkWorkContent.Size = new System.Drawing.Size(75, 21);
            this.chkWorkContent.TabIndex = 7;
            this.chkWorkContent.Text = "工序内容";
            this.chkWorkContent.UseVisualStyleBackColor = true;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(422, 113);
            this.Controls.Add(this.chkWorkContent);
            this.Controls.Add(this.chkTool);
            this.Controls.Add(this.chkOperation);
            this.Controls.Add(this.butSql);
            this.Controls.Add(this.comboBox1);
            this.Controls.Add(this.LoadBtn);
            this.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "工艺导入";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button LoadBtn;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.ComboBox comboBox1;
        private System.Windows.Forms.Button butSql;
        private System.Windows.Forms.CheckBox chkOperation;
        private System.Windows.Forms.CheckBox chkTool;
        private System.Windows.Forms.CheckBox chkWorkContent;
    }
}

