namespace ExcelToXML
{
    partial class Main
    {
        /// <summary>
        /// Требуется переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Освободить все используемые ресурсы.
        /// </summary>
        /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Код, автоматически созданный конструктором форм Windows

        /// <summary>
        /// Обязательный метод для поддержки конструктора - не изменяйте
        /// содержимое данного метода при помощи редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
            this.txtExcelFile = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.btExcelChoose = new System.Windows.Forms.Button();
            this.btTransfer = new System.Windows.Forms.Button();
            this.ofdExcelFile = new System.Windows.Forms.OpenFileDialog();
            this.sfdXMLFile = new System.Windows.Forms.SaveFileDialog();
            this.label2 = new System.Windows.Forms.Label();
            this.txtXMLFile = new System.Windows.Forms.TextBox();
            this.btXMLChoose = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // txtExcelFile
            // 
            this.txtExcelFile.Location = new System.Drawing.Point(12, 26);
            this.txtExcelFile.Name = "txtExcelFile";
            this.txtExcelFile.Size = new System.Drawing.Size(268, 20);
            this.txtExcelFile.TabIndex = 0;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(36, 13);
            this.label1.TabIndex = 1;
            this.label1.Text = "Файл";
            // 
            // btExcelChoose
            // 
            this.btExcelChoose.Location = new System.Drawing.Point(289, 26);
            this.btExcelChoose.Name = "btExcelChoose";
            this.btExcelChoose.Size = new System.Drawing.Size(27, 20);
            this.btExcelChoose.TabIndex = 2;
            this.btExcelChoose.Text = "...";
            this.btExcelChoose.UseVisualStyleBackColor = true;
            this.btExcelChoose.Click += new System.EventHandler(this.btExcelChoose_Click);
            // 
            // btTransfer
            // 
            this.btTransfer.Location = new System.Drawing.Point(12, 116);
            this.btTransfer.Name = "btTransfer";
            this.btTransfer.Size = new System.Drawing.Size(98, 23);
            this.btTransfer.TabIndex = 3;
            this.btTransfer.Text = "Преобразовать";
            this.btTransfer.UseVisualStyleBackColor = true;
            this.btTransfer.Click += new System.EventHandler(this.btTransfer_Click);
            // 
            // ofdExcelFile
            // 
            this.ofdExcelFile.DefaultExt = "Excel";
            this.ofdExcelFile.Filter = "Excel (*.xl*)|*.xl*|All files (*.*)|*.*";
            // 
            // sfdXMLFile
            // 
            this.sfdXMLFile.DefaultExt = "XML";
            this.sfdXMLFile.Filter = "XML|*.xml";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(12, 62);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(86, 13);
            this.label2.TabIndex = 4;
            this.label2.Text = "Преобразовать";
            // 
            // txtXMLFile
            // 
            this.txtXMLFile.Location = new System.Drawing.Point(12, 78);
            this.txtXMLFile.Name = "txtXMLFile";
            this.txtXMLFile.Size = new System.Drawing.Size(268, 20);
            this.txtXMLFile.TabIndex = 5;
            // 
            // btXMLChoose
            // 
            this.btXMLChoose.Location = new System.Drawing.Point(289, 78);
            this.btXMLChoose.Name = "btXMLChoose";
            this.btXMLChoose.Size = new System.Drawing.Size(27, 20);
            this.btXMLChoose.TabIndex = 6;
            this.btXMLChoose.Text = "...";
            this.btXMLChoose.UseVisualStyleBackColor = true;
            this.btXMLChoose.Click += new System.EventHandler(this.btXMLChoose_Click);
            // 
            // Main
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(319, 273);
            this.Controls.Add(this.btXMLChoose);
            this.Controls.Add(this.txtXMLFile);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.btTransfer);
            this.Controls.Add(this.btExcelChoose);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.txtExcelFile);
            this.Name = "Main";
            this.Text = "Экспорт в XML";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox txtExcelFile;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btExcelChoose;
        private System.Windows.Forms.Button btTransfer;
        private System.Windows.Forms.OpenFileDialog ofdExcelFile;
        private System.Windows.Forms.SaveFileDialog sfdXMLFile;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtXMLFile;
        private System.Windows.Forms.Button btXMLChoose;
    }
}

