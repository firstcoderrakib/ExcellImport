namespace ImportExcellExe
{
    partial class Form1
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
            this.textBoxCompanyName = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.CompanyIDlistBox = new System.Windows.Forms.ListBox();
            this.FileTextBox = new System.Windows.Forms.TextBox();
            this.buttonSelectFile = new System.Windows.Forms.Button();
            this.buttonResetMessage = new System.Windows.Forms.Button();
            this.buttonImportHSCode = new System.Windows.Forms.Button();
            this.buttonCostLedger = new System.Windows.Forms.Button();
            this.buttonImportServiceCode = new System.Windows.Forms.Button();
            this.buttonMeasurement = new System.Windows.Forms.Button();
            this.buttonInputOutCoEfficient = new System.Windows.Forms.Button();
            this.buttonItemName = new System.Windows.Forms.Button();
            this.buttonCustomer = new System.Windows.Forms.Button();
            this.buttonSupplier = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // textBoxCompanyName
            // 
            this.textBoxCompanyName.Location = new System.Drawing.Point(164, 30);
            this.textBoxCompanyName.Name = "textBoxCompanyName";
            this.textBoxCompanyName.Size = new System.Drawing.Size(293, 20);
            this.textBoxCompanyName.TabIndex = 0;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(15, 25);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(122, 19);
            this.label1.TabIndex = 1;
            this.label1.Text = "Company Name";
            // 
            // CompanyIDlistBox
            // 
            this.CompanyIDlistBox.FormattingEnabled = true;
            this.CompanyIDlistBox.Location = new System.Drawing.Point(164, 51);
            this.CompanyIDlistBox.Name = "CompanyIDlistBox";
            this.CompanyIDlistBox.Size = new System.Drawing.Size(293, 95);
            this.CompanyIDlistBox.TabIndex = 2;
            // 
            // FileTextBox
            // 
            this.FileTextBox.Location = new System.Drawing.Point(42, 152);
            this.FileTextBox.Name = "FileTextBox";
            this.FileTextBox.Size = new System.Drawing.Size(415, 20);
            this.FileTextBox.TabIndex = 3;
            // 
            // buttonSelectFile
            // 
            this.buttonSelectFile.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.buttonSelectFile.Location = new System.Drawing.Point(474, 143);
            this.buttonSelectFile.Name = "buttonSelectFile";
            this.buttonSelectFile.Size = new System.Drawing.Size(184, 33);
            this.buttonSelectFile.TabIndex = 4;
            this.buttonSelectFile.Text = "Select File";
            this.buttonSelectFile.UseVisualStyleBackColor = true;
            this.buttonSelectFile.Click += new System.EventHandler(this.buttonSelectFile_Click);
            // 
            // buttonResetMessage
            // 
            this.buttonResetMessage.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.buttonResetMessage.Location = new System.Drawing.Point(692, 137);
            this.buttonResetMessage.Name = "buttonResetMessage";
            this.buttonResetMessage.Size = new System.Drawing.Size(181, 44);
            this.buttonResetMessage.TabIndex = 5;
            this.buttonResetMessage.Text = "Reset  Message";
            this.buttonResetMessage.UseVisualStyleBackColor = true;
            // 
            // buttonImportHSCode
            // 
            this.buttonImportHSCode.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.buttonImportHSCode.Location = new System.Drawing.Point(474, 188);
            this.buttonImportHSCode.Name = "buttonImportHSCode";
            this.buttonImportHSCode.Size = new System.Drawing.Size(184, 44);
            this.buttonImportHSCode.TabIndex = 6;
            this.buttonImportHSCode.Text = "Import HS Code";
            this.buttonImportHSCode.UseVisualStyleBackColor = true;
            // 
            // buttonCostLedger
            // 
            this.buttonCostLedger.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.buttonCostLedger.Location = new System.Drawing.Point(692, 188);
            this.buttonCostLedger.Name = "buttonCostLedger";
            this.buttonCostLedger.Size = new System.Drawing.Size(181, 44);
            this.buttonCostLedger.TabIndex = 7;
            this.buttonCostLedger.Text = "Cost Ledger";
            this.buttonCostLedger.UseVisualStyleBackColor = true;
            // 
            // buttonImportServiceCode
            // 
            this.buttonImportServiceCode.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.buttonImportServiceCode.Location = new System.Drawing.Point(474, 238);
            this.buttonImportServiceCode.Name = "buttonImportServiceCode";
            this.buttonImportServiceCode.Size = new System.Drawing.Size(184, 44);
            this.buttonImportServiceCode.TabIndex = 8;
            this.buttonImportServiceCode.Text = "Import Service Code";
            this.buttonImportServiceCode.UseVisualStyleBackColor = true;
            // 
            // buttonMeasurement
            // 
            this.buttonMeasurement.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.buttonMeasurement.Location = new System.Drawing.Point(692, 238);
            this.buttonMeasurement.Name = "buttonMeasurement";
            this.buttonMeasurement.Size = new System.Drawing.Size(181, 54);
            this.buttonMeasurement.TabIndex = 9;
            this.buttonMeasurement.Text = "Measurement Unit";
            this.buttonMeasurement.UseVisualStyleBackColor = true;
            // 
            // buttonInputOutCoEfficient
            // 
            this.buttonInputOutCoEfficient.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.buttonInputOutCoEfficient.Location = new System.Drawing.Point(474, 288);
            this.buttonInputOutCoEfficient.Name = "buttonInputOutCoEfficient";
            this.buttonInputOutCoEfficient.Size = new System.Drawing.Size(184, 56);
            this.buttonInputOutCoEfficient.TabIndex = 10;
            this.buttonInputOutCoEfficient.Text = "Input Output Co-Efficient";
            this.buttonInputOutCoEfficient.UseVisualStyleBackColor = true;
            // 
            // buttonItemName
            // 
            this.buttonItemName.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.buttonItemName.Location = new System.Drawing.Point(692, 298);
            this.buttonItemName.Name = "buttonItemName";
            this.buttonItemName.Size = new System.Drawing.Size(181, 46);
            this.buttonItemName.TabIndex = 11;
            this.buttonItemName.Text = "Item Name";
            this.buttonItemName.UseVisualStyleBackColor = true;
            // 
            // buttonCustomer
            // 
            this.buttonCustomer.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.buttonCustomer.Location = new System.Drawing.Point(474, 350);
            this.buttonCustomer.Name = "buttonCustomer";
            this.buttonCustomer.Size = new System.Drawing.Size(184, 44);
            this.buttonCustomer.TabIndex = 12;
            this.buttonCustomer.Text = "Customer";
            this.buttonCustomer.UseVisualStyleBackColor = true;
            // 
            // buttonSupplier
            // 
            this.buttonSupplier.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.buttonSupplier.Location = new System.Drawing.Point(692, 350);
            this.buttonSupplier.Name = "buttonSupplier";
            this.buttonSupplier.Size = new System.Drawing.Size(181, 44);
            this.buttonSupplier.TabIndex = 13;
            this.buttonSupplier.Text = "Supplier";
            this.buttonSupplier.UseVisualStyleBackColor = true;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(901, 450);
            this.Controls.Add(this.buttonSupplier);
            this.Controls.Add(this.buttonCustomer);
            this.Controls.Add(this.buttonItemName);
            this.Controls.Add(this.buttonInputOutCoEfficient);
            this.Controls.Add(this.buttonMeasurement);
            this.Controls.Add(this.buttonImportServiceCode);
            this.Controls.Add(this.buttonCostLedger);
            this.Controls.Add(this.buttonImportHSCode);
            this.Controls.Add(this.buttonResetMessage);
            this.Controls.Add(this.buttonSelectFile);
            this.Controls.Add(this.FileTextBox);
            this.Controls.Add(this.CompanyIDlistBox);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.textBoxCompanyName);
            this.Name = "Form1";
            this.Text = "Form1";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox textBoxCompanyName;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ListBox CompanyIDlistBox;
        private System.Windows.Forms.TextBox FileTextBox;
        private System.Windows.Forms.Button buttonSelectFile;
        private System.Windows.Forms.Button buttonResetMessage;
        private System.Windows.Forms.Button buttonImportHSCode;
        private System.Windows.Forms.Button buttonCostLedger;
        private System.Windows.Forms.Button buttonImportServiceCode;
        private System.Windows.Forms.Button buttonMeasurement;
        private System.Windows.Forms.Button buttonInputOutCoEfficient;
        private System.Windows.Forms.Button buttonItemName;
        private System.Windows.Forms.Button buttonCustomer;
        private System.Windows.Forms.Button buttonSupplier;
    }
}

