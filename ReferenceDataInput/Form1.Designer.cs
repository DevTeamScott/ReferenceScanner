namespace ReferenceDataInput
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
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.panel2 = new System.Windows.Forms.Panel();
            this.Output = new System.Windows.Forms.Label();
            this.DtOutput = new System.Windows.Forms.Button();
            this.ExlOutput = new System.Windows.Forms.Button();
            this.CSVOutput = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.panel2.SuspendLayout();
            this.SuspendLayout();
            // 
            // dataGridView1
            // 
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Location = new System.Drawing.Point(7, 7);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.Size = new System.Drawing.Size(701, 404);
            this.dataGridView1.TabIndex = 0;
            // 
            // panel2
            // 
            this.panel2.Controls.Add(this.Output);
            this.panel2.Controls.Add(this.DtOutput);
            this.panel2.Controls.Add(this.ExlOutput);
            this.panel2.Controls.Add(this.CSVOutput);
            this.panel2.Location = new System.Drawing.Point(7, 417);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(700, 117);
            this.panel2.TabIndex = 3;
            // 
            // Output
            // 
            this.Output.AutoSize = true;
            this.Output.Location = new System.Drawing.Point(11, 11);
            this.Output.Name = "Output";
            this.Output.Size = new System.Drawing.Size(39, 13);
            this.Output.TabIndex = 3;
            this.Output.Text = "Output";
            // 
            // DtOutput
            // 
            this.DtOutput.Location = new System.Drawing.Point(383, 31);
            this.DtOutput.Name = "DtOutput";
            this.DtOutput.Size = new System.Drawing.Size(134, 44);
            this.DtOutput.TabIndex = 2;
            this.DtOutput.Text = "Datatable";
            this.DtOutput.UseVisualStyleBackColor = true;
            this.DtOutput.Click += new System.EventHandler(this.DtOutput_Click);
            // 
            // ExlOutput
            // 
            this.ExlOutput.Location = new System.Drawing.Point(207, 31);
            this.ExlOutput.Name = "ExlOutput";
            this.ExlOutput.Size = new System.Drawing.Size(134, 44);
            this.ExlOutput.TabIndex = 1;
            this.ExlOutput.Text = "Excel";
            this.ExlOutput.UseVisualStyleBackColor = true;
            this.ExlOutput.Click += new System.EventHandler(this.ExlOutput_Click);
            // 
            // CSVOutput
            // 
            this.CSVOutput.Location = new System.Drawing.Point(34, 31);
            this.CSVOutput.Name = "CSVOutput";
            this.CSVOutput.Size = new System.Drawing.Size(134, 44);
            this.CSVOutput.TabIndex = 0;
            this.CSVOutput.Text = "CSV";
            this.CSVOutput.UseVisualStyleBackColor = true;
            this.CSVOutput.Click += new System.EventHandler(this.CSVOutput_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(715, 535);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.dataGridView1);
            this.Name = "Form1";
            this.Text = "Form1";
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.panel2.ResumeLayout(false);
            this.panel2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.Label Output;
        private System.Windows.Forms.Button DtOutput;
        private System.Windows.Forms.Button ExlOutput;
        private System.Windows.Forms.Button CSVOutput;
    }
}

