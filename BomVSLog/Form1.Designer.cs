namespace BomVSLog
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
            this.label3 = new System.Windows.Forms.Label();
            this.SelBomB = new System.Windows.Forms.Button();
            this.Select_BomTB = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.SelLogB = new System.Windows.Forms.Button();
            this.Select_LogTB = new System.Windows.Forms.TextBox();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.ExitB = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.SelTOB = new System.Windows.Forms.Button();
            this.Sel_TOTB = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.StatusL = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            this.label5 = new System.Windows.Forms.Label();
            this.SelTPB = new System.Windows.Forms.Button();
            this.Sel_TPTB = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(18, 133);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(82, 16);
            this.label3.TabIndex = 13;
            this.label3.Text = "Select BOM:";
            // 
            // SelBomB
            // 
            this.SelBomB.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.SelBomB.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.SelBomB.Location = new System.Drawing.Point(562, 152);
            this.SelBomB.Name = "SelBomB";
            this.SelBomB.Size = new System.Drawing.Size(96, 26);
            this.SelBomB.TabIndex = 12;
            this.SelBomB.Text = "Browse..";
            this.SelBomB.UseVisualStyleBackColor = true;
            this.SelBomB.Click += new System.EventHandler(this.SelBomB_Click);
            // 
            // Select_BomTB
            // 
            this.Select_BomTB.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Select_BomTB.Location = new System.Drawing.Point(20, 152);
            this.Select_BomTB.Name = "Select_BomTB";
            this.Select_BomTB.Size = new System.Drawing.Size(536, 26);
            this.Select_BomTB.TabIndex = 11;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(18, 195);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(75, 16);
            this.label1.TabIndex = 16;
            this.label1.Text = "Select Log:";
            this.label1.Click += new System.EventHandler(this.label1_Click);
            // 
            // SelLogB
            // 
            this.SelLogB.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.SelLogB.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.SelLogB.Location = new System.Drawing.Point(562, 214);
            this.SelLogB.Name = "SelLogB";
            this.SelLogB.Size = new System.Drawing.Size(96, 26);
            this.SelLogB.TabIndex = 15;
            this.SelLogB.Text = "Browse..";
            this.SelLogB.UseVisualStyleBackColor = true;
            this.SelLogB.Click += new System.EventHandler(this.SelLogB_Click);
            // 
            // Select_LogTB
            // 
            this.Select_LogTB.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Select_LogTB.Location = new System.Drawing.Point(20, 214);
            this.Select_LogTB.Name = "Select_LogTB";
            this.Select_LogTB.Size = new System.Drawing.Size(536, 26);
            this.Select_LogTB.TabIndex = 14;
            this.Select_LogTB.TextChanged += new System.EventHandler(this.Select_LogTB_TextChanged);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.AddExtension = false;
            this.openFileDialog1.ValidateNames = false;
            // 
            // ExitB
            // 
            this.ExitB.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.ExitB.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.ExitB.Location = new System.Drawing.Point(167, 273);
            this.ExitB.Name = "ExitB";
            this.ExitB.Size = new System.Drawing.Size(96, 26);
            this.ExitB.TabIndex = 18;
            this.ExitB.Text = "Exit";
            this.ExitB.UseVisualStyleBackColor = true;
            this.ExitB.Click += new System.EventHandler(this.ExitB_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(16, 20);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(111, 16);
            this.label2.TabIndex = 21;
            this.label2.Text = "Select Testorder:";
            // 
            // SelTOB
            // 
            this.SelTOB.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.SelTOB.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.SelTOB.Location = new System.Drawing.Point(560, 39);
            this.SelTOB.Name = "SelTOB";
            this.SelTOB.Size = new System.Drawing.Size(96, 26);
            this.SelTOB.TabIndex = 20;
            this.SelTOB.Text = "Browse..";
            this.SelTOB.UseVisualStyleBackColor = true;
            this.SelTOB.Click += new System.EventHandler(this.SelTOB_Click);
            // 
            // Sel_TOTB
            // 
            this.Sel_TOTB.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Sel_TOTB.Location = new System.Drawing.Point(18, 39);
            this.Sel_TOTB.Name = "Sel_TOTB";
            this.Sel_TOTB.Size = new System.Drawing.Size(536, 26);
            this.Sel_TOTB.TabIndex = 19;
            this.Sel_TOTB.TextChanged += new System.EventHandler(this.Sel_TOTB_TextChanged);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.Location = new System.Drawing.Point(17, 319);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(48, 16);
            this.label4.TabIndex = 22;
            this.label4.Text = "Status:";
            // 
            // StatusL
            // 
            this.StatusL.AutoSize = true;
            this.StatusL.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.StatusL.Location = new System.Drawing.Point(29, 347);
            this.StatusL.Name = "StatusL";
            this.StatusL.Size = new System.Drawing.Size(0, 16);
            this.StatusL.TabIndex = 23;
            // 
            // button1
            // 
            this.button1.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button1.Location = new System.Drawing.Point(18, 273);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(131, 26);
            this.button1.TabIndex = 24;
            this.button1.Text = "Extract";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label5.Location = new System.Drawing.Point(18, 77);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(105, 16);
            this.label5.TabIndex = 27;
            this.label5.Text = "Select Testplan:";
            // 
            // SelTPB
            // 
            this.SelTPB.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.SelTPB.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.SelTPB.Location = new System.Drawing.Point(562, 96);
            this.SelTPB.Name = "SelTPB";
            this.SelTPB.Size = new System.Drawing.Size(96, 26);
            this.SelTPB.TabIndex = 26;
            this.SelTPB.Text = "Browse..";
            this.SelTPB.UseVisualStyleBackColor = true;
            this.SelTPB.Click += new System.EventHandler(this.SelTPB_Click);
            // 
            // Sel_TPTB
            // 
            this.Sel_TPTB.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Sel_TPTB.Location = new System.Drawing.Point(20, 96);
            this.Sel_TPTB.Name = "Sel_TPTB";
            this.Sel_TPTB.Size = new System.Drawing.Size(536, 26);
            this.Sel_TPTB.TabIndex = 25;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(681, 370);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.SelTPB);
            this.Controls.Add(this.Sel_TPTB);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.StatusL);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.SelTOB);
            this.Controls.Add(this.Sel_TOTB);
            this.Controls.Add(this.ExitB);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.SelLogB);
            this.Controls.Add(this.Select_LogTB);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.SelBomB);
            this.Controls.Add(this.Select_BomTB);
            this.Name = "Form1";
            this.Text = "ICT Program Verification";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button SelBomB;
        private System.Windows.Forms.TextBox Select_BomTB;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button SelLogB;
        private System.Windows.Forms.TextBox Select_LogTB;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.Button ExitB;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button SelTOB;
        private System.Windows.Forms.TextBox Sel_TOTB;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label StatusL;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Button SelTPB;
        private System.Windows.Forms.TextBox Sel_TPTB;
    }
}

