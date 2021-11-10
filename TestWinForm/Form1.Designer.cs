
namespace TestWinForm
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
            this.btn_to_base64 = new System.Windows.Forms.Button();
            this.richTextBox1 = new System.Windows.Forms.RichTextBox();
            this.btn_send_file = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // btn_to_base64
            // 
            this.btn_to_base64.Location = new System.Drawing.Point(135, 58);
            this.btn_to_base64.Name = "btn_to_base64";
            this.btn_to_base64.Size = new System.Drawing.Size(180, 23);
            this.btn_to_base64.TabIndex = 0;
            this.btn_to_base64.Text = "Convert to base64";
            this.btn_to_base64.UseVisualStyleBackColor = true;
            this.btn_to_base64.Click += new System.EventHandler(this.btn_to_base64_Click);
            // 
            // richTextBox1
            // 
            this.richTextBox1.Location = new System.Drawing.Point(56, 125);
            this.richTextBox1.Name = "richTextBox1";
            this.richTextBox1.Size = new System.Drawing.Size(737, 275);
            this.richTextBox1.TabIndex = 1;
            this.richTextBox1.Text = "";
            // 
            // btn_send_file
            // 
            this.btn_send_file.Location = new System.Drawing.Point(334, 58);
            this.btn_send_file.Name = "btn_send_file";
            this.btn_send_file.Size = new System.Drawing.Size(180, 23);
            this.btn_send_file.TabIndex = 2;
            this.btn_send_file.Text = "Send File";
            this.btn_send_file.UseVisualStyleBackColor = true;
            this.btn_send_file.Click += new System.EventHandler(this.btn_send_file_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.btn_send_file);
            this.Controls.Add(this.richTextBox1);
            this.Controls.Add(this.btn_to_base64);
            this.Name = "Form1";
            this.Text = "Form1";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button btn_to_base64;
        private System.Windows.Forms.RichTextBox richTextBox1;
        private System.Windows.Forms.Button btn_send_file;
    }
}

