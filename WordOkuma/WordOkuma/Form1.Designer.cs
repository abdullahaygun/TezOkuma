
namespace WordOkuma
{
    partial class Form1
    {
        /// <summary>
        ///Gerekli tasarımcı değişkeni.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///Kullanılan tüm kaynakları temizleyin.
        /// </summary>
        ///<param name="disposing">yönetilen kaynaklar dispose edilmeliyse doğru; aksi halde yanlış.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer üretilen kod

        /// <summary>
        /// Tasarımcı desteği için gerekli metot - bu metodun 
        ///içeriğini kod düzenleyici ile değiştirmeyin.
        /// </summary>
        private void InitializeComponent()
        {
            this.richTextBox1 = new System.Windows.Forms.RichTextBox();
            this.richTextBox2 = new System.Windows.Forms.RichTextBox();
            this.btnBrowse = new System.Windows.Forms.Button();
            this.btnOku = new System.Windows.Forms.Button();
            this.btnPageMarginControl = new System.Windows.Forms.Button();
            this.btnBaslikKontrol = new System.Windows.Forms.Button();
            this.pBar = new System.Windows.Forms.ProgressBar();
            this.SuspendLayout();
            // 
            // richTextBox1
            // 
            this.richTextBox1.Location = new System.Drawing.Point(4, 4);
            this.richTextBox1.Name = "richTextBox1";
            this.richTextBox1.Size = new System.Drawing.Size(873, 515);
            this.richTextBox1.TabIndex = 0;
            this.richTextBox1.Text = "";
            // 
            // richTextBox2
            // 
            this.richTextBox2.Location = new System.Drawing.Point(883, 12);
            this.richTextBox2.Name = "richTextBox2";
            this.richTextBox2.Size = new System.Drawing.Size(230, 415);
            this.richTextBox2.TabIndex = 1;
            this.richTextBox2.Text = "";
            // 
            // btnBrowse
            // 
            this.btnBrowse.Location = new System.Drawing.Point(887, 515);
            this.btnBrowse.Name = "btnBrowse";
            this.btnBrowse.Size = new System.Drawing.Size(96, 35);
            this.btnBrowse.TabIndex = 2;
            this.btnBrowse.Text = "Göz At";
            this.btnBrowse.UseVisualStyleBackColor = true;
            this.btnBrowse.Click += new System.EventHandler(this.btnBrowse_Click);
            // 
            // btnOku
            // 
            this.btnOku.Location = new System.Drawing.Point(995, 515);
            this.btnOku.Name = "btnOku";
            this.btnOku.Size = new System.Drawing.Size(122, 35);
            this.btnOku.TabIndex = 3;
            this.btnOku.Text = "Oku";
            this.btnOku.UseVisualStyleBackColor = true;
            this.btnOku.Click += new System.EventHandler(this.btnOku_Click);
            // 
            // btnPageMarginControl
            // 
            this.btnPageMarginControl.Location = new System.Drawing.Point(887, 474);
            this.btnPageMarginControl.Name = "btnPageMarginControl";
            this.btnPageMarginControl.Size = new System.Drawing.Size(230, 35);
            this.btnPageMarginControl.TabIndex = 4;
            this.btnPageMarginControl.Text = "Sayfa Kenar Boşluğu";
            this.btnPageMarginControl.UseVisualStyleBackColor = true;
            this.btnPageMarginControl.Click += new System.EventHandler(this.btnPageMarginControl_Click);
            // 
            // btnBaslikKontrol
            // 
            this.btnBaslikKontrol.Location = new System.Drawing.Point(995, 433);
            this.btnBaslikKontrol.Name = "btnBaslikKontrol";
            this.btnBaslikKontrol.Size = new System.Drawing.Size(122, 35);
            this.btnBaslikKontrol.TabIndex = 5;
            this.btnBaslikKontrol.Text = "Başlıklar";
            this.btnBaslikKontrol.UseVisualStyleBackColor = true;
            this.btnBaslikKontrol.Click += new System.EventHandler(this.btnBaslikKontrol_Click);
            // 
            // pBar
            // 
            this.pBar.Location = new System.Drawing.Point(4, 525);
            this.pBar.Name = "pBar";
            this.pBar.Size = new System.Drawing.Size(873, 23);
            this.pBar.Step = 1;
            this.pBar.TabIndex = 6;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1129, 560);
            this.Controls.Add(this.pBar);
            this.Controls.Add(this.btnBaslikKontrol);
            this.Controls.Add(this.btnPageMarginControl);
            this.Controls.Add(this.btnOku);
            this.Controls.Add(this.btnBrowse);
            this.Controls.Add(this.richTextBox2);
            this.Controls.Add(this.richTextBox1);
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Form1";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.Form1_FormClosing);
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.RichTextBox richTextBox1;
        private System.Windows.Forms.RichTextBox richTextBox2;
        private System.Windows.Forms.Button btnBrowse;
        private System.Windows.Forms.Button btnOku;
        private System.Windows.Forms.Button btnPageMarginControl;
        private System.Windows.Forms.Button btnBaslikKontrol;
        private System.Windows.Forms.ProgressBar pBar;
    }
}

