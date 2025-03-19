namespace Sistema_de_mail_para_Bridgestone___Thalamus
{
    partial class Form1
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
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
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            selecArchBtn = new Button();
            selecFcsBtn = new Button();
            panel1 = new Panel();
            label1 = new Label();
            pictureBox2 = new PictureBox();
            button1 = new Button();
            pictureBox1 = new PictureBox();
            button2 = new Button();
            button3 = new Button();
            panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)pictureBox2).BeginInit();
            ((System.ComponentModel.ISupportInitialize)pictureBox1).BeginInit();
            SuspendLayout();
            // 
            // selecArchBtn
            // 
            selecArchBtn.BackColor = Color.FromArgb(3, 77, 119);
            selecArchBtn.FlatAppearance.BorderSize = 0;
            selecArchBtn.FlatStyle = FlatStyle.Flat;
            selecArchBtn.Font = new Font("Segoe UI Semibold", 9.75F, FontStyle.Bold);
            selecArchBtn.ForeColor = SystemColors.ButtonHighlight;
            selecArchBtn.Location = new Point(43, 79);
            selecArchBtn.Name = "selecArchBtn";
            selecArchBtn.Size = new Size(135, 26);
            selecArchBtn.TabIndex = 43;
            selecArchBtn.Text = "Subir excel";
            selecArchBtn.UseVisualStyleBackColor = false;
            selecArchBtn.Click += selecArchBtn_Click;
            // 
            // selecFcsBtn
            // 
            selecFcsBtn.BackColor = Color.FromArgb(3, 77, 119);
            selecFcsBtn.FlatAppearance.BorderSize = 0;
            selecFcsBtn.FlatStyle = FlatStyle.Flat;
            selecFcsBtn.Font = new Font("Segoe UI Semibold", 9.75F, FontStyle.Bold);
            selecFcsBtn.ForeColor = SystemColors.ButtonHighlight;
            selecFcsBtn.Location = new Point(43, 179);
            selecFcsBtn.Name = "selecFcsBtn";
            selecFcsBtn.Size = new Size(135, 26);
            selecFcsBtn.TabIndex = 42;
            selecFcsBtn.Text = "Subida de facturas";
            selecFcsBtn.UseVisualStyleBackColor = false;
            selecFcsBtn.Click += selecFcsBtn_Click;
            // 
            // panel1
            // 
            panel1.BackColor = Color.FromArgb(3, 77, 119);
            panel1.Controls.Add(label1);
            panel1.Controls.Add(pictureBox2);
            panel1.Controls.Add(button1);
            panel1.Dock = DockStyle.Top;
            panel1.Location = new Point(0, 0);
            panel1.Name = "panel1";
            panel1.Size = new Size(788, 27);
            panel1.TabIndex = 45;
            panel1.DoubleClick += panel1_DoubleClick;
            panel1.MouseDown += panel1_MouseDown;
            panel1.MouseMove += panel1_MouseMove;
            panel1.MouseUp += panel1_MouseUp;
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Font = new Font("Segoe UI Semibold", 9.75F, FontStyle.Bold);
            label1.ForeColor = SystemColors.ButtonHighlight;
            label1.Location = new Point(23, 4);
            label1.Name = "label1";
            label1.Size = new Size(180, 17);
            label1.TabIndex = 2;
            label1.Text = "Sistema de mails - Thalamus";
            // 
            // pictureBox2
            // 
            pictureBox2.BackColor = Color.Transparent;
            pictureBox2.Image = (Image)resources.GetObject("pictureBox2.Image");
            pictureBox2.Location = new Point(4, 4);
            pictureBox2.Name = "pictureBox2";
            pictureBox2.Size = new Size(18, 18);
            pictureBox2.SizeMode = PictureBoxSizeMode.Zoom;
            pictureBox2.TabIndex = 1;
            pictureBox2.TabStop = false;
            // 
            // button1
            // 
            button1.BackgroundImage = (Image)resources.GetObject("button1.BackgroundImage");
            button1.BackgroundImageLayout = ImageLayout.Zoom;
            button1.Dock = DockStyle.Right;
            button1.FlatAppearance.BorderSize = 0;
            button1.FlatStyle = FlatStyle.Flat;
            button1.Location = new Point(760, 0);
            button1.Name = "button1";
            button1.Size = new Size(28, 27);
            button1.TabIndex = 0;
            button1.UseVisualStyleBackColor = true;
            button1.Click += button1_Click;
            // 
            // pictureBox1
            // 
            pictureBox1.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            pictureBox1.Image = (Image)resources.GetObject("pictureBox1.Image");
            pictureBox1.Location = new Point(-133, 24);
            pictureBox1.Name = "pictureBox1";
            pictureBox1.Size = new Size(921, 444);
            pictureBox1.SizeMode = PictureBoxSizeMode.StretchImage;
            pictureBox1.TabIndex = 46;
            pictureBox1.TabStop = false;
            // 
            // button2
            // 
            button2.BackColor = Color.FromArgb(3, 77, 119);
            button2.FlatAppearance.BorderSize = 0;
            button2.FlatStyle = FlatStyle.Flat;
            button2.Font = new Font("Segoe UI Semibold", 9.75F, FontStyle.Bold);
            button2.ForeColor = SystemColors.ButtonHighlight;
            button2.Location = new Point(43, 279);
            button2.Name = "button2";
            button2.Size = new Size(135, 26);
            button2.TabIndex = 47;
            button2.Text = "Subir excel";
            button2.UseVisualStyleBackColor = false;
            // 
            // button3
            // 
            button3.BackColor = Color.FromArgb(3, 77, 119);
            button3.FlatAppearance.BorderSize = 0;
            button3.FlatStyle = FlatStyle.Flat;
            button3.Font = new Font("Segoe UI Semibold", 9.75F, FontStyle.Bold);
            button3.ForeColor = SystemColors.ButtonHighlight;
            button3.Location = new Point(43, 379);
            button3.Name = "button3";
            button3.Size = new Size(135, 26);
            button3.TabIndex = 48;
            button3.Text = "Subir excel";
            button3.UseVisualStyleBackColor = false;
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(788, 467);
            Controls.Add(button3);
            Controls.Add(button2);
            Controls.Add(panel1);
            Controls.Add(selecArchBtn);
            Controls.Add(selecFcsBtn);
            Controls.Add(pictureBox1);
            Font = new Font("Segoe UI", 9F, FontStyle.Regular, GraphicsUnit.Point, 0);
            FormBorderStyle = FormBorderStyle.None;
            Icon = (Icon)resources.GetObject("$this.Icon");
            Name = "Form1";
            RightToLeft = RightToLeft.No;
            Text = "Sistema de mails a Bridgestone - Thalamus";
            Load += Form1_Load;
            panel1.ResumeLayout(false);
            panel1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)pictureBox2).EndInit();
            ((System.ComponentModel.ISupportInitialize)pictureBox1).EndInit();
            ResumeLayout(false);
        }

        #endregion
        private Button selecArchBtn;
        private Button selecFcsBtn;
        private Panel panel1;
        private Button button1;
        private PictureBox pictureBox1;
        private PictureBox pictureBox2;
        private Label label1;
        private Button button2;
        private Button button3;
    }
}
