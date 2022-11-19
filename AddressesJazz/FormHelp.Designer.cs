namespace AddressesJazz
{
    partial class FormHelp
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FormHelp));
            this.m_button_help_exit = new System.Windows.Forms.Button();
            this.m_rich_text_box_help = new System.Windows.Forms.RichTextBox();
            this.m_button_test_protokoll = new System.Windows.Forms.Button();
            this.m_button_program_documentation = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // m_button_help_exit
            // 
            this.m_button_help_exit.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.m_button_help_exit.Location = new System.Drawing.Point(614, 304);
            this.m_button_help_exit.Name = "m_button_help_exit";
            this.m_button_help_exit.Size = new System.Drawing.Size(72, 22);
            this.m_button_help_exit.TabIndex = 2;
            this.m_button_help_exit.Text = "Exit";
            this.m_button_help_exit.UseVisualStyleBackColor = true;
            this.m_button_help_exit.Click += new System.EventHandler(this.m_button_help_exit_Click);
            // 
            // m_rich_text_box_help
            // 
            this.m_rich_text_box_help.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.m_rich_text_box_help.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.m_rich_text_box_help.Location = new System.Drawing.Point(2, 12);
            this.m_rich_text_box_help.Name = "m_rich_text_box_help";
            this.m_rich_text_box_help.Size = new System.Drawing.Size(699, 284);
            this.m_rich_text_box_help.TabIndex = 3;
            this.m_rich_text_box_help.Text = "";
            // 
            // m_button_test_protokoll
            // 
            this.m_button_test_protokoll.Location = new System.Drawing.Point(317, 304);
            this.m_button_test_protokoll.Name = "m_button_test_protokoll";
            this.m_button_test_protokoll.Size = new System.Drawing.Size(97, 22);
            this.m_button_test_protokoll.TabIndex = 4;
            this.m_button_test_protokoll.Text = "Test Protokoll";
            this.m_button_test_protokoll.UseVisualStyleBackColor = true;
            this.m_button_test_protokoll.Click += new System.EventHandler(this.m_button_test_protokoll_Click);
            // 
            // m_button_program_documentation
            // 
            this.m_button_program_documentation.Location = new System.Drawing.Point(12, 304);
            this.m_button_program_documentation.Name = "m_button_program_documentation";
            this.m_button_program_documentation.Size = new System.Drawing.Size(122, 22);
            this.m_button_program_documentation.TabIndex = 5;
            this.m_button_program_documentation.Text = "Programm Dokumentation";
            this.m_button_program_documentation.UseVisualStyleBackColor = true;
            this.m_button_program_documentation.Click += new System.EventHandler(this.m_button_program_documentation_Click);
            // 
            // FormHelp
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(5F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(698, 338);
            this.Controls.Add(this.m_button_program_documentation);
            this.Controls.Add(this.m_button_test_protokoll);
            this.Controls.Add(this.m_rich_text_box_help);
            this.Controls.Add(this.m_button_help_exit);
            this.Font = new System.Drawing.Font("Arial Narrow", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(2, 3, 2, 3);
            this.Name = "FormHelp";
            this.Text = "Help Adressen";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button m_button_help_exit;
        private System.Windows.Forms.RichTextBox m_rich_text_box_help;
        private System.Windows.Forms.Button m_button_test_protokoll;
        private System.Windows.Forms.Button m_button_program_documentation;
    }
}