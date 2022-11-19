namespace AddressesJazz
{
    partial class ResetForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ResetForm));
            this.m_combo_box_backups = new System.Windows.Forms.ComboBox();
            this.m_combo_box_backups_label = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // m_combo_box_backups
            // 
            this.m_combo_box_backups.FormattingEnabled = true;
            this.m_combo_box_backups.Location = new System.Drawing.Point(12, 28);
            this.m_combo_box_backups.Name = "m_combo_box_backups";
            this.m_combo_box_backups.Size = new System.Drawing.Size(398, 21);
            this.m_combo_box_backups.TabIndex = 0;
            this.m_combo_box_backups.SelectedIndexChanged += new System.EventHandler(this.m_combo_box_backups_SelectedIndexChanged);
            // 
            // m_combo_box_backups_label
            // 
            this.m_combo_box_backups_label.AutoSize = true;
            this.m_combo_box_backups_label.Font = new System.Drawing.Font("Arial Narrow", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.m_combo_box_backups_label.Location = new System.Drawing.Point(21, 10);
            this.m_combo_box_backups_label.Name = "m_combo_box_backups_label";
            this.m_combo_box_backups_label.Size = new System.Drawing.Size(79, 15);
            this.m_combo_box_backups_label.TabIndex = 1;
            this.m_combo_box_backups_label.Text = "Select backup file";
            // 
            // ResetForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(422, 57);
            this.Controls.Add(this.m_combo_box_backups_label);
            this.Controls.Add(this.m_combo_box_backups);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "ResetForm";
            this.Text = "Reset addresses";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ComboBox m_combo_box_backups;
        private System.Windows.Forms.Label m_combo_box_backups_label;
    }
}