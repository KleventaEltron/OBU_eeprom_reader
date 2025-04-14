namespace EepromReader.Forms
{
    partial class xml_loader
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
            this.textBox_xml_location = new System.Windows.Forms.TextBox();
            this.button_get_xml = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // textBox_xml_location
            // 
            this.textBox_xml_location.Font = new System.Drawing.Font("Microsoft Sans Serif", 11F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBox_xml_location.Location = new System.Drawing.Point(59, 249);
            this.textBox_xml_location.Name = "textBox_xml_location";
            this.textBox_xml_location.Size = new System.Drawing.Size(606, 28);
            this.textBox_xml_location.TabIndex = 0;
            // 
            // button_get_xml
            // 
            this.button_get_xml.Location = new System.Drawing.Point(671, 249);
            this.button_get_xml.Name = "button_get_xml";
            this.button_get_xml.Size = new System.Drawing.Size(89, 26);
            this.button_get_xml.TabIndex = 1;
            this.button_get_xml.Text = "Open";
            this.button_get_xml.UseVisualStyleBackColor = true;
            this.button_get_xml.Click += new System.EventHandler(this.button_get_xml_Click);
            // 
            // xml_loader
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.button_get_xml);
            this.Controls.Add(this.textBox_xml_location);
            this.Name = "xml_loader";
            this.Text = "xml_loader";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox textBox_xml_location;
        private System.Windows.Forms.Button button_get_xml;
    }
}