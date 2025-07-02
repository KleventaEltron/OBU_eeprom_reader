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
            this.save_eeprom_btn = new System.Windows.Forms.Button();
            this.save_eeprom_textbox = new System.Windows.Forms.TextBox();
            this.checkBox1 = new System.Windows.Forms.CheckBox();
            this.SuspendLayout();
            // 
            // textBox_xml_location
            // 
            this.textBox_xml_location.Font = new System.Drawing.Font("Microsoft Sans Serif", 11F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBox_xml_location.Location = new System.Drawing.Point(44, 202);
            this.textBox_xml_location.Margin = new System.Windows.Forms.Padding(2);
            this.textBox_xml_location.Name = "textBox_xml_location";
            this.textBox_xml_location.Size = new System.Drawing.Size(456, 24);
            this.textBox_xml_location.TabIndex = 0;
            // 
            // button_get_xml
            // 
            this.button_get_xml.Location = new System.Drawing.Point(503, 202);
            this.button_get_xml.Margin = new System.Windows.Forms.Padding(2);
            this.button_get_xml.Name = "button_get_xml";
            this.button_get_xml.Size = new System.Drawing.Size(67, 21);
            this.button_get_xml.TabIndex = 1;
            this.button_get_xml.Text = "Open";
            this.button_get_xml.UseVisualStyleBackColor = true;
            this.button_get_xml.Click += new System.EventHandler(this.button_get_xml_Click);
            // 
            // save_eeprom_btn
            // 
            this.save_eeprom_btn.Location = new System.Drawing.Point(503, 281);
            this.save_eeprom_btn.Name = "save_eeprom_btn";
            this.save_eeprom_btn.Size = new System.Drawing.Size(67, 21);
            this.save_eeprom_btn.TabIndex = 2;
            this.save_eeprom_btn.Text = "Save";
            this.save_eeprom_btn.UseVisualStyleBackColor = true;
            this.save_eeprom_btn.Click += new System.EventHandler(this.save_eeprom_btn_Click);
            // 
            // save_eeprom_textbox
            // 
            this.save_eeprom_textbox.Font = new System.Drawing.Font("Microsoft Sans Serif", 11F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.save_eeprom_textbox.Location = new System.Drawing.Point(44, 278);
            this.save_eeprom_textbox.Margin = new System.Windows.Forms.Padding(2);
            this.save_eeprom_textbox.Name = "save_eeprom_textbox";
            this.save_eeprom_textbox.Size = new System.Drawing.Size(456, 24);
            this.save_eeprom_textbox.TabIndex = 3;
            // 
            // checkBox1
            // 
            this.checkBox1.AutoSize = true;
            this.checkBox1.Location = new System.Drawing.Point(44, 34);
            this.checkBox1.Name = "checkBox1";
            this.checkBox1.Size = new System.Drawing.Size(69, 17);
            this.checkBox1.TabIndex = 4;
            this.checkBox1.Text = "Unfolded";
            this.checkBox1.UseVisualStyleBackColor = true;
            this.checkBox1.CheckedChanged += new System.EventHandler(this.checkBox1_CheckedChanged);
            // 
            // xml_loader
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(600, 366);
            this.Controls.Add(this.checkBox1);
            this.Controls.Add(this.save_eeprom_textbox);
            this.Controls.Add(this.save_eeprom_btn);
            this.Controls.Add(this.button_get_xml);
            this.Controls.Add(this.textBox_xml_location);
            this.Margin = new System.Windows.Forms.Padding(2);
            this.Name = "xml_loader";
            this.Text = "xml_loader";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox textBox_xml_location;
        private System.Windows.Forms.Button button_get_xml;
        private System.Windows.Forms.Button save_eeprom_btn;
        private System.Windows.Forms.TextBox save_eeprom_textbox;
        private System.Windows.Forms.CheckBox checkBox1;
    }
}