namespace Шаблон_программы_Димы
{
    partial class Form2
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
            this.btnCancelProp = new System.Windows.Forms.Button();
            this.btnSaveProp = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // btnCancelProp
            // 
            this.btnCancelProp.Font = new System.Drawing.Font("Arial", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.btnCancelProp.Location = new System.Drawing.Point(432, 166);
            this.btnCancelProp.Name = "btnCancelProp";
            this.btnCancelProp.Size = new System.Drawing.Size(105, 28);
            this.btnCancelProp.TabIndex = 12;
            this.btnCancelProp.Text = "Отмена";
            this.btnCancelProp.UseVisualStyleBackColor = true;
            this.btnCancelProp.Click += new System.EventHandler(this.btnCancelProp_Click);
            // 
            // btnSaveProp
            // 
            this.btnSaveProp.Font = new System.Drawing.Font("Arial", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.btnSaveProp.Location = new System.Drawing.Point(321, 166);
            this.btnSaveProp.Name = "btnSaveProp";
            this.btnSaveProp.Size = new System.Drawing.Size(105, 28);
            this.btnSaveProp.TabIndex = 13;
            this.btnSaveProp.Text = "Сохранить";
            this.btnSaveProp.UseVisualStyleBackColor = true;
            // 
            // Form2
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(549, 206);
            this.Controls.Add(this.btnSaveProp);
            this.Controls.Add(this.btnCancelProp);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.KeyPreview = true;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "Form2";
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Добавление свойств";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button btnCancelProp;
        private System.Windows.Forms.Button btnSaveProp;
    }
}