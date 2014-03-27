namespace Chemtech.Autoproj.DocGenerator.Interface
{
    partial class SelectTemplateFormForGen
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
            this.label1 = new System.Windows.Forms.Label();
            this.cancel = new System.Windows.Forms.Button();
            this.Templates_Liberados = new System.Windows.Forms.ListBox();
            this.btnConfirm = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(13, 8);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(326, 13);
            this.label1.TabIndex = 17;
            this.label1.Text = "Selecione o template liberado para o qual deseja preparar a planilha";
            // 
            // cancel
            // 
            this.cancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.cancel.Location = new System.Drawing.Point(310, 167);
            this.cancel.Name = "cancel";
            this.cancel.Size = new System.Drawing.Size(75, 23);
            this.cancel.TabIndex = 16;
            this.cancel.Text = "Cancel";
            this.cancel.UseVisualStyleBackColor = true;
            this.cancel.Click += new System.EventHandler(this.CancelClick);
            // 
            // Templates_Liberados
            // 
            this.Templates_Liberados.FormattingEnabled = true;
            this.Templates_Liberados.Location = new System.Drawing.Point(12, 30);
            this.Templates_Liberados.Name = "Templates_Liberados";
            this.Templates_Liberados.Size = new System.Drawing.Size(292, 160);
            this.Templates_Liberados.TabIndex = 15;
            this.Templates_Liberados.SelectedIndexChanged += new System.EventHandler(this.Templates_Liberados_SelectedIndexChanged);
            // 
            // btnConfirm
            // 
            this.btnConfirm.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.btnConfirm.Location = new System.Drawing.Point(310, 138);
            this.btnConfirm.Name = "btnConfirm";
            this.btnConfirm.Size = new System.Drawing.Size(75, 23);
            this.btnConfirm.TabIndex = 14;
            this.btnConfirm.Text = "OK";
            this.btnConfirm.UseVisualStyleBackColor = true;
            this.btnConfirm.Click += new System.EventHandler(this.BtnConfirmClick);
            // 
            // SelectTemplateFormForGen
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(395, 203);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.cancel);
            this.Controls.Add(this.Templates_Liberados);
            this.Controls.Add(this.btnConfirm);
            this.Name = "SelectTemplateFormForGen";
            this.Text = "Macro de preenchimento de templates";
            this.Load += new System.EventHandler(this.SelectTemplateFormForGenLoad);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button cancel;
        private System.Windows.Forms.ListBox Templates_Liberados;
        private System.Windows.Forms.Button btnConfirm;
    }
}