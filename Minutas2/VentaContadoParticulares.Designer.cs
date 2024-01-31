namespace Minutas2
{
    partial class VentaContadoParticulares
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
            this.btnvalidar = new System.Windows.Forms.Button();
            this.txtnumescritura = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // btnvalidar
            // 
            this.btnvalidar.Location = new System.Drawing.Point(102, 338);
            this.btnvalidar.Name = "btnvalidar";
            this.btnvalidar.Size = new System.Drawing.Size(75, 23);
            this.btnvalidar.TabIndex = 0;
            this.btnvalidar.Text = "validar";
            this.btnvalidar.UseVisualStyleBackColor = true;
            this.btnvalidar.Click += new System.EventHandler(this.btnvalidar_Click);
            // 
            // txtnumescritura
            // 
            this.txtnumescritura.Location = new System.Drawing.Point(135, 164);
            this.txtnumescritura.Name = "txtnumescritura";
            this.txtnumescritura.Size = new System.Drawing.Size(100, 20);
            this.txtnumescritura.TabIndex = 1;
            // 
            // VentaContadoParticulares
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.txtnumescritura);
            this.Controls.Add(this.btnvalidar);
            this.Name = "VentaContadoParticulares";
            this.Text = "VentaContadoParticulares";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnvalidar;
        private System.Windows.Forms.TextBox txtnumescritura;
    }
}