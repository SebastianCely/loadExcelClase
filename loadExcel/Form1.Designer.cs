namespace loadExcel
{
    partial class Form1
    {
        /// <summary>
        /// Variable del diseñador necesaria.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Limpiar los recursos que se estén usando.
        /// </summary>
        /// <param name="disposing">true si los recursos administrados se deben desechar; false en caso contrario.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Código generado por el Diseñador de Windows Forms

        /// <summary>
        /// Método necesario para admitir el Diseñador. No se puede modificar
        /// el contenido de este método con el editor de código.
        /// </summary>
        private void InitializeComponent()
        {
            this.btnLoadExcel = new System.Windows.Forms.Button();
            this.dgvDatos = new System.Windows.Forms.DataGridView();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.dgvResumen = new System.Windows.Forms.DataGridView();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.dgvEstadisticas = new System.Windows.Forms.DataGridView();
            this.btnExportar = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dgvDatos)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgvResumen)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgvEstadisticas)).BeginInit();
            this.SuspendLayout();
            // 
            // btnLoadExcel
            // 
            this.btnLoadExcel.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnLoadExcel.Location = new System.Drawing.Point(45, 31);
            this.btnLoadExcel.Name = "btnLoadExcel";
            this.btnLoadExcel.Size = new System.Drawing.Size(139, 36);
            this.btnLoadExcel.TabIndex = 0;
            this.btnLoadExcel.Text = "Cargar Excel...";
            this.btnLoadExcel.UseVisualStyleBackColor = true;
            this.btnLoadExcel.Click += new System.EventHandler(this.btnLoadExcel_Click);
            // 
            // dgvDatos
            // 
            this.dgvDatos.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvDatos.Location = new System.Drawing.Point(45, 106);
            this.dgvDatos.Name = "dgvDatos";
            this.dgvDatos.Size = new System.Drawing.Size(479, 170);
            this.dgvDatos.TabIndex = 1;
            // 
            // textBox1
            // 
            this.textBox1.BackColor = System.Drawing.SystemColors.Menu;
            this.textBox1.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.textBox1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBox1.Location = new System.Drawing.Point(45, 299);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(222, 14);
            this.textBox1.TabIndex = 2;
            this.textBox1.Text = "AGRUPAR TIPOS DOCUMENTOS";
            // 
            // dgvResumen
            // 
            this.dgvResumen.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvResumen.Location = new System.Drawing.Point(45, 338);
            this.dgvResumen.Name = "dgvResumen";
            this.dgvResumen.Size = new System.Drawing.Size(240, 150);
            this.dgvResumen.TabIndex = 3;
            // 
            // textBox2
            // 
            this.textBox2.BackColor = System.Drawing.SystemColors.Menu;
            this.textBox2.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.textBox2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBox2.Location = new System.Drawing.Point(377, 299);
            this.textBox2.Name = "textBox2";
            this.textBox2.Size = new System.Drawing.Size(100, 13);
            this.textBox2.TabIndex = 4;
            this.textBox2.Text = "ESTADÍSTICAS";
            // 
            // dgvEstadisticas
            // 
            this.dgvEstadisticas.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvEstadisticas.Location = new System.Drawing.Point(377, 338);
            this.dgvEstadisticas.Name = "dgvEstadisticas";
            this.dgvEstadisticas.Size = new System.Drawing.Size(493, 150);
            this.dgvEstadisticas.TabIndex = 5;
            // 
            // btnExportar
            // 
            this.btnExportar.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnExportar.Location = new System.Drawing.Point(562, 31);
            this.btnExportar.Name = "btnExportar";
            this.btnExportar.Size = new System.Drawing.Size(149, 23);
            this.btnExportar.TabIndex = 6;
            this.btnExportar.Text = "Exportar CSV...";
            this.btnExportar.UseVisualStyleBackColor = true;
            this.btnExportar.Click += new System.EventHandler(this.btnExportar_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(882, 503);
            this.Controls.Add(this.btnExportar);
            this.Controls.Add(this.dgvEstadisticas);
            this.Controls.Add(this.textBox2);
            this.Controls.Add(this.dgvResumen);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.dgvDatos);
            this.Controls.Add(this.btnLoadExcel);
            this.Name = "Form1";
            this.Text = "Form1";
            ((System.ComponentModel.ISupportInitialize)(this.dgvDatos)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgvResumen)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgvEstadisticas)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnLoadExcel;
        private System.Windows.Forms.DataGridView dgvDatos;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.DataGridView dgvResumen;
        private System.Windows.Forms.TextBox textBox2;
        private System.Windows.Forms.DataGridView dgvEstadisticas;
        private System.Windows.Forms.Button btnExportar;
    }
}

