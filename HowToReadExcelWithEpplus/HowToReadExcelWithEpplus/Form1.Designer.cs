namespace HowToReadExcelWithEpplus
{
    partial class Form1
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
            this.components = new System.ComponentModel.Container();
            this.btnOpenExcel = new System.Windows.Forms.Button();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.servicioDTOBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.servicioDTOBindingSource1 = new System.Windows.Forms.BindingSource(this.components);
            this.numeroDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.serviciosDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.litrosDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.mesDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.anioDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.servicioDTOBindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.servicioDTOBindingSource1)).BeginInit();
            this.SuspendLayout();
            // 
            // btnOpenExcel
            // 
            this.btnOpenExcel.Location = new System.Drawing.Point(615, 38);
            this.btnOpenExcel.Name = "btnOpenExcel";
            this.btnOpenExcel.Size = new System.Drawing.Size(173, 23);
            this.btnOpenExcel.TabIndex = 0;
            this.btnOpenExcel.Text = "Leer Documento";
            this.btnOpenExcel.UseVisualStyleBackColor = true;
            this.btnOpenExcel.Click += new System.EventHandler(this.btnOpenExcel_Click);
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(12, 12);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(776, 20);
            this.textBox1.TabIndex = 1;
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // dataGridView1
            // 
            this.dataGridView1.AutoGenerateColumns = false;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.numeroDataGridViewTextBoxColumn,
            this.serviciosDataGridViewTextBoxColumn,
            this.litrosDataGridViewTextBoxColumn,
            this.mesDataGridViewTextBoxColumn,
            this.anioDataGridViewTextBoxColumn});
            this.dataGridView1.DataSource = this.servicioDTOBindingSource1;
            this.dataGridView1.Location = new System.Drawing.Point(12, 124);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.Size = new System.Drawing.Size(743, 150);
            this.dataGridView1.TabIndex = 2;
            // 
            // servicioDTOBindingSource
            // 
            this.servicioDTOBindingSource.DataSource = typeof(HowToReadExcelWithEpplus.ServicioDTO);
            // 
            // servicioDTOBindingSource1
            // 
            this.servicioDTOBindingSource1.DataSource = typeof(HowToReadExcelWithEpplus.ServicioDTO);
            // 
            // numeroDataGridViewTextBoxColumn
            // 
            this.numeroDataGridViewTextBoxColumn.DataPropertyName = "Numero";
            this.numeroDataGridViewTextBoxColumn.HeaderText = "Numero";
            this.numeroDataGridViewTextBoxColumn.Name = "numeroDataGridViewTextBoxColumn";
            // 
            // serviciosDataGridViewTextBoxColumn
            // 
            this.serviciosDataGridViewTextBoxColumn.DataPropertyName = "Servicios";
            this.serviciosDataGridViewTextBoxColumn.HeaderText = "Servicios";
            this.serviciosDataGridViewTextBoxColumn.Name = "serviciosDataGridViewTextBoxColumn";
            // 
            // litrosDataGridViewTextBoxColumn
            // 
            this.litrosDataGridViewTextBoxColumn.DataPropertyName = "Litros";
            this.litrosDataGridViewTextBoxColumn.HeaderText = "Litros";
            this.litrosDataGridViewTextBoxColumn.Name = "litrosDataGridViewTextBoxColumn";
            // 
            // mesDataGridViewTextBoxColumn
            // 
            this.mesDataGridViewTextBoxColumn.DataPropertyName = "Mes";
            this.mesDataGridViewTextBoxColumn.HeaderText = "Mes";
            this.mesDataGridViewTextBoxColumn.Name = "mesDataGridViewTextBoxColumn";
            // 
            // anioDataGridViewTextBoxColumn
            // 
            this.anioDataGridViewTextBoxColumn.DataPropertyName = "Anio";
            this.anioDataGridViewTextBoxColumn.HeaderText = "Anio";
            this.anioDataGridViewTextBoxColumn.Name = "anioDataGridViewTextBoxColumn";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.btnOpenExcel);
            this.Name = "Form1";
            this.Text = "Leer Excel Picolita";
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.servicioDTOBindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.servicioDTOBindingSource1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnOpenExcel;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.DataGridViewTextBoxColumn servicioDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn añoDataGridViewTextBoxColumn;
        private System.Windows.Forms.BindingSource servicioDTOBindingSource;
        private System.Windows.Forms.DataGridViewTextBoxColumn numeroDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn serviciosDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn litrosDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn mesDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn anioDataGridViewTextBoxColumn;
        private System.Windows.Forms.BindingSource servicioDTOBindingSource1;
    }
}

