namespace ExcelSplitter
{
    partial class ExcelSplitter
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
            this.pnl_top_panel = new System.Windows.Forms.Panel();
            this.lbl_txtsavepath = new System.Windows.Forms.Label();
            this.lbl_txtfilename = new System.Windows.Forms.Label();
            this.lbl_txtn = new System.Windows.Forms.Label();
            this.lbl_savepath = new System.Windows.Forms.Label();
            this.lbl_n = new System.Windows.Forms.Label();
            this.lbl_filename = new System.Windows.Forms.Label();
            this.txt_n = new System.Windows.Forms.TextBox();
            this.btn_split = new System.Windows.Forms.Button();
            this.btn_savepath = new System.Windows.Forms.Button();
            this.btn_file = new System.Windows.Forms.Button();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.panel2 = new System.Windows.Forms.Panel();
            this.lbl_error = new System.Windows.Forms.Label();
            this.lbl_log = new System.Windows.Forms.Label();
            this.lbl_status = new System.Windows.Forms.Label();
            this.tab_control = new System.Windows.Forms.TabControl();
            this.btn_openfolder = new System.Windows.Forms.Button();
            this.pnl_top_panel.SuspendLayout();
            this.panel2.SuspendLayout();
            this.SuspendLayout();
            // 
            // pnl_top_panel
            // 
            this.pnl_top_panel.Controls.Add(this.btn_openfolder);
            this.pnl_top_panel.Controls.Add(this.lbl_filename);
            this.pnl_top_panel.Controls.Add(this.lbl_txtsavepath);
            this.pnl_top_panel.Controls.Add(this.lbl_txtfilename);
            this.pnl_top_panel.Controls.Add(this.lbl_txtn);
            this.pnl_top_panel.Controls.Add(this.lbl_savepath);
            this.pnl_top_panel.Controls.Add(this.lbl_n);
            this.pnl_top_panel.Controls.Add(this.txt_n);
            this.pnl_top_panel.Controls.Add(this.btn_split);
            this.pnl_top_panel.Controls.Add(this.btn_savepath);
            this.pnl_top_panel.Controls.Add(this.btn_file);
            this.pnl_top_panel.Location = new System.Drawing.Point(21, 22);
            this.pnl_top_panel.Name = "pnl_top_panel";
            this.pnl_top_panel.Size = new System.Drawing.Size(755, 133);
            this.pnl_top_panel.TabIndex = 0;
            // 
            // lbl_txtsavepath
            // 
            this.lbl_txtsavepath.AutoSize = true;
            this.lbl_txtsavepath.Location = new System.Drawing.Point(34, 74);
            this.lbl_txtsavepath.Name = "lbl_txtsavepath";
            this.lbl_txtsavepath.Size = new System.Drawing.Size(60, 17);
            this.lbl_txtsavepath.TabIndex = 3;
            this.lbl_txtsavepath.Text = "Save at:";
            // 
            // lbl_txtfilename
            // 
            this.lbl_txtfilename.AutoSize = true;
            this.lbl_txtfilename.Location = new System.Drawing.Point(34, 44);
            this.lbl_txtfilename.Name = "lbl_txtfilename";
            this.lbl_txtfilename.Size = new System.Drawing.Size(34, 17);
            this.lbl_txtfilename.TabIndex = 3;
            this.lbl_txtfilename.Text = "File:";
            // 
            // lbl_txtn
            // 
            this.lbl_txtn.AutoSize = true;
            this.lbl_txtn.Location = new System.Drawing.Point(34, 13);
            this.lbl_txtn.Name = "lbl_txtn";
            this.lbl_txtn.Size = new System.Drawing.Size(79, 17);
            this.lbl_txtn.TabIndex = 3;
            this.lbl_txtn.Text = "No. to split:";
            // 
            // lbl_savepath
            // 
            this.lbl_savepath.AutoSize = true;
            this.lbl_savepath.Location = new System.Drawing.Point(204, 74);
            this.lbl_savepath.Name = "lbl_savepath";
            this.lbl_savepath.Size = new System.Drawing.Size(46, 17);
            this.lbl_savepath.TabIndex = 0;
            this.lbl_savepath.Text = "label1";
            // 
            // lbl_n
            // 
            this.lbl_n.AutoSize = true;
            this.lbl_n.Location = new System.Drawing.Point(204, 13);
            this.lbl_n.Name = "lbl_n";
            this.lbl_n.Size = new System.Drawing.Size(46, 17);
            this.lbl_n.TabIndex = 0;
            this.lbl_n.Text = "label1";
            // 
            // lbl_filename
            // 
            this.lbl_filename.AutoSize = true;
            this.lbl_filename.Location = new System.Drawing.Point(204, 44);
            this.lbl_filename.Name = "lbl_filename";
            this.lbl_filename.Size = new System.Drawing.Size(46, 17);
            this.lbl_filename.TabIndex = 0;
            this.lbl_filename.Text = "label1";
            // 
            // txt_n
            // 
            this.txt_n.Location = new System.Drawing.Point(123, 10);
            this.txt_n.Name = "txt_n";
            this.txt_n.Size = new System.Drawing.Size(75, 22);
            this.txt_n.TabIndex = 2;
            // 
            // btn_split
            // 
            this.btn_split.Location = new System.Drawing.Point(123, 100);
            this.btn_split.Name = "btn_split";
            this.btn_split.Size = new System.Drawing.Size(75, 25);
            this.btn_split.TabIndex = 1;
            this.btn_split.Text = "Split";
            this.btn_split.UseVisualStyleBackColor = true;
            this.btn_split.Click += new System.EventHandler(this.btn_split_Click);
            // 
            // btn_savepath
            // 
            this.btn_savepath.Location = new System.Drawing.Point(123, 70);
            this.btn_savepath.Name = "btn_savepath";
            this.btn_savepath.Size = new System.Drawing.Size(75, 25);
            this.btn_savepath.TabIndex = 0;
            this.btn_savepath.Text = "Browse";
            this.btn_savepath.UseVisualStyleBackColor = true;
            this.btn_savepath.Click += new System.EventHandler(this.btn_savepath_Click);
            // 
            // btn_file
            // 
            this.btn_file.Location = new System.Drawing.Point(123, 40);
            this.btn_file.Name = "btn_file";
            this.btn_file.Size = new System.Drawing.Size(75, 25);
            this.btn_file.TabIndex = 0;
            this.btn_file.Text = "Browse";
            this.btn_file.UseVisualStyleBackColor = true;
            this.btn_file.Click += new System.EventHandler(this.btn_file_Click);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // panel2
            // 
            this.panel2.Controls.Add(this.lbl_error);
            this.panel2.Controls.Add(this.lbl_log);
            this.panel2.Controls.Add(this.lbl_status);
            this.panel2.Location = new System.Drawing.Point(21, 161);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(755, 69);
            this.panel2.TabIndex = 0;
            // 
            // lbl_error
            // 
            this.lbl_error.AutoSize = true;
            this.lbl_error.Location = new System.Drawing.Point(34, 43);
            this.lbl_error.Name = "lbl_error";
            this.lbl_error.Size = new System.Drawing.Size(46, 17);
            this.lbl_error.TabIndex = 0;
            this.lbl_error.Text = "label1";
            // 
            // lbl_log
            // 
            this.lbl_log.AutoSize = true;
            this.lbl_log.Location = new System.Drawing.Point(34, 26);
            this.lbl_log.Name = "lbl_log";
            this.lbl_log.Size = new System.Drawing.Size(46, 17);
            this.lbl_log.TabIndex = 0;
            this.lbl_log.Text = "label1";
            // 
            // lbl_status
            // 
            this.lbl_status.AutoSize = true;
            this.lbl_status.Location = new System.Drawing.Point(34, 9);
            this.lbl_status.Name = "lbl_status";
            this.lbl_status.Size = new System.Drawing.Size(46, 17);
            this.lbl_status.TabIndex = 0;
            this.lbl_status.Text = "label1";
            // 
            // tab_control
            // 
            this.tab_control.Location = new System.Drawing.Point(21, 234);
            this.tab_control.Name = "tab_control";
            this.tab_control.SelectedIndex = 0;
            this.tab_control.Size = new System.Drawing.Size(755, 233);
            this.tab_control.TabIndex = 2;
            // 
            // btn_fileopen
            // 
            this.btn_openfolder.Location = new System.Drawing.Point(257, 70);
            this.btn_openfolder.Name = "btn_fileopen";
            this.btn_openfolder.Size = new System.Drawing.Size(30, 25);
            this.btn_openfolder.TabIndex = 4;
            this.btn_openfolder.Text = "...";
            this.btn_openfolder.UseVisualStyleBackColor = true;
            // 
            // ExcelSplitter
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 480);
            this.Controls.Add(this.tab_control);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.pnl_top_panel);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Name = "ExcelSplitter";
            this.Text = "Form 1";
            this.pnl_top_panel.ResumeLayout(false);
            this.pnl_top_panel.PerformLayout();
            this.panel2.ResumeLayout(false);
            this.panel2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel pnl_top_panel;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.Button btn_file;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.Label lbl_filename;
        private System.Windows.Forms.Label lbl_status;
        private System.Windows.Forms.Label lbl_log;
        private System.Windows.Forms.Label lbl_error;
        private System.Windows.Forms.TabControl tab_control;
        private System.Windows.Forms.Button btn_split;
        private System.Windows.Forms.Button btn_savepath;
        private System.Windows.Forms.TextBox txt_n;
        private System.Windows.Forms.Label lbl_txtsavepath;
        private System.Windows.Forms.Label lbl_txtfilename;
        private System.Windows.Forms.Label lbl_txtn;
        private System.Windows.Forms.Label lbl_savepath;
        private System.Windows.Forms.Label lbl_n;
        private System.Windows.Forms.Button btn_openfolder;
    }
}

