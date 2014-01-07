namespace VirtualClipBoard
{
    partial class VirtualClipBoard
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(VirtualClipBoard));
            this.list_clipboard = new System.Windows.Forms.ListBox();
            this.exit = new System.Windows.Forms.Button();
            this.clear = new System.Windows.Forms.Button();
            this.autoload = new System.Windows.Forms.CheckBox();
            this.history_size = new System.Windows.Forms.NumericUpDown();
            this.label1 = new System.Windows.Forms.Label();
            this.size_tray = new System.Windows.Forms.NumericUpDown();
            this.label2 = new System.Windows.Forms.Label();
            this._notifyIcon = new System.Windows.Forms.NotifyIcon(this.components);
            ((System.ComponentModel.ISupportInitialize)(this.history_size)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.size_tray)).BeginInit();
            this.SuspendLayout();
            // 
            // list_clipboard
            // 
            this.list_clipboard.FormattingEnabled = true;
            resources.ApplyResources(this.list_clipboard, "list_clipboard");
            this.list_clipboard.Name = "list_clipboard";
            this.list_clipboard.SelectedIndexChanged += new System.EventHandler(this.list_clipboard_SelectedIndexChanged);
            // 
            // exit
            // 
            resources.ApplyResources(this.exit, "exit");
            this.exit.Name = "exit";
            this.exit.UseVisualStyleBackColor = true;
            this.exit.Click += new System.EventHandler(this.exit_Click);
            // 
            // clear
            // 
            resources.ApplyResources(this.clear, "clear");
            this.clear.Name = "clear";
            this.clear.UseVisualStyleBackColor = true;
            this.clear.Click += new System.EventHandler(this.clear_Click);
            // 
            // autoload
            // 
            resources.ApplyResources(this.autoload, "autoload");
            this.autoload.Name = "autoload";
            this.autoload.UseVisualStyleBackColor = true;
            this.autoload.CheckedChanged += new System.EventHandler(this.autoload_CheckedChanged);
            // 
            // history_size
            // 
            resources.ApplyResources(this.history_size, "history_size");
            this.history_size.Maximum = new decimal(new int[] {
            10000,
            0,
            0,
            0});
            this.history_size.Minimum = new decimal(new int[] {
            2,
            0,
            0,
            0});
            this.history_size.Name = "history_size";
            this.history_size.Value = new decimal(new int[] {
            50,
            0,
            0,
            0});
            this.history_size.ValueChanged += new System.EventHandler(this.history_size_ValueChanged);
            // 
            // label1
            // 
            resources.ApplyResources(this.label1, "label1");
            this.label1.Name = "label1";
            // 
            // size_tray
            // 
            resources.ApplyResources(this.size_tray, "size_tray");
            this.size_tray.Maximum = new decimal(new int[] {
            1000,
            0,
            0,
            0});
            this.size_tray.Minimum = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.size_tray.Name = "size_tray";
            this.size_tray.Value = new decimal(new int[] {
            20,
            0,
            0,
            0});
            this.size_tray.ValueChanged += new System.EventHandler(this.size_tray_ValueChanged);
            // 
            // label2
            // 
            resources.ApplyResources(this.label2, "label2");
            this.label2.Name = "label2";
            // 
            // _notifyIcon
            // 
            resources.ApplyResources(this._notifyIcon, "_notifyIcon");
            // 
            // VirtualClipBoard
            // 
            resources.ApplyResources(this, "$this");
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.label2);
            this.Controls.Add(this.size_tray);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.history_size);
            this.Controls.Add(this.autoload);
            this.Controls.Add(this.clear);
            this.Controls.Add(this.exit);
            this.Controls.Add(this.list_clipboard);
            this.Name = "VirtualClipBoard";
            ((System.ComponentModel.ISupportInitialize)(this.history_size)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.size_tray)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ListBox list_clipboard;
        private System.Windows.Forms.Button exit;
        private System.Windows.Forms.Button clear;
        private System.Windows.Forms.CheckBox autoload;
        private System.Windows.Forms.NumericUpDown history_size;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.NumericUpDown size_tray;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.NotifyIcon _notifyIcon;
    }
}

