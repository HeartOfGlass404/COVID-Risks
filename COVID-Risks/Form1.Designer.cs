namespace COVID_Risks
{
    partial class Form1
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
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
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            flowLayoutPanel1 = new FlowLayoutPanel();
            toolStrip1 = new ToolStrip();
            FileMenuButton = new ToolStripDropDownButton();
            OpenButton = new ToolStripMenuItem();
            SaveMenuButton = new ToolStripMenuItem();
            сохранитьКакToolStripMenuItem = new ToolStripMenuItem();
            toolStripSeparator1 = new ToolStripSeparator();
            EditMenuButton = new ToolStripDropDownButton();
            AddButton = new ToolStripMenuItem();
            DeleteButton = new ToolStripMenuItem();
            изменитьЗаписьToolStripMenuItem = new ToolStripMenuItem();
            HelpMenuButton = new ToolStripButton();
            personBindingSource = new BindingSource(components);
            PersonList = new ListView();
            FIOColumn = new ColumnHeader();
            PersonView1 = new ListView();
            columnHeader1 = new ColumnHeader();
            columnHeader2 = new ColumnHeader();
            splitContainer1 = new SplitContainer();
            flowLayoutPanel1.SuspendLayout();
            toolStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)personBindingSource).BeginInit();
            ((System.ComponentModel.ISupportInitialize)splitContainer1).BeginInit();
            splitContainer1.Panel1.SuspendLayout();
            splitContainer1.Panel2.SuspendLayout();
            splitContainer1.SuspendLayout();
            SuspendLayout();
            // 
            // flowLayoutPanel1
            // 
            flowLayoutPanel1.Controls.Add(toolStrip1);
            flowLayoutPanel1.Dock = DockStyle.Top;
            flowLayoutPanel1.Location = new Point(0, 0);
            flowLayoutPanel1.Name = "flowLayoutPanel1";
            flowLayoutPanel1.Size = new Size(800, 31);
            flowLayoutPanel1.TabIndex = 0;
            // 
            // toolStrip1
            // 
            toolStrip1.ImageScalingSize = new Size(20, 20);
            toolStrip1.Items.AddRange(new ToolStripItem[] { FileMenuButton, toolStripSeparator1, EditMenuButton, HelpMenuButton });
            toolStrip1.Location = new Point(0, 0);
            toolStrip1.Name = "toolStrip1";
            toolStrip1.Size = new Size(155, 27);
            toolStrip1.TabIndex = 0;
            toolStrip1.Text = "toolStrip1";
            // 
            // FileMenuButton
            // 
            FileMenuButton.DisplayStyle = ToolStripItemDisplayStyle.Image;
            FileMenuButton.DropDownItems.AddRange(new ToolStripItem[] { OpenButton, SaveMenuButton, сохранитьКакToolStripMenuItem });
            FileMenuButton.Image = (Image)resources.GetObject("FileMenuButton.Image");
            FileMenuButton.ImageTransparentColor = Color.Magenta;
            FileMenuButton.Name = "FileMenuButton";
            FileMenuButton.Size = new Size(34, 24);
            FileMenuButton.Text = "Файл";
            // 
            // OpenButton
            // 
            OpenButton.Name = "OpenButton";
            OpenButton.Size = new Size(224, 26);
            OpenButton.Text = "Открыть";
            OpenButton.Click += OpenButton_Click;
            // 
            // SaveMenuButton
            // 
            SaveMenuButton.Name = "SaveMenuButton";
            SaveMenuButton.Size = new Size(224, 26);
            SaveMenuButton.Text = "Сохранить";
            SaveMenuButton.Click += SaveMenuButton_Click;
            // 
            // сохранитьКакToolStripMenuItem
            // 
            сохранитьКакToolStripMenuItem.Name = "сохранитьКакToolStripMenuItem";
            сохранитьКакToolStripMenuItem.Size = new Size(224, 26);
            сохранитьКакToolStripMenuItem.Text = "Сохранить как...";
            сохранитьКакToolStripMenuItem.Click += SaveAsMenuButton_Click;
            // 
            // toolStripSeparator1
            // 
            toolStripSeparator1.Name = "toolStripSeparator1";
            toolStripSeparator1.Size = new Size(6, 27);
            // 
            // EditMenuButton
            // 
            EditMenuButton.DisplayStyle = ToolStripItemDisplayStyle.Image;
            EditMenuButton.DropDownItems.AddRange(new ToolStripItem[] { AddButton, DeleteButton, изменитьЗаписьToolStripMenuItem });
            EditMenuButton.Image = (Image)resources.GetObject("EditMenuButton.Image");
            EditMenuButton.ImageTransparentColor = Color.Magenta;
            EditMenuButton.Name = "EditMenuButton";
            EditMenuButton.Size = new Size(34, 24);
            EditMenuButton.Text = "Редактировать";
            // 
            // AddButton
            // 
            AddButton.Name = "AddButton";
            AddButton.Size = new Size(213, 26);
            AddButton.Text = "Добавить запись";
            AddButton.Click += AddButton_Click;
            // 
            // DeleteButton
            // 
            DeleteButton.Name = "DeleteButton";
            DeleteButton.Size = new Size(213, 26);
            DeleteButton.Text = "Удалить запись";
            DeleteButton.Click += DeleteButton_Click;
            // 
            // изменитьЗаписьToolStripMenuItem
            // 
            изменитьЗаписьToolStripMenuItem.Name = "изменитьЗаписьToolStripMenuItem";
            изменитьЗаписьToolStripMenuItem.Size = new Size(213, 26);
            изменитьЗаписьToolStripMenuItem.Text = "Изменить запись";
            // 
            // HelpMenuButton
            // 
            HelpMenuButton.DisplayStyle = ToolStripItemDisplayStyle.Image;
            HelpMenuButton.Image = (Image)resources.GetObject("HelpMenuButton.Image");
            HelpMenuButton.ImageTransparentColor = Color.Magenta;
            HelpMenuButton.Name = "HelpMenuButton";
            HelpMenuButton.Size = new Size(29, 24);
            HelpMenuButton.Text = "Помощь";
            // 
            // personBindingSource
            // 
            personBindingSource.DataSource = typeof(WinFormsApp1.Models.Person);
            // 
            // PersonList
            // 
            PersonList.Columns.AddRange(new ColumnHeader[] { FIOColumn });
            PersonList.DataBindings.Add(new Binding("DataContext", personBindingSource, "Fio", true));
            PersonList.Dock = DockStyle.Fill;
            PersonList.GridLines = true;
            PersonList.Location = new Point(0, 0);
            PersonList.Margin = new Padding(3, 4, 3, 4);
            PersonList.Name = "PersonList";
            PersonList.Size = new Size(265, 420);
            PersonList.TabIndex = 0;
            PersonList.UseCompatibleStateImageBehavior = false;
            PersonList.View = View.Details;
            PersonList.SelectedIndexChanged += PersonList_SelectedIndexChanged;
            // 
            // FIOColumn
            // 
            FIOColumn.Text = "Пациенты";
            FIOColumn.Width = 261;
            // 
            // PersonView1
            // 
            PersonView1.Columns.AddRange(new ColumnHeader[] { columnHeader1, columnHeader2 });
            PersonView1.Dock = DockStyle.Fill;
            PersonView1.GridLines = true;
            PersonView1.Location = new Point(0, 0);
            PersonView1.Name = "PersonView1";
            PersonView1.Size = new Size(530, 420);
            PersonView1.TabIndex = 0;
            PersonView1.UseCompatibleStateImageBehavior = false;
            PersonView1.View = View.Details;
            // 
            // columnHeader1
            // 
            columnHeader1.Text = "";
            columnHeader1.Width = 250;
            // 
            // columnHeader2
            // 
            columnHeader2.Text = "";
            columnHeader2.Width = 250;
            // 
            // splitContainer1
            // 
            splitContainer1.Dock = DockStyle.Fill;
            splitContainer1.Location = new Point(0, 31);
            splitContainer1.Name = "splitContainer1";
            // 
            // splitContainer1.Panel1
            // 
            splitContainer1.Panel1.Controls.Add(PersonList);
            // 
            // splitContainer1.Panel2
            // 
            splitContainer1.Panel2.Controls.Add(PersonView1);
            splitContainer1.Size = new Size(800, 420);
            splitContainer1.SplitterDistance = 265;
            splitContainer1.SplitterWidth = 5;
            splitContainer1.TabIndex = 1;
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(8F, 20F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(800, 451);
            Controls.Add(splitContainer1);
            Controls.Add(flowLayoutPanel1);
            Name = "Form1";
            Text = "Covid-Risk";
            flowLayoutPanel1.ResumeLayout(false);
            flowLayoutPanel1.PerformLayout();
            toolStrip1.ResumeLayout(false);
            toolStrip1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)personBindingSource).EndInit();
            splitContainer1.Panel1.ResumeLayout(false);
            splitContainer1.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)splitContainer1).EndInit();
            splitContainer1.ResumeLayout(false);
            ResumeLayout(false);
        }

        #endregion

        private FlowLayoutPanel flowLayoutPanel1;
        private BindingSource personBindingSource;
        private ToolStrip toolStrip1;
        private ToolStripDropDownButton FileMenuButton;
        private ToolStripMenuItem OpenButton;
        private ToolStripMenuItem SaveMenuButton;
        private ToolStripMenuItem сохранитьКакToolStripMenuItem;
        private ToolStripSeparator toolStripSeparator1;
        private ToolStripDropDownButton EditMenuButton;
        private ToolStripMenuItem AddButton;
        private ToolStripMenuItem DeleteButton;
        private ToolStripMenuItem изменитьЗаписьToolStripMenuItem;
        private ToolStripButton HelpMenuButton;
        private ListView PersonList;
        private ColumnHeader FIOColumn;
        private ListView PersonView1;
        private ColumnHeader columnHeader1;
        private ColumnHeader columnHeader2;
        private SplitContainer splitContainer1;
    }
}