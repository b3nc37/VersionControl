﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using UserMaintenance.Entities;

namespace UserMaintenance
{
    public partial class Form1 : Form
 {
        BindingList<User> users = new BindingList<User>();
        public Form1()
        {
            InitializeComponent();
            label1.Text = Resource.FullName;
            
            button1.Text = Resource.Add;
            button2.Text=Resource.Write;
            button3.Text = Resource.Delete;

            listBox1.DataSource = users;
            listBox1.ValueMember = "ID";
            listBox1.DisplayMember = "FullName";

           
        }

        private void button1_Click(object sender, EventArgs e)
        {
            var u = new User()
            {
                Fullname = textBox1.Text,
                
            };
            users.Add(u);
        }

        private void button2_Click(object sender, EventArgs e)
        {
           
            SaveFileDialog sfd = new SaveFileDialog();
            if (sfd.ShowDialog()!=DialogResult.OK)
            {
                return;
            }
            StreamWriter sw = new StreamWriter(sfd.FileName, false, Encoding.Default);

            foreach (var x in users)
            {
                sw.Write(x.ID);
                sw.Write(";");
                sw.Write(x.Fullname);
                sw.Write(";");
                sw.WriteLine();
            }
            
            
        }

        private void button3_Click(object sender, EventArgs e)
        {
            
            User torlendo = (User)listBox1.SelectedItem;
            users.Remove(torlendo);
        }
    }
}
