using System;
using Dapper;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Extensions.Configuration;
using System.Data.SqlClient;
using System.IO;
using System.Configuration;
using Microsoft.AspNetCore.Mvc;
using PMS_V4_SAP_integration.Helper;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace PMS_V4_SAP_integration
{
    public partial class Form1 : Form
    {
        private readonly IConfiguration _configuration;
        private System.Windows.Forms.ListView logListView;


        public Form1(string[] args)
        {
            //rread parameter invoice/cn/ commission
            //connect to database 
            //use dapper get things from the database
            //get array

            InitializeComponent();
            this.Text = "PMS_V4_SAP_integration";

            logListView = listView;
            logListView.View = View.Details;
            logListView.FullRowSelect = true;



            // Add the logListView to the form's Controls collection
            Controls.Add(logListView);

            var builder = new ConfigurationBuilder()
                .SetBasePath(Directory.GetCurrentDirectory())
                .AddJsonFile("appsettings.json", optional: true, reloadOnChange: true);

            _configuration = builder.Build();

        }

        private void menuStrip_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {
            System.Windows.Forms.ListView.SelectedListViewItemCollection selectedItems =
            listView.SelectedItems;
            if (e.ClickedItem.Text == "Copy")
            {
                String text = "";
                foreach (ListViewItem item in selectedItems)
                {
                    text += item.SubItems[1].Text;
                }
                Clipboard.SetText(text);
            }
        }


        private void Form1_Load(object sender, EventArgs e)
        {
            //this logListView references to your logListView define in Form1.cs, not from the task helper.cs
            // Call ConnectSAPAsync to get the SAP connectSQL
            ConnectSAP sapConnection = TaskHelper.ConnectSAP(_configuration, logListView);

            // Check if the SAP connectSQL is successful before proceeding
            if (sapConnection != null)
            {
                // Pass the SAP connectSQL to ConnectToDatabaseAsync
                TaskHelper.ConnectToDatabase(_configuration, sapConnection, logListView, textBox1);
            }
            else
            {
                // Handle failure to connect to SAP
                // Log or handle the error accordingly
            }

        }


    }
}