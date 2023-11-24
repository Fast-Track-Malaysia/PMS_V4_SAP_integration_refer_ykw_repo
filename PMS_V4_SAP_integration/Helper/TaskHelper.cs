using Dapper;
using Microsoft.Extensions.Configuration;
using Newtonsoft.Json;
using PMS_V4_SAP_integration.Models;
using SAPbobsCOM;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PMS_V4_SAP_integration.Helper
{
    public static class TaskHelper
    {
        public static void Log(ListView logListView, string message)
        {
            // Get the current time
            string currentTime = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

            // Create ListViewItem with time and message
            ListViewItem item = new ListViewItem(new[] { currentTime, message });

            // Add the ListViewItem to the ListView
            logListView.Items.Add(item).EnsureVisible();

            // credit memo and commission go to ORIN table in SAP (in the table, the series referes to DOCNUM )
            // invoice go to OINV table in SAP
        }

        public static async Task ConnectToDatabaseAsync(IConfiguration _configuration, ConnectSAP connectSAP, ListView logListView, TextBox textboxView)
        {
            await Task.Delay(2000);
            Log(logListView, "Attempting to connect SQL database");
            var connectionString = _configuration.GetConnectionString("Default");

            try
            {
                using (var connection = new SqlConnection(connectionString))
                {
                    await connection.OpenAsync(); // Asynchronous connection open
                    Log(logListView, "Successfully connected to the database.");

                    // Pass the open connection to Execute_Invoice_ScriptAsync
                    Log(logListView, "Adding Invoice...");
                    await Execute_Invoice_ScriptAsync(connection, logListView, textboxView, connectSAP);

                    Log(logListView, "Adding Commission...");
                    await Execute_Commission_ScriptAsync(connection, logListView, textboxView);

                    Log(logListView, "Adding Credit Note...");
                    await Execute_CreditNote_ScriptAsync(connection, logListView, textboxView);

                    Log(logListView, "Process completed! ");
                    connectSAP.oCompany.Disconnect();

                    // Check if disconnected successfully
                    if (!connectSAP.oCompany.Connected)
                    {
                        Log(logListView, "Disconnected from SAP B1.");
                    }
                    else
                    {
                        Log(logListView, "Failed to disconnect from SAP B1.");
                    }
                }
            }
            catch (Exception ex)
            {
                Log(logListView, $"Error connecting to the database: {ex.Message}");
                return;
            }
        }

        public static async Task<ConnectSAP> ConnectSAPAsync(IConfiguration _configuration, ListView logListView)
        {
            try { await Task.Delay(2000);
                Log(logListView, "Attempting to connect SAP database");

                ConnectSAP connection = new ConnectSAP(_configuration);

                int connectResult = await Task.Run(() => connection.SAPConnect()); // Asynchronous execution of Connect method

                if (connectResult == 0)
                {
                    Log(logListView, "Successfully connected to the SAP database.");
                    return connection;
                }

                else if (connectResult == 100000085)
                {
                    Log(logListView, "Already log into SAP.");
                    return connection;
                }
                else
                {
                    Log(logListView, "Failed to connect to the SAP database. Connection Result: " + connectResult);
                    return null;
                }
            }
            catch (Exception ex)
            {
                // Log the exception message or details when the connection fails
                Log(logListView, $"Failed to connect to SAP: {ex.Message}");
                return null;
            }

        }

        //make all function accept ListView as a parameter, that way when you use log, you can able to access the list view properties
        public static async Task Execute_Invoice_ScriptAsync(SqlConnection connection, ListView logListView, TextBox textboxView, ConnectSAP connectSAP)
        {
            // First asynchronous operation
            await Task.Delay(1000);

            var query = "exec SP_invoicePosting";
            var jsonResult = connection.QueryFirstOrDefault<string>(query); // Retrieve the JSON string result

            if (!string.IsNullOrEmpty(jsonResult))
            {
                //Log(logListView, "Retrieved JSON:");
                //Log(logListView, jsonResult);

                // Deserialize the JSON string into a list of Invoice objects
                List<Invoice> invoices = JsonConvert.DeserializeObject<List<Invoice>>(jsonResult);

                // Accessing the deserialized data and logging
                foreach (var invoice in invoices)
                {
                    //insert SAP DI API here to put into sap.
                    Log(logListView, $"CardCode: {invoice.CardCode}, NumAtCard: {invoice.NumAtCard}");
                    textboxView.Text = "Invoice " + invoice.CardCode;
                    await Insert_Invoice_To_SAP_Async(invoice, logListView, connectSAP);
                }
            }
            else
            {
                Log(logListView, "No data retrieved from the stored procedure.");
            }
            // Further processing with the retrieved invoices if needed
        }

        public static async Task Execute_Commission_ScriptAsync(SqlConnection connection, ListView logListView, TextBox textboxView)
        {
            // First asynchronous operation
            await Task.Delay(1000);

            var query = "exec SP_commissionPosting";
            // query should check whether the postflag is 1, if already 1 that means already post, and to prevent duplicate.

            var jsonResult = connection.QueryFirstOrDefault<string>(query); // Retrieve the JSON string result

            if (!string.IsNullOrEmpty(jsonResult))
            {
                //Log(logListView, "Retrieved JSON:");
                //Log(logListView, jsonResult);

                // Deserialize the JSON string into a list of Invoice objects
                List<Commission> commissions = JsonConvert.DeserializeObject<List<Commission>>(jsonResult);

                // Accessing the deserialized data and logging
                foreach (var commission in commissions)
                {
                    Log(logListView, $"Commission: CardCode: {commission.CardCode}, NumAtCard: {commission.NumAtCard}");
                    textboxView.Text = "Commission " + commission.CardCode;
                    await Insert_Commission_To_SAP_Async(commission, logListView);
                }
            }
            else
            {
                Log(logListView, "No data retrieved from the stored procedure.");
            }
            // Further processing with the retrieved invoices if needed
        }

        public static async Task Execute_CreditNote_ScriptAsync(SqlConnection connection, ListView logListView, TextBox textboxView)
        {
            // First asynchronous operation
            await Task.Delay(1000);

            var query = "exec SP_creditnotePosting";
            var jsonResult = connection.QueryFirstOrDefault<string>(query); // Retrieve the JSON string result

            if (!string.IsNullOrEmpty(jsonResult))
            {
                //Log(logListView, "Retrieved JSON:");
                //Log(logListView, jsonResult);

                // Deserialize the JSON string into a list of Invoice objects
                List<CreditNote> creditNotes = JsonConvert.DeserializeObject<List<CreditNote>>(jsonResult);

                // Accessing the deserialized data and logging
                foreach (var creditNote in creditNotes)
                {
                    Log(logListView, $"CardCode: {creditNote.CardCode}, NumAtCard: {creditNote.NumAtCard}");
                    textboxView.Text = "Credit Note " + creditNote.CardCode;
                    await Insert_CreditNote_To_SAP_Async(creditNote, logListView);
                }
            }
            else
            {
                Log(logListView, "No data retrieved from the stored procedure.");
            }

        }

        public static async Task Insert_Invoice_To_SAP_Async(Invoice invoice, ListView logListView, ConnectSAP connectSAP)
        {
            await Task.Delay(100);
            Log(logListView, "Attempt to insert " + $"CardCode: {invoice.CardCode}" + " into SAP");
            //Documents oINVOICE = null;

            //oINVOICE = (Documents)connectSAP.oCompany.GetBusinessObject(BoObjectTypes.oInvoices);

            //oINVOICE.CardCode = invoice.CardCode;
            //oINVOICE.CardName = invoice.NumAtCard;


            //foreach (var line in invoice.Lines)
            //{
            //    //Log(logListView, $"ItemCode: {line.ItemCode}, Quantity: {line.Quantity}, UnitPrice: {line.UnitPrice}");
            //    //Log(logListView, "Attempt to insert " + $"ItemCode: {line.ItemCode}" + " into SAP");
            //    oINVOICE.Lines.ItemCode = line.ItemCode;
            //    oINVOICE.Lines.ItemDescription = line.ProjectCode;
            //    oINVOICE.Lines.UnitPrice = (double)line.UnitPrice;
            //    oINVOICE.Lines.Quantity = line.Quantity;
            //}

            Documents oINV = null;

            try
            {
                if (connectSAP.SAPConnect() == 0)
                {
                   
                    oINV = (Documents)connectSAP.oCompany.GetBusinessObject(BoObjectTypes.oInvoices);

                    oINV.CardCode = invoice.CardCode;
                    oINV.CardName = invoice.NumAtCard;
                        //add order items
                        foreach (var lines in invoice.Lines)
                        {

                        oINV.Lines.ItemCode = lines.ItemCode;
                        oINV.Lines.Quantity = (double)lines.Quantity;
                        oINV.Lines.Add();

                        }
                    
                    if (oINV.Add() != 0) //if 0 is success, else is fail
                    {

                        string errmsg = $"{connectSAP.oCompany.GetLastErrorCode()} - {connectSAP.oCompany.GetLastErrorDescription()}";
                        Log(logListView, errmsg);
                    }
                    else
                    {
                        Log(logListView,"successfully added invoice");
                    }
                }
                else
                {
                    // Return an error response if the connection fails
                    Log(logListView, "Failed to connect to SAP");
                }
            }
            finally
            {
                Marshal.ReleaseComObject(connectSAP); //this function exclusive for windows 
                connectSAP = null; //need to clear the object, or else will cause lag
            }

        }


        public static async Task Insert_Commission_To_SAP_Async(Commission commission, ListView logListView)
        {
            // First asynchronous operation
            await Task.Delay(100);
            Log(logListView, "Commission: Attempt to insert " + $"CardCode: {commission.CardCode}" + " into SAP");
            foreach (var line in commission.Lines)
            {
                //Log(logListView, $"Commission: ItemDescription: {line.ItemDescription}, Quantity: {line.Quantity}, UnitPrice: {line.UnitPrice}");
                Log(logListView, "Attempt to insert " + $"ItemDescription: {line.ItemDescription}" + " into SAP");
            }
        }

        public static async Task Insert_CreditNote_To_SAP_Async(CreditNote creditNote, ListView logListView)
        {
            // First asynchronous operation
            await Task.Delay(100);
            Log(logListView, "Attempt to insert " + $"CardCode: {creditNote.CardCode}" + " into SAP");
            foreach (var line in creditNote.Lines)
            {
                //Log(logListView, $"ItemCode: {line.ItemCode}, Quantity: {line.Quantity}, UnitPrice: {line.UnitPrice}");
                Log(logListView, "Attempt to insert " + $"ItemCode: {line.ItemCode}" + " into SAP");
            }
        }



    }
}
