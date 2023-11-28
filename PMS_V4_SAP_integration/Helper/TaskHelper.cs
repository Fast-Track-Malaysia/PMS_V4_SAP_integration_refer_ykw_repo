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

        //connect SQL
        public static async Task ConnectToDatabaseAsync(IConfiguration _configuration, ConnectSAP connectSAP, ListView logListView, TextBox textboxView)
        {
            await Task.Delay(2000);
            Log(logListView, "Attempting to connect SQL database");
            var connectionString = _configuration.GetConnectionString("Default");

            try
            {
                using (var connectSQL = new SqlConnection(connectionString))
                {
                    await connectSQL.OpenAsync(); // Asynchronous connectSQL open
                    Log(logListView, "Successfully connected to the database.");

                    // Pass the open connectSQL to Execute_Invoice_ScriptAsync
                    Log(logListView, "Adding Invoice...");
                    await Execute_Invoice_ScriptAsync(connectSQL, logListView, textboxView, connectSAP);  //disabled for testing
                    
                    Log(logListView, "Adding Commission...");
                    await Execute_Commission_ScriptAsync(connectSQL, logListView, textboxView, connectSAP);

                    Log(logListView, "Adding Credit Note...");
                    await Execute_CreditNote_ScriptAsync(connectSQL, logListView, textboxView, connectSAP);

                    Log(logListView, "Process completed! ");
                    textboxView.Text = "Process completed!";
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

        //connect SAP
        public static async Task<ConnectSAP> ConnectSAPAsync(IConfiguration _configuration, ListView logListView)
        {
            try { await Task.Delay(2000);
                Log(logListView, "Attempting to connect SAP database");

                ConnectSAP connectSQL = new ConnectSAP(_configuration);

                int connectResult = await Task.Run(() => connectSQL.SAPConnect()); // Asynchronous execution of Connect method

                if (connectResult == 0)
                {
                    Log(logListView, "Successfully connected to the SAP database.");
                    return connectSQL;
                }

                else if (connectResult == 100000085)
                {
                    Log(logListView, "Already log into SAP.");
                    return connectSQL;
                }
                else
                {
                    Log(logListView, "Failed to connect to the SAP database. Connection Result: " + connectResult);
                    return null;
                }
            }
            catch (Exception ex)
            {
                // Log the exception message or details when the connectSQL fails
                Log(logListView, $"Failed to connect to SAP: {ex.Message}");
                return null;
            }

        }

        //make all function accept ListView as a parameter, that way when you use log, you can able to access the list view properties

        //Exec Invoice
        public static async Task Execute_Invoice_ScriptAsync(SqlConnection connectSQL, ListView logListView, TextBox textboxView, ConnectSAP connectSAP)
        {
            // First asynchronous operation
            await Task.Delay(1000);

            var query = "exec SP_invoicePosting";
            var jsonResult = connectSQL.QueryFirstOrDefault<string>(query); // Retrieve the JSON string result

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
                    //Log(logListView, $"CardCode: {invoice.CardCode}, NumAtCard: {invoice.NumAtCard}");
                    textboxView.Text = "Invoice " + invoice.CardCode;
                    await Insert_Invoice_To_SAP_Async(invoice, logListView, connectSAP, connectSQL);
                }
            }
            else
            {
                Log(logListView, "No data retrieved from the stored procedure.");
            }
            // Further processing with the retrieved invoices if needed
        }

        //Exec Commission
        public static async Task Execute_Commission_ScriptAsync(SqlConnection connectSQL, ListView logListView, TextBox textboxView, ConnectSAP connectSAP)
        {
            // First asynchronous operation
            await Task.Delay(1000);

            var query = "exec SP_commissionPosting";
            // query should check whether the postflag is 1, if already 1 that means already post, and to prevent duplicate.

            var jsonResult = connectSQL.QueryFirstOrDefault<string>(query); // Retrieve the JSON string result

            if (!string.IsNullOrEmpty(jsonResult))
            {
                //Log(logListView, "Retrieved JSON:");
                //Log(logListView, jsonResult);

                // Deserialize the JSON string into a list of Invoice objects
                List<Commission> commissions = JsonConvert.DeserializeObject<List<Commission>>(jsonResult);

                // Accessing the deserialized data and logging
                foreach (var commission in commissions)
                {
                    //Log(logListView, $"Commission: CardCode: {commission.CardCode}, NumAtCard: {commission.NumAtCard}");
                    textboxView.Text = "Commission " + commission.CardCode;
                    await Insert_Commission_To_SAP_Async(commission, logListView, connectSAP, connectSQL);
                }
            }
            else
            {
                Log(logListView, "No data retrieved from the stored procedure.");
            }
            // Further processing with the retrieved invoices if needed
        }

        //Exec Credit Note
        public static async Task Execute_CreditNote_ScriptAsync(SqlConnection connectSQL, ListView logListView, TextBox textboxView, ConnectSAP connectSAP)
        {
            await Task.Delay(1000);

            var query = "exec SP_creditnotePosting"; //i exchange SP for testing
            var jsonResult = connectSQL.QueryFirstOrDefault<string>(query); // Retrieve the JSON string result

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
                    await Insert_CreditNote_To_SAP_Async(creditNote, logListView, connectSAP, connectSQL);
                }
            }
            else
            {
                Log(logListView, "No data retrieved from the stored procedure.");
            }

        }

        //Add invoice to SAP
        public static async Task Insert_Invoice_To_SAP_Async(Invoice invoice, ListView logListView, ConnectSAP connectSAP, SqlConnection connectSQL)
        {
            await Task.Delay(100);
            Log(logListView, "Inserting " + $"Card Code: {invoice.CardCode}" + $" Document Number: {invoice.NumAtCard}" + " into SAP");
            Documents oINV = null;

            try
            {             
                oINV = (Documents)connectSAP.oCompany.GetBusinessObject(BoObjectTypes.oInvoices);
                
                oINV.CardCode = invoice.CardCode;
                oINV.Comments = invoice.Comments;
                oINV.DocDate = invoice.DocDate;
                oINV.DocDueDate = invoice.DocDueDate;  // disabled dates for testing as sample data exceed posting periods
                oINV.NumAtCard = invoice.NumAtCard;
                oINV.Project = invoice.Project;
                oINV.Series = invoice.Series;          
                oINV.TaxDate = invoice.TaxDate;

                //add lines
                foreach (var lines in invoice.Lines)
                {

                oINV.Lines.ItemCode = lines.ItemCode;
                oINV.Lines.ProjectCode = lines.ProjectCode;
                oINV.Lines.VatGroup = lines.VatGroup;
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

                connectSAP.oCompany.GetNewObjectCode(out string docEntry);
                    Log(logListView,"Successfully added invoice. DocEntry: " + docEntry);
                    // run the post flag update script here, you need to pass a parameter though
                    await Update_Invoice_PostFlag_Async(connectSQL, logListView, invoice);
                }
                
            }
            finally
            {
                //Marshal.ReleaseComObject(connectSAP); //this function exclusive for windows 
                
                connectSAP = null; //need to clear the object, or else will cause lag
            }

        }
        //Add Commission to SAP
        public static async Task Insert_Commission_To_SAP_Async(Commission commission, ListView logListView, ConnectSAP connectSAP, SqlConnection connectSQL)
        {
            await Task.Delay(100);
            Log(logListView, "Inserting " + $"CardCode: {commission.CardCode} " + $" Document Number: {commission.NumAtCard}" + " into SAP");
            Documents oCOM = null;

            try
            {
                oCOM = (Documents)connectSAP.oCompany.GetBusinessObject(BoObjectTypes.oCreditNotes);

                oCOM.CardCode = commission.CardCode;
                oCOM.Comments = commission.Comments;
                oCOM.DocDate = commission.DocDate;
                oCOM.DocDueDate = commission.DocDueDate;
                oCOM.DocType = commission.DocType;
                oCOM.HandWritten = commission.Handwritten;
                oCOM.NumAtCard = commission.NumAtCard;
                oCOM.Project = commission.Project;
                oCOM.Series = commission.Series;
                oCOM.TaxDate = commission.TaxDate;

                //add lines
                //int line = -1;
                foreach (var lines in commission.Lines)
                {
                    //line++;
                    //if (line > 0) 
                    //    oCOM.Lines.Add();
                    //oCOM.Lines.SetCurrentLine(line);
                    oCOM.Lines.ItemDescription = lines.ItemDescription;
                    oCOM.Lines.UnitPrice = lines.UnitPrice;
                    oCOM.Lines.AccountCode = lines.AccountCode;  //SAP requires Account code insert to lines
                    oCOM.Lines.ProjectCode = lines.ProjectCode;
                    oCOM.Lines.VatGroup = lines.VatGroup;
                    //oCOM.Lines.Quantity = lines.Quantity;  //quantity not require in service document
                    oCOM.Lines.UserFields.Fields.Item("U_FRef").Value = lines.U_FRef; //user defined fields

                }

                if (oCOM.Add() != 0) //if 0 is success, else is fail
                {

                    string errmsg = $"{connectSAP.oCompany.GetLastErrorCode()} - {connectSAP.oCompany.GetLastErrorDescription()}";
                    Log(logListView, errmsg);
                }
                else
                {

                    connectSAP.oCompany.GetNewObjectCode(out string docEntry);
                    Log(logListView, "Successfully added commission. DocEntry: " + docEntry);
                    await Update_Commission_PostFlag_Async(connectSQL, logListView, commission);
                }

            }
            finally
            {
                //Marshal.ReleaseComObject(connectSAP); //this function exclusive for windows 
                connectSAP = null; //need to clear the object, or else will cause lag
            }

        }

        //Add CreditNote to SAP
        public static async Task Insert_CreditNote_To_SAP_Async(CreditNote creditNote, ListView logListView, ConnectSAP connectSAP, SqlConnection connectSQL)
        {
            await Task.Delay(100);
            Log(logListView, "Inserting " + $"CardCode: {creditNote.CardCode} " + $" Document Number: {creditNote.NumAtCard}" + " into SAP");
            Documents oCOM = null;

            try
            {
                oCOM = (Documents)connectSAP.oCompany.GetBusinessObject(BoObjectTypes.oCreditNotes);

                oCOM.CardCode = creditNote.CardCode;
                oCOM.Comments = creditNote.Comments;
                oCOM.DocDate = creditNote.DocDate;
                oCOM.DocDueDate = creditNote.DocDueDate;
                oCOM.NumAtCard = creditNote.NumAtCard;
                oCOM.Project = creditNote.Project;
                oCOM.Series = creditNote.Series;
                oCOM.TaxDate = creditNote.TaxDate;

                //add lines
                foreach (var lines in creditNote.Lines)
                {

                    oCOM.Lines.ItemCode = lines.ItemCode;
                    oCOM.Lines.UnitPrice = lines.UnitPrice;
                    oCOM.Lines.ProjectCode = lines.ProjectCode;
                    oCOM.Lines.VatGroup = lines.VatGroup; 
                    oCOM.Lines.Quantity = lines.Quantity;
                    oCOM.Lines.UserFields.Fields.Item("U_FRef").Value = "0";//lines.U_FRef; //user defined fields
                    oCOM.Lines.Add();

                }

                if (oCOM.Add() != 0) //if 0 is success, else is fail
                {

                    string errmsg = $"{connectSAP.oCompany.GetLastErrorCode()} - {connectSAP.oCompany.GetLastErrorDescription()}";
                    Log(logListView, errmsg);
                }
                else
                {

                    connectSAP.oCompany.GetNewObjectCode(out string docEntry);
                    Log(logListView, "Successfully added credit note. DocEntry: " + docEntry);
                    await Update_CreditNote_PostFlag_Async(connectSQL, logListView, creditNote);
                }

            }
            finally
            {
                //Marshal.ReleaseComObject(connectSAP); //this function exclusive for windows 
                connectSAP = null; //need to clear the object, or else will cause lag
            }

        }

        // get some sample credit note rows using this query below, this is for testing only
        //        select* from creditnote order by sk_hdr desc
        //update creditnote set postflag = 0 where sk_hdr >= 213

        //credit note dont have UREF
        //commission have UREF

        //set post flag to 1 after posting //updated kw script to retrieve SKHDR
        public static async Task Update_Invoice_PostFlag_Async(SqlConnection connectSQL, ListView logListView, Invoice invoice)
        {
            await Task.Delay(100);

            var query = "update invoice set postflag = 1 where postflag = 0 and chkflag = 1 and sk_hdr = @SKHDR"; //i exchange SP for testing
            connectSQL.Query(query, new { SKHDR = invoice.sk_hdr });
            Log(logListView, "Invoice postflag sk_hdr: " + invoice.sk_hdr + " updated.");
        }

        public static async Task Update_Commission_PostFlag_Async(SqlConnection connectSQL, ListView logListView, Commission commission)
        {
            await Task.Delay(100);

            var query = "update comm_iv set postflag = 1 where postflag = 0 and chkflag = 1 and sk_hdr = @SKHDR"; //i exchange SP for testing
            connectSQL.Query(query, new { SKHDR = commission.sk_hdr });
            Log(logListView, "Commission postflag sk_hdr: " + commission.sk_hdr + " updated.");
        }

        public static async Task Update_CreditNote_PostFlag_Async(SqlConnection connectSQL, ListView logListView, CreditNote creditNote)
        {
            await Task.Delay(100);

            var query = "update creditnote set postflag = 1 where postflag = 0 and chkflag = 1 and sk_hdr = @SKHDR"; //i exchange SP for testing
            connectSQL.Query(query, new { SKHDR = creditNote.sk_hdr });
            Log(logListView, "Credit Note postflag sk_hdr: " + creditNote.sk_hdr + " updated.");
        }

    }
}
