using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using QBFC13Lib;
using QuickBooksToLocalV2.Models;
using QuickBooksToLocalV2.QuickbooksDAL;

namespace QuickBooksToLocalV2.QuickbooksDAL
{
    //class QBinvoicelineitem
    //{
    //    bool sessionBegun = false;
    //    bool connectionOpen = false;
    //    QBSessionManager sessionManager = null;

    //    public void DbGetInvoices()
    //    {
    //        var qbe = new synncquickbooksEntities();
    //        var Invoicelineitem = new List<invoicelineitem>();

    //        try
    //        {
    //            //Create the session Manager object
    //            sessionManager = new QBSessionManager();

    //            //Create the message set request object to hold our request
    //            IMsgSetRequest requestMsgSet = sessionManager.CreateMsgSetRequest("US", 8, 0);
    //            requestMsgSet.Attributes.OnError = ENRqOnError.roeContinue;

    //            //Connect to QuickBooks and begin a session
    //            sessionManager.OpenConnection(@"", "Sync Quickbooks");
    //            connectionOpen = true;
    //            sessionManager.BeginSession(@"", ENOpenMode.omDontCare);
    //            sessionBegun = true;

    //            IInvoiceQuery invoiceQueryRq = requestMsgSet.AppendInvoiceQueryRq();

    //            //Send the request and get the response from QuickBooks
    //            IMsgSetResponse responseMsgSet = sessionManager.DoRequests(requestMsgSet);
    //            IResponse response = responseMsgSet.ResponseList.GetAt(0);
    //            IInvoiceRetList invoiceRetList = (IInvoiceRetList)response.Detail;



    //            if (invoiceRetList != null)
    //            {
    //                for (int i = 0; i < invoiceRetList.Count; i++)
    //                {
    //                    IInvoiceRet invoiceRet = invoiceRetList.GetAt(i);







    //                    Invoices.Add(Invoice);
    //                }
    //            }
    //            qbe.invoices.AddRange(Invoices);
    //            qbe.SaveChanges();
    //        }
    //        catch (Exception ex)
    //        {
    //            MessageBox.Show(ex.Message, "Error");
    //        }
    //        finally
    //        {
    //            //End the session and close the connection to QuickBooks
    //            if (sessionBegun)
    //            {
    //                sessionManager.EndSession();
    //            }
    //            if (connectionOpen)
    //            {
    //                sessionManager.CloseConnection();
    //            }
    //        }



    //    }
    //}
}
