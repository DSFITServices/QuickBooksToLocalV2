using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Threading.Tasks;
using QBFC13Lib;
using QuickBooksToLocalV2.Models;
using QuickBooksToLocalV2.QuickbooksDAL;

namespace QuickBooksToLocalV2.QuickbooksDAL
{
    class QBcustomers
    {
        bool sessionBegun = false;
        bool connectionOpen = false;
        QBSessionManager sessionManager = null;

        public void DbGetCustomers()
        {
            var qbe = new synncquickbooksEntities();
            var Customers = new List<customer>();

            try
            {
                //Create the session Manager object
                sessionManager = new QBSessionManager();

                //Create the message set request object to hold our request
                IMsgSetRequest requestMsgSet = sessionManager.CreateMsgSetRequest("US", 8, 0);
                requestMsgSet.Attributes.OnError = ENRqOnError.roeContinue;

                //Connect to QuickBooks and begin a session
                sessionManager.OpenConnection(@"", "Sync Quickbooks");
                connectionOpen = true;
                sessionManager.BeginSession(@"", ENOpenMode.omDontCare);
                sessionBegun = true;

                ICustomerQuery customerQueryRq = requestMsgSet.AppendCustomerQueryRq();

                //Send the request and get the response from QuickBooks
                IMsgSetResponse responseMsgSet = sessionManager.DoRequests(requestMsgSet);
                IResponse response = responseMsgSet.ResponseList.GetAt(0);
                ICustomerRetList customerRetList = (ICustomerRetList)response.Detail;



                if (customerRetList != null)
                {
                    for (int i = 0; i < customerRetList.Count; i++)
                    {
                        ICustomerRet customerRet = customerRetList.GetAt(i);

                        var Customer = new customer();

                        Customer.ID = customerRet.ListID.GetValue();
                        Customer.Name = customerRet.Name.GetValue();
                        Customer.IsActive = customerRet.IsActive.GetValue();
                        Customer.FullName = customerRet.FullName.GetValue();
                        Customer.TimeCreated = customerRet.TimeCreated.GetValue();
                        Customer.TimeModified = customerRet.TimeModified.GetValue();
                        if (customerRet.ParentRef != null)
                        {
                            if (customerRet.ParentRef.ListID != null)
                            {
                                Customer.ParentId = customerRet.ParentRef.ListID.GetValue();
                            }
                            if (customerRet.ParentRef.FullName != null)
                            {
                                Customer.ParentName = customerRet.ParentRef.FullName.GetValue();
                            }
                        }
                        Customer.EditSequence = customerRet.EditSequence.GetValue();
                        if(customerRet.CustomerTypeRef != null)
                        {
                            if(customerRet.CustomerTypeRef.ListID != null)
                            {
                                Customer.TypeId = customerRet.CustomerTypeRef.ListID.GetValue();
                            }
                            if (customerRet.CustomerTypeRef.FullName != null)
                            {
                                Customer.Type = customerRet.CustomerTypeRef.FullName.GetValue();
                            }
                        }



                        Customers.Add(Customer);
                    }
                }
                qbe.customers.AddRange(Customers);
                qbe.SaveChanges();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error");
            }
            finally
            {
                //End the session and close the connection to QuickBooks
                if (sessionBegun)
                {
                    sessionManager.EndSession();
                }
                if (connectionOpen)
                {
                    sessionManager.CloseConnection();
                }
            }



        }
    }
}
