using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using QBFC13Lib;
using QuickBooksToLocalV2.Models;

namespace QuickBooksToLocalV2.QuickbooksDAL
{
    class QBcustomerTypes
    {
        bool sessionBegun = false;
        bool connectionOpen = false;
        QBSessionManager sessionManager = null;

        public void DbGetCustomerTypes()
        {
            qbEntities qbe = new qbEntities();
            var CustomerTypes = new List<customertype>();

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

                ICustomerTypeQuery customerTypeQueryRq = requestMsgSet.AppendCustomerTypeQueryRq();

                //Send the request and get the response from QuickBooks
                IMsgSetResponse responseMsgSet = sessionManager.DoRequests(requestMsgSet);
                IResponse response = responseMsgSet.ResponseList.GetAt(0);
                ICustomerTypeRetList customerTypeRetList = (ICustomerTypeRetList)response.Detail;

                

                if (customerTypeRetList != null)
                {
                    for (int i = 0; i < customerTypeRetList.Count; i++)
                    {
                        ICustomerTypeRet customerTypeRet = customerTypeRetList.GetAt(i);

                        var CustomerType = new customertype();

                        CustomerType.ID = customerTypeRet.ListID.GetValue();
                        CustomerType.Name = customerTypeRet.Name.GetValue();
                        CustomerType.IsActive = customerTypeRet.IsActive.GetValue();
                        CustomerType.FullName = customerTypeRet.FullName.GetValue();
                        CustomerType.TimeCreated = customerTypeRet.TimeCreated.GetValue();
                        CustomerType.TimeModified = customerTypeRet.TimeModified.GetValue();
                        if(customerTypeRet.ParentRef != null)
                        {
                            if(customerTypeRet.ParentRef.ListID != null)
                            {
                                CustomerType.ParentId = customerTypeRet.ParentRef.ListID.GetValue();
                            }
                            if(customerTypeRet.ParentRef.FullName != null)
                            {
                                CustomerType.ParentName = customerTypeRet.ParentRef.FullName.GetValue();
                            }
                        }
                        CustomerType.EditSequence = customerTypeRet.EditSequence.GetValue();
                        



                        CustomerTypes.Add(CustomerType);
                    }
                }
                qbe.customertypes.AddRange(CustomerTypes);
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
