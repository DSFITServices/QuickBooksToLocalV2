using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using QBFC13Lib;
using QuickBooksToLocalV2.Interfaces;
using QuickBooksToLocalV2.Models;
using QuickBooksToLocalV2.QuickbooksDAL.CommonDAL;


namespace QuickBooksToLocalV2.QuickbooksDAL
{
    public class Dbcustomertype : DBCommon , IDBCurd<customertype>
    {


        public ObservableCollection<customertype> DbRead(Nullable<DateTime> startDate, Nullable<DateTime> endDate)
        {
            QBSessionManager sessionManager = null;

            try
            {
                //Create the session Manager object
                sessionManager = new QBSessionManager();

                //Create the message set request object to hold our request
                IMsgSetRequest requestMsgSet = sessionManager.CreateMsgSetRequest("UK", 10, 0);
                requestMsgSet.Attributes.OnError = ENRqOnError.roeContinue;

                //Connect to QuickBooks and begin a session
                sessionManager.OpenConnection(@"", "synncquickbooks");
                connectionOpen = true;
                sessionManager.BeginSession(@"", ENOpenMode.omDontCare);
                sessionBegun = true;


                ICustomerTypeQuery customerTypeQuery = requestMsgSet.AppendCustomerTypeQueryRq();

                customerTypeQuery.metaData.SetAsString("MetaDataAndResponseData");

                if (startDate == null) { startDate = new DateTime(2010, 01, 01, 00, 00, 00); }
                if (endDate == null) { endDate = DateTime.Now; }

                if (startDate != null)
                {
                    customerTypeQuery.ORListQuery.ListFilter.FromModifiedDate.SetValue((DateTime)startDate, false);
                }
                if (endDate != null)
                {
                    customerTypeQuery.ORListQuery.ListFilter.ToModifiedDate.SetValue((DateTime)endDate, false);
                }















                //Send the request and get the response from QuickBooks
                IMsgSetResponse responseMsgSet = sessionManager.DoRequests(requestMsgSet);
                IResponse response = responseMsgSet.ResponseList.GetAt(0);



                ICustomerTypeRetList customerTypeRetList = (ICustomerTypeRetList)response.Detail;



                ObservableCollection<customertype> Customertypes = new ObservableCollection<customertype>();


                if (customerTypeRetList != null)
                {
                    for (int i = 0; i < customerTypeRetList.Count; i++)
                    {
                        ICustomerTypeRet CustomerTypeRet = customerTypeRetList.GetAt(i);

                        var Customertype = new customertype();




                        if (CustomerTypeRet == null) return null;

                        //Go through all the elements of ICustomerTypeRetList
                        //Get value of ListID
                        Customertype.ListID = (string)CustomerTypeRet.ListID.GetValue();
                        //Get value of TimeCreated
                        Customertype.TimeCreated = (DateTime)CustomerTypeRet.TimeCreated.GetValue();
                        //Get value of TimeModified
                        Customertype.TimeModified = (DateTime)CustomerTypeRet.TimeModified.GetValue();
                        //Get value of EditSequence
                        Customertype.EditSequence = (string)CustomerTypeRet.EditSequence.GetValue();
                        //Get value of Name
                        Customertype.Name = (string)CustomerTypeRet.Name.GetValue();
                        //Get value of FullName
                        Customertype.FullName = (string)CustomerTypeRet.FullName.GetValue();
                        //Get value of IsActive
                        if (CustomerTypeRet.IsActive != null)
                        {
                            Customertype.IsActive = (bool)CustomerTypeRet.IsActive.GetValue();
                        }
                        if (CustomerTypeRet.ParentRef != null)
                        {
                            //Get value of ListID
                            if (CustomerTypeRet.ParentRef.ListID != null)
                            {
                                Customertype.ParentId = (string)CustomerTypeRet.ParentRef.ListID.GetValue();
                            }
                            //Get value of FullName
                            if (CustomerTypeRet.ParentRef.FullName != null)
                            {
                                Customertype.ParentName = (string)CustomerTypeRet.ParentRef.FullName.GetValue();
                            }
                        }
                        //Get value of Sublevel
                        //Customertype. = (int)CustomerTypeRet.Sublevel.GetValue();









                        Customertypes.Add(Customertype);
                    }
                }
                DbClose(ref sessionManager);
                return Customertypes;

                //synncquickbooksEntities oContext = new synncquickbooksEntities();

                //oContext.customertypes.AddRange(Customertypes);
                //oContext.SaveChanges();





            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error");
                DbClose(ref sessionManager);
                return null;
            }
            finally
            {
                DbClose(ref sessionManager);
            }











        }



    }
}
