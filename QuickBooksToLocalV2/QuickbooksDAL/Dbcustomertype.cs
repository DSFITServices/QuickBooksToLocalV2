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

        private Nullable<DateTime> _startDate;
        private Nullable<DateTime> _endDate;



        public Dbcustomertype(DateTime? startDate, DateTime? endDate)
        {
            if (startDate != null)
            {
                StartDate = startDate;
            }
            


            if (endDate != null)
            {
                EndDate = endDate;
            }
        }

        public DateTime? StartDate { get => _startDate; set => _startDate = value; }
        public DateTime? EndDate { get => _endDate; set => _endDate = value; }

        
        public ObservableCollection<customertype> DbRead()
        {
            QBSessionManager sessionManager = null;
            IMsgSetRequest requestMsgSet = null;

            try
            {
                //Create the session Manager object
                sessionManager = DbSession();

                //Create the message set request object to hold our request
                DbCreateMessageSet(ref sessionManager, ref requestMsgSet);

                //Connect to QuickBooks and begin a session
                DbConnect(ref sessionManager);


                ICustomerTypeQuery customerTypeQuery = requestMsgSet.AppendCustomerTypeQueryRq();

                customerTypeQuery.metaData.SetAsString("MetaDataAndResponseData");

                
                if (StartDate != null)
                {
                    customerTypeQuery.ORListQuery.ListFilter.FromModifiedDate.SetValue((DateTime)StartDate, false);
                }
                if (EndDate != null)
                {
                    customerTypeQuery.ORListQuery.ListFilter.ToModifiedDate.SetValue((DateTime)EndDate, false);
                }



                //Send the request and get the response from QuickBooks
                IResponse response = DbGetResponseSet(ref sessionManager, requestMsgSet);



                ICustomerTypeRetList customerTypeRetList = (ICustomerTypeRetList)response.Detail;
                
                ObservableCollection<customertype> Customertypes = new ObservableCollection<customertype>();
                
                if (customerTypeRetList != null)
                {
                    for (int i = 0; i < customerTypeRetList.Count; i++)
                    {
                        customertype Customertype = GetCustomerType(customerTypeRetList.GetAt(i));
                        Customertypes.Add(Customertype);
                    }
                }
                DbClose(ref sessionManager);
                return Customertypes;

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

        public customertype GetCustomerType(ICustomerTypeRet CustomerTypeRet)
        {
            var Customertype = new customertype();

            if (CustomerTypeRet == null) return null;

            //Go through all the elements of ICustomerTypeRetList
            //Get value of ListID
            Customertype.ID = (string)CustomerTypeRet.ListID.GetValue();
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
            if (CustomerTypeRet.Sublevel != null)
            {
                //Get value of Sublevel
                Customertype.Sublevel = (int)CustomerTypeRet.Sublevel.GetValue();
            }

            return Customertype;
        }
    }
}
