using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using QBFC13Lib;
using QuickBooksToLocalV2;
using QuickBooksToLocalV2.Interfaces;
using QuickBooksToLocalV2.Models;
using QuickBooksToLocalV2.QuickbooksDAL.CommonDAL;


namespace QuickBooksToLocalV2.QuickbooksDAL.CommonDAL
{
    public class DBCommon : DbConnection
    {
        public QBSessionManager DbSession()
        {
            return new QBSessionManager();
        }

        public void DbCreateMessageSet(ref QBSessionManager sessionManager, ref IMsgSetRequest requestMsgSet)
        {
            //Create the message set request object to hold our request
            requestMsgSet = sessionManager.CreateMsgSetRequest("UK", 10, 0);
            requestMsgSet.Attributes.OnError = ENRqOnError.roeContinue;
        }

        public IResponse DbGetResponseSet(ref QBSessionManager sessionManager, IMsgSetRequest requestMsgSet)
        {
            //Send the request and get the response from QuickBooks
            IMsgSetResponse responseMsgSet = sessionManager.DoRequests(requestMsgSet);
            return responseMsgSet.ResponseList.GetAt(0);
        }


    }
}
