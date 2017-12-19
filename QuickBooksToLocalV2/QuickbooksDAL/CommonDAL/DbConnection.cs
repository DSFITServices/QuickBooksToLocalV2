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
using QuickBooksToLocalV2.QuickbooksDAL;
using QuickBooksToLocalV2.QuickbooksDAL.CommonDAL;

namespace QuickBooksToLocalV2.QuickbooksDAL
{
    public class DbConnection : IConnect<QBSessionManager>
    {
        public bool connectionOpen = false;
        public bool sessionBegun = false;

        public void DbConnect(ref QBSessionManager sessionManager)
        {
            //Connect to QuickBooks and begin a session
            sessionManager.OpenConnection(@"", "synncquickbooks");
            connectionOpen = true;
            sessionManager.BeginSession(@"", ENOpenMode.omDontCare);
            sessionBegun = true;
        }

        public bool DbClose(ref QBSessionManager sessionManager)
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
            return true;
        }
    }
}
