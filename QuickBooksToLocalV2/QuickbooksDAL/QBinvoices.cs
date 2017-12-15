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
    class QBinvoices
    {
        bool sessionBegun = false;
        bool connectionOpen = false;
        QBSessionManager sessionManager = null;

        public void DbGetInvoices()
        {
            var qbe = new synncquickbooksEntities();
            var Invoices = new List<invoice>();

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

                IInvoiceQuery invoiceQueryRq = requestMsgSet.AppendInvoiceQueryRq();

                //Send the request and get the response from QuickBooks
                IMsgSetResponse responseMsgSet = sessionManager.DoRequests(requestMsgSet);
                IResponse response = responseMsgSet.ResponseList.GetAt(0);
                IInvoiceRetList invoiceRetList = (IInvoiceRetList)response.Detail;



                if (invoiceRetList != null)
                {
                    for (int i = 0; i < invoiceRetList.Count; i++)
                    {
                        IInvoiceRet invoiceRet = invoiceRetList.GetAt(i);

                        var Invoice = new invoice();
                        if (invoiceRet == null) return;

                        //Go through all the elements of IInvoiceRetList
                        //Get value of TxnID
                        Invoice.ID = invoiceRet.TxnID.GetValue();
                        //Get value of TimeCreated
                        Invoice.TimeCreated = invoiceRet.TimeCreated.GetValue();
                        //Get value of TimeModified
                        Invoice.TimeModified = invoiceRet.TimeModified.GetValue();
                        //Get value of EditSequence
                        Invoice.EditSequence = invoiceRet.EditSequence.GetValue();
                        //Get value of TxnNumber
                        if (invoiceRet.TxnNumber != null)
                        {
                            Invoice.TxnNumber = invoiceRet.TxnNumber.GetValue();
                        }
                        //Get value of ListID
                        if (invoiceRet.CustomerRef.ListID != null)
                        {
                            Invoice.CustomerId = invoiceRet.CustomerRef.ListID.GetValue();
                        }
                        //Get value of FullName
                        if (invoiceRet.CustomerRef.FullName != null)
                        {
                            Invoice.CustomerName = invoiceRet.CustomerRef.FullName.GetValue();
                        }
                        if (invoiceRet.ClassRef != null)
                        {
                            //Get value of ListID
                            if (invoiceRet.ClassRef.ListID != null)
                            {
                                Invoice.ClassId = invoiceRet.ClassRef.ListID.GetValue();
                            }
                            //Get value of FullName
                            if (invoiceRet.ClassRef.FullName != null)
                            {
                                Invoice.Class = invoiceRet.ClassRef.FullName.GetValue();
                            }
                        }
                        if (invoiceRet.ARAccountRef != null)
                        {
                            //Get value of ListID
                            if (invoiceRet.ARAccountRef.ListID != null)
                            {
                                Invoice.AccountId = invoiceRet.ARAccountRef.ListID.GetValue();
                            }
                            //Get value of FullName
                            if (invoiceRet.ARAccountRef.FullName != null)
                            {
                                Invoice.Account = invoiceRet.ARAccountRef.FullName.GetValue();
                            }
                        }
                        if (invoiceRet.TemplateRef != null)
                        {
                            //Get value of ListID
                            if (invoiceRet.TemplateRef.ListID != null)
                            {
                                Invoice.TemplateId = invoiceRet.TemplateRef.ListID.GetValue();
                            }
                            //Get value of FullName
                            if (invoiceRet.TemplateRef.FullName != null)
                            {
                                Invoice.Template = invoiceRet.TemplateRef.FullName.GetValue();
                            }
                        }
                        //Get value of TxnDate
                        Invoice.Date = invoiceRet.TxnDate.GetValue();
                        //Get value of RefNumber
                        if (invoiceRet.RefNumber != null)
                        {
                            Invoice.ReferenceNumber = invoiceRet.RefNumber.GetValue();
                        }
                        if (invoiceRet.BillAddress != null)
                        {
                            //Get value of Addr1
                            if (invoiceRet.BillAddress.Addr1 != null)
                            {
                                Invoice.BillingLine1 = invoiceRet.BillAddress.Addr1.GetValue();
                            }
                            //Get value of Addr2
                            if (invoiceRet.BillAddress.Addr2 != null)
                            {
                                Invoice.BillingLine2 = invoiceRet.BillAddress.Addr2.GetValue();
                            }
                            //Get value of Addr3
                            if (invoiceRet.BillAddress.Addr3 != null)
                            {
                                Invoice.BillingLine3 = invoiceRet.BillAddress.Addr3.GetValue();
                            }
                            //Get value of Addr4
                            if (invoiceRet.BillAddress.Addr4 != null)
                            {
                                Invoice.BillingLine4 = invoiceRet.BillAddress.Addr4.GetValue();
                            }
                            //Get value of Addr5
                            if (invoiceRet.BillAddress.Addr5 != null)
                            {
                                Invoice.BillingLine5 = invoiceRet.BillAddress.Addr5.GetValue();
                            }
                            //Get value of City
                            if (invoiceRet.BillAddress.City != null)
                            {
                                Invoice.BillingCity = invoiceRet.BillAddress.City.GetValue();
                            }
                            //Get value of State
                            if (invoiceRet.BillAddress.State != null)
                            {
                                Invoice.BillingState = invoiceRet.BillAddress.State.GetValue();
                            }
                            //Get value of PostalCode
                            if (invoiceRet.BillAddress.PostalCode != null)
                            {
                                Invoice.BillingPostalCode = invoiceRet.BillAddress.PostalCode.GetValue();
                            }
                            //Get value of Country
                            if (invoiceRet.BillAddress.Country != null)
                            {
                                Invoice.BillingCountry = invoiceRet.BillAddress.Country.GetValue();
                            }
                            //Get value of Note
                            if (invoiceRet.BillAddress.Note != null)
                            {
                                Invoice.BillingNote = invoiceRet.BillAddress.Note.GetValue();
                            }
                        }
                        ////////if (invoiceRet.BillAddressBlock != null)
                        ////////{
                        ////////    //Get value of Addr1
                        ////////    if (invoiceRet.BillAddressBlock.Addr1 != null)
                        ////////    {
                        ////////        Invoice.Addr133 = invoiceRet.BillAddressBlock.Addr1.GetValue();
                        ////////    }
                        ////////    //Get value of Addr2
                        ////////    if (invoiceRet.BillAddressBlock.Addr2 != null)
                        ////////    {
                        ////////        Invoice.Addr234 = invoiceRet.BillAddressBlock.Addr2.GetValue();
                        ////////    }
                        ////////    //Get value of Addr3
                        ////////    if (invoiceRet.BillAddressBlock.Addr3 != null)
                        ////////    {
                        ////////        Invoice.Addr335 = invoiceRet.BillAddressBlock.Addr3.GetValue();
                        ////////    }
                        ////////    //Get value of Addr4
                        ////////    if (invoiceRet.BillAddressBlock.Addr4 != null)
                        ////////    {
                        ////////        Invoice.Addr436 = invoiceRet.BillAddressBlock.Addr4.GetValue();
                        ////////    }
                        ////////    //Get value of Addr5
                        ////////    if (invoiceRet.BillAddressBlock.Addr5 != null)
                        ////////    {
                        ////////        Invoice.Addr537 = invoiceRet.BillAddressBlock.Addr5.GetValue();
                        ////////    }
                        ////////}
                        if (invoiceRet.ShipAddress != null)
                        {
                            //Get value of Addr1
                            if (invoiceRet.ShipAddress.Addr1 != null)
                            {
                                Invoice.ShippingLine1 = invoiceRet.ShipAddress.Addr1.GetValue();
                            }
                            //Get value of Addr2
                            if (invoiceRet.ShipAddress.Addr2 != null)
                            {
                                Invoice.ShippingLine2 = invoiceRet.ShipAddress.Addr2.GetValue();
                            }
                            //Get value of Addr3
                            if (invoiceRet.ShipAddress.Addr3 != null)
                            {
                                Invoice.ShippingLine3 = invoiceRet.ShipAddress.Addr3.GetValue();
                            }
                            //Get value of Addr4
                            if (invoiceRet.ShipAddress.Addr4 != null)
                            {
                                Invoice.ShippingLine4 = invoiceRet.ShipAddress.Addr4.GetValue();
                            }
                            //Get value of Addr5
                            if (invoiceRet.ShipAddress.Addr5 != null)
                            {
                                Invoice.ShippingLine5 = invoiceRet.ShipAddress.Addr5.GetValue();
                            }
                            //Get value of City
                            if (invoiceRet.ShipAddress.City != null)
                            {
                                Invoice.ShippingCity = invoiceRet.ShipAddress.City.GetValue();
                            }
                            //Get value of State
                            if (invoiceRet.ShipAddress.State != null)
                            {
                                Invoice.ShippingState = invoiceRet.ShipAddress.State.GetValue();
                            }
                            //Get value of PostalCode
                            if (invoiceRet.ShipAddress.PostalCode != null)
                            {
                                Invoice.ShippingPostalCode = invoiceRet.ShipAddress.PostalCode.GetValue();
                            }
                            //Get value of Country
                            if (invoiceRet.ShipAddress.Country != null)
                            {
                                Invoice.ShippingCountry = invoiceRet.ShipAddress.Country.GetValue();
                            }
                            //Get value of Note
                            if (invoiceRet.ShipAddress.Note != null)
                            {
                                Invoice.ShippingNote = invoiceRet.ShipAddress.Note.GetValue();
                            }
                        }
                        //////////////if (invoiceRet.ShipAddressBlock != null)
                        //////////////{
                        //////////////    //Get value of Addr1
                        //////////////    if (invoiceRet.ShipAddressBlock.Addr1 != null)
                        //////////////    {
                        //////////////        Invoice.Addr148 = invoiceRet.ShipAddressBlock.Addr1.GetValue();
                        //////////////    }
                        //////////////    //Get value of Addr2
                        //////////////    if (invoiceRet.ShipAddressBlock.Addr2 != null)
                        //////////////    {
                        //////////////        Invoice.Addr249 = invoiceRet.ShipAddressBlock.Addr2.GetValue();
                        //////////////    }
                        //////////////    //Get value of Addr3
                        //////////////    if (invoiceRet.ShipAddressBlock.Addr3 != null)
                        //////////////    {
                        //////////////        Invoice.Addr350 = invoiceRet.ShipAddressBlock.Addr3.GetValue();
                        //////////////    }
                        //////////////    //Get value of Addr4
                        //////////////    if (invoiceRet.ShipAddressBlock.Addr4 != null)
                        //////////////    {
                        //////////////        Invoice.Addr451 = invoiceRet.ShipAddressBlock.Addr4.GetValue();
                        //////////////    }
                        //////////////    //Get value of Addr5
                        //////////////    if (invoiceRet.ShipAddressBlock.Addr5 != null)
                        //////////////    {
                        //////////////        Invoice.Addr552 = invoiceRet.ShipAddressBlock.Addr5.GetValue();
                        //////////////    }
                        //////////////}
                        //Get value of IsPending
                        if (invoiceRet.IsPending != null)
                        {
                            Invoice.IsPending = invoiceRet.IsPending.GetValue();
                        }
                        //Get value of IsFinanceCharge
                        //////if (invoiceRet.IsFinanceCharge != null)
                        //////{
                        //////    Invoice.IsFinanceCharge = invoiceRet.IsFinanceCharge.ToString();
                        //////}
                        //Get value of PONumber
                        if (invoiceRet.PONumber != null)
                        {
                            Invoice.POnumber = invoiceRet.PONumber.GetValue();
                        }
                        if (invoiceRet.TermsRef != null)
                        {
                            //Get value of ListID
                            if (invoiceRet.TermsRef.ListID != null)
                            {
                                Invoice.TermsId = invoiceRet.TermsRef.ListID.GetValue();
                            }
                            //Get value of FullName
                            if (invoiceRet.TermsRef.FullName != null)
                            {
                                Invoice.Terms = invoiceRet.TermsRef.FullName.GetValue();
                            }
                        }
                        //Get value of DueDate
                        if (invoiceRet.DueDate != null)
                        {
                            Invoice.DueDate = invoiceRet.DueDate.GetValue();
                        }
                        if (invoiceRet.SalesRepRef != null)
                        {
                            //Get value of ListID
                            if (invoiceRet.SalesRepRef.ListID != null)
                            {
                                Invoice.SalesRepId = invoiceRet.SalesRepRef.ListID.GetValue();
                            }
                            //Get value of FullName
                            if (invoiceRet.SalesRepRef.FullName != null)
                            {
                                Invoice.SalesRep = invoiceRet.SalesRepRef.FullName.GetValue();
                            }
                        }
                        //Get value of FOB
                        if (invoiceRet.FOB != null)
                        {
                            Invoice.FOB = invoiceRet.FOB.GetValue();
                        }
                        //Get value of ShipDate
                        if (invoiceRet.ShipDate != null)
                        {
                            Invoice.ShipDate = invoiceRet.ShipDate.GetValue();
                        }
                        if (invoiceRet.ShipMethodRef != null)
                        {
                            //Get value of ListID
                            if (invoiceRet.ShipMethodRef.ListID != null)
                            {
                                Invoice.ShipMethodId = invoiceRet.ShipMethodRef.ListID.GetValue();
                            }
                            //Get value of FullName
                            if (invoiceRet.ShipMethodRef.FullName != null)
                            {
                                Invoice.ShipMethod = invoiceRet.ShipMethodRef.FullName.GetValue();
                            }
                        }
                        //Get value of Subtotal
                        if (invoiceRet.Subtotal != null)
                        {
                            Invoice.Subtotal = invoiceRet.Subtotal.GetValue();
                        }
                        if (invoiceRet.ItemSalesTaxRef != null)
                        {
                            //Get value of ListID
                            if (invoiceRet.ItemSalesTaxRef.ListID != null)
                            {
                                Invoice.TaxItemId = invoiceRet.ItemSalesTaxRef.ListID.GetValue();
                            }
                            //Get value of FullName
                            if (invoiceRet.ItemSalesTaxRef.FullName != null)
                            {
                                Invoice.TaxItem = invoiceRet.ItemSalesTaxRef.FullName.GetValue();
                            }
                        }
                        //Get value of SalesTaxPercentage
                        if (invoiceRet.SalesTaxPercentage != null)
                        {
                            Invoice.TaxPercent = invoiceRet.SalesTaxPercentage.GetValue();
                        }
                        //Get value of SalesTaxTotal
                        if (invoiceRet.SalesTaxTotal != null)
                        {
                            Invoice.Tax = invoiceRet.SalesTaxTotal.GetValue();
                        }
                        //Get value of AppliedAmount
                        if (invoiceRet.AppliedAmount != null)
                        {
                            Invoice.AppliedAmount = invoiceRet.AppliedAmount.GetValue();
                        }
                        //Get value of BalanceRemaining
                        if (invoiceRet.BalanceRemaining != null)
                        {
                            Invoice.Balance = invoiceRet.BalanceRemaining.GetValue();
                        }
                        if (invoiceRet.CurrencyRef != null)
                        {
                            //Get value of ListID
                            if (invoiceRet.CurrencyRef.ListID != null)
                            {
                                Invoice.CurrencyId = invoiceRet.CurrencyRef.ListID.GetValue();
                            }
                            //Get value of FullName
                            if (invoiceRet.CurrencyRef.FullName != null)
                            {
                                Invoice.CurrencyName = invoiceRet.CurrencyRef.FullName.GetValue();
                            }
                        }
                        //Get value of ExchangeRate
                        if (invoiceRet.ExchangeRate != null)
                        {
                            Invoice.ExchangeRate = invoiceRet.ExchangeRate.GetValue();
                        }
                        //Get value of BalanceRemainingInHomeCurrency
                        if (invoiceRet.BalanceRemainingInHomeCurrency != null)
                        {
                            Invoice.BalanceInHomeCurrency = invoiceRet.BalanceRemainingInHomeCurrency.GetValue();
                        }
                        //Get value of Memo
                        if (invoiceRet.Memo != null)
                        {
                            Invoice.Memo = invoiceRet.Memo.GetValue();
                        }
                        //Get value of IsPaid
                        if (invoiceRet.IsPaid != null)
                        {
                            Invoice.IsPaid = invoiceRet.IsPaid.GetValue();
                        }
                        if (invoiceRet.CustomerMsgRef != null)
                        {
                            //Get value of ListID
                            if (invoiceRet.CustomerMsgRef.ListID != null)
                            {
                                Invoice.MessageId = invoiceRet.CustomerMsgRef.ListID.GetValue();
                            }
                            //Get value of FullName
                            if (invoiceRet.CustomerMsgRef.FullName != null)
                            {
                                Invoice.Message = invoiceRet.CustomerMsgRef.FullName.GetValue();
                            }
                        }
                        //Get value of IsToBePrinted
                        if (invoiceRet.IsToBePrinted != null)
                        {
                            Invoice.IsToBePrinted = invoiceRet.IsToBePrinted.GetValue();
                        }
                        //Get value of IsToBeEmailed
                        if (invoiceRet.IsToBeEmailed != null)
                        {
                            Invoice.IsToBeEmailed = invoiceRet.IsToBeEmailed.GetValue();
                        }
                        if (invoiceRet.CustomerSalesTaxCodeRef != null)
                        {
                            //Get value of ListID
                            if (invoiceRet.CustomerSalesTaxCodeRef.ListID != null)
                            {
                                Invoice.CustomerTaxCodeId = invoiceRet.CustomerSalesTaxCodeRef.ListID.GetValue();
                            }
                            //Get value of FullName
                            if (invoiceRet.CustomerSalesTaxCodeRef.FullName != null)
                            {
                                Invoice.CustomerTaxCode = invoiceRet.CustomerSalesTaxCodeRef.FullName.GetValue();
                            }
                        }
                        //Get value of SuggestedDiscountAmount
                        if (invoiceRet.SuggestedDiscountAmount != null)
                        {
                            Invoice.SuggestedDiscountAmount = invoiceRet.SuggestedDiscountAmount.GetValue();
                        }
                        //Get value of SuggestedDiscountDate
                        if (invoiceRet.SuggestedDiscountDate != null)
                        {
                            Invoice.SuggestedDiscountDate = invoiceRet.SuggestedDiscountDate.GetValue();
                        }
                        //Get value of Other
                        if (invoiceRet.Other != null)
                        {
                            Invoice.Other = invoiceRet.Other.GetValue();
                        }
                        //Get value of ExternalGUID
                        //////if (invoiceRet.ExternalGUID != null)
                        //////{
                        //////    Invoice.ExternalGUID87 = invoiceRet.ExternalGUID.GetValue();
                        //////}
                        //////if (invoiceRet.LinkedTxnList != null)
                        //////{
                        //////    for (int i = 0; i < invoiceRet.LinkedTxnList.Count; i++)
                        //////    {
                        //////        ILinkedTxn LinkedTxn = invoiceRet.LinkedTxnList.GetAt(i);
                        //////        //Get value of TxnID
                        //////        Invoice.L = LinkedTxn.TxnID.GetValue();
                        //////        //Get value of TxnType
                        //////        ENTxnType TxnType90 = (ENTxnType)LinkedTxn.TxnType.GetValue();
                        //////        //Get value of TxnDate
                        //////        Invoice.TxnDate91 = LinkedTxn.TxnDate.GetValue();
                        //////        //Get value of RefNumber
                        //////        if (LinkedTxn.RefNumber != null)
                        //////        {
                        //////            Invoice.RefNumber92 = LinkedTxn.RefNumber.GetValue();
                        //////        }
                        //////        //Get value of LinkType
                        //////        if (LinkedTxn.LinkType != null)
                        //////        {
                        //////            ENLinkType LinkType93 = (ENLinkType)LinkedTxn.LinkType.GetValue();
                        //////        }
                        //////        //Get value of Amount
                        //////        Invoice.Amount94 = LinkedTxn.Amount.GetValue();
                        //////    }
                        //////}
                        //////if (invoiceRet.ORInvoiceLineRetList != null)
                        //////{
                        //////    for (int i95 = 0; i95 < invoiceRet.ORInvoiceLineRetList.Count; i95++)
                        //////    {
                        //////        IORInvoiceLineRet ORInvoiceLineRet = invoiceRet.ORInvoiceLineRetList.GetAt(i95);
                        //////        if (ORInvoiceLineRet.InvoiceLineRet != null)
                        //////        {
                        //////            if (ORInvoiceLineRet.InvoiceLineRet != null)
                        //////            {
                        //////                //Get value of TxnLineID
                        //////                Invoice.TxnLineID96 = ORInvoiceLineRet.InvoiceLineRet.TxnLineID.GetValue();
                        //////                if (ORInvoiceLineRet.InvoiceLineRet.ItemRef != null)
                        //////                {
                        //////                    //Get value of ListID
                        //////                    if (ORInvoiceLineRet.InvoiceLineRet.ItemRef.ListID != null)
                        //////                    {
                        //////                        Invoice.ListID97 = ORInvoiceLineRet.InvoiceLineRet.ItemRef.ListID.GetValue();
                        //////                    }
                        //////                    //Get value of FullName
                        //////                    if (ORInvoiceLineRet.InvoiceLineRet.ItemRef.FullName != null)
                        //////                    {
                        //////                        Invoice.FullName98 = ORInvoiceLineRet.InvoiceLineRet.ItemRef.FullName.GetValue();
                        //////                    }
                        //////                }
                        //////                //Get value of Desc
                        //////                if (ORInvoiceLineRet.InvoiceLineRet.Desc != null)
                        //////                {
                        //////                    Invoice.Desc99 = ORInvoiceLineRet.InvoiceLineRet.Desc.GetValue();
                        //////                }
                        //////                //Get value of Quantity
                        //////                if (ORInvoiceLineRet.InvoiceLineRet.Quantity != null)
                        //////                {
                        //////                    Invoice.Quantity100 = ORInvoiceLineRet.InvoiceLineRet.Quantity.GetValue();
                        //////                }
                        //////                //Get value of UnitOfMeasure
                        //////                if (ORInvoiceLineRet.InvoiceLineRet.UnitOfMeasure != null)
                        //////                {
                        //////                    Invoice.UnitOfMeasure101 = ORInvoiceLineRet.InvoiceLineRet.UnitOfMeasure.GetValue();
                        //////                }
                        //////                if (ORInvoiceLineRet.InvoiceLineRet.OverrideUOMSetRef != null)
                        //////                {
                        //////                    //Get value of ListID
                        //////                    if (ORInvoiceLineRet.InvoiceLineRet.OverrideUOMSetRef.ListID != null)
                        //////                    {
                        //////                        Invoice.ListID102 = ORInvoiceLineRet.InvoiceLineRet.OverrideUOMSetRef.ListID.GetValue();
                        //////                    }
                        //////                    //Get value of FullName
                        //////                    if (ORInvoiceLineRet.InvoiceLineRet.OverrideUOMSetRef.FullName != null)
                        //////                    {
                        //////                        Invoice.FullName103 = ORInvoiceLineRet.InvoiceLineRet.OverrideUOMSetRef.FullName.GetValue();
                        //////                    }
                        //////                }
                        //////                if (ORInvoiceLineRet.InvoiceLineRet.ORRate != null)
                        //////                {
                        //////                    if (ORInvoiceLineRet.InvoiceLineRet.ORRate.Rate != null)
                        //////                    {
                        //////                        //Get value of Rate
                        //////                        if (ORInvoiceLineRet.InvoiceLineRet.ORRate.Rate != null)
                        //////                        {
                        //////                            Invoice.Rate104 = ORInvoiceLineRet.InvoiceLineRet.ORRate.Rate.GetValue();
                        //////                        }
                        //////                    }
                        //////                    if (ORInvoiceLineRet.InvoiceLineRet.ORRate.RatePercent != null)
                        //////                    {
                        //////                        //Get value of RatePercent
                        //////                        if (ORInvoiceLineRet.InvoiceLineRet.ORRate.RatePercent != null)
                        //////                        {
                        //////                            Invoice.RatePercent105 = ORInvoiceLineRet.InvoiceLineRet.ORRate.RatePercent.GetValue();
                        //////                        }
                        //////                    }
                        //////                }
                        //////                if (ORInvoiceLineRet.InvoiceLineRet.ClassRef != null)
                        //////                {
                        //////                    //Get value of ListID
                        //////                    if (ORInvoiceLineRet.InvoiceLineRet.ClassRef.ListID != null)
                        //////                    {
                        //////                        Invoice.ListID106 = ORInvoiceLineRet.InvoiceLineRet.ClassRef.ListID.GetValue();
                        //////                    }
                        //////                    //Get value of FullName
                        //////                    if (ORInvoiceLineRet.InvoiceLineRet.ClassRef.FullName != null)
                        //////                    {
                        //////                        Invoice.FullName107 = ORInvoiceLineRet.InvoiceLineRet.ClassRef.FullName.GetValue();
                        //////                    }
                        //////                }
                        //////                //Get value of Amount
                        //////                if (ORInvoiceLineRet.InvoiceLineRet.Amount != null)
                        //////                {
                        //////                    Invoice.Amount108 = ORInvoiceLineRet.InvoiceLineRet.Amount.GetValue();
                        //////                }
                        //////                if (ORInvoiceLineRet.InvoiceLineRet.InventorySiteRef != null)
                        //////                {
                        //////                    //Get value of ListID
                        //////                    if (ORInvoiceLineRet.InvoiceLineRet.InventorySiteRef.ListID != null)
                        //////                    {
                        //////                        Invoice.ListID109 = ORInvoiceLineRet.InvoiceLineRet.InventorySiteRef.ListID.GetValue();
                        //////                    }
                        //////                    //Get value of FullName
                        //////                    if (ORInvoiceLineRet.InvoiceLineRet.InventorySiteRef.FullName != null)
                        //////                    {
                        //////                        Invoice.FullName110 = ORInvoiceLineRet.InvoiceLineRet.InventorySiteRef.FullName.GetValue();
                        //////                    }
                        //////                }
                        //////                if (ORInvoiceLineRet.InvoiceLineRet.InventorySiteLocationRef != null)
                        //////                {
                        //////                    //Get value of ListID
                        //////                    if (ORInvoiceLineRet.InvoiceLineRet.InventorySiteLocationRef.ListID != null)
                        //////                    {
                        //////                        Invoice.ListID111 = ORInvoiceLineRet.InvoiceLineRet.InventorySiteLocationRef.ListID.GetValue();
                        //////                    }
                        //////                    //Get value of FullName
                        //////                    if (ORInvoiceLineRet.InvoiceLineRet.InventorySiteLocationRef.FullName != null)
                        //////                    {
                        //////                        Invoice.FullName112 = ORInvoiceLineRet.InvoiceLineRet.InventorySiteLocationRef.FullName.GetValue();
                        //////                    }
                        //////                }
                        //////                if (ORInvoiceLineRet.InvoiceLineRet.ORSerialLotNumber != null)
                        //////                {
                        //////                    if (ORInvoiceLineRet.InvoiceLineRet.ORSerialLotNumber.SerialNumber != null)
                        //////                    {
                        //////                        //Get value of SerialNumber
                        //////                        if (ORInvoiceLineRet.InvoiceLineRet.ORSerialLotNumber.SerialNumber != null)
                        //////                        {
                        //////                            Invoice.SerialNumber113 = ORInvoiceLineRet.InvoiceLineRet.ORSerialLotNumber.SerialNumber.GetValue();
                        //////                        }
                        //////                    }
                        //////                    if (ORInvoiceLineRet.InvoiceLineRet.ORSerialLotNumber.LotNumber != null)
                        //////                    {
                        //////                        //Get value of LotNumber
                        //////                        if (ORInvoiceLineRet.InvoiceLineRet.ORSerialLotNumber.LotNumber != null)
                        //////                        {
                        //////                            Invoice.LotNumber114 = ORInvoiceLineRet.InvoiceLineRet.ORSerialLotNumber.LotNumber.GetValue();
                        //////                        }
                        //////                    }
                        //////                }
                        //////                //Get value of ServiceDate
                        //////                if (ORInvoiceLineRet.InvoiceLineRet.ServiceDate != null)
                        //////                {
                        //////                    Invoice.ServiceDate115 = ORInvoiceLineRet.InvoiceLineRet.ServiceDate.GetValue();
                        //////                }
                        //////                if (ORInvoiceLineRet.InvoiceLineRet.SalesTaxCodeRef != null)
                        //////                {
                        //////                    //Get value of ListID
                        //////                    if (ORInvoiceLineRet.InvoiceLineRet.SalesTaxCodeRef.ListID != null)
                        //////                    {
                        //////                        Invoice.ListID116 = ORInvoiceLineRet.InvoiceLineRet.SalesTaxCodeRef.ListID.GetValue();
                        //////                    }
                        //////                    //Get value of FullName
                        //////                    if (ORInvoiceLineRet.InvoiceLineRet.SalesTaxCodeRef.FullName != null)
                        //////                    {
                        //////                        Invoice.FullName117 = ORInvoiceLineRet.InvoiceLineRet.SalesTaxCodeRef.FullName.GetValue();
                        //////                    }
                        //////                }
                        //////                //Get value of Other1
                        //////                if (ORInvoiceLineRet.InvoiceLineRet.Other1 != null)
                        //////                {
                        //////                    Invoice.Other1118 = ORInvoiceLineRet.InvoiceLineRet.Other1.GetValue();
                        //////                }
                        //////                //Get value of Other2
                        //////                if (ORInvoiceLineRet.InvoiceLineRet.Other2 != null)
                        //////                {
                        //////                    Invoice.Other2119 = ORInvoiceLineRet.InvoiceLineRet.Other2.GetValue();
                        //////                }
                        //////                if (ORInvoiceLineRet.InvoiceLineRet.DataExtRetList != null)
                        //////                {
                        //////                    for (Invoice.i120 = 0; i120 < ORInvoiceLineRet.InvoiceLineRet.DataExtRetList.Count; i120++)
                        //////                    {
                        //////                        IDataExtRet DataExtRet = ORInvoiceLineRet.InvoiceLineRet.DataExtRetList.GetAt(i120);
                        //////                        //Get value of OwnerID
                        //////                        if (DataExtRet.OwnerID != null)
                        //////                        {
                        //////                            Invoice.OwnerID121 = DataExtRet.OwnerID.GetValue();
                        //////                        }
                        //////                        //Get value of DataExtName
                        //////                        Invoice.DataExtName122 = DataExtRet.DataExtName.GetValue();
                        //////                        //Get value of DataExtType
                        //////                        ENDataExtType DataExtType123 = (ENDataExtType)DataExtRet.DataExtType.GetValue();
                        //////                        //Get value of DataExtValue
                        //////                        Invoice.DataExtValue124 = DataExtRet.DataExtValue.GetValue();
                        //////                    }
                        //////                }
                        //////            }
                        //////        }
                        //////        if (ORInvoiceLineRet.InvoiceLineGroupRet != null)
                        //////        {
                        //////            if (ORInvoiceLineRet.InvoiceLineGroupRet != null)
                        //////            {
                        //////                //Get value of TxnLineID
                        //////                Invoice.TxnLineID125 = ORInvoiceLineRet.InvoiceLineGroupRet.TxnLineID.GetValue();
                        //////                //Get value of ListID
                        //////                if (ORInvoiceLineRet.InvoiceLineGroupRet.ItemGroupRef.ListID != null)
                        //////                {
                        //////                    Invoice.ListID126 = ORInvoiceLineRet.InvoiceLineGroupRet.ItemGroupRef.ListID.GetValue();
                        //////                }
                        //////                //Get value of FullName
                        //////                if (ORInvoiceLineRet.InvoiceLineGroupRet.ItemGroupRef.FullName != null)
                        //////                {
                        //////                    Invoice.FullName127 = ORInvoiceLineRet.InvoiceLineGroupRet.ItemGroupRef.FullName.GetValue();
                        //////                }
                        //////                //Get value of Desc
                        //////                if (ORInvoiceLineRet.InvoiceLineGroupRet.Desc != null)
                        //////                {
                        //////                    Invoice.Desc128 = ORInvoiceLineRet.InvoiceLineGroupRet.Desc.GetValue();
                        //////                }
                        //////                //Get value of Quantity
                        //////                if (ORInvoiceLineRet.InvoiceLineGroupRet.Quantity != null)
                        //////                {
                        //////                    Invoice.Quantity129 = ORInvoiceLineRet.InvoiceLineGroupRet.Quantity.GetValue();
                        //////                }
                        //////                //Get value of UnitOfMeasure
                        //////                if (ORInvoiceLineRet.InvoiceLineGroupRet.UnitOfMeasure != null)
                        //////                {
                        //////                    Invoice.UnitOfMeasure130 = ORInvoiceLineRet.InvoiceLineGroupRet.UnitOfMeasure.GetValue();
                        //////                }
                        //////                if (ORInvoiceLineRet.InvoiceLineGroupRet.OverrideUOMSetRef != null)
                        //////                {
                        //////                    //Get value of ListID
                        //////                    if (ORInvoiceLineRet.InvoiceLineGroupRet.OverrideUOMSetRef.ListID != null)
                        //////                    {
                        //////                        Invoice.ListID131 = ORInvoiceLineRet.InvoiceLineGroupRet.OverrideUOMSetRef.ListID.GetValue();
                        //////                    }
                        //////                    //Get value of FullName
                        //////                    if (ORInvoiceLineRet.InvoiceLineGroupRet.OverrideUOMSetRef.FullName != null)
                        //////                    {
                        //////                        Invoice.FullName132 = ORInvoiceLineRet.InvoiceLineGroupRet.OverrideUOMSetRef.FullName.GetValue();
                        //////                    }
                        //////                }
                        //////                //Get value of IsPrintItemsInGroup
                        //////                Invoice.IsPrintItemsInGroup133 = ORInvoiceLineRet.InvoiceLineGroupRet.IsPrintItemsInGroup.GetValue();
                        //////                //Get value of TotalAmount
                        //////                Invoice.TotalAmount134 = ORInvoiceLineRet.InvoiceLineGroupRet.TotalAmount.GetValue();
                        //////                if (ORInvoiceLineRet.InvoiceLineGroupRet.InvoiceLineRetList != null)
                        //////                {
                        //////                    for (int i135 = 0; i135 < ORInvoiceLineRet.InvoiceLineGroupRet.InvoiceLineRetList.Count; i135++)
                        //////                    {
                        //////                        IInvoiceLineRet InvoiceLineRet = ORInvoiceLineRet.InvoiceLineGroupRet.InvoiceLineRetList.GetAt(i135);
                        //////                        //Get value of TxnLineID
                        //////                        Invoice.TxnLineID136 = InvoiceLineRet.TxnLineID.GetValue();
                        //////                        if (InvoiceLineRet.ItemRef != null)
                        //////                        {
                        //////                            //Get value of ListID
                        //////                            if (InvoiceLineRet.ItemRef.ListID != null)
                        //////                            {
                        //////                                Invoice.ListID137 = InvoiceLineRet.ItemRef.ListID.GetValue();
                        //////                            }
                        //////                            //Get value of FullName
                        //////                            if (InvoiceLineRet.ItemRef.FullName != null)
                        //////                            {
                        //////                                Invoice.FullName138 = InvoiceLineRet.ItemRef.FullName.GetValue();
                        //////                            }
                        //////                        }
                        //////                        //Get value of Desc
                        //////                        if (InvoiceLineRet.Desc != null)
                        //////                        {
                        //////                            Invoice.Desc139 = InvoiceLineRet.Desc.GetValue();
                        //////                        }
                        //////                        //Get value of Quantity
                        //////                        if (InvoiceLineRet.Quantity != null)
                        //////                        {
                        //////                            Invoice.Quantity140 = InvoiceLineRet.Quantity.GetValue();
                        //////                        }
                        //////                        //Get value of UnitOfMeasure
                        //////                        if (InvoiceLineRet.UnitOfMeasure != null)
                        //////                        {
                        //////                            Invoice.UnitOfMeasure141 = InvoiceLineRet.UnitOfMeasure.GetValue();
                        //////                        }
                        //////                        if (InvoiceLineRet.OverrideUOMSetRef != null)
                        //////                        {
                        //////                            //Get value of ListID
                        //////                            if (InvoiceLineRet.OverrideUOMSetRef.ListID != null)
                        //////                            {
                        //////                                Invoice.ListID142 = InvoiceLineRet.OverrideUOMSetRef.ListID.GetValue();
                        //////                            }
                        //////                            //Get value of FullName
                        //////                            if (InvoiceLineRet.OverrideUOMSetRef.FullName != null)
                        //////                            {
                        //////                                Invoice.FullName143 = InvoiceLineRet.OverrideUOMSetRef.FullName.GetValue();
                        //////                            }
                        //////                        }
                        //////                        if (InvoiceLineRet.ORRate != null)
                        //////                        {
                        //////                            if (InvoiceLineRet.ORRate.Rate != null)
                        //////                            {
                        //////                                //Get value of Rate
                        //////                                if (InvoiceLineRet.ORRate.Rate != null)
                        //////                                {
                        //////                                    Invoice.Rate144 = InvoiceLineRet.ORRate.Rate.GetValue();
                        //////                                }
                        //////                            }
                        //////                            if (InvoiceLineRet.ORRate.RatePercent != null)
                        //////                            {
                        //////                                //Get value of RatePercent
                        //////                                if (InvoiceLineRet.ORRate.RatePercent != null)
                        //////                                {
                        //////                                    Invoice.RatePercent145 = InvoiceLineRet.ORRate.RatePercent.GetValue();
                        //////                                }
                        //////                            }
                        //////                        }
                        //////                        if (InvoiceLineRet.ClassRef != null)
                        //////                        {
                        //////                            //Get value of ListID
                        //////                            if (InvoiceLineRet.ClassRef.ListID != null)
                        //////                            {
                        //////                                Invoice.ListID146 = InvoiceLineRet.ClassRef.ListID.GetValue();
                        //////                            }
                        //////                            //Get value of FullName
                        //////                            if (InvoiceLineRet.ClassRef.FullName != null)
                        //////                            {
                        //////                                Invoice.FullName147 = InvoiceLineRet.ClassRef.FullName.GetValue();
                        //////                            }
                        //////                        }
                        //////                        //Get value of Amount
                        //////                        if (InvoiceLineRet.Amount != null)
                        //////                        {
                        //////                            Invoice.Amount148 = InvoiceLineRet.Amount.GetValue();
                        //////                        }
                        //////                        if (InvoiceLineRet.InventorySiteRef != null)
                        //////                        {
                        //////                            //Get value of ListID
                        //////                            if (InvoiceLineRet.InventorySiteRef.ListID != null)
                        //////                            {
                        //////                                Invoice.ListID149 = InvoiceLineRet.InventorySiteRef.ListID.GetValue();
                        //////                            }
                        //////                            //Get value of FullName
                        //////                            if (InvoiceLineRet.InventorySiteRef.FullName != null)
                        //////                            {
                        //////                                Invoice.FullName150 = InvoiceLineRet.InventorySiteRef.FullName.GetValue();
                        //////                            }
                        //////                        }
                        //////                        if (InvoiceLineRet.InventorySiteLocationRef != null)
                        //////                        {
                        //////                            //Get value of ListID
                        //////                            if (InvoiceLineRet.InventorySiteLocationRef.ListID != null)
                        //////                            {
                        //////                                Invoice.ListID151 = InvoiceLineRet.InventorySiteLocationRef.ListID.GetValue();
                        //////                            }
                        //////                            //Get value of FullName
                        //////                            if (InvoiceLineRet.InventorySiteLocationRef.FullName != null)
                        //////                            {
                        //////                                Invoice.FullName152 = InvoiceLineRet.InventorySiteLocationRef.FullName.GetValue();
                        //////                            }
                        //////                        }
                        //////                        if (InvoiceLineRet.ORSerialLotNumber != null)
                        //////                        {
                        //////                            if (InvoiceLineRet.ORSerialLotNumber.SerialNumber != null)
                        //////                            {
                        //////                                //Get value of SerialNumber
                        //////                                if (InvoiceLineRet.ORSerialLotNumber.SerialNumber != null)
                        //////                                {
                        //////                                    Invoice.SerialNumber153 = InvoiceLineRet.ORSerialLotNumber.SerialNumber.GetValue();
                        //////                                }
                        //////                            }
                        //////                            if (InvoiceLineRet.ORSerialLotNumber.LotNumber != null)
                        //////                            {
                        //////                                //Get value of LotNumber
                        //////                                if (InvoiceLineRet.ORSerialLotNumber.LotNumber != null)
                        //////                                {
                        //////                                    Invoice.LotNumber154 = InvoiceLineRet.ORSerialLotNumber.LotNumber.GetValue();
                        //////                                }
                        //////                            }
                        //////                        }
                        //////                        //Get value of ServiceDate
                        //////                        if (InvoiceLineRet.ServiceDate != null)
                        //////                        {
                        //////                            Invoice.ServiceDate155 = InvoiceLineRet.ServiceDate.GetValue();
                        //////                        }
                        //////                        if (InvoiceLineRet.SalesTaxCodeRef != null)
                        //////                        {
                        //////                            //Get value of ListID
                        //////                            if (InvoiceLineRet.SalesTaxCodeRef.ListID != null)
                        //////                            {
                        //////                                Invoice.ListID156 = InvoiceLineRet.SalesTaxCodeRef.ListID.GetValue();
                        //////                            }
                        //////                            //Get value of FullName
                        //////                            if (InvoiceLineRet.SalesTaxCodeRef.FullName != null)
                        //////                            {
                        //////                                Invoice.FullName157 = InvoiceLineRet.SalesTaxCodeRef.FullName.GetValue();
                        //////                            }
                        //////                        }
                        //////                        //Get value of Other1
                        //////                        if (InvoiceLineRet.Other1 != null)
                        //////                        {
                        //////                            Invoice.Other1158 = InvoiceLineRet.Other1.GetValue();
                        //////                        }
                        //////                        //Get value of Other2
                        //////                        if (InvoiceLineRet.Other2 != null)
                        //////                        {
                        //////                            Invoice.Other2159 = InvoiceLineRet.Other2.GetValue();
                        //////                        }
                        //////                        if (InvoiceLineRet.DataExtRetList != null)
                        //////                        {
                        //////                            for (int i160 = 0; i160 < InvoiceLineRet.DataExtRetList.Count; i160++)
                        //////                            {
                        //////                                IDataExtRet DataExtRet = InvoiceLineRet.DataExtRetList.GetAt(i160);
                        //////                                //Get value of OwnerID
                        //////                                if (DataExtRet.OwnerID != null)
                        //////                                {
                        //////                                    Invoice.OwnerID161 = DataExtRet.OwnerID.GetValue();
                        //////                                }
                        //////                                //Get value of DataExtName
                        //////                                Invoice.DataExtName162 = DataExtRet.DataExtName.GetValue();
                        //////                                //Get value of DataExtType
                        //////                                ENDataExtType DataExtType163 = (ENDataExtType)DataExtRet.DataExtType.GetValue();
                        //////                                //Get value of DataExtValue
                        //////                                Invoice.DataExtValue164 = DataExtRet.DataExtValue.GetValue();
                        //////                            }
                        //////                        }
                        //////                    }
                        //////                }
                        //////                if (ORInvoiceLineRet.InvoiceLineGroupRet.DataExtRetList != null)
                        //////                {
                        //////                    for (int i165 = 0; i165 < ORInvoiceLineRet.InvoiceLineGroupRet.DataExtRetList.Count; i165++)
                        //////                    {
                        //////                        IDataExtRet DataExtRet = ORInvoiceLineRet.InvoiceLineGroupRet.DataExtRetList.GetAt(i165);
                        //////                        //Get value of OwnerID
                        //////                        if (DataExtRet.OwnerID != null)
                        //////                        {
                        //////                            Invoice.OwnerID166 = DataExtRet.OwnerID.GetValue();
                        //////                        }
                        //////                        //Get value of DataExtName
                        //////                        Invoice.DataExtName167 = DataExtRet.DataExtName.GetValue();
                        //////                        //Get value of DataExtType
                        //////                        ENDataExtType DataExtType168 = (ENDataExtType)DataExtRet.DataExtType.GetValue();
                        //////                        //Get value of DataExtValue
                        //////                        Invoice.DataExtValue169 = DataExtRet.DataExtValue.GetValue();
                        //////                    }
                        //////                }
                        //////            }
                        //////        }
                        //////    }
                        //////}
                        ////if (invoiceRet.DataExtRetList != null)
                        ////{
                        ////    for (int i = 0; i < invoiceRet.DataExtRetList.Count; i++)
                        ////    {
                        ////        IDataExtRet DataExtRet = invoiceRet.DataExtRetList.GetAt(i);
                        ////        //Get value of OwnerID
                        ////        //Get value of DataExtName
                        ////        Invoice.DataExtName172 = DataExtRet.DataExtName.GetValue();
                        ////        //Get value of DataExtType
                        ////        ENDataExtType DataExtType173 = (ENDataExtType)DataExtRet.DataExtType.GetValue();
                        ////        //Get value of DataExtValue
                        ////        Invoice.DataExtValue174 = DataExtRet.DataExtValue.GetValue();
                        ////    }
                        ////}






                        Invoices.Add(Invoice);
                    }
                }
                qbe.invoices.AddRange(Invoices);
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

