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
                        Invoice.ID = (string)invoiceRet.TxnID.GetValue();
                        //Get value of TimeCreated
                        Invoice.TimeCreated = (DateTime)invoiceRet.TimeCreated.GetValue();
                        //Get value of TimeModified
                        Invoice.TimeModified = (DateTime)invoiceRet.TimeModified.GetValue();
                        //Get value of EditSequence
                        Invoice.EditSequence = (string)invoiceRet.EditSequence.GetValue();
                        //Get value of TxnNumber
                        if (invoiceRet.TxnNumber != null)
                        {
                            Invoice.TxnNumber = (int)invoiceRet.TxnNumber.GetValue();
                        }
                        //Get value of ListID
                        if (invoiceRet.CustomerRef.ListID != null)
                        {
                            Invoice.CustomerId = (string)invoiceRet.CustomerRef.ListID.GetValue();
                        }
                        //Get value of FullName
                        if (invoiceRet.CustomerRef.FullName != null)
                        {
                            Invoice.CustomerName = (string)invoiceRet.CustomerRef.FullName.GetValue();
                        }
                        if (invoiceRet.ClassRef != null)
                        {
                            //Get value of ListID
                            if (invoiceRet.ClassRef.ListID != null)
                            {
                                Invoice.ClassId = (string)invoiceRet.ClassRef.ListID.GetValue();
                            }
                            //Get value of FullName
                            if (invoiceRet.ClassRef.FullName != null)
                            {
                                Invoice.Class = (string)invoiceRet.ClassRef.FullName.GetValue();
                            }
                        }
                        if (invoiceRet.ARAccountRef != null)
                        {
                            //Get value of ListID
                            if (invoiceRet.ARAccountRef.ListID != null)
                            {
                                Invoice.AccountId = (string)invoiceRet.ARAccountRef.ListID.GetValue();
                            }
                            //Get value of FullName
                            if (invoiceRet.ARAccountRef.FullName != null)
                            {
                                Invoice.Account = (string)InvoiceRet.ARAccountRef.FullName.GetValue();
                            }
                        }
                        if (invoiceRet.TemplateRef != null)
                        {
                            //Get value of ListID
                            if (invoiceRet.TemplateRef.ListID != null)
                            {
                                Invoice.TemplateId = (string)invoiceRet.TemplateRef.ListID.GetValue();
                            }
                            //Get value of FullName
                            if (invoiceRet.TemplateRef.FullName != null)
                            {
                                Invoice.Template = (string)invoiceRet.TemplateRef.FullName.GetValue();
                            }
                        }
                        //Get value of TxnDate
                        Invoice.Date = (DateTime)invoiceRet.TxnDate.GetValue();
                        //Get value of RefNumber
                        if (invoiceRet.RefNumber != null)
                        {
                            Invoice.ReferenceNumber = (string)invoiceRet.RefNumber.GetValue();
                        }
                        if (invoiceRet.BillAddress != null)
                        {
                            //Get value of Addr1
                            if (invoiceRet.BillAddress.Addr1 != null)
                            {
                                Invoice.BillingLine1 = (string)invoiceRet.BillAddress.Addr1.GetValue();
                            }
                            //Get value of Addr2
                            if (invoiceRet.BillAddress.Addr2 != null)
                            {
                                Invoice.BillingLine2 = (string)invoiceRet.BillAddress.Addr2.GetValue();
                            }
                            //Get value of Addr3
                            if (invoiceRet.BillAddress.Addr3 != null)
                            {
                                Invoice.BillingLine3 = (string)invoiceRet.BillAddress.Addr3.GetValue();
                            }
                            //Get value of Addr4
                            if (invoiceRet.BillAddress.Addr4 != null)
                            {
                                Invoice.BillingLine4 = (string)invoiceRet.BillAddress.Addr4.GetValue();
                            }
                            //Get value of Addr5
                            if (invoiceRet.BillAddress.Addr5 != null)
                            {
                                Invoice.BillingLine5 = (string)invoiceRet.BillAddress.Addr5.GetValue();
                            }
                            //Get value of City
                            if (invoiceRet.BillAddress.City != null)
                            {
                                Invoice.BillingCity = (string)invoiceRet.BillAddress.City.GetValue();
                            }
                            //Get value of State
                            if (invoiceRet.BillAddress.State != null)
                            {
                                Invoice.BillingState = (string)invoiceRet.BillAddress.State.GetValue();
                            }
                            //Get value of PostalCode
                            if (invoiceRet.BillAddress.PostalCode != null)
                            {
                                Invoice.BillingPostalCode = (string)invoiceRet.BillAddress.PostalCode.GetValue();
                            }
                            //Get value of Country
                            if (invoiceRet.BillAddress.Country != null)
                            {
                                Invoice.BillingCountry = (string)invoiceRet.BillAddress.Country.GetValue();
                            }
                            //Get value of Note
                            if (invoiceRet.BillAddress.Note != null)
                            {
                                Invoice.BillingNote = (string)invoiceRet.BillAddress.Note.GetValue();
                            }
                        }
                        //////if (invoiceRet.BillAddressBlock != null)
                        //////{
                        //////    //Get value of Addr1
                        //////    if (invoiceRet.BillAddressBlock.Addr1 != null)
                        //////    {
                        //////        Invoice.Addr133 = (string)invoiceRet.BillAddressBlock.Addr1.GetValue();
                        //////    }
                        //////    //Get value of Addr2
                        //////    if (invoiceRet.BillAddressBlock.Addr2 != null)
                        //////    {
                        //////        Invoice.Addr234 = (string)invoiceRet.BillAddressBlock.Addr2.GetValue();
                        //////    }
                        //////    //Get value of Addr3
                        //////    if (invoiceRet.BillAddressBlock.Addr3 != null)
                        //////    {
                        //////        Invoice.Addr335 = (string)invoiceRet.BillAddressBlock.Addr3.GetValue();
                        //////    }
                        //////    //Get value of Addr4
                        //////    if (invoiceRet.BillAddressBlock.Addr4 != null)
                        //////    {
                        //////        Invoice.Addr436 = (string)invoiceRet.BillAddressBlock.Addr4.GetValue();
                        //////    }
                        //////    //Get value of Addr5
                        //////    if (invoiceRet.BillAddressBlock.Addr5 != null)
                        //////    {
                        //////        Invoice.Addr537 = (string)invoiceRet.BillAddressBlock.Addr5.GetValue();
                        //////    }
                        //////}
                        if (invoiceRet.ShipAddress != null)
                        {
                            //Get value of Addr1
                            if (invoiceRet.ShipAddress.Addr1 != null)
                            {
                                Invoice.ShippingLine1 = (string)invoiceRet.ShipAddress.Addr1.GetValue();
                            }
                            //Get value of Addr2
                            if (invoiceRet.ShipAddress.Addr2 != null)
                            {
                                Invoice.ShippingLine2 = (string)invoiceRet.ShipAddress.Addr2.GetValue();
                            }
                            //Get value of Addr3
                            if (invoiceRet.ShipAddress.Addr3 != null)
                            {
                                Invoice.ShippingLine3 = (string)invoiceRet.ShipAddress.Addr3.GetValue();
                            }
                            //Get value of Addr4
                            if (invoiceRet.ShipAddress.Addr4 != null)
                            {
                                Invoice.ShippingLine4 = (string)invoiceRet.ShipAddress.Addr4.GetValue();
                            }
                            //Get value of Addr5
                            if (invoiceRet.ShipAddress.Addr5 != null)
                            {
                                Invoice.ShippingLine5 = (string)invoiceRet.ShipAddress.Addr5.GetValue();
                            }
                            //Get value of City
                            if (invoiceRet.ShipAddress.City != null)
                            {
                                Invoice.ShippingCity = (string)invoiceRet.ShipAddress.City.GetValue();
                            }
                            //Get value of State
                            if (invoiceRet.ShipAddress.State != null)
                            {
                                Invoice.ShippingState = (string)invoiceRet.ShipAddress.State.GetValue();
                            }
                            //Get value of PostalCode
                            if (invoiceRet.ShipAddress.PostalCode != null)
                            {
                                Invoice.ShippingPostalCode = (string)invoiceRet.ShipAddress.PostalCode.GetValue();
                            }
                            //Get value of Country
                            if (invoiceRet.ShipAddress.Country != null)
                            {
                                Invoice.ShippingCountry = (string)invoiceRet.ShipAddress.Country.GetValue();
                            }
                            //Get value of Note
                            if (invoiceRet.ShipAddress.Note != null)
                            {
                                Invoice.ShippingNote = (string)invoiceRet.ShipAddress.Note.GetValue();
                            }
                        }
                        ////////////if (invoiceRet.ShipAddressBlock != null)
                        ////////////{
                        ////////////    //Get value of Addr1
                        ////////////    if (invoiceRet.ShipAddressBlock.Addr1 != null)
                        ////////////    {
                        ////////////        Invoice.Addr148 = (string)invoiceRet.ShipAddressBlock.Addr1.GetValue();
                        ////////////    }
                        ////////////    //Get value of Addr2
                        ////////////    if (invoiceRet.ShipAddressBlock.Addr2 != null)
                        ////////////    {
                        ////////////        Invoice.Addr249 = (string)invoiceRet.ShipAddressBlock.Addr2.GetValue();
                        ////////////    }
                        ////////////    //Get value of Addr3
                        ////////////    if (invoiceRet.ShipAddressBlock.Addr3 != null)
                        ////////////    {
                        ////////////        Invoice.Addr350 = (string)invoiceRet.ShipAddressBlock.Addr3.GetValue();
                        ////////////    }
                        ////////////    //Get value of Addr4
                        ////////////    if (invoiceRet.ShipAddressBlock.Addr4 != null)
                        ////////////    {
                        ////////////        Invoice.Addr451 = (string)invoiceRet.ShipAddressBlock.Addr4.GetValue();
                        ////////////    }
                        ////////////    //Get value of Addr5
                        ////////////    if (invoiceRet.ShipAddressBlock.Addr5 != null)
                        ////////////    {
                        ////////////        Invoice.Addr552 = (string)invoiceRet.ShipAddressBlock.Addr5.GetValue();
                        ////////////    }
                        ////////////}
                        //Get value of IsPending
                        if (invoiceRet.IsPending != null)
                        {
                            Invoice.IsPending = (bool)invoiceRet.IsPending.GetValue();
                        }
                        //Get value of IsFinanceCharge
                        if (invoiceRet.IsFinanceCharge != null)
                        {
                            Invoice.IsFinanceCharge = invoiceRet.IsFinanceCharge.ToString();
                        }
                        //Get value of PONumber
                        if (invoiceRet.PONumber != null)
                        {
                            Invoice.POnumber = (string)invoiceRet.PONumber.GetValue();
                        }
                        if (invoiceRet.TermsRef != null)
                        {
                            //Get value of ListID
                            if (invoiceRet.TermsRef.ListID != null)
                            {
                                Invoice.TermsId = (string)invoiceRet.TermsRef.ListID.GetValue();
                            }
                            //Get value of FullName
                            if (invoiceRet.TermsRef.FullName != null)
                            {
                                Invoice.Terms = (string)invoiceRet.TermsRef.FullName.GetValue();
                            }
                        }
                        //Get value of DueDate
                        if (invoiceRet.DueDate != null)
                        {
                            Invoice.DueDate = (DateTime)invoiceRet.DueDate.GetValue();
                        }
                        if (invoiceRet.SalesRepRef != null)
                        {
                            //Get value of ListID
                            if (invoiceRet.SalesRepRef.ListID != null)
                            {
                                Invoice.SalesRepId = (string)invoiceRet.SalesRepRef.ListID.GetValue();
                            }
                            //Get value of FullName
                            if (invoiceRet.SalesRepRef.FullName != null)
                            {
                                Invoice.SalesRep = (string)invoiceRet.SalesRepRef.FullName.GetValue();
                            }
                        }
                        //Get value of FOB
                        if (invoiceRet.FOB != null)
                        {
                            Invoice.FOB = (string)invoiceRet.FOB.GetValue();
                        }
                        //Get value of ShipDate
                        if (invoiceRet.ShipDate != null)
                        {
                            Invoice.ShipDate = (DateTime)invoiceRet.ShipDate.GetValue();
                        }
                        if (invoiceRet.ShipMethodRef != null)
                        {
                            //Get value of ListID
                            if (invoiceRet.ShipMethodRef.ListID != null)
                            {
                                Invoice.ShipMethodId = (string)invoiceRet.ShipMethodRef.ListID.GetValue();
                            }
                            //Get value of FullName
                            if (invoiceRet.ShipMethodRef.FullName != null)
                            {
                                Invoice.ShipMethod = (string)invoiceRet.ShipMethodRef.FullName.GetValue();
                            }
                        }
                        //Get value of Subtotal
                        if (invoiceRet.Subtotal != null)
                        {
                            Invoice.Subtotal = (double)invoiceRet.Subtotal.GetValue();
                        }
                        if (invoiceRet.ItemSalesTaxRef != null)
                        {
                            //Get value of ListID
                            if (invoiceRet.ItemSalesTaxRef.ListID != null)
                            {
                                Invoice.TaxItemId = (string)invoiceRet.ItemSalesTaxRef.ListID.GetValue();
                            }
                            //Get value of FullName
                            if (invoiceRet.ItemSalesTaxRef.FullName != null)
                            {
                                Invoice.TaxItem = (string)invoiceRet.ItemSalesTaxRef.FullName.GetValue();
                            }
                        }
                        //Get value of SalesTaxPercentage
                        if (invoiceRet.SalesTaxPercentage != null)
                        {
                            Invoice.TaxPercent = (double)invoiceRet.SalesTaxPercentage.GetValue();
                        }
                        //Get value of SalesTaxTotal
                        if (invoiceRet.SalesTaxTotal != null)
                        {
                            Invoice.Tax = (double)invoiceRet.SalesTaxTotal.GetValue();
                        }
                        //Get value of AppliedAmount
                        if (invoiceRet.AppliedAmount != null)
                        {
                            Invoice.AppliedAmount = (double)invoiceRet.AppliedAmount.GetValue();
                        }
                        //Get value of BalanceRemaining
                        if (invoiceRet.BalanceRemaining != null)
                        {
                            Invoice.Balance = (double)invoiceRet.BalanceRemaining.GetValue();
                        }
                        if (invoiceRet.CurrencyRef != null)
                        {
                            //Get value of ListID
                            if (invoiceRet.CurrencyRef.ListID != null)
                            {
                                Invoice.CurrencyId = (string)invoiceRet.CurrencyRef.ListID.GetValue();
                            }
                            //Get value of FullName
                            if (invoiceRet.CurrencyRef.FullName != null)
                            {
                                Invoice.CurrencyName = (string)invoiceRet.CurrencyRef.FullName.GetValue();
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
                            Invoice.BalanceInHomeCurrency = (double)invoiceRet.BalanceRemainingInHomeCurrency.GetValue();
                        }
                        //Get value of Memo
                        if (invoiceRet.Memo != null)
                        {
                            Invoice.Memo = (string)invoiceRet.Memo.GetValue();
                        }
                        //Get value of IsPaid
                        if (invoiceRet.IsPaid != null)
                        {
                            Invoice.IsPaid = (bool)invoiceRet.IsPaid.GetValue();
                        }
                        if (invoiceRet.CustomerMsgRef != null)
                        {
                            //Get value of ListID
                            if (invoiceRet.CustomerMsgRef.ListID != null)
                            {
                                Invoice.MessageId = (string)invoiceRet.CustomerMsgRef.ListID.GetValue();
                            }
                            //Get value of FullName
                            if (invoiceRet.CustomerMsgRef.FullName != null)
                            {
                                Invoice.Message = (string)invoiceRet.CustomerMsgRef.FullName.GetValue();
                            }
                        }
                        //Get value of IsToBePrinted
                        if (invoiceRet.IsToBePrinted != null)
                        {
                            Invoice.IsToBePrinted = (bool)invoiceRet.IsToBePrinted.GetValue();
                        }
                        //Get value of IsToBeEmailed
                        if (invoiceRet.IsToBeEmailed != null)
                        {
                            Invoice.IsToBeEmailed = (bool)invoiceRet.IsToBeEmailed.GetValue();
                        }
                        if (invoiceRet.CustomerSalesTaxCodeRef != null)
                        {
                            //Get value of ListID
                            if (invoiceRet.CustomerSalesTaxCodeRef.ListID != null)
                            {
                                Invoice.CustomerTaxCodeId = (string)invoiceRet.CustomerSalesTaxCodeRef.ListID.GetValue();
                            }
                            //Get value of FullName
                            if (invoiceRet.CustomerSalesTaxCodeRef.FullName != null)
                            {
                                Invoice.CustomerTaxCode = (string)invoiceRet.CustomerSalesTaxCodeRef.FullName.GetValue();
                            }
                        }
                        //Get value of SuggestedDiscountAmount
                        if (invoiceRet.SuggestedDiscountAmount != null)
                        {
                            Invoice.SuggestedDiscountAmount = (double)invoiceRet.SuggestedDiscountAmount.GetValue();
                        }
                        //Get value of SuggestedDiscountDate
                        if (invoiceRet.SuggestedDiscountDate != null)
                        {
                            Invoice.SuggestedDiscountDate = (DateTime)invoiceRet.SuggestedDiscountDate.GetValue();
                        }
                        //Get value of Other
                        if (invoiceRet.Other != null)
                        {
                            Invoice.Other = (string)invoiceRet.Other.GetValue();
                        }
                        //Get value of ExternalGUID
                        //////if (invoiceRet.ExternalGUID != null)
                        //////{
                        //////    Invoice.ExternalGUID87 = (string)invoiceRet.ExternalGUID.GetValue();
                        //////}
                        //////if (invoiceRet.LinkedTxnList != null)
                        //////{
                        //////    for (int i = 0; i < invoiceRet.LinkedTxnList.Count; i++)
                        //////    {
                        //////        ILinkedTxn LinkedTxn = invoiceRet.LinkedTxnList.GetAt(i);
                        //////        //Get value of TxnID
                        //////        Invoice.L = (string)LinkedTxn.TxnID.GetValue();
                        //////        //Get value of TxnType
                        //////        ENTxnType TxnType90 = (ENTxnType)LinkedTxn.TxnType.GetValue();
                        //////        //Get value of TxnDate
                        //////        Invoice.TxnDate91 = (DateTime)LinkedTxn.TxnDate.GetValue();
                        //////        //Get value of RefNumber
                        //////        if (LinkedTxn.RefNumber != null)
                        //////        {
                        //////            Invoice.RefNumber92 = (string)LinkedTxn.RefNumber.GetValue();
                        //////        }
                        //////        //Get value of LinkType
                        //////        if (LinkedTxn.LinkType != null)
                        //////        {
                        //////            ENLinkType LinkType93 = (ENLinkType)LinkedTxn.LinkType.GetValue();
                        //////        }
                        //////        //Get value of Amount
                        //////        Invoice.Amount94 = (double)LinkedTxn.Amount.GetValue();
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
                        //////                Invoice.TxnLineID96 = (string)ORInvoiceLineRet.InvoiceLineRet.TxnLineID.GetValue();
                        //////                if (ORInvoiceLineRet.InvoiceLineRet.ItemRef != null)
                        //////                {
                        //////                    //Get value of ListID
                        //////                    if (ORInvoiceLineRet.InvoiceLineRet.ItemRef.ListID != null)
                        //////                    {
                        //////                        Invoice.ListID97 = (string)ORInvoiceLineRet.InvoiceLineRet.ItemRef.ListID.GetValue();
                        //////                    }
                        //////                    //Get value of FullName
                        //////                    if (ORInvoiceLineRet.InvoiceLineRet.ItemRef.FullName != null)
                        //////                    {
                        //////                        Invoice.FullName98 = (string)ORInvoiceLineRet.InvoiceLineRet.ItemRef.FullName.GetValue();
                        //////                    }
                        //////                }
                        //////                //Get value of Desc
                        //////                if (ORInvoiceLineRet.InvoiceLineRet.Desc != null)
                        //////                {
                        //////                    Invoice.Desc99 = (string)ORInvoiceLineRet.InvoiceLineRet.Desc.GetValue();
                        //////                }
                        //////                //Get value of Quantity
                        //////                if (ORInvoiceLineRet.InvoiceLineRet.Quantity != null)
                        //////                {
                        //////                    Invoice.Quantity100 = (int)ORInvoiceLineRet.InvoiceLineRet.Quantity.GetValue();
                        //////                }
                        //////                //Get value of UnitOfMeasure
                        //////                if (ORInvoiceLineRet.InvoiceLineRet.UnitOfMeasure != null)
                        //////                {
                        //////                    Invoice.UnitOfMeasure101 = (string)ORInvoiceLineRet.InvoiceLineRet.UnitOfMeasure.GetValue();
                        //////                }
                        //////                if (ORInvoiceLineRet.InvoiceLineRet.OverrideUOMSetRef != null)
                        //////                {
                        //////                    //Get value of ListID
                        //////                    if (ORInvoiceLineRet.InvoiceLineRet.OverrideUOMSetRef.ListID != null)
                        //////                    {
                        //////                        Invoice.ListID102 = (string)ORInvoiceLineRet.InvoiceLineRet.OverrideUOMSetRef.ListID.GetValue();
                        //////                    }
                        //////                    //Get value of FullName
                        //////                    if (ORInvoiceLineRet.InvoiceLineRet.OverrideUOMSetRef.FullName != null)
                        //////                    {
                        //////                        Invoice.FullName103 = (string)ORInvoiceLineRet.InvoiceLineRet.OverrideUOMSetRef.FullName.GetValue();
                        //////                    }
                        //////                }
                        //////                if (ORInvoiceLineRet.InvoiceLineRet.ORRate != null)
                        //////                {
                        //////                    if (ORInvoiceLineRet.InvoiceLineRet.ORRate.Rate != null)
                        //////                    {
                        //////                        //Get value of Rate
                        //////                        if (ORInvoiceLineRet.InvoiceLineRet.ORRate.Rate != null)
                        //////                        {
                        //////                            Invoice.Rate104 = (double)ORInvoiceLineRet.InvoiceLineRet.ORRate.Rate.GetValue();
                        //////                        }
                        //////                    }
                        //////                    if (ORInvoiceLineRet.InvoiceLineRet.ORRate.RatePercent != null)
                        //////                    {
                        //////                        //Get value of RatePercent
                        //////                        if (ORInvoiceLineRet.InvoiceLineRet.ORRate.RatePercent != null)
                        //////                        {
                        //////                            Invoice.RatePercent105 = (double)ORInvoiceLineRet.InvoiceLineRet.ORRate.RatePercent.GetValue();
                        //////                        }
                        //////                    }
                        //////                }
                        //////                if (ORInvoiceLineRet.InvoiceLineRet.ClassRef != null)
                        //////                {
                        //////                    //Get value of ListID
                        //////                    if (ORInvoiceLineRet.InvoiceLineRet.ClassRef.ListID != null)
                        //////                    {
                        //////                        Invoice.ListID106 = (string)ORInvoiceLineRet.InvoiceLineRet.ClassRef.ListID.GetValue();
                        //////                    }
                        //////                    //Get value of FullName
                        //////                    if (ORInvoiceLineRet.InvoiceLineRet.ClassRef.FullName != null)
                        //////                    {
                        //////                        Invoice.FullName107 = (string)ORInvoiceLineRet.InvoiceLineRet.ClassRef.FullName.GetValue();
                        //////                    }
                        //////                }
                        //////                //Get value of Amount
                        //////                if (ORInvoiceLineRet.InvoiceLineRet.Amount != null)
                        //////                {
                        //////                    Invoice.Amount108 = (double)ORInvoiceLineRet.InvoiceLineRet.Amount.GetValue();
                        //////                }
                        //////                if (ORInvoiceLineRet.InvoiceLineRet.InventorySiteRef != null)
                        //////                {
                        //////                    //Get value of ListID
                        //////                    if (ORInvoiceLineRet.InvoiceLineRet.InventorySiteRef.ListID != null)
                        //////                    {
                        //////                        Invoice.ListID109 = (string)ORInvoiceLineRet.InvoiceLineRet.InventorySiteRef.ListID.GetValue();
                        //////                    }
                        //////                    //Get value of FullName
                        //////                    if (ORInvoiceLineRet.InvoiceLineRet.InventorySiteRef.FullName != null)
                        //////                    {
                        //////                        Invoice.FullName110 = (string)ORInvoiceLineRet.InvoiceLineRet.InventorySiteRef.FullName.GetValue();
                        //////                    }
                        //////                }
                        //////                if (ORInvoiceLineRet.InvoiceLineRet.InventorySiteLocationRef != null)
                        //////                {
                        //////                    //Get value of ListID
                        //////                    if (ORInvoiceLineRet.InvoiceLineRet.InventorySiteLocationRef.ListID != null)
                        //////                    {
                        //////                        Invoice.ListID111 = (string)ORInvoiceLineRet.InvoiceLineRet.InventorySiteLocationRef.ListID.GetValue();
                        //////                    }
                        //////                    //Get value of FullName
                        //////                    if (ORInvoiceLineRet.InvoiceLineRet.InventorySiteLocationRef.FullName != null)
                        //////                    {
                        //////                        Invoice.FullName112 = (string)ORInvoiceLineRet.InvoiceLineRet.InventorySiteLocationRef.FullName.GetValue();
                        //////                    }
                        //////                }
                        //////                if (ORInvoiceLineRet.InvoiceLineRet.ORSerialLotNumber != null)
                        //////                {
                        //////                    if (ORInvoiceLineRet.InvoiceLineRet.ORSerialLotNumber.SerialNumber != null)
                        //////                    {
                        //////                        //Get value of SerialNumber
                        //////                        if (ORInvoiceLineRet.InvoiceLineRet.ORSerialLotNumber.SerialNumber != null)
                        //////                        {
                        //////                            Invoice.SerialNumber113 = (string)ORInvoiceLineRet.InvoiceLineRet.ORSerialLotNumber.SerialNumber.GetValue();
                        //////                        }
                        //////                    }
                        //////                    if (ORInvoiceLineRet.InvoiceLineRet.ORSerialLotNumber.LotNumber != null)
                        //////                    {
                        //////                        //Get value of LotNumber
                        //////                        if (ORInvoiceLineRet.InvoiceLineRet.ORSerialLotNumber.LotNumber != null)
                        //////                        {
                        //////                            Invoice.LotNumber114 = (string)ORInvoiceLineRet.InvoiceLineRet.ORSerialLotNumber.LotNumber.GetValue();
                        //////                        }
                        //////                    }
                        //////                }
                        //////                //Get value of ServiceDate
                        //////                if (ORInvoiceLineRet.InvoiceLineRet.ServiceDate != null)
                        //////                {
                        //////                    Invoice.ServiceDate115 = (DateTime)ORInvoiceLineRet.InvoiceLineRet.ServiceDate.GetValue();
                        //////                }
                        //////                if (ORInvoiceLineRet.InvoiceLineRet.SalesTaxCodeRef != null)
                        //////                {
                        //////                    //Get value of ListID
                        //////                    if (ORInvoiceLineRet.InvoiceLineRet.SalesTaxCodeRef.ListID != null)
                        //////                    {
                        //////                        Invoice.ListID116 = (string)ORInvoiceLineRet.InvoiceLineRet.SalesTaxCodeRef.ListID.GetValue();
                        //////                    }
                        //////                    //Get value of FullName
                        //////                    if (ORInvoiceLineRet.InvoiceLineRet.SalesTaxCodeRef.FullName != null)
                        //////                    {
                        //////                        Invoice.FullName117 = (string)ORInvoiceLineRet.InvoiceLineRet.SalesTaxCodeRef.FullName.GetValue();
                        //////                    }
                        //////                }
                        //////                //Get value of Other1
                        //////                if (ORInvoiceLineRet.InvoiceLineRet.Other1 != null)
                        //////                {
                        //////                    Invoice.Other1118 = (string)ORInvoiceLineRet.InvoiceLineRet.Other1.GetValue();
                        //////                }
                        //////                //Get value of Other2
                        //////                if (ORInvoiceLineRet.InvoiceLineRet.Other2 != null)
                        //////                {
                        //////                    Invoice.Other2119 = (string)ORInvoiceLineRet.InvoiceLineRet.Other2.GetValue();
                        //////                }
                        //////                if (ORInvoiceLineRet.InvoiceLineRet.DataExtRetList != null)
                        //////                {
                        //////                    for (Invoice.i120 = 0; i120 < ORInvoiceLineRet.InvoiceLineRet.DataExtRetList.Count; i120++)
                        //////                    {
                        //////                        IDataExtRet DataExtRet = ORInvoiceLineRet.InvoiceLineRet.DataExtRetList.GetAt(i120);
                        //////                        //Get value of OwnerID
                        //////                        if (DataExtRet.OwnerID != null)
                        //////                        {
                        //////                            Invoice.OwnerID121 = (string)DataExtRet.OwnerID.GetValue();
                        //////                        }
                        //////                        //Get value of DataExtName
                        //////                        Invoice.DataExtName122 = (string)DataExtRet.DataExtName.GetValue();
                        //////                        //Get value of DataExtType
                        //////                        ENDataExtType DataExtType123 = (ENDataExtType)DataExtRet.DataExtType.GetValue();
                        //////                        //Get value of DataExtValue
                        //////                        Invoice.DataExtValue124 = (string)DataExtRet.DataExtValue.GetValue();
                        //////                    }
                        //////                }
                        //////            }
                        //////        }
                        //////        if (ORInvoiceLineRet.InvoiceLineGroupRet != null)
                        //////        {
                        //////            if (ORInvoiceLineRet.InvoiceLineGroupRet != null)
                        //////            {
                        //////                //Get value of TxnLineID
                        //////                Invoice.TxnLineID125 = (string)ORInvoiceLineRet.InvoiceLineGroupRet.TxnLineID.GetValue();
                        //////                //Get value of ListID
                        //////                if (ORInvoiceLineRet.InvoiceLineGroupRet.ItemGroupRef.ListID != null)
                        //////                {
                        //////                    Invoice.ListID126 = (string)ORInvoiceLineRet.InvoiceLineGroupRet.ItemGroupRef.ListID.GetValue();
                        //////                }
                        //////                //Get value of FullName
                        //////                if (ORInvoiceLineRet.InvoiceLineGroupRet.ItemGroupRef.FullName != null)
                        //////                {
                        //////                    Invoice.FullName127 = (string)ORInvoiceLineRet.InvoiceLineGroupRet.ItemGroupRef.FullName.GetValue();
                        //////                }
                        //////                //Get value of Desc
                        //////                if (ORInvoiceLineRet.InvoiceLineGroupRet.Desc != null)
                        //////                {
                        //////                    Invoice.Desc128 = (string)ORInvoiceLineRet.InvoiceLineGroupRet.Desc.GetValue();
                        //////                }
                        //////                //Get value of Quantity
                        //////                if (ORInvoiceLineRet.InvoiceLineGroupRet.Quantity != null)
                        //////                {
                        //////                    Invoice.Quantity129 = (int)ORInvoiceLineRet.InvoiceLineGroupRet.Quantity.GetValue();
                        //////                }
                        //////                //Get value of UnitOfMeasure
                        //////                if (ORInvoiceLineRet.InvoiceLineGroupRet.UnitOfMeasure != null)
                        //////                {
                        //////                    Invoice.UnitOfMeasure130 = (string)ORInvoiceLineRet.InvoiceLineGroupRet.UnitOfMeasure.GetValue();
                        //////                }
                        //////                if (ORInvoiceLineRet.InvoiceLineGroupRet.OverrideUOMSetRef != null)
                        //////                {
                        //////                    //Get value of ListID
                        //////                    if (ORInvoiceLineRet.InvoiceLineGroupRet.OverrideUOMSetRef.ListID != null)
                        //////                    {
                        //////                        Invoice.ListID131 = (string)ORInvoiceLineRet.InvoiceLineGroupRet.OverrideUOMSetRef.ListID.GetValue();
                        //////                    }
                        //////                    //Get value of FullName
                        //////                    if (ORInvoiceLineRet.InvoiceLineGroupRet.OverrideUOMSetRef.FullName != null)
                        //////                    {
                        //////                        Invoice.FullName132 = (string)ORInvoiceLineRet.InvoiceLineGroupRet.OverrideUOMSetRef.FullName.GetValue();
                        //////                    }
                        //////                }
                        //////                //Get value of IsPrintItemsInGroup
                        //////                Invoice.IsPrintItemsInGroup133 = (bool)ORInvoiceLineRet.InvoiceLineGroupRet.IsPrintItemsInGroup.GetValue();
                        //////                //Get value of TotalAmount
                        //////                Invoice.TotalAmount134 = (double)ORInvoiceLineRet.InvoiceLineGroupRet.TotalAmount.GetValue();
                        //////                if (ORInvoiceLineRet.InvoiceLineGroupRet.InvoiceLineRetList != null)
                        //////                {
                        //////                    for (int i135 = 0; i135 < ORInvoiceLineRet.InvoiceLineGroupRet.InvoiceLineRetList.Count; i135++)
                        //////                    {
                        //////                        IInvoiceLineRet InvoiceLineRet = ORInvoiceLineRet.InvoiceLineGroupRet.InvoiceLineRetList.GetAt(i135);
                        //////                        //Get value of TxnLineID
                        //////                        Invoice.TxnLineID136 = (string)InvoiceLineRet.TxnLineID.GetValue();
                        //////                        if (InvoiceLineRet.ItemRef != null)
                        //////                        {
                        //////                            //Get value of ListID
                        //////                            if (InvoiceLineRet.ItemRef.ListID != null)
                        //////                            {
                        //////                                Invoice.ListID137 = (string)InvoiceLineRet.ItemRef.ListID.GetValue();
                        //////                            }
                        //////                            //Get value of FullName
                        //////                            if (InvoiceLineRet.ItemRef.FullName != null)
                        //////                            {
                        //////                                Invoice.FullName138 = (string)InvoiceLineRet.ItemRef.FullName.GetValue();
                        //////                            }
                        //////                        }
                        //////                        //Get value of Desc
                        //////                        if (InvoiceLineRet.Desc != null)
                        //////                        {
                        //////                            Invoice.Desc139 = (string)InvoiceLineRet.Desc.GetValue();
                        //////                        }
                        //////                        //Get value of Quantity
                        //////                        if (InvoiceLineRet.Quantity != null)
                        //////                        {
                        //////                            Invoice.Quantity140 = (int)InvoiceLineRet.Quantity.GetValue();
                        //////                        }
                        //////                        //Get value of UnitOfMeasure
                        //////                        if (InvoiceLineRet.UnitOfMeasure != null)
                        //////                        {
                        //////                            Invoice.UnitOfMeasure141 = (string)InvoiceLineRet.UnitOfMeasure.GetValue();
                        //////                        }
                        //////                        if (InvoiceLineRet.OverrideUOMSetRef != null)
                        //////                        {
                        //////                            //Get value of ListID
                        //////                            if (InvoiceLineRet.OverrideUOMSetRef.ListID != null)
                        //////                            {
                        //////                                Invoice.ListID142 = (string)InvoiceLineRet.OverrideUOMSetRef.ListID.GetValue();
                        //////                            }
                        //////                            //Get value of FullName
                        //////                            if (InvoiceLineRet.OverrideUOMSetRef.FullName != null)
                        //////                            {
                        //////                                Invoice.FullName143 = (string)InvoiceLineRet.OverrideUOMSetRef.FullName.GetValue();
                        //////                            }
                        //////                        }
                        //////                        if (InvoiceLineRet.ORRate != null)
                        //////                        {
                        //////                            if (InvoiceLineRet.ORRate.Rate != null)
                        //////                            {
                        //////                                //Get value of Rate
                        //////                                if (InvoiceLineRet.ORRate.Rate != null)
                        //////                                {
                        //////                                    Invoice.Rate144 = (double)InvoiceLineRet.ORRate.Rate.GetValue();
                        //////                                }
                        //////                            }
                        //////                            if (InvoiceLineRet.ORRate.RatePercent != null)
                        //////                            {
                        //////                                //Get value of RatePercent
                        //////                                if (InvoiceLineRet.ORRate.RatePercent != null)
                        //////                                {
                        //////                                    Invoice.RatePercent145 = (double)InvoiceLineRet.ORRate.RatePercent.GetValue();
                        //////                                }
                        //////                            }
                        //////                        }
                        //////                        if (InvoiceLineRet.ClassRef != null)
                        //////                        {
                        //////                            //Get value of ListID
                        //////                            if (InvoiceLineRet.ClassRef.ListID != null)
                        //////                            {
                        //////                                Invoice.ListID146 = (string)InvoiceLineRet.ClassRef.ListID.GetValue();
                        //////                            }
                        //////                            //Get value of FullName
                        //////                            if (InvoiceLineRet.ClassRef.FullName != null)
                        //////                            {
                        //////                                Invoice.FullName147 = (string)InvoiceLineRet.ClassRef.FullName.GetValue();
                        //////                            }
                        //////                        }
                        //////                        //Get value of Amount
                        //////                        if (InvoiceLineRet.Amount != null)
                        //////                        {
                        //////                            Invoice.Amount148 = (double)InvoiceLineRet.Amount.GetValue();
                        //////                        }
                        //////                        if (InvoiceLineRet.InventorySiteRef != null)
                        //////                        {
                        //////                            //Get value of ListID
                        //////                            if (InvoiceLineRet.InventorySiteRef.ListID != null)
                        //////                            {
                        //////                                Invoice.ListID149 = (string)InvoiceLineRet.InventorySiteRef.ListID.GetValue();
                        //////                            }
                        //////                            //Get value of FullName
                        //////                            if (InvoiceLineRet.InventorySiteRef.FullName != null)
                        //////                            {
                        //////                                Invoice.FullName150 = (string)InvoiceLineRet.InventorySiteRef.FullName.GetValue();
                        //////                            }
                        //////                        }
                        //////                        if (InvoiceLineRet.InventorySiteLocationRef != null)
                        //////                        {
                        //////                            //Get value of ListID
                        //////                            if (InvoiceLineRet.InventorySiteLocationRef.ListID != null)
                        //////                            {
                        //////                                Invoice.ListID151 = (string)InvoiceLineRet.InventorySiteLocationRef.ListID.GetValue();
                        //////                            }
                        //////                            //Get value of FullName
                        //////                            if (InvoiceLineRet.InventorySiteLocationRef.FullName != null)
                        //////                            {
                        //////                                Invoice.FullName152 = (string)InvoiceLineRet.InventorySiteLocationRef.FullName.GetValue();
                        //////                            }
                        //////                        }
                        //////                        if (InvoiceLineRet.ORSerialLotNumber != null)
                        //////                        {
                        //////                            if (InvoiceLineRet.ORSerialLotNumber.SerialNumber != null)
                        //////                            {
                        //////                                //Get value of SerialNumber
                        //////                                if (InvoiceLineRet.ORSerialLotNumber.SerialNumber != null)
                        //////                                {
                        //////                                    Invoice.SerialNumber153 = (string)InvoiceLineRet.ORSerialLotNumber.SerialNumber.GetValue();
                        //////                                }
                        //////                            }
                        //////                            if (InvoiceLineRet.ORSerialLotNumber.LotNumber != null)
                        //////                            {
                        //////                                //Get value of LotNumber
                        //////                                if (InvoiceLineRet.ORSerialLotNumber.LotNumber != null)
                        //////                                {
                        //////                                    Invoice.LotNumber154 = (string)InvoiceLineRet.ORSerialLotNumber.LotNumber.GetValue();
                        //////                                }
                        //////                            }
                        //////                        }
                        //////                        //Get value of ServiceDate
                        //////                        if (InvoiceLineRet.ServiceDate != null)
                        //////                        {
                        //////                            Invoice.ServiceDate155 = (DateTime)InvoiceLineRet.ServiceDate.GetValue();
                        //////                        }
                        //////                        if (InvoiceLineRet.SalesTaxCodeRef != null)
                        //////                        {
                        //////                            //Get value of ListID
                        //////                            if (InvoiceLineRet.SalesTaxCodeRef.ListID != null)
                        //////                            {
                        //////                                Invoice.ListID156 = (string)InvoiceLineRet.SalesTaxCodeRef.ListID.GetValue();
                        //////                            }
                        //////                            //Get value of FullName
                        //////                            if (InvoiceLineRet.SalesTaxCodeRef.FullName != null)
                        //////                            {
                        //////                                Invoice.FullName157 = (string)InvoiceLineRet.SalesTaxCodeRef.FullName.GetValue();
                        //////                            }
                        //////                        }
                        //////                        //Get value of Other1
                        //////                        if (InvoiceLineRet.Other1 != null)
                        //////                        {
                        //////                            Invoice.Other1158 = (string)InvoiceLineRet.Other1.GetValue();
                        //////                        }
                        //////                        //Get value of Other2
                        //////                        if (InvoiceLineRet.Other2 != null)
                        //////                        {
                        //////                            Invoice.Other2159 = (string)InvoiceLineRet.Other2.GetValue();
                        //////                        }
                        //////                        if (InvoiceLineRet.DataExtRetList != null)
                        //////                        {
                        //////                            for (int i160 = 0; i160 < InvoiceLineRet.DataExtRetList.Count; i160++)
                        //////                            {
                        //////                                IDataExtRet DataExtRet = InvoiceLineRet.DataExtRetList.GetAt(i160);
                        //////                                //Get value of OwnerID
                        //////                                if (DataExtRet.OwnerID != null)
                        //////                                {
                        //////                                    Invoice.OwnerID161 = (string)DataExtRet.OwnerID.GetValue();
                        //////                                }
                        //////                                //Get value of DataExtName
                        //////                                Invoice.DataExtName162 = (string)DataExtRet.DataExtName.GetValue();
                        //////                                //Get value of DataExtType
                        //////                                ENDataExtType DataExtType163 = (ENDataExtType)DataExtRet.DataExtType.GetValue();
                        //////                                //Get value of DataExtValue
                        //////                                Invoice.DataExtValue164 = (string)DataExtRet.DataExtValue.GetValue();
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
                        //////                            Invoice.OwnerID166 = (string)DataExtRet.OwnerID.GetValue();
                        //////                        }
                        //////                        //Get value of DataExtName
                        //////                        Invoice.DataExtName167 = (string)DataExtRet.DataExtName.GetValue();
                        //////                        //Get value of DataExtType
                        //////                        ENDataExtType DataExtType168 = (ENDataExtType)DataExtRet.DataExtType.GetValue();
                        //////                        //Get value of DataExtValue
                        //////                        Invoice.DataExtValue169 = (string)DataExtRet.DataExtValue.GetValue();
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
                        ////        Invoice.DataExtName172 = (string)DataExtRet.DataExtName.GetValue();
                        ////        //Get value of DataExtType
                        ////        ENDataExtType DataExtType173 = (ENDataExtType)DataExtRet.DataExtType.GetValue();
                        ////        //Get value of DataExtValue
                        ////        Invoice.DataExtValue174 = (string)DataExtRet.DataExtValue.GetValue();
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

