using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;
using Sellars.DRCLib;

namespace Sellars.DRCLibTest
{
    public class QueuedWebSalesOrderTests
    {
        private string _shopfloorConnection = ConfigurationManager.ConnectionStrings["Shopfloor"].ConnectionString;

        public void RunAddRecordTest()
        {
            QueuedWebSalesOrderMaster webSalesOrderMaster = new QueuedWebSalesOrderMaster();
            QueuedWebSalesOrderDetail webSalesOrderDetail = new QueuedWebSalesOrderDetail();
            SalesOrderDetailNotes orderDetailNote = new SalesOrderDetailNotes();

            webSalesOrderMaster.CUSTID = "ELCOR";
            webSalesOrderMaster.CUSTPO = "TESTPO-DA-01";
            webSalesOrderMaster.CartID = 100098;
            webSalesOrderMaster.ORDDTE = DateTime.Now;
            webSalesOrderMaster.COMNT1 = "test comnt 1";
            webSalesOrderMaster.COMNT2 = "test comnt 2";
            webSalesOrderMaster.COMNT3 = "test comnt 3";
            webSalesOrderMaster.NAME = "Test Addr Name";
            webSalesOrderMaster.ADDR1 = "Test addr1";
            webSalesOrderMaster.ADDR2 = "Test addr2";
            webSalesOrderMaster.ADDR3 = "Test addr3";
            webSalesOrderMaster.ADDR4 = "Test addr4";
            webSalesOrderMaster.ADDR5 = "Test addr5";
            webSalesOrderMaster.ADDR6 = "Test addr6";
            webSalesOrderMaster.CITY = "Test city";
            webSalesOrderMaster.STATE = "Test state";
            webSalesOrderMaster.CNTRY = "Test country";
            webSalesOrderMaster.ZIPCD = "Test zipcode";
            webSalesOrderMaster.ContactPhone = "1-555-123-1222";
            webSalesOrderMaster.ContactName = "ContactName";
            webSalesOrderMaster.ContactEmail = "contact@email.test";
            webSalesOrderMaster.Notes = "this is a note that is used for testing. it is nothing more than a test note. no information should be used from this test note.";
            webSalesOrderMaster.Carrier = 1;
            webSalesOrderMaster.CarrierMethod = "GND";
            webSalesOrderMaster.CarrierThirdParty = "AZ3809428320984";
            webSalesOrderMaster.CarrierName = "MyCarrier";
            webSalesOrderMaster.FOB = "DESTINATION";
            webSalesOrderMaster.DefaultSTK = "DSC1";
            webSalesOrderMaster.ShippingContactName = "ShippingContactName";
            webSalesOrderMaster.ShippingContactPhone = "1 (555) 234-9393";
            webSalesOrderMaster.ShippingContactEmail = "shippingemail@testing.tests";
            webSalesOrderMaster.EstimatedShipping = "444.22";
            webSalesOrderMaster.WebsiteTotal = "349290.33";
            webSalesOrderMaster.CCType = "c3";
            webSalesOrderMaster.CCNumber = "123456789123011";
            webSalesOrderMaster.CCExpiration = new DateTime(2025, 12, 30);
            webSalesOrderMaster.CCVN = "3434";
            webSalesOrderMaster.CCBTFirstName = "CCBTFirstNameTest";
            webSalesOrderMaster.CCBTLastName = "CCBTLastNameTest";
            webSalesOrderMaster.CCBTAddress = "CCBTAddressTest";
            webSalesOrderMaster.CCBTCity = "CCBTCityTest";
            webSalesOrderMaster.CCBTState = "CCBTStateTest";
            webSalesOrderMaster.CCBTZip = "CCBTZipTest";
            webSalesOrderMaster.ShippingNotes = "Shipping NOTES! NOTES NOTES NOTES NOTES NOTES!!!!!! there are 6 notes in this string";
            webSalesOrderMaster.FixCrLF = true;

            Guid masterID = webSalesOrderMaster.AddRecordToQueue(_shopfloorConnection);

            webSalesOrderDetail.PRTNUM = "20421";
            webSalesOrderDetail.DUEDATE = new DateTime(2020, 12, 31);
            webSalesOrderDetail.SLSUOM = "CS";
            webSalesOrderDetail.PRICE = 46.67M;
            webSalesOrderDetail.QuantityPurchased = 5;
            webSalesOrderDetail.QuantityFree = 0;
            webSalesOrderDetail.DISC = 5;
            webSalesOrderDetail.BuyerPartNumber = "buyerpartnumberTest";
            webSalesOrderDetail.PromotionCode = "PROMO2020";
            webSalesOrderDetail.PromotionAddedLine = true;
            webSalesOrderDetail.FullPrice = 50.33M;
            webSalesOrderDetail.DiscountPercent = 3M;
            webSalesOrderDetail.GLXREF = "000401002402000";

            Guid detailID = webSalesOrderDetail.AddRecordToQueue(_shopfloorConnection, masterID);

            orderDetailNote.DoNotPrint = "DO not print this";
            orderDetailNote.PrintOnInvoiceOnly = "Only print this on the invoice please";
            orderDetailNote.PrintOnOrderOnly = "This better only be on the order...";
            orderDetailNote.PrintOnBoth = "Do whatever you want";

            orderDetailNote.AddRecordToQueue(_shopfloorConnection, detailID);


            List<QueuedWebSalesOrderMaster> masters = QueuedWebSalesOrderMaster.GetQueuedRecords(_shopfloorConnection, false, false);

            if (masters.Count > 0)
            {
                // huzzah!
            }
        }
    }
}
