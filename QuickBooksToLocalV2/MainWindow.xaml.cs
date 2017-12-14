using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using QuickBooksToLocalV2.QuickbooksDAL;

namespace QuickBooksToLocalV2
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();

            //var qbcusttype = new QBcustomerTypes();
            //qbcusttype.DbGetCustomerTypes();

            //var qbCustomers = new QBcustomers();
            //qbCustomers.DbGetCustomers();

            var qbInvoices = new QBinvoices();
            qbInvoices.DbGetInvoices();








        }
    }
}
