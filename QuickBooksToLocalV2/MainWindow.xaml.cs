using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
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
using QuickBooksToLocalV2.Models;

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
            synncquickbooksEntities qbct = new synncquickbooksEntities();

            Dbcustomertype qbCtype = new Dbcustomertype(null,null);

            ObservableCollection<customertype> CustomerTypes = qbCtype.DbRead();

            qbct.customertypes.AddRange(CustomerTypes);
            qbct.SaveChanges();


            





        }
    }
}
