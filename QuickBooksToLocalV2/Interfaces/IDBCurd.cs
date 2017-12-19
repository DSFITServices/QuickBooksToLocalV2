using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace QuickBooksToLocalV2.Interfaces
{
    public interface IDBCurd<T>
    {
        //T DbCreate(T obj);

        //T DbUpdate(T obj);

        ObservableCollection<T> DbRead(Nullable<DateTime> startDate,  Nullable<DateTime> endDate);

        //bool DbDelete(string ID);

    }
}
