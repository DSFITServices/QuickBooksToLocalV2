using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using QBFC13Lib;

namespace QuickBooksToLocalV2.Interfaces
{
    public interface IConnect<T>
    {
        T DbConnect(T obj);

        bool DbClose(T obj);
    }
}
