using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Database
{
    interface IDataBaseRequestsHandler
    {
        string[,] GetInfoFromRequest(string request, int colonRequest);
    }
}
