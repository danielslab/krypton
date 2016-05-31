using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Krypton
{
    public interface IKryptonLogger
    {
        void WriteLog(string message);
    }
}
