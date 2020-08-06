using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SuzuOffice
{
    public class Class1
    {
        /// <summary>
        /// 
        /// </summary>
        /// <param name="i"></param>
        /// <returns></returns>
        /// <exception cref="Exception"></exception>
        public int A(int i = 0)
        {
            if (i > 0)
            {
                throw new Exception("");
            }

            return 0;
        }
    }
}
