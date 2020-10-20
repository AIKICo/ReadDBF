using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReadDBF
{
    public static  class ExtentionMethodCollection
    {
        public static string Inverse(this string @base)
        {
            return new string(@base.Reverse().ToArray());
        }
    }
}
