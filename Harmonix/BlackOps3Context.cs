using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Harmonix
{
    internal class BlackOps3Context
    {
        public string Name { get; set; }

        public string CtxType { get; set; }

        public bool DefaultAlwaysPlay { get; set; }

        public bool Entity { get; set; }

        public bool Ambient { get; set; }

        public bool Global { get; set; }

        public List<string> CtxValues { get; set; }
    }
}
