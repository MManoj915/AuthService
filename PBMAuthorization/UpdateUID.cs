using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PBMAuthorizationService
{
    class UpdateUID
    {
        public long Member_Code { get; set; }
        public string Passport_No { get; set; }
        public string DateofBirth { get; set; }
        public string Emirates_ID { get; set; }
    }

    class AuthBatchID
    {
        public string BatchID { get; set; }
    }
}

