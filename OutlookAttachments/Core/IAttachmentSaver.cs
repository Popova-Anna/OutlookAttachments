using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OutlookAttachments.Core
{
    public interface IAttachmentSaver
    {
        void SaveAttachments(DateTime startDate, DateTime endDate, string saveLocation);
    }
}
