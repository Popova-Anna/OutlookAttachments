using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OutlookAttachments.Core
{
    public interface IAttachmentSaver
    {
        /// <summary>
        /// Метод получает письма из Outlook, создает папки для сохранения вложений на основе темы письма и даты получения, и сохраняет вложения в соответствующие папки.
        /// </summary>
        /// <param name="startDate">Дата начала выборки</param>
        /// <param name="endDate">Дата окончания выборки</param>
        /// <param name="saveLocation">Путь сохранения</param>
        void SaveAttachments(DateTime startDate, DateTime endDate, string saveLocation);
    }
}
