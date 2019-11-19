using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;

namespace SharepointTemplate
{
   public class DropZoneInfo
    {
        public string siteUrl { get; set; }
        public string listTitle { get; set; }
        public string listId { get; set; }
        public string currentFolder { get; set; }
    }
}
