using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Mail;
using Telerik.Reporting.Processing;
using System.IO;
using System.Threading.Tasks;

public class AttachDocuments
{
    public MemoryStream Document { get; set; }
    public string Name { get; set; }
    public Extension Extension { get; set; }
}