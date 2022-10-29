using ExcelDataReader;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq.Expressions;
using System.Net;
using System.Text;
using static System.Net.Mime.MediaTypeNames;

namespace ExcelDemo.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ExcelController : ControllerBase
    {

        public ExcelController()
        {

        }




        [HttpPost]
        public IActionResult Index(IFormFile formFile)
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            string message = "";

            Stream stream = formFile.OpenReadStream();

            IExcelDataReader reader = null;

            if (formFile.FileName.EndsWith(".xls"))
            {
                reader = ExcelReaderFactory.CreateBinaryReader(stream);
            }
            else if (formFile.FileName.EndsWith(".xlsx"))
            {
                reader = ExcelReaderFactory.CreateOpenXmlReader(stream);
            }
            else
            {
                message = "This file format is not supported";
            }

            DataSet excelRecords = reader.AsDataSet();
            reader.Close();

            var finalRecords = excelRecords.Tables[0];
            int InvoiceId = 1, InvoiceDetailID = 1;



            List<Invoice> InvoiceList = new List<Invoice>();
           

            for (int i = 1; i < finalRecords.Rows.Count; i++)
            {
                Invoice invoice = new Invoice();
                invoice.InvoiceId = InvoiceId.ToString();
                invoice.InvoiceDate = finalRecords.Rows[i][1].ToString();
                invoice.CustomerName = finalRecords.Rows[i][2].ToString();
                invoice.InvoiceDetails= ReadDetails(finalRecords.Rows[i].Table, invoice.InvoiceId, i);
                InvoiceList.Add(invoice);
                InvoiceId++;
            }

         
            return Ok(InvoiceList);
        }
     //   List<List<InvoiceDetail>> DataResult = new List<List<InvoiceDetail>>();
        private List<InvoiceDetail> ReadDetails(DataTable indata, string invoiceID, int index)
        {

            List<InvoiceDetail> invoiceDetailList = new List<InvoiceDetail>();
            int skip = 0;
            for (int i = 0; i < indata.Rows.Count+1; i++)
            {
                var columsHeader = indata.Rows[index].ItemArray.Skip(3);
                var result = columsHeader.Skip(skip).Take(3).ToList();
                InvoiceDetail InvoiceDetails = new InvoiceDetail
                {

                    InvoiceDetailID = result[0].ToString(),
                    Item = result[1].ToString(),
                    Quantity = result[2].ToString(),
                    InvoiceId = invoiceID


                };

                skip = skip + 3;

                invoiceDetailList.Add(InvoiceDetails);

                //DataResult.Add(invoiceDetailList);


               
            }
            return invoiceDetailList;

        }


        public class Invoice
        {
            public string InvoiceId { get; set; }
            public string InvoiceDate { get; set; }

            public string CustomerName { get; set; }

            public List<InvoiceDetail> InvoiceDetails { get; set; } = new List<InvoiceDetail>();

        }

        public class InvoiceDetail
        {
            public string InvoiceDetailID { get; set; }
            public string Item { get; set; }
            public string InvoiceId { get; set; }
            public string Quantity { get; set; }

        }



    }
}
