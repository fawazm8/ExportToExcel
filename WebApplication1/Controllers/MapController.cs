using Microsoft.AspNetCore.Mvc;
using System.Net.Http;
using Newtonsoft.Json;
using OfficeOpenXml;
using ExportToExcel.Models;
using OfficeOpenXml.Style;
using System.Drawing;
using System.Net.Mail;
using Microsoft.Extensions.Options;
using System.IO;
using System.Data;
using System.Reflection.Metadata;
using System.Reflection;

namespace ExportToExcel.Controllers
{
    public class MapController : Controller
    {
        private readonly IHttpClientFactory _clientFactory;
        private readonly EmailSettings _emailSettings;

        public MapController(IHttpClientFactory clientFactory, IOptions<EmailSettings> emailSettings)
        {
            _clientFactory = clientFactory;
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            //_emailSettings = "";// (EmailSettings)emailSettings;
        }

        [HttpPost]
        public async Task<IActionResult> ExportReport(string[] selectedColumns)
        {
            var queryFields = string.Join(",", selectedColumns);
            var queryUrl = $"https://services.arcgis.com/V6ZHFr6zdgNZuVG0/arcgis/rest/services/Landscape_Trees/FeatureServer/0/query?where=1%3D1&outFields={queryFields}&outSR=4326&f=json";


            var response = await _clientFactory.CreateClient().GetAsync(queryUrl);
            if (response.IsSuccessStatusCode)
            {
                var responseStream = await response.Content.ReadAsStringAsync();
                var featureResponse = JsonConvert.DeserializeObject<FeatureResponse>(responseStream);

                // Use EPPlus to create the Excel file with the data
                var memoryStream = await GenerateExcelFile(featureResponse.Features, selectedColumns);

                memoryStream.Position = 0;
                string excelName = "Report.xlsx";
                return File(memoryStream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", excelName);
            }

            return View("Error");
        }

        private async Task<MemoryStream> GenerateExcelFile(List<TreeFeature> features, string[] selectedColumns)
        {
            var memoryStream = new MemoryStream();

            using (var package = new ExcelPackage(memoryStream))
            {
                var workSheet = package.Workbook.Worksheets.Add("FeatureData");

                // Define the headers based on selected columns
                workSheet.Cells["A1"].LoadFromCollection(new List<string[]> { selectedColumns }, true);

                // Load the data into the worksheet
                int row = 2; // Start from the second row because the first row has headers
                foreach (var feature in features)
                {
                    for (int i = 0; i < selectedColumns.Length; i++)
                    {
                        string value = GetPropertyValue(feature.Attributes, selectedColumns[i])?.ToString() ?? string.Empty;
                        workSheet.Cells[row, i + 1].Value = value;
                    }
                    row++;
                }


                // Apply any additional formatting you need for the worksheet here
                // For example, autofit columns for all cells
                workSheet.Cells[workSheet.Dimension.Address].AutoFitColumns();

                package.Save();
            }

            return memoryStream;
        }


        public async Task<IActionResult> ExportToExcel()
        {
            List<Attributes> allAttributes = new List<Attributes>();
            int recordOffset = 0;
            bool moreRecordsAvailable = true;
            string baseQueryUrl = "https://webgis.momra.gov.sa/server/rest/services/RealEstate/Parcel_Border/MapServer/2/query";

            // Continue fetching pages of data until no more records are available
            while (moreRecordsAvailable)
            {
                var queryUrl = $"{baseQueryUrl}?where=1%3D1&outFields=*&outSR=4326&resultOffset={recordOffset}&f=json";
                var response = await _clientFactory.CreateClient().GetAsync(queryUrl);

                if (response.IsSuccessStatusCode)
                {
                    var responseStream = await response.Content.ReadAsStringAsync();
                    var featureResponse = JsonConvert.DeserializeObject<FeatureResponse>(responseStream);

                    var features = featureResponse.Features;
                    if (features != null && features.Any())
                    {
                        allAttributes.AddRange(features.Select(f => f.Attributes));
                        recordOffset += features.Count; // Increment the offset by the number of features returned
                    }
                    else
                    {
                        moreRecordsAvailable = false; // No more records, exit the loop
                    }
                }
                else
                {
                    return View("Error"); // Handle the error accordingly
                }
            }

            // Use EPPlus to create an Excel file with all the data collected
            var stream = new MemoryStream();
            using (var package = new ExcelPackage(stream))
            {
                var workSheet = package.Workbook.Worksheets.Add("FeatureData");

                // تطبيق تنسيق العنوان
                var title = workSheet.Cells["A1"];
                title.Value = "عنوان البيانات";
                title.Style.Font.Size = 20;
                title.Style.Font.Bold = true;
                workSheet.Cells[1, 1, 1, 5].Merge = true; // تغيير 5 إلى عدد الأعمدة الفعلي

                // تحميل البيانات
                workSheet.Cells["A2"].LoadFromCollection(allAttributes, true);

                // تحديد عدد الصفوف بناءً على البيانات
                int rowCount = allAttributes.Count + 1; // +1 لأن الصف الأول هو العنوان

                // تطبيق تنسيق رأس الجدول
                workSheet.Row(2).Style.Font.Bold = true;
                workSheet.Row(2).Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Row(2).Style.Fill.BackgroundColor.SetColor(Color.LightBlue);

                // ضبط عرض الأعمدة
                for (int col = 1; col <= 3; col++)
                {
                    workSheet.Column(col).Width = 20; // يمكنك تعديل العرض حسب الحاجة
                }

                // تطبيق تنسيقات الحدود
                var borderRange = workSheet.Cells[$"A2:Z{rowCount}"];
                borderRange.Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Medium);

                // تطبيق خلفية للجدول
                for (int row = 3; row <= rowCount; row++)
                {
                    if (row % 2 == 0)
                    {
                        workSheet.Cells[$"A{row}:Z{row}"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        workSheet.Cells[$"A{row}:Z{row}"].Style.Fill.BackgroundColor.SetColor(Color.LightGray);
                    }
                }


                // تطبيق تنسيقات أخرى
                var dataRange = workSheet.Cells[$"A2:Z{rowCount}"];
                dataRange.Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                dataRange.Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                dataRange.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                dataRange.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;

                // ... التنسيقات الأخرى ...

                package.Save();
            }
            stream.Position = 0;
            string excelName = $"FeatureData-{DateTime.Now:yyyyMMddHHmmss}.xlsx";

            return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", excelName);
        }

        [HttpPost]
        public async Task<IActionResult> ExportAndEmail()
        {
            // Generate the Excel file and store it in a MemoryStream
            var memoryStream = await GenerateExcelFile();

            // Ensure that the MemoryStream is at the beginning
            memoryStream.Position = 0;

            // Send the email
            using (var message = new MailMessage())
            {
                message.To.Add(new MailAddress("fawazm8@gmail.com"));
                message.From = new MailAddress(_emailSettings.FromEmail);
                message.Subject = "Your Exported Data";
                message.Body = "Please find the attached Excel file.";
                message.Attachments.Add(new Attachment(memoryStream, "FeatureData.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"));

                using (var smtpClient = new SmtpClient(_emailSettings.PrimaryDomain, (_emailSettings.PrimaryPort)))
                {
                    smtpClient.Credentials = new System.Net.NetworkCredential(_emailSettings.UsernameEmail, _emailSettings.UsernamePassword);
                    smtpClient.EnableSsl = true;
                    await smtpClient.SendMailAsync(message);
                }
            }

            // Cleanup
            memoryStream.Close();

            return RedirectToAction("Index");
        }
      

        private async Task<MemoryStream> GenerateExcelFile()
        {
            var memoryStream = new MemoryStream();

            // Fetch the data - this should be the same logic as in your ExportToExcel method
            var response = await _clientFactory.CreateClient().GetAsync("your-api-url");
            if (response.IsSuccessStatusCode)
            {
                var responseStream = await response.Content.ReadAsStringAsync();
                var featureResponse = JsonConvert.DeserializeObject<FeatureResponse>(responseStream);
                var features = featureResponse.Features;

                var attributesList = features.Select(f => f.Attributes).ToList();

                using (var package = new ExcelPackage(memoryStream))
                {
                    var workSheet = package.Workbook.Worksheets.Add("FeatureData");

                    // Apply your formatting to the worksheet here
                    // ...

                    // Load the data into the worksheet
                    workSheet.Cells["A2"].LoadFromCollection(attributesList, false);

                    // Save the package to the MemoryStream
                    package.Save();
                }
            }

            // Reset the position of the MemoryStream to the beginning
            memoryStream.Position = 0;

            return memoryStream;
        }

        private DataTable ConvertToDataTable(List<TreeFeature> features, string[] selectedColumns)
        {
            {
                var dataTable = new DataTable();

                // Add columns to the DataTable
                foreach (var column in selectedColumns)
                {
                    dataTable.Columns.Add(column);
                }

                // Add rows to the DataTable
                foreach (var feature in features)
                {
                    var row = dataTable.NewRow();
                    foreach (var column in selectedColumns)
                    {
                        // Use reflection or a switch statement to set the column values based on the property name
                        var value = GetPropertyValue(feature.Attributes, column);
                        row[column] = value ?? DBNull.Value;
                    }
                    dataTable.Rows.Add(row);
                }

                return dataTable;
            }
        }

        private object GetPropertyValue(Attributes attributes, string propertyName)
        {
            // Use a switch statement for properties you know
            switch (propertyName)
            {
                case "REQUEST_NO":
                    return attributes.REQUEST_NO;
                case "Sci_Name":
                    return attributes.REQUEST_TYPE;
                // Add cases for other known properties
                default:
                    // If the property name is not known, use reflection or throw an exception
                    PropertyInfo propInfo = typeof(Attributes).GetProperty(propertyName);
                    if (propInfo != null)
                    {
                        return propInfo.GetValue(attributes, null);
                    }
                    else
                    {
                        throw new ArgumentException("Property not found", propertyName);
                    }
            }
        }
    }
}

