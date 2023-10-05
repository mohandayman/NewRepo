
#r "nuget: EPPlus, 5.5.3"
#r "nuget: MimeKit, 2.10.1"

using System;
using System.IO;
using System.Net;
using System.Net.Mail;
using OfficeOpenXml;


		string excelFilePath = "C:\\Users\\mohand\\Desktop\\Egypt_Software_Companiesxlsx.xlsx"; // Replace with your Excel file path

		using (var package = new ExcelPackage(new FileInfo(excelFilePath)))
		{
			var worksheet = package.Workbook.Worksheets[0]; // Assuming data is in the first worksheet
	
			for (int row = 1273; row <= worksheet.Dimension.Rows; row++)
			{
				string email = worksheet.Cells[row, 3]?.Value?.ToString(); // Column C contains email addresses
				string subject = " Application For FUll Stack | .NET Developer | Angular Developer | Front-end | Backend  | Power BI | React | Nodejs  Developer Position "; // Set your subject
				string body = "I am writing to express my strong interest in the Software Developer position at your esteemed organization" +
					",Sincerely,\r\nMohand Ayman\r\nAlexandria, Egypt\r\nEmail: mohandayman0127@gmail.com\r\nPhone: +201212046990\r\n "; // Set your email body
				string attachmentPath = @"C:\Users\mohand\Desktop\My CV\MohandAymanFullStack.pdf"; // Path to CV attachment

				if (!string.IsNullOrWhiteSpace(email))
				{
					SendEmail(email, subject, body, attachmentPath);
				}
			}
		}


static void SendEmail(string toEmail, string subject, string body, string attachmentPath)
{
	using (SmtpClient smtpClient = new SmtpClient("smtp.gmail.com"))
	{
		smtpClient.Port = 587;
		smtpClient.EnableSsl = true;
		smtpClient.Credentials = new NetworkCredential("mohandayman0127@gmail.com", "gnxpbwispvlrkobo");

		using (MailMessage mailMessage = new MailMessage())
		{
			mailMessage.From = new MailAddress("mohandayman0127@gmail.com");
			mailMessage.Subject = subject;
			mailMessage.Body = body;
			mailMessage.To.Add(toEmail);
			if (!string.IsNullOrEmpty(attachmentPath))
			{
				Attachment attachment = new Attachment(attachmentPath);
				mailMessage.Attachments.Add(attachment);
			}

			smtpClient.Send(mailMessage);
			Console.WriteLine($"Mail Was Sent To {toEmail}");



		}
	}
}



