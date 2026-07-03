using DXTools;
using SaltireAPI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Net.Mime;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace BudgetExcelSheets.Classes
{
   internal class SupportMail
   {
		private const string MouleName = "Saltire.Classes.SupportMail";

      internal void Send(string message, string subject, int MaxEmailLength = 0)
      {
         Send(message, subject, "ITSupport@visionprofiles.co.uk", string.Empty, MaxEmailLength);
      }

      internal void Send(string message, string subject, string EmailTo, int MaxEmailLength = 0)
      {
         Send(message, subject, EmailTo, string.Empty, MaxEmailLength);
      }

      internal void Send(string message, string subject, string EmailTo, string AttachmentPath, int MaxEmailLength = 0)
      {
			List<Attachment> attachments = new List<Attachment>();
			if (!string.IsNullOrEmpty(AttachmentPath))
				attachments.Add(new Attachment(AttachmentPath));
			Send(message, subject, EmailTo, attachments, MaxEmailLength);
		}

		internal void Send(string message, string subject, string EmailTo, List<Stream> AttachmentdataStreams, string Filename, int MaxEmailLength = 0)
		{
			string MediaType = string.Empty;
			switch (Path.GetExtension(Filename).ToLower())
         {
				case ".pdf":
					MediaType = MediaTypeNames.Application.Pdf;
					break;
				case ".zip":
					MediaType = MediaTypeNames.Application.Zip;
					break;
				case ".xls":
				case ".xlsx":
				case ".csv":
					MediaType = MediaTypeNames.Application.Octet;
					break;
				case ".rtf":
					MediaType = MediaTypeNames.Application.Rtf;
					break;
				default:
					MediaType = MediaTypeNames.Application.Octet;
					break;
			}
			List<Attachment> data = new List<Attachment>();
			if (AttachmentdataStreams != null)
			{
				foreach (Stream AttachmentdataStream in AttachmentdataStreams)
				{
					data.Add(new Attachment(AttachmentdataStream, Filename, MediaType));
				}
			}

			Send(message, subject, EmailTo, data, MaxEmailLength);
		}

		internal void Send(string message, string subject, string EmailTo, List<Attachment> attachments, int MaxEmailLength = 0)
      {
			clsConfig Config = new clsConfig();
			Config.Retrieve("WHERE ConfigID ='TDL7'");

         VPSSecurity.Global.TDL7 = "KL" + Config.cValue + "YT";
         VPSSecurity.Global.Salt = Encoding.UTF8.GetBytes(Config.cValue.Substring(20, 17));
         VPSSecurity.Global.ProgramID = Config.cValue.Substring(8, 6);

         clsConfigs ConfigList = new clsConfigs();
			ConfigList.RetrieveList();

         if (ConfigList.Count > 0)
         {
            foreach (clsConfig Setting in ConfigList)
            {
               switch (Setting.ConfigID)
               {
                  case "EmailHost":
                     Global.EmailHost = Setting.cValue;
                     break;
                  case "EmailHostUser":
                     Global.EmailHostUser = Setting.cValue;
                     break;
                  case "E_EmailHostPassword":
                     // After next publish change this back to EmailHostPassword
                     Global.EmailHostPassword = new VPSSecurity.Encryption().Decrypt(Setting.cValue);
                     break;
                  case "EmailHostPort":
                     Global.EmailHostPort = CInt(Setting.cValue);
                     break;
                  case "EmailHostSSL":
                     Global.EmailHostSSL = CBol(Setting.cValue);
                     break;
                  case "EmailHostUsePort":
                     Global.EmailHostUsePort = CBol(Setting.cValue);
                     break;
               }

            }
         }

         MailMessage sendEnhancedStockFeedEmail = new MailMessage();
			sendEnhancedStockFeedEmail.From = new MailAddress(Global.EmailHostUser);

			string[] DistributionList = EmailTo.Split(';');

			foreach (string DistMailTo in DistributionList)
			{
				sendEnhancedStockFeedEmail.To.Add(DistMailTo);
			}

			if (attachments != null)
         {
				foreach (Attachment attachment in attachments)
				{
					sendEnhancedStockFeedEmail.Attachments.Add(attachment);
				}
         }

			if (MaxEmailLength > 0)
				message = message.Substring(0, MaxEmailLength) + " ...(Truncated)";

			sendEnhancedStockFeedEmail.Subject = subject;
			sendEnhancedStockFeedEmail.IsBodyHtml = true;
			sendEnhancedStockFeedEmail.Body = message;

			Thread t = new Thread(new ParameterizedThreadStart(Thread_Email));
			t.Start(sendEnhancedStockFeedEmail);
		}

		private static void Thread_Email(object oPass)
		{
			try
			{
				MailMessage ePass = (MailMessage)oPass;

				SmtpClient client = new SmtpClient(Global.EmailHost);
				client.UseDefaultCredentials = false;

            /************************************************************************************
				 * For Auth 2 this credentials section will change
				 * var credentials = new OAuthCredentials(accessToken);
				 * client.Credentials = credentials;
				 ***********************************************************************************/
            client.Credentials = new System.Net.NetworkCredential(Global.EmailHostUser, Global.EmailHostPassword);
				
				if (Global.EmailHostUsePort)
					client.Port = Global.EmailHostPort;
				else
					client.Port = 25;

				client.EnableSsl = Global.EmailHostSSL;

				client.Send(ePass);
			}
			catch (Exception ex)
			{
				// As this is in a seperate thread we dont want to try and return a messagebox so ths will just log in the database if there is an issue
				ProcessError.Return_Error(MouleName, "Thread_Email", ex, new List<string>() { "oPass - MailMessage", Newtonsoft.Json.JsonConvert.SerializeObject(oPass) });
			}
		}

		internal void PopupEmail(string Message, string Subject, List<string> EmailTo, List<string> CC, List<Attachment> attachments)
		{
			try
			{
            MailMessage message = new MailMessage();
            foreach (Attachment AttachmentFile in attachments)
            {
               message.Attachments.Add(AttachmentFile);
            }

            foreach (string EmailToAddress in EmailTo)
            {
               message.To.Add(new MailAddress(EmailToAddress));
            }

            foreach (string CCAddress in CC)
            {
               message.CC.Add(new MailAddress(CCAddress));
            }

            message.From = new MailAddress("fake2@fake.com");
            message.IsBodyHtml = true;
            message.Subject = Subject;
				message.Body = Message + "<br/><br/><br/>" + Global.ReadSignature();

            string tempfilename = Path.GetTempPath() + Guid.NewGuid().ToString() + ".eml";
            SaveExt(message, tempfilename);

            string text = File.ReadAllText(tempfilename);
            text = text.Replace("fake2@fake.com", "");
            File.WriteAllText(tempfilename, text);

            string Application_str = "outlook.exe";
            string Parameters_str = "/eml \"" + tempfilename + "\"";
            System.Diagnostics.Process p = System.Diagnostics.Process.Start(Application_str, Parameters_str);
         }
         catch (Exception ex)
			{
				ProcessError.Show("", "", ex);
			}
		}

      public void SaveExt(MailMessage Message, string FileName)
      {
         Assembly assembly = typeof(SmtpClient).Assembly;
         Type _mailWriterType =
           assembly.GetType("System.Net.Mail.MailWriter");

         using (FileStream _fileStream =
                new FileStream(FileName, FileMode.Create))
         {
            var binaryWriter = new BinaryWriter(_fileStream);
            //Write the Unsent header to the file so the mail client knows this mail must be presented in "New message" mode
            binaryWriter.Write(System.Text.Encoding.UTF8.GetBytes("X-Unsent: 1" + Environment.NewLine));

            // Get reflection info for MailWriter contructor
            ConstructorInfo _mailWriterContructor =
                _mailWriterType.GetConstructor(
                    BindingFlags.Instance | BindingFlags.NonPublic,
                    null,
                    new Type[] { typeof(Stream) },
                    null);

            // Construct MailWriter object with our FileStream
            object _mailWriter =
              _mailWriterContructor.Invoke(new object[] { _fileStream });

            // Get reflection info for Send() method on MailMessage
            MethodInfo _sendMethod =
                typeof(MailMessage).GetMethod(
                    "Send",
                    BindingFlags.Instance | BindingFlags.NonPublic);



            // Call method passing in MailWriter
            //_sendMethod.Invoke(
            //   Message,
            //   BindingFlags.Instance | BindingFlags.NonPublic,
            //   null,
            //   new object[] { _mailWriter, true },
            //   null);
            if (_sendMethod.GetParameters().Length == 2)
            {
               _sendMethod.Invoke(Message, BindingFlags.Instance | BindingFlags.NonPublic, null, new object[] { _mailWriter, true }, null);
            }
            else
            {
               _sendMethod.Invoke(Message, BindingFlags.Instance | BindingFlags.NonPublic, null, new object[] { _mailWriter, true, true }, null);
            }

            // Finally get reflection info for Close() method on our MailWriter
            MethodInfo _closeMethod =
               _mailWriter.GetType().GetMethod(
                   "Close",
                   BindingFlags.Instance | BindingFlags.NonPublic);

            // Call close method
            _closeMethod.Invoke(
                _mailWriter,
                BindingFlags.Instance | BindingFlags.NonPublic,
                null,
                new object[] { },
                null);
         }
      }

      public string StandardTemplate(string body)
		{
			string message = "<html>";
			message += @"<head>
                            <meta http-equiv=""Content-Type"" content=""text/html; charset=us-ascii"">
	                        <style>
		                        body {
			                        font-family: ""Roboto"", ""Helvetica Neue"", Helvetica, Arial, sans-serif;
			                        font-size: 13px;
		                        }
                                small {
                                    font-size: 90%;
                                }
		                        table {
			                        font-family: ""Roboto"", ""Helvetica Neue"", Helvetica, Arial, sans-serif;
			                        font-size: 13px;
			                        border: 1px solid #c0c0c0;
			                        border-collapse: collapse;
		                        }
		                        table th {
			                        font-weight: bold;
		                        }
		                        table th, table td {
			                        text-align: left;
			                        margin: 0px;
			                        padding: 7px 9px 5px 9px;
			                        border: 1px solid #c0c0c0;
		                        }
		                        .text-left {
			                        text-align: left;
		                        }
		                        .text-right {
			                        text-align: right;
		                        }
		                        .text-center {
			                        text-align: center;
		                        }
		                        .alert-info {
			                        padding: 15px;
			                        margin-bottom: 20px;
			                        border: 1px solid transparent;
			                        border-radius: 0px;
			                        box-shadow: 0 2px 2px 0 rgba(0, 0, 0, 0.14), 0 3px 1px -2px rgba(0, 0, 0, 0.2), 0 1px 5px 0 rgba(0, 0, 0, 0.12);
			                        color: #31708f;
			                        background-color: #d9edf7;
			                        border-color: #bce8f1;
		                        }
		                        .alert-warning {
			                        padding: 15px;
			                        margin-bottom: 20px;
			                        border: 1px solid transparent;
			                        border-radius: 0px;
			                        box-shadow: 0 2px 2px 0 rgba(0, 0, 0, 0.14), 0 3px 1px -2px rgba(0, 0, 0, 0.2), 0 1px 5px 0 rgba(0, 0, 0, 0.12);
			                        color: #8a6d3b;
			                        background-color: #fcf8e3;
			                        border-color: #faebcc;
		                        }
		                        .alert-danger {
			                        padding: 15px;
			                        margin-bottom: 20px;
			                        border: 1px solid transparent;
			                        border-radius: 0px;
			                        box-shadow: 0 2px 2px 0 rgba(0, 0, 0, 0.14), 0 3px 1px -2px rgba(0, 0, 0, 0.2), 0 1px 5px 0 rgba(0, 0, 0, 0.12);
		                            color: #a94442;
			                        background-color: #f2dede;
			                        border-color: #ebccd1;
		                        }
	                        </style>
                        </head>
                        <body>";
			message += body;

         message += "</body>";
         message += "</html>";

			return message;
		}

      private static bool CBol(object value)
      {
         try
         {
            string strValue = value.ToString().ToLower();

            switch (strValue)
            {
               case "true":
               case "1":
                  return true;
               case "false":
               case "0":
                  return false;
               default:
                  return false;
            }
         }
         catch
         {
            return false;
         }
      }

      private static int CInt(object value)
      {
         try
         {
            int iout = 0;

            int.TryParse(value.ToString(), out iout);

            return iout;
         }
         catch
         {
            return 0;
         }
      }
   }
}
