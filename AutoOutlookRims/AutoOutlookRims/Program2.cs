using System;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Threading.Tasks;

namespace AutoOutlookRims //.SendMail
{
    public partial class Program
    {

        // Метод отправки письма , нужно сформировать отчет для письма.
        //public async Task<bool> SendM(IMySettings configiniP, long countDB)
        async static Task<bool> SendM()
        {
            DateTime dt = DateTime.Now;
            string htmlH = @"<!DOCTYPE HTML PUBLIC '-//W3C//DTD HTML 4.01//EN' 'http://www.w3.org/TR/html4/strict.dtd'>
                            <html><head>
                            <meta http-equiv='Content-Type' content='text/html; charset=utf-8'>
                            <title> Отчет AutoOutlookRims </title>
                            <style>
                            .layer1 { font: normal 12pt/10pt serif;} 
                            .cap { font: bold italic 12pt serif; }
                             </style></head>";

            string htmlT1 = @"<table width='700' cellpadding='5' cellspacing='1' border='0'>
                            <tr>
	                        <td align='center' class='cap'>форма отчета программы AutoOutlookRims<td>
                            </tr>
                            </table>";

            string htmlT2 = $@"<table width='700' cellpadding='5' cellspacing='1' border='1'>
	                        <tr bgcolor='#81B764'>
                            <td colspan= '2' class='layer1' align='left'>Отчет по пунктам: </td>
		                    <td align='center'>{dt:dd-MM-yyyy}</td>
                            </tr>
                            <tr>
		                    <td colspan='2' align='left'>1. Количество пользователей в базе MS SQL</td>
		                    <td align='center'>{countDB}</td>
	                        </tr>
	                        <tr>
		                    <td colspan='2' align='left'>2. Количество пользователей со своим Автоответом</td>
		                    <td align='center'>{statusAUsers}</td>
	                        </tr>
	                        <tr>
		                    <td colspan='2' align='left'>3. Количество используемых потоков</td>
		                    <td align='center'>{QThred}</td>
	                        </tr>
	                        <tr>
		                    <td colspan='2' align='left'>4. Время выполнения получения статуса Автоответа для всех пользователей</td>
		                    <td align='center'>{elapsedTimeGetStatus}</td>
	                        </tr>
	                        <tr>
		                    <td colspan='2' align='left'>5. Время применения запроса на установку Автоотвкта пользователям</td>
		                    <td align='center'>{elapsedTimeSetAutoAnswer}</td>
	                        </tr>
	                        <tr>
		                    <td colspan='2' align='left'>5. Список пользователей исключенных из обработки программы:</td>
		                    <td align='center'>{exlistusers.Count()}</td>
	                        </tr>
	                        <tr>
		                    <td colspan='3' align='left'>{string.Join(",", exlistusers.Select(p => p)).TrimEnd(',')}</td>
	                        </tr>                            
                            </table>";
            string htmlF = @"</html>";


            MailAddress from = new MailAddress(configiniP.UserFrom, configiniP.UserFrom.Split('@')[0]);
            MailAddress to = new MailAddress(configiniP.UserTo);
            MailAddress cc = new MailAddress(configiniP.UserCc);
            MailMessage m = new MailMessage(from, to);
            m.CC.Add(cc);
            // письмо представляет код html
            m.IsBodyHtml = true;
            m.Subject = $"Отчет AutoOutlookRims {dt:dd.MM.yyyy}";

            string pathARH = $"logs\\{dt:yyyy-MM-dd}.zip";
            if (File.Exists(pathARH))
            {
                try
                {
                    logger.Info($"файл {pathARH}, прикреплен к письму");
                    m.Attachments.Add(new Attachment(pathARH));
                    Console.WriteLine($"файл {pathARH}, прикреплен к письму");
                }
                catch (Exception ex)
                {
                    logger.Info($"{ex.Message}");
                }
            }
            else
            {
                htmlF = $@"<p style='font-size: 10px' align='right'>Файл {pathARH} не найден :( </p>
                </html>";
            }
            m.Body = htmlH + htmlT1 + htmlT2 + htmlF;

            SmtpClient smtp = new SmtpClient(configiniP.ServerPost, configiniP.Port);
            smtp.UseDefaultCredentials = true;
            smtp.EnableSsl = false;
            await smtp.SendMailAsync(m);

            m.Dispose();
            return Task.CompletedTask.IsCompleted;
        }


    }
}

