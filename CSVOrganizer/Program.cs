using System;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Net.Mail;
using System.Text;

class Program
{
    private static List<string> links;
    static string filePath = "./mails.csv"; // replace with the path to your CSV file
    static string outputFilePath = "./output.csv"; // replace with the path where you want to save the output file

    static void Main(string[] args)
    {
        if (!CheckForDuplicatedLink())
        {
            ReadCSVandSend();
        }
    }

    static bool CheckForDuplicatedLink()
    {
        links = new List<string>();
        var i = 1;
        try
        {
            using (StreamReader sr = new StreamReader(filePath))
            using (StreamWriter sw = new StreamWriter(outputFilePath))
            {
                string line;
                while ((line = sr.ReadLine()) != null)
                {
                    string[] fields = line.Split(',');


                    string link = fields[2];

                    if (links.Exists(s => s == link))
                    {
                        Console.WriteLine($"Duplicated Link Esists at {i}");
                        return true;
                    }

                    if (links.Exists(s => s.Equals(link)))
                    {
                        Console.WriteLine($"Duplicated Link Esists at {i}");
                        return true;
                    }
                    i++;

                }
            }

            Console.WriteLine("Done");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }

        return false;
    }
    private static void ReadCSVandSend()
    {
       
        var i = 43;
        try
        {
            using (StreamReader sr = new StreamReader(filePath))
            using (StreamWriter sw = new StreamWriter(outputFilePath))
            {
                string line;
                while ((line = sr.ReadLine()) != null)
                {
                    string[] fields = line.Split(',');

                    // extract name from first column
                    string name = fields[0];

                    // extract email from second column
                    string email = fields[1];

                    // extract link from third column
                    string link = fields[2];

                    if (string.IsNullOrEmpty(link)) continue;
                    // create text for fourth column
                    string text =
                        $"Hi {name} \nHere is your Tim Hortons Gift card: {link} \n\nThank you for participating in the hackathon \nWe are looking forward see you again in our future events. \nStay tuned for our upcoming announcements by following our Instagram: https://www.instagram.com/schulichignite/ \n\nSincerely, \nPayam \nJr VP External Schulich Ignite";

                    Console.WriteLine($"to {email}");
                    Console.WriteLine($"#{i} {text}");
                    i++;
                    SendEmail(email, text);
                    // break;
                    // write output line to output file
                    StringBuilder outputLine = new StringBuilder();
                    outputLine.Append(fields[0]).Append(",");
                    outputLine.Append(fields[1]).Append(",");
                    outputLine.Append(fields[2]).Append(",");
                    outputLine.Append(text);
                    // sw.WriteLine(outputLine.ToString());
                }
            }

            Console.WriteLine("Done");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }

    static void SendEmail(string toAddress, string body)
    {
        string fromAddress = "payam.ranjbar@live.com"; // replace with your own email address
        string password = ""; // replace with your own email password
        // string password = "#"; // replace with your own email password
        // toAddress = "pm.rastdast@gmail.com";

        MailMessage message = new MailMessage(fromAddress, toAddress);
        message.Subject = "CANIS Hackathon Gift Cards";
        message.Body = body;
        message.Bcc.Add("info@schulichignite.com");
        message.Bcc.Add("pm.rastdast@gmail.com");


        
        SmtpClient smtp = new SmtpClient("smtp.office365.com", 587); 
        smtp.UseDefaultCredentials = false;
        smtp.EnableSsl = true;
        smtp.Credentials = new NetworkCredential(fromAddress, password);
        smtp.Send(message);
    }

}