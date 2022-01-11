using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Threading;
using EAGetMail;
using EASendMail;

class Program
{
    private static string[] _settingsContent;
    private static string TutorEmail { get; set; }
    private static string TutorPassword { get; set; }
    private static string TuteeFormLink { get; set; }
    private static string LicenseKey { get; set; }
    private static MailServer InServer { get; set; }
    private static MailClient InClient { get; set; }
    private static SmtpMail ToMail { get; set; }
    private static SmtpServer ToServer { get; set; }
    private static SmtpClient ToSmtp { get; set; }
    private static SortedDictionary<string, string> TutorTime { get; set; }
    static void Main(string[] args)
    {
        try
        {
            Console.ForegroundColor = ConsoleColor.White;
            ReadSettings();
            GetTuteeForm();
            GetTutorLogin();
            GetTutorTime();


            while (true)
            {
                PrepareReceive();
                MailInfo[] infos = InClient.GetMailInfos();
                SortedDictionary<string, string> studentToSession = new SortedDictionary<string, string>();

                for (int i = 0; i < infos.Length; i++)
                {
                    MailInfo info = infos[i];

                    Mail inMail = InClient.GetMail(info);

                    if (inMail.From.ToString() == "\"WARC Mail Informer\" <no-reply-warc@auca.kg>" &&
                        inMail.Subject == "New Booking Alert")
                    {
                        string emailOfStudent = GetEmailOfStudent(inMail.TextBody);
                        string dateOfSession = GetDateOfSession(inMail.TextBody);


                        if (checkRegisteredTime(dateOfSession))
                        {
                            studentToSession[dateOfSession] = emailOfStudent;
                            InClient.MarkAsRead(info, true);
                            PrepareSend();
                            SendToTutor(dateOfSession, emailOfStudent);
                        }
                        else
                        {
                            Console.ForegroundColor = ConsoleColor.Red;
                            Console.Error.WriteLine("Tutor registered for this time: " + GetTimeOfSession(dateOfSession) +
                                              " , but there is no such time in settings. Please update settings.txt and restart the program.");
                            Console.ForegroundColor = ConsoleColor.White;
                        }
                    }
                }

                InClient.Quit();
                if (studentToSession.Count > 0)
                {
                    PrepareSend();
                    SendToTutors(studentToSession);
                }

                Thread.Sleep(5000);
            }
        }
        catch (Exception e)
        {
            Console.WriteLine(e.Message);
        }
        finally
        {
            Console.WriteLine("Program stopped. Press any key to exit.");
            Console.ReadKey();
        }
    }

    public static string Reverse(string text)
    {
        if (text == null) return null;
        char[] array = text.ToCharArray();
        Array.Reverse(array);
        return new String(array);
    }

    public static void ReadSettings()
    {
        try
        {
            _settingsContent = System.IO.File.ReadAllLines(Path.Combine(System.AppDomain.CurrentDomain.BaseDirectory, "settings.txt"));
        }
        catch (Exception e)
        {
            throw new Exception("Can't find settings.txt in current folder...Please create the settings.txt");
        }
    }

    public static void GetTuteeForm()
    {
        const int formLinkRow = 0;
        for (int i = 1; i < _settingsContent[formLinkRow].Length; i++)
        {
            if (_settingsContent[formLinkRow][i] == ']') break;
            TuteeFormLink += _settingsContent[formLinkRow][i];
        }
        Console.WriteLine("Link for tutee : {0}", TuteeFormLink);
    }

    public static void GetTutorLogin()
    {
        const int loginRow = 1;
        for (int i = 1; i < _settingsContent[loginRow].Length; i++)
        {
            if (_settingsContent[loginRow][i] == ']')
            {
                for (int j = i + 2; j < _settingsContent[loginRow].Length; j++)
                {
                    if (_settingsContent[loginRow][j] == ']') break;
                    TutorPassword += _settingsContent[loginRow][j];
                }

                break;
            }
            TutorEmail += _settingsContent[loginRow][i];
        }
        Console.WriteLine("Tutor's email: {0}", TutorEmail);
        Console.WriteLine("Tutor's password: {0}", TutorPassword);
        Console.WriteLine("-------------------------------------");
    }

    public static void GetTutorTime()
    {
        const int timeRow = 2;
        TutorTime = new SortedDictionary<string, string>();

        for (int i = timeRow; i < _settingsContent.Length; i++)
        {
            string time = "", link = "";
            for (int j = 1; j < _settingsContent[i].Length; j++)
            {
                if (_settingsContent[i][j] == ']')
                {
                    for (int k = j + 2; k < _settingsContent[i].Length; k++)
                    {
                        if (_settingsContent[i][k] == ']') break;
                        link += _settingsContent[i][k];
                    }
                    break;
                }
                time += _settingsContent[i][j];
                
            }

            TutorTime[time] = link;

            Console.WriteLine("Time: {0} link: {1}", time, link);
        }
        Console.WriteLine("-------------------------------------");
    }

    public static void PrepareReceive()
    {
        LicenseKey = "TEAM BEAN 2014-00666-a8fa1baf7e92e8e38c737232cd1a81ff";

        InServer = new MailServer("imap.gmail.com",
            TutorEmail,
            TutorPassword,
            EAGetMail.ServerProtocol.Imap4);

        InServer.SSLConnection = true;
        InServer.Port = 993;

        InClient = new MailClient(LicenseKey);
        InClient.Connect(InServer);

        // retrieve unread/new email only
        InClient.GetMailInfosParam.Reset();
        InClient.GetMailInfosParam.GetMailInfosOptions = GetMailInfosOptionType.NewOnly;
    }

    public static void PrepareSend()
    {
        ToMail = new SmtpMail();
        ToMail.From = TutorEmail;
        ToMail.Subject = "WARC Booking Informer";

        ToServer = new SmtpServer("smtp.gmail.com");
        ToServer.User = TutorEmail;
        ToServer.Password = TutorPassword;
        ToServer.Port = 465;
        ToServer.ConnectType = SmtpConnectType.ConnectSSLAuto;

        ToSmtp = new SmtpClient();
    }
    public static void SendToTutor(string dateOfSession, string emailOfStudent)
    {
        ToMail.To = emailOfStudent;
        ToMail.TextBody = "Dear " + emailOfStudent + ",\n\t\nPlease find the link for the WARC session on " + dateOfSession + "here: " + TutorTime[GetTimeOfSession(dateOfSession)];
        ToMail.TextBody += "\n\nAfter the session, please complete this form: " + TuteeFormLink;

        ToSmtp.SendMail(ToServer, ToMail);
        Console.WriteLine("Sent link to {0}, date: {1}", emailOfStudent, dateOfSession);
    }

    public static void SendToTutors(SortedDictionary<string, string> SessionToStudent)
    {
        foreach (var session in SessionToStudent)
        {
            ToMail.To = session.Value;
            ToMail.TextBody = "Dear " + session.Value + ",\n\t\nPlease find the link for the WARC session on " + session.Key + "here: " + TutorTime[GetTimeOfSession(session.Key)];
            ToMail.TextBody += "\n\nAfter the session, please complete this form: " + TuteeFormLink;
            
            ToSmtp.SendMail(ToServer, ToMail);
            Console.WriteLine("Sent link to {0}, date: {1}", session.Value, session.Key);
        }
    }

    public static bool checkRegisteredTime(string dateOfSession)
    {
        return TutorTime.ContainsKey(GetTimeOfSession(dateOfSession));
    }

    public static string GetTimeOfSession(string dateOfSession)
    {
        string result = "";
        for (int i = 11; i < 16; i++)
        {
            if (dateOfSession[i] == ' ') break;
            result += dateOfSession[i];
        }

        return result;
    }

    public static string GetEmailOfStudent(string bodyMessage)
    {
        string result = "";
        for (int j = bodyMessage.Length - 9; j >= 0; j--)
        {
            if (bodyMessage[j] == ' ')
                break;
            result += bodyMessage[j];
        }

        return Reverse(result);
    }

    public static string GetDateOfSession(string bodyMessage)
    {
        string result = "";
        for (int j = 7; j < bodyMessage.Length; j++)
        {
            if (char.IsLetter(bodyMessage[j]))
                break;
            result += bodyMessage[j];
        }

        return result;
    }
}
