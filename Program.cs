using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Newtonsoft.Json;
using ParserOpenXML;
using System.Net;
using System.Net.Mail;
using System.Text.RegularExpressions;
using Run = DocumentFormat.OpenXml.Wordprocessing.Run;
using RunProperties = DocumentFormat.OpenXml.Wordprocessing.RunProperties;


string? path = "";

CheckingExistFile(ref path);

//string path = "C:/Users/serg/Desktop/_Приложение.docx";

List<Chapter> chapterList = SelectChapters(path);

var json = JsonConvert.SerializeObject(chapterList);
string? js = JsonConvert.DeserializeObject(json).ToString();

string pathJson = @"MyTest.json";

if (CreateFileJson(js, pathJson))
{
    Console.WriteLine("Файл c json создан");
    Attachment item = new Attachment(pathJson);

    string pattern = @"^(?("")(""[^""]+?""@)|(([0-9a-z]((\.(?!\.))|[-!#\$%&'\*\+/=\?\^`\{\}\|~\w])*)(?<=[0-9a-z])@))" +
                    @"(?(\[)(\[(\d{1,3}\.){3}\d{1,3}\])|(([0-9a-z][-\w]*[0-9a-z]*\.)+[a-z0-9]{2,17}))$";
    string? to = "";

    do
    {
        Console.WriteLine("Введите адрес для отправки сообщения: ");
        to = Console.ReadLine();
    } while (!Regex.IsMatch(to, pattern, RegexOptions.IgnoreCase));

    string mesMail = SendMail(to, item);
    Console.WriteLine(mesMail);
}


List<Chapter> SelectChapters(string _path)
{
    string reg = @"^Глава [0-9]\w*";
    List<string> resultChapter = new List<string>();

    using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(_path, false))
    {
        Body body = wordDoc.MainDocumentPart.Document.Body;
        foreach (Paragraph paragraf in body.Descendants<Paragraph>())
        {
            if (Regex.IsMatch(paragraf.InnerText, reg))
            {
                if (paragraf.InnerText.Length < 500)
                    resultChapter.Add(paragraf.InnerText);
            }
            else
                try
                {
                    if (((paragraf.ParagraphProperties != null) && paragraf.ParagraphProperties.ParagraphStyleId != null) && (paragraf.ParagraphProperties.ParagraphStyleId.Val != null) &&
                        (paragraf.ParagraphProperties.ParagraphStyleId.Val == "a3"))
                    {

                        bool isBold = false;
                        foreach (Run r in paragraf.Descendants<Run>())
                        {
                            if (r.RunProperties != null)
                            {
                                RunProperties rProp = r.RunProperties;

                                if (rProp.Bold != null)
                                    isBold = true;
                                else
                                    isBold = false;
                            }
                            else
                            {
                                isBold = false;
                            }
                        }

                        if (isBold)
                            resultChapter.Add(paragraf.InnerText);
                    }
                }
                catch (Exception)
                {
                    throw;
                }
        }
    }

    List<Chapter> _chapterList = new List<Chapter>();

    for (int i = 0; i < resultChapter.Count; i++)
        if (!Regex.IsMatch(resultChapter[i], reg))
            _chapterList.Add(new Chapter("Глава " + (i + 1).ToString(), resultChapter[i]));
        else
            _chapterList.Add(new Chapter(resultChapter[i].Substring(0, 7), resultChapter[i].Substring(8)));

    return _chapterList;
}


void CheckingExistFile(ref string pathFile)
{
    do
    {
        Console.WriteLine("Введите путь до файла");
        Console.WriteLine("Путь должен быть следующего формата: С:/Users/Name/Desktop/_Приложение.docx");
        pathFile = Console.ReadLine();

        if ((pathFile == null) || pathFile.Length < 5)
        {
            Console.WriteLine("Путь указан не верно");
            continue;
        }

        if (File.Exists(pathFile))
        {
            Console.WriteLine("Файл с текстом найден");
            break;
        }
        else
        {
            Console.WriteLine("Файл с текстом не найден");
        }
    } while (true);
}


bool CreateFileJson(string text, string pathFile)
{
    try
    {
        FileInfo fileInf = new FileInfo(pathFile);
        StreamWriter sw = fileInf.CreateText();
        sw.WriteLine(text);
        sw.Close();
    }
    catch (Exception ex)
    {
        Console.WriteLine("Ошибка создания файла: ");
        Console.WriteLine(ex.ToString());
        return false;
    }

    return true;
}


string SendMail(string to, Attachment file)
{
    try
    {
        string from = @"Test-Email132@yandex.ru";
        string pass = "hgqxutksqrpsrmmv";
        MailMessage mess = new MailMessage();
        mess.To.Add(to);
        mess.From = new MailAddress(from);
        mess.Subject = "Список глав";
        mess.Body = "";
        mess.Attachments.Add(file);
        SmtpClient client = new SmtpClient();
        client.Host = "smtp.yandex.ru";
        client.Port = 587;
        client.EnableSsl = true;
        client.Credentials = new NetworkCredential(from.Split('@')[0], pass);
        client.DeliveryMethod = SmtpDeliveryMethod.Network;

        client.Send(mess);
        mess.Dispose();
    }
    catch (Exception e)
    {
        return "Произошла ошибка отправки " + e.Message;
    }

    return "Сообщение отправлено";
}
