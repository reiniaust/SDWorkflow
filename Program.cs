using System;
using System.Reflection;
using System.Text.Json;
using Outlook = Microsoft.Office.Interop.Outlook;
using Aspose.Email;
using Aspose.Email.Clients.Smtp;
using Aspose.Email.Clients;
using System.Text;
using Aspose.Email.Clients.Pop3;

// Konfiguration lesen
string jsonString = File.ReadAllText("config.json");
Config? config = JsonSerializer.Deserialize<Config>(jsonString);




// E-Mail senden (Test)
if (config is not null)
{
    //SenEmail(config, "Test aus SDWorkflow", "Hallo Welt", "reinhard.austermeier@gmail.com");
    foreach (MailMessage message in ReadEmails(config))
    {
        Console.WriteLine(message.From);
        Console.WriteLine(message.Subject);
        Console.WriteLine(message.Body);
        Console.WriteLine("a Beantworten");
        Console.WriteLine("b Überspringen");
        string? input = Console.ReadLine();
        if (input == "a")
        {
            SenEmail(config, message.From.Address, "AW: " + message.Subject, "Hallo, ich leite es weiter.");
        }
    }
}


var help = new
{
    a_SDWorkflow = new
    {
        In_Planung = new
        {
            Aufruf_mit_Parametern = new
            {
                Parmeter_anzeigen = new { }
            },
            Objekte_aus_Datenbank = new
            {
            },
            Objekte_mit_Daten = new { },
            Email = new
            {
                Email_schreiben = new { },
                Email_lesen = new { },
                Email_beantworten = new { },
            },
            Aufgaben_verwalten = new
            {
                In_Asana_speichern = new { },
                Zustaendiger = new { },
                Termin = new { },
                Prioritaet = new { }
            }
        },
        Umgesetzt = new
        {
            User_verwalten = new
            {
                Eigener_User = "UserEMail in Config.json",
            },
        }
    },
    b_GUSS_info = new { }
};

var problemSolving = new
{
};

var priority = new
{
    Hoch = new { },
    Mittel = new { },
    Niedrig = new { }
};

InfoNode programs = new() 
{
    Caption = "Programme/Module",
    ChildNodes = new List<InfoNode>(){
        new() { 
            Caption = "Modellverwaltung (Guss-mv)",
            ChildNodes = new List<InfoNode>(){
                new() {Caption = "Chronik"},
                new() {Caption = "Dokumente"}
            }
        },
        new() { Caption = "Kalkulation" }
    }
};

InfoNode errorProblem = new InfoNode() {
    Caption = "Fehler/Problem",
    ChildNodes = new List<InfoNode>() {
        new() {
            Caption = "Nur bei einem User",
            ChildNodes = new List<InfoNode>(){
                new() {Caption = "Userupdate"}
            }
        },
        programs,
        new() {
            Caption = "Zentrales Problem, seit wann",
            ChildNodes = new List<InfoNode>(){
                new() {Caption = "Heute"},
                new() {Caption = "Gestern"},
                new() {Caption = "Tagen"},
                new() {Caption = "Wochen"},
                new() {Caption = "Unklar"}
            }
        },
    }
};

List<InfoNode> questionTypes = new() {
    errorProblem
};

InfoNode customers = new InfoNode() 
{
    Caption = "Kunden",
    ChildNodes = new List<InfoNode>(){
        new() { 
            Caption = "166 SLR1",
            ChildNodes = new List<InfoNode>()
            {
                new() {
                    Caption = "Frau Külsch",
                    ChildNodes = questionTypes
                },
                new() {
                    Caption = "Herr Clerici",
                    ChildNodes = questionTypes
                }
            }
        },
        new() { Caption = "275 Eurotech" }
    }
};

InfoNode mainNode = new InfoNode()
{
    Caption = "SDWorkflow",
    ChildNodes =
    new List<InfoNode> {
        new() { Caption = "Hilfe/Anleitungen" },
        customers
    }
}
;

InfoNode currentNode = mainNode;

while (currentNode.ChildNodes.Count > 0)
{
    Console.WriteLine();

    int ascii = 97;
    foreach (InfoNode infoNode in currentNode.ChildNodes)
    {
        Console.WriteLine(((char) ascii).ToString() + " " + infoNode.Caption);
        ascii++;
    }

    string? choice = "";
    string? input = Console.ReadLine();
    if (input is not null)
    {
        choice = input;
    }

    if (choice.Length == 0)
    {
        if (currentNode == mainNode)
        {
            currentNode = new InfoNode();
        }
        else
        {
            currentNode = mainNode;
        }
    }
    else
    {
        ascii = 97;
        foreach (InfoNode infoNode in currentNode.ChildNodes)
        {
            if (infoNode.Caption is not null && infoNode.ChildNodes is not null)
            {
                if (choice.Length > 1 && infoNode.Caption.ToUpper().Contains(choice.ToUpper()))
                {
                    currentNode = infoNode;
                }
                else
                {
                    string label = ((char) ascii).ToString() + infoNode.Caption;
                    if (label.StartsWith(choice))
                    {
                        currentNode = infoNode;
                    }
                    ascii++;
                }
            }
        }
    }
}

static List<MailMessage> ReadEmails(Config config) {

    List<MailMessage> eMails = new();

    Pop3Client client = new() {
        Host = config.PopHost,
        Port = config.PopPort,
        Username = config.Username,
        Password = config.SmtpPassword
    };

    int messageCount = client.GetMessageCount();

    for (int i = 1; i <= messageCount; i++)
    {
        var messageInfo = client.GetMessageInfo(i);
        MailMessage message = client.FetchMessage(messageInfo.SequenceNumber);
        eMails.Add(message);
        /*
        Console.WriteLine("Subject: " + message.Subject);
        Console.WriteLine("From: " + message.From);
        Console.WriteLine("----------");
        */
    }
    return eMails;
}


#pragma warning disable CS8321 // Local function is declared but never used
static void SenEmail(Config config, string to, string subject, string body)
{
    MailMessage message = new()
    {
        Subject = subject,
        Body = body,
        From = new MailAddress(config.Username)
    };

    // Zu Empfängern und CC-Empfängern hinzufügen
    message.To.Add(new MailAddress(to));
    //message.CC.Add(new MailAddress("reinhard.austermeier@gmx.de", "Reinhard Austermeier", false));

    // Anhänge hinzufügen
    //message.Attachments.Add(new Attachment("word.docx"));

    // MailMessage-Instanz erstellen. Sie können eine neue Nachricht erstellen oder eine bereits erstellte Nachrichtendatei (eml, msg usw.)
    //MailMessage msg = MailMessage.Load("EmailMessage.msg");

    SmtpClient client = new()
    {
        // Geben Sie Ihren Mailing-Host, Benutzernamen, Passwort, Portnummer und Sicherheitsoption an
        Host = config.SmtpHost,
        Port = config.SmtpPort,
        Username = config.Username,
        Password = config.SmtpPassword
    };
    //client.SecurityOptions = SecurityOptions.SSLExplicit;
    try
    {
        // Senden Sie diese Email
        client.Send(message);
    }
    catch (Exception ex)
    {
        Console.WriteLine(ex.ToString());
    }
}
#pragma warning restore CS8321 // Local function is declared but never used

class Config
{
    public string? Username { get; set; }
    public string? SmtpHost { get; set; }
    public string? SmtpPassword { get; set; }
    public int SmtpPort { get; set; }
    public string? PopHost { get; set; }
    public int PopPort { get; set; }
}

class InfoNode
{
    public string Caption { get; set; } = "";
    public object? InfoObject { get; set; }
    public List<InfoNode> ChildNodes { get; set; } = new List<InfoNode>();
}