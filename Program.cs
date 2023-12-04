using System;
using System.Reflection;
using System.Text.Json;

//string jsonString = File.ReadAllText("QuestionPath.json");

string jsonString = File.ReadAllText("config.json");
Config? config = JsonSerializer.Deserialize<Config>(jsonString);

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

var dateQuestion = new
{
    Seit_heute = new { },
    Seit_Datum = new { },
    Seit_letztem_Update = new { },
};

var programs = new
{
    Teilestammverwaltung = dateQuestion,
    Modellverwaltung = dateQuestion
};

var priority = new
{
    Hoch = new { },
    Mittel = new { },
    Niedrig = new { }
};


var customerQuestions = new
{
    a_Fehler = new
    {
        Nur_bei_einem_User = new
        {
            j = new
            {
                Userupdate = new { }
            },
            n = programs
        }
    },
    b_Frage = new { },
    c_Anfrage = new { }
};

var customers = new
{
    a_Unbekannt = customerQuestions,
    b_SLR1_Frau_Kuelsch = customerQuestions,
    c_Gienanth_Steyr_Herr_Mueller = customerQuestions
};

var questionPath = new
{
    a_Hilfe = help,
    b_Kundenauswahl = customers
};

InfoNode mainNode = new InfoNode()
{
    Caption = "SDWorkflow",
    ChildNodes =
    new List<InfoNode> {
        new() { Caption = "Problem/Fehler" },
        new() { Caption = "Hilfe", ChildNodes =
    new List<InfoNode> {
        new() { Caption = "SDWorkflow" },
        new() { Caption = "GUSS info" }
    } }
    }
}
;

InfoNode currentNode = mainNode;
//currentObj = questionPath;
//currentObj = mainNode;
//object? yesObj = new { };
//object? noObj = new { };

bool quit = false;

while (currentNode.ChildNodes.Count > 0 && !quit)
{
    Console.WriteLine();

    //PropertyInfo[] properties = currentObj.GetType().GetProperties();

    //foreach (PropertyInfo property in properties)
    foreach (InfoNode infoNode in currentNode.ChildNodes)
    {
        //string propertyName = property.Name;
        //string showLine = propertyName.Replace("_", " ");

        // Wenn Ja/Nein Frage, dann (j/n) anhängen
        /*
        object? childObj = property.GetValue(currentObj);
        yesObj = null;
        noObj = null;
        if (childObj is not null)
        {
            PropertyInfo[] childProperties = childObj.GetType().GetProperties();
            foreach (PropertyInfo prop in childProperties)
            {
                if (prop.Name == "j")
                {
                    showLine += " (j";
                    yesObj = prop.GetValue(childObj);
                }
                if (prop.Name == "n")
                {
                    showLine += "/n)";
                    noObj = prop.GetValue(childObj);
                }
            }
        }
        */

        Console.WriteLine(infoNode.Caption);
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
        //foreach (PropertyInfo property in properties)
        foreach (InfoNode infoNode in mainNode.ChildNodes)
        {
            /*
            if (yesObj is not null)
            {
                if (choice == "j")
                {
                    currentObj = yesObj;
                }
                else
                {
                    currentObj = noObj;
                }
            }
            else
            {*/
            if (infoNode.Caption is not null && infoNode.ChildNodes is not null)
            {
                if (choice.Length > 1 && infoNode.Caption.ToUpper().Contains(choice.ToUpper()))
                {
                    currentNode = infoNode;
                }
                else
                {
                    if (infoNode.Caption.StartsWith(choice))
                    {
                        //currentObj = property.GetValue(currentObj);
                        currentNode = infoNode;
                    }
                }
            }
            //}
        }
    }
}

class Config
{
    public string? UserEmail { get; set; }
}

class InfoNode
{
    public string Caption { get; set; } = "";
    public object? InfoObject { get; set; }
    public List<InfoNode> ChildNodes { get; set; } = new List<InfoNode>();
}