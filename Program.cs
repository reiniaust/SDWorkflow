using System;
using System.Reflection;
using System.Text.Json;

//string jsonString = File.ReadAllText("QuestionPath.json");
//var questionPath = JsonSerializer.Deserialize<Object>(jsonString);

var help = new
{
    a_SDWorkflow = new
    {
        Aufruf_mit_Parametern = new
        {
            Parmeter_anzeigen = new { }
        },
        Objekte_aus_Datenbank = new { },
        Aufgaben_verwalten = new
        {
            In_Asana_speichern = new { },
            Zustaendiger = new { },
            Termin = new { },
            Prioritaet = new { }
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


object? currentObj;
currentObj = questionPath;
object? yesObj = new { };
object? noObj = new { };

while (currentObj is not null)
{
    //Console.WriteLine("Bitte wählen:");
    Console.WriteLine();

    PropertyInfo[] properties = currentObj.GetType().GetProperties();

    //int i = 0;
    foreach (PropertyInfo property in properties)
    {
        //i += 1;
        string propertyName = property.Name;
        //Console.WriteLine(i.ToString() + ". " + propertyName);
        string showLine = propertyName.Replace("_", " ");

        // Wenn Ja/Nein Frage, dann (j/n) anhängen
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

        Console.WriteLine(showLine);
    }

    //ConsoleKeyInfo keyInfo = Console.ReadKey();
    string? choice = "";
    string? input = Console.ReadLine();
    if (input is not null)
    {
        choice = input;
    }
    //int choice = int.Parse(keyInfo.KeyChar.ToString());

    //i = 0;
    if (choice.Length == 0)
    {
        if (currentObj == questionPath)
        {
            currentObj = null;
        }
        else
        {
            currentObj = questionPath;
        }
    }
    else
    {
        foreach (PropertyInfo property in properties)
        {
            //i += 1;
            //if (i == choice)
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
            {
                if (choice.Length > 1 && property.Name.ToUpper().Contains(choice.ToUpper()))
                {
                    currentObj = property.GetValue(currentObj);
                }
                else
                {
                    if (property.Name.StartsWith(choice))
                    {
                        currentObj = property.GetValue(currentObj);
                    }
                }
            }
        }
    }
}
