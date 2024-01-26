using Newtonsoft.Json;
using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Collections.Generic;
using MimeKit;
using MsgReader;
using System.Security.Cryptography;
using static System.Runtime.InteropServices.JavaScript.JSType;
using RtfPipe.Tokens;

MyClass currentItem;
MyClass newItem = null;
MyClass modifyItem = null;
MyClass cutItem = null;
MyClass dependenceItem = null;
List<MyClass> currentList;
string json;
List<MyClass> data;
string? input = "s";
string filePath = @"S:\SDWorkflow\";
string fileName = "data.json";
string fileNameToDay = "data" + DateTime.Today.ToString().Split(" ")[0].Replace(".", "") + ".json";
string varName = "";
string varValue = "";
Dictionary<string, string> paramList = new Dictionary<string, string>();

int weekdayNumber = (int)DateTime.Now.DayOfWeek;

// Wochentage
Dictionary<int, string> daysOfWeek = new Dictionary<int, string>();
{
    daysOfWeek.Add(1, "Montag");
    daysOfWeek.Add(2, "Dienstag");
    daysOfWeek.Add(3, "Mittwoch");
    daysOfWeek.Add(4, "Donnerstag");
    daysOfWeek.Add(5, "Freitag");
    daysOfWeek.Add(6, "Samstag");
    daysOfWeek.Add(0, "Sonntag");
}
string weekdayName = daysOfWeek[weekdayNumber];

//HelperClass.SetForeground("Outlook");
executeCommand("subst s: \"g:\\Meine Ablage\"");


while (input == "s")
{
    try
    {
        json = File.ReadAllText(filePath + fileName);
        data = JsonConvert.DeserializeObject<List<MyClass>>(json);
    }
    catch (System.Exception)
    {
        data = new();
        data.Add(new MyClass() { Id = 1, Name = "SDWorkflow" });
    }

    MyClass dayItem;
    dayItem = data.Find(item => item.Name == weekdayName);
    if (dayItem != null)
    {
        dayItem.Name = "heute";
        dayItem.TimeStamp = DateTime.Now;
    }
    // morgigen Wochentag auf "morgen" ändern
    dayItem = data.Find(item => item.Name == daysOfWeek[weekdayNumber + 1]);
    if (dayItem != null)
    {
        dayItem.Name = "morgen";
        dayItem.TimeStamp = DateTime.Now;
    }
    // dann "heute" auf "gestern"
    dayItem = data.Find(item => item.Name == "heute" && item.TimeStamp.AddDays(1).Date == DateTime.Today);
    if (dayItem != null)
    {
        dayItem.Name = "gestern";
        dayItem.TimeStamp = DateTime.Now;
    }
    // dann "morgen" auf "heute"
    dayItem = data.Find(item => item.Name == "morgen" && item.TimeStamp.AddDays(1).Date == DateTime.Today);
    if (dayItem != null)
    {
        dayItem.Name = "heute";
        dayItem.TimeStamp = DateTime.Now;
    }
    if (dayItem != null && dayItem.TimeStamp == DateTime.Now)
    {
        saveData();
    }

    currentItem = data[0];

    // E-Mails aus dem Ordner lesen, Punkte anlegen und verschieben
    {
        //string downloadFolderPath = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + "\\Downloads\\";
        string downloadFolderPath = filePath;
        string[] files = Directory.GetFiles(downloadFolderPath, "*.msg");
        foreach (string file in files)
        {
            currentItem = data.Find(item => item.Name == "Vorgänge");
            var message = new MsgReader.Outlook.Storage.Message(file);
            setAndSaveNewItem(message.Subject);
            message.Dispose();
            string newFolder = filePath + newItem.Id;
            Directory.CreateDirectory(newFolder);
            MyClass emailItem = newItem;

            string newPath = newFolder + "\\" + Path.GetFileName(file);
            Directory.Move(file, newPath);

            newPath = "\"" + newPath + "\"";
            currentItem = newItem;
            setAndSaveNewItem(newPath);
            executeCommand(newPath);

            Console.WriteLine(message.Subject);

            Console.WriteLine("Termin:");
            string dateString = Console.ReadLine();
            MyClass dateItem = data.Find(item => item.Name == dateString);
            if (dateItem is null)
            {
                currentItem = data.Find(item => item.Name == "Termine");
                setAndSaveNewItem(dateString);
                dateItem = newItem;
            }
            emailItem.DependenceIds.Add(dateItem.Id);
            //Console.WriteLine("der Vorgang wird bearbeitet und bis " + dateString + " ein Feedback gegeben.");
            Console.WriteLine("Mögliche Antwort: wir werden es bearbeiten und " + dateString + " ein Feedback gegeben.");
            Console.ReadKey();

            newItem = null;

        }
    }


    currentList = data.Where(item => item.ParentId == currentItem.Id).ToList();

    showList();

    saveData();
}

void saveData()
{
    json = JsonConvert.SerializeObject(data, Formatting.Indented);
    File.WriteAllText(filePath + fileName, json);
    File.WriteAllText(filePath + fileNameToDay, json);
}

void showList()
{
    try
    {
        Console.Clear();
    }
    catch (System.Exception)
    { }

    if (input == "")
    {
        setVariableOrExecuteCommand(currentItem.Name);
    }


    // Daten anzeigen
    {
        string done = "";
        if (currentItem.Done)
        {
            done = "Erledigt ";
        }
        Console.WriteLine(currentItem.Name + " (" + done + currentItem.TimeStamp + ")");
        Console.WriteLine();

        if (currentItem.Name.StartsWith("Termine"))
        {
            currentList = currentList.OrderBy(i => dateFromWord(i.Name)).ToList();
        }
        else
        {
            currentList = currentList.OrderBy(i => i.TimeStamp).ToList();
        }


        int i = 0;
        foreach (MyClass item in currentList)
        {
            // Pluszeigen setzen, wenn Unterpunkte, Abhängikeiten oder Verwendungen im Unterpunkt sind
            string plus = "";
            i += 1;
            item.Position = i;
            if (data.Where(child => child.ParentId == item.Id).ToList().Count > 0 || item.DependenceIds.Count > 0)
            {
                plus = " +";
            }
            if (plus == "")
            {
                foreach (var depItem in data.Where(item => !item.Done))
                {
                    if (depItem.DependenceIds.Contains(item.Id))
                    {
                        plus = " +";
                        break;
                    }
                }
            }

            setParameter(item.Name);

            done = "";
            if (item.Done)
            {
                done = " (Erledigt)";
            }
            Console.WriteLine(item.Position + " " + replaceParameterValueInText(item.Name) + done + plus);
        }

        // Anhängigkeiten anzeigen
        if (currentItem.DependenceIds.Count > 0)
        {
            Console.WriteLine();
            Console.WriteLine("Anhängigkeiten:");
            foreach (int id in currentItem.DependenceIds)
            {
                MyClass depItem = data.Find(item => item.Id == id);
                if (depItem != null)
                {
                    i += 1;
                    depItem.Position = i;
                    Console.WriteLine(depItem.Position + " " + depItem.Name);
                }
            }
        }

        // Verwendungen anzeigen
        bool uses = false;
        foreach (var item in data.Where(item => !item.Done)) // Nur nicht erledigte Punkte anzeigen
        {
            if (item.DependenceIds.Contains(currentItem.Id))
            {
                if (!uses)
                {
                    uses = true;
                    Console.WriteLine();
                    Console.WriteLine("Verwendungen:");
                }
                i += 1;
                item.Position = i;
                Console.WriteLine(item.Position + " " + item.Name);
            }
        }
    }

    Console.WriteLine();
    Console.WriteLine("h Hinzufügen");
    if (currentItem.Id != 1)
    {
        Console.WriteLine("ä Ändern");
        if (cutItem is null)
        {
            Console.WriteLine("a Ausschneiden");
            Console.WriteLine("l Löschen");
        }
        else
        {
            Console.WriteLine("a Ausgeschnittenen Punkt einfügen");
        }

        if (currentItem.Done)
        {
            Console.WriteLine("e Auf Nicht Erledigt setzen");
        }
        else
        {
            Console.WriteLine("e Auf Erledigt setzen");
        }

        if (dependenceItem is null)
        {
            Console.WriteLine("v Verknüpfung/Abhängigkeit hinzufügen");
        }
        else
        {
            Console.WriteLine("v Verknüpfung/Abhängigkeit setzen");
        }

    }

    if (currentList.Where(item => item.Name.StartsWith("Befehl:") || item.Name.StartsWith("[")).Count() > 0)
    {
        Console.WriteLine("b Befehle ausführen und/oder Platzhalter zuweisen");
    }
    if (currentItem.Id != 1)
    {
        Console.WriteLine("z Zurück");
    }
    if (currentItem.Id != 1 || cutItem is not null)
    {
        Console.WriteLine("s Startseite");
    }
    Console.WriteLine("x Ende");
    Console.WriteLine();

    if (input == "h" || input == "ä")
    {
        // Ändern
        Console.WriteLine("Neuer Titel:");
    }

    input = Console.ReadLine();

    if (input != "x" && input != "s")
    {
        if (input == "z" && currentItem.ParentId != 0)
        {
            currentItem = data.Find(item => item.Id == currentItem.ParentId);
        }
        else
        {
            if (newItem is not null)
            {
                // Text hinzufügen und speichern
                setAndSaveNewItem(input);
                newItem = null;
            }
            else
            {

                if (modifyItem is not null)
                {
                    // Änderung speichern
                    modifyItem.Name = input;
                    saveData();
                    modifyItem = null;
                }
                else
                {
                    if (input == "h")
                    {
                        // Hinzufügen
                        newItem = new();
                    }
                    else
                    {
                        if (input == "v")
                        {
                            // Verknüpfung/Abhängigkeit
                            if (dependenceItem is null)
                            {
                                dependenceItem = currentItem;
                            }
                            else
                            {
                                dependenceItem = data.Find(item => item.Id == dependenceItem.Id);
                                dependenceItem.DependenceIds.Add(currentItem.Id);
                                currentItem = dependenceItem;
                                saveData();
                                dependenceItem = null;
                            }
                        }
                        else
                        {
                            if (input == "ä")
                            {
                                // Ändern
                                modifyItem = currentItem;
                            }
                            else
                            {
                                if (input == "a")
                                {
                                    // Ausschneiden
                                    if (cutItem is null)
                                    {
                                        cutItem = currentItem;
                                        currentItem = data.Find(item => item.Id == currentItem.ParentId);
                                        cutItem.ParentId = 0;
                                    }
                                    else
                                    {
                                        // Ausgeschnittenen Punkt einfügen
                                        cutItem.ParentId = currentItem.Id;
                                        cutItem.TimeStamp = DateTime.Now;
                                        saveData();
                                        cutItem = null;
                                    }
                                }
                                else
                                {
                                    if (input == "l")
                                    {
                                        // Löschen
                                        data.Remove(currentItem);
                                        saveData();
                                        currentItem = data.Find(item => item.Id == currentItem.ParentId);
                                    }
                                    else
                                    {
                                        if (input == "e")
                                        {
                                            // Erledigt/Unerledigt
                                            currentItem.Done = !currentItem.Done;
                                            currentItem.TimeStamp = DateTime.Now;
                                            saveData();
                                        }
                                        else
                                        {
                                            // Befehle ausführen
                                            if (input == "b")
                                            {
                                                foreach (var item in currentList.Where(item => item.Name.StartsWith("Befehl:") || item.Name.StartsWith("[")).ToList())
                                                {
                                                    setVariableOrExecuteCommand(item.Name);
                                                }
                                                Console.WriteLine("Beliebige Taste drücken...");
                                                Console.ReadKey();
                                            }
                                            else
                                            {
                                                if (currentItem.Name.Contains("]="))
                                                {
                                                    // Parameter setzen (z.B. [Faelligkeit]=morgen)
                                                    currentItem.Name = currentItem.Name.Split('=')[0] + "=" + input;
                                                    setParameter(currentItem.Name);
                                                    saveData();
                                                    // eine Ebene zurückgehen
                                                    currentItem = data.Find(item => item.Id == currentItem.ParentId);
                                                }
                                                else
                                                {
                                                    int number;
                                                    bool isNumber = int.TryParse(input, out number);
                                                    bool found = false;
                                                    foreach (MyClass item in currentList)
                                                    {
                                                        if (isNumber && number == item.Position || !isNumber && item.Name.ToLower().Contains(input.ToLower()))
                                                        {
                                                            currentItem = item;
                                                            found = true;
                                                            break;
                                                        }
                                                    }
                                                    if (!found)
                                                    {
                                                        // Verknüpfung/Abhängigkeit suchen/aufrufen
                                                        foreach (int id in currentItem.DependenceIds)
                                                        {
                                                            MyClass depItem = data.Find(item => item.Id == id);
                                                            if (input == depItem.Position.ToString() || depItem.Name.ToLower().Contains(input.ToLower()))
                                                            {
                                                                currentItem = depItem;
                                                                found = true;
                                                                break;
                                                            }
                                                        }
                                                    }
                                                    if (!found)
                                                    {
                                                        foreach (var item in data)
                                                        {
                                                            if (item.DependenceIds.Contains(currentItem.Id))
                                                            {
                                                                if (input == item.Position.ToString() || item.Name.ToLower().Contains(input.ToLower()))
                                                                {
                                                                    currentItem = item;
                                                                    found = true;
                                                                    break;
                                                                }
                                                            }
                                                        }
                                                    }
                                                    if (!found)
                                                    {
                                                        MyClass foundItem = FoundChildItem(currentList, input);

                                                        // Neuen Punkt anlegen/hinzufügen
                                                        if (foundItem is not null)
                                                            currentItem = foundItem;
                                                    }

                                                    input = "";
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
        if (currentItem != null)
        {
            currentList = data.Where(item => item.ParentId == currentItem.Id).ToList();
        }
        showList();
    }

    MyClass FoundChildItem(List<MyClass> list, string search)
    {
        MyClass item = null;
        foreach (MyClass child in list)
        {
            if (child.Name.ToLower().Contains(search.ToLower()))
            {
                item = child;
                break;
            }
        }
        if (item is null)
        {
            foreach (MyClass child in list)
            {
                if (item is null)
                {
                    item = FoundChildItem(data.Where(i => i.ParentId == child.Id).ToList(), search);
                }
            }
        }
        return item;
    }

    string Path(MyClass item)
    {
        string path = "";
        while (item.ParentId != 0)
        {
            item = data.FirstOrDefault(i => i.Id == i.ParentId);
            path += item.Name;
        }
        return path;
    }
}

void setAndSaveNewItem(string name)
{
    newItem = new()
    {
        Id = data.Max(item => item.Id) + 1,
        Name = name,
        ParentId = currentItem.Id,
        TimeStamp = DateTime.Now
    };
    data.Add(newItem);
    saveData();
}


// Variable/Parameter/Platzhalter lesen
void setParameter(string itemName)
{
    if (itemName.StartsWith("["))
    {
        varName = itemName.Split("=")[0];
        varValue = itemName.Split("=")[1];
        if (paramList.ContainsKey(varName))
        {
            paramList.Remove(varName);
        }
        paramList.Add(varName, varValue);
    }
}

// Paramaternamen (in eckigen Klammern) durch den Wert/Inhalt ersetzen
string replaceParameterValueInText(string itemName)
{
    string result = itemName;
    if (!itemName.Contains("]="))
    {
        foreach (KeyValuePair<string, string> entry in paramList)
        {
            if (entry.Key != "")
            {
                result = result.Replace(entry.Key, entry.Value);
            }
        }
    }
    return result;
}

void setVariableOrExecuteCommand(string itemName)
{
    setParameter(itemName);

    // Befehl ausführen
    if (itemName.StartsWith("Befehl:"))
    {
        string cmd = itemName.Split("Befehl:")[1] + replaceParameterValueInText(itemName);

        executeCommand(cmd);
    }
    if (itemName.Contains(@":\"))
    {
        executeCommand(itemName);
    }
}

void executeCommand(string command)
{
    //string command = "subst s: \"g:\\Meine Ablage\""; // Replace with the CMD command you want to run
    //Console.WriteLine(command);

    Process process = new Process();
    ProcessStartInfo startInfo = new ProcessStartInfo();

    startInfo.FileName = "cmd.exe"; // Specify the CMD executable
    startInfo.Arguments = "/C " + command; // Specify the command to execute
    startInfo.RedirectStandardOutput = true;
    startInfo.UseShellExecute = false;
    startInfo.CreateNoWindow = true;

    process.StartInfo = startInfo;
    process.Start();

    // Read the output of the command
    string output = process.StandardOutput.ReadToEnd();

    // Wait for the command to finish executing
    process.WaitForExit();

    Console.WriteLine(output);

    //Console.WriteLine("Beliebige Taste drücken...");
    //Console.ReadKey();
}


DateTime dateFromWord(string dateString)
{
    DateTime dateTime = DateTime.Today;

    dateString = dateString.ToLower();
    if (dateString == "heute")
    {
        dateTime = DateTime.Today;
    }
    if (dateString == "morgen")
    {
        dateTime = DateTime.Today.AddDays(1);
    }

    int dayNumber = getWeekdayNumber(dateString);
    if (dayNumber != 0)
    {
        int daysAdd = dayNumber - weekdayNumber;
        if (daysAdd < 0)
        {
            daysAdd = daysAdd + 7;
        }
        dateTime = DateTime.Today.AddDays(daysAdd);
    }

    if (dateString.EndsWith("."))
    {
        string format = "d.M.yyyy";
        string fullDateString = dateString + DateTime.Today.Year.ToString();
        DateTime.TryParseExact(fullDateString, format, null, System.Globalization.DateTimeStyles.None, out dateTime);
    }
    return dateTime;
}

int getWeekdayNumber(string dateString)
{
    try
    {
        var day = daysOfWeek.First(d => d.Value.ToLower() == dateString.ToLower());
        return day.Key;
    }
    catch (Exception)
    {
        return 0;
    }
}

class MyClass
{
    public int Id { get; set; }
    public string Name { get; set; } = "";
    public int ParentId { get; set; }
    public int Position { get; set; }
    public List<int> DependenceIds { get; set; } = new List<int>();
    public bool Done { get; set; } = false;
    public DateTime TimeStamp { get; set; } = DateTime.Now;
}

class HelperClass
{
    [DllImport("user32.dll")]
    private static extern bool SetForegroundWindow(IntPtr hWnd);

    public static void SetForeground(string appName)
    {
        Process[] processes = Process.GetProcessesByName(appName);

        if (processes.Length > 0)
        {
            IntPtr mainWindowHandle = processes[0].MainWindowHandle;
            SetForegroundWindow(mainWindowHandle);
        }
        else
        {
            Console.WriteLine("Outlook is not running.");
        }
    }
}
