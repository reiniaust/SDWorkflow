using Newtonsoft.Json;
using System;
using System.Globalization;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Collections.Generic;
using MimeKit;
using MsgReader;
using System.Security.Cryptography;
using static System.Runtime.InteropServices.JavaScript.JSType;
using System.Xml.Linq;
using Microsoft.Office.Interop.Outlook;
using System.Data.OleDb;
using System.Data.Odbc;


bool showHelp = false;
MyClass currentItem;
string currentUserName = Environment.GetEnvironmentVariable("USERNAME");
MyClass newItem = null;
MyClass modifyItem = null;
MyClass cutItem = null;
MyClass dependenceItem = null;
List<MyClass> currentList;
string json;
List<MyClass> data = new();
List<MyClass> tempData = new();
string? input = "";
//string filePath[0] = @"S:\SDWorkflow\";
List<string> filePathArray;
//string fileName = "data.json";
string fileNameToDay = DateTime.Today.ToString().Split(" ")[0].Replace(".", "");
string varName = "";
string varValue = "";
Dictionary<string, string> paramList = new Dictionary<string, string>();
int searchCounter;
string lastSearch;

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
    daysOfWeek.Add(7, "diese Woche");
    daysOfWeek.Add(14, "nächste Woche");
}
string weekdayNameToday = daysOfWeek[weekdayNumber];

//HelperClass.SetForeground("Outlook");
//executeCommand("subst s: \"g:\\Meine Ablage\"");

//ReadDataFromAcceessDb();

while (input == "s" || input == "")
{

    searchCounter = 1;
    lastSearch = "";

    // Daten lesen        
    readData();

    // Termin-Worte ändern
    {
        currentItem = data.Find(item => item.Name == weekdayNameToday);
        if (currentItem != null)
        {
            currentItem.Name = "heute";
            saveDataItem();
        }
        // morgigen Wochentag auf "morgen" ändern
        int nextDayNumber = weekdayNumber + 1;
        if (nextDayNumber == 7) nextDayNumber = 0;
        currentItem = data.Find(item => item.Name == daysOfWeek[nextDayNumber]);
        if (currentItem != null)
        {
            currentItem.Name = "morgen";
            saveDataItem();
        }
        // dann "heute" auf "gestern"
        currentItem = data.Find(item => item.Name == "heute" && item.TimeStamp.AddDays(1).Date == DateTime.Today);
        if (currentItem != null)
        {
            currentItem.Name = "Rückstand";
            saveDataItem();
        }
        // dann "morgen" auf "heute"
        currentItem = data.Find(item => item.Name == "morgen" && item.TimeStamp.AddDays(1).Date == DateTime.Today);
        if (currentItem != null)
        {
            currentItem.Name = "heute";
            saveDataItem();
        }

        currentItem = data.Find(item => item.Name == "diese Woche" && GetIso8601WeekOfYear(item.TimeStamp) < GetIso8601WeekOfYear(DateTime.Today));
        if (currentItem != null)
        {
            currentItem.Name = "Rückstand";
            saveDataItem();
        }
        currentItem = data.Find(item => item.Name == "nächste Woche" && GetIso8601WeekOfYear(item.TimeStamp) < GetIso8601WeekOfYear(DateTime.Today));
        if (currentItem != null)
        {
            currentItem.Name = "diese Woche";
            saveDataItem();
        }
    }

    if (input == "s")
    {
        // wenn man Startseite gewählt hat
        currentItem = data.Find(i => i.Id == 1); 
    }
    if (input == "")
    {
        currentItem = searchItem("");
    }

    // E-Mails aus dem Ordner lesen, Punkte anlegen und verschieben
    foreach (var filePath in filePathArray)
    {
        string path = Path.GetDirectoryName(filePath) + "\\";

        if (!Directory.Exists(path))
        {
            Directory.CreateDirectory(path);
        }

        //string downloadFolderPath = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + "\\Downloads\\";
        string[] files = Directory.GetFiles(path, "*.msg");
        foreach (string file in files)
        {
            currentItem = data.Find(item => item.Name == "Vorgänge" && item.File == filePath);
            var message = new MsgReader.Outlook.Storage.Message(file);
            newItem = new() { Name = message.Subject };
            // Kunde (Absender)
            MyClass customerItem = data.Find(item => item.Name.ToLower().Contains(message.Headers.From.Address.ToLower()));
            string salutation = "";
            if (customerItem != null)
            {
                newItem.DependenceIds.Add(customerItem.Id);
                customerItem = data.Find(item => item.ParentId == customerItem.Id && item.Name.StartsWith("Anrede:"));
                if (customerItem != null)
                {
                    salutation = customerItem.Name.Split("Anrede: ")[1] + ", ";
                }
            }
            saveDataItem();
            MyClass subjectItem = newItem;


            message.Dispose();
            string newFolder = path + newItem.Id;
            Directory.CreateDirectory(newFolder);
            //MyClass emailItem = newItem;

            string newPath = newFolder + "\\" + Path.GetFileName(file);
            Directory.Move(file, newPath);
            newPath = "\"" + newPath + "\"";
            newItem = new() { Name = newPath };
            currentItem = subjectItem;
            saveDataItem();
            executeCommand(newPath);

            Console.WriteLine(message.Subject);

            Console.WriteLine("Termin:");
            string dateString = Console.ReadLine();
            MyClass dateItem = data.Find(item => item.Name == dateString);
            if (dateItem is null)
            {
                currentItem = data.Find(item => item.Name == "Termine");
                newItem = new() { Name = dateString };
                saveDataItem();
                dateItem = newItem;
                newItem = null;
            }
            subjectItem.DependenceIds.Add(dateItem.Id);

            currentItem = data.Find(item => item.Name == "Kollegen");
            if (currentItem != null)
            {
                Console.WriteLine("Zuständiger Kollege?");
                string responsible = Console.ReadLine();
                if (responsible != "")
                {
                    MyClass colleaguesItem = data.Find(item => item.Name.ToLower().Contains(responsible.ToLower()));
                    if (colleaguesItem is null)
                    {
                        newItem = new() { Name = responsible };
                        saveDataItem();
                        colleaguesItem = newItem;
                    }
                    subjectItem.DependenceIds.Add(colleaguesItem.Id);
                    Console.WriteLine("Weiterleitung an " + responsible + ": bitte kümmere dich bis " + dateString + " darum und gib mir dann ein Feedback.");
                }
            }

            modifyItem = subjectItem;
            saveDataItem();
            modifyItem = null;

            Console.WriteLine("Antwort an Kunden: " + salutation + " wir kümmern uns darum und geben bis " + dateString + " ein Feedback.");
            Console.ReadKey();

            currentItem = subjectItem;
            newItem = null;

        }
    }


    currentList = data.Where(item => item.ParentId == currentItem.Id && item.File == filePathArray[0]).ToList();

    showList();

    //saveData();
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


    // Hilfe und Befehle anzeigen
    if (showHelp)
    {
        Console.WriteLine("? Hilfe ausblenden");
        Console.WriteLine("+ Hinzufügen");
        if (currentItem.Id != 1)
        {
            Console.WriteLine("* Ändern");
            if (cutItem is null)
            {
                Console.WriteLine("< Ausschneiden");
                Console.WriteLine("- Löschen");
            }
            else
            {
                Console.WriteLine("> Ausgeschnittenen Punkt einfügen");
            }

            if (currentItem.Done)
            {
                Console.WriteLine(". Auf Nicht Erledigt setzen");
            }
            else
            {
                Console.WriteLine(". Auf Erledigt setzen");
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
    }
    else
    {
        Console.WriteLine("? Hilfe einblenden");
    }
    Console.WriteLine();
    Console.WriteLine();


    // Daten anzeigen
    {
        foreach (var item in data)
        {
            item.Position = 0;
        }

        // Pfad/Oberpunkte anzeigen
        Console.WriteLine(getItemPath(currentItem));
        Console.WriteLine();


        string done = "";
        if (currentItem != null && currentItem.Done)
        {
            done = "Erledigt ";
        }
        string userString = "";
        if (currentItem.UserName != "" && currentItem.UserName != currentUserName)
        {
            userString = currentItem.UserName + " ";
        }
        Console.WriteLine(currentItem.Name + " (" + userString + done + currentItem.TimeStamp + ")");
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
            if (data.Where(child => matchList(child, item)).ToList().Count > 0 || item.DependenceIds.Count > 0)
            {
                plus = " +";
            }
            if (plus == "")
            {
                // Verwendungen prüfen
                foreach (var depItem in data.Where(item => item.File == currentItem.File && !item.Done))
                {
                    if (depItem.DependenceIds.Contains(item.Id))
                    {
                        if (isInfoForUser(depItem, item))
                        {
                            plus = " +";
                            break;
                        }
                    }
                }
            }

            // Wenn eine Anhängigkeit hinteregt ist, dann direkt mit anzeigen
            string depString = "";
            if (item.DependenceIds.Count == 1)
            {
                MyClass depItem = data.Find(i => i.File == currentItem.File && i.Id == item.DependenceIds[0]);
                if (depItem != null)
                {
                    depString = " -> " + depItem.Name;
                }
            }

            setParameter(item.Name);

            done = "";
            if (item.Done)
            {
                done = " (Erledigt)";
            }
            Console.WriteLine(item.Position + "   " + replaceParameterValueInText(item.Name) + depString + done + plus);
        }

        // Anhängigkeiten anzeigen
        if (currentItem.DependenceIds.Count > 0)
        {
            Console.WriteLine();
            Console.WriteLine("Anhängigkeiten:");
            foreach (int id in currentItem.DependenceIds)
            {
                MyClass depItem = data.Find(item => item.File == currentItem.File && item.Id == id);
                if (depItem != null)
                {
                    i += 1;
                    depItem.Position = i;
                    Console.WriteLine(depItem.Position + "   " + depItem.Name);
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
                Console.WriteLine(item.Position + "   " + item.Name);
            }
        }
    }

    Console.WriteLine();


    if (input == "+" || input == "*")
    {
        // Ändern
        Console.WriteLine("Neuer Titel:");
    }

    input = Console.ReadLine();

    if (input == "")
    {
        input = lastSearch;
        searchCounter += 1;
        currentList = data.Where(item => item.ParentId == 1 && item.File == filePathArray[0]).ToList();
    }
    else
    {
        searchCounter = 1;
    }

    if (input != "x" && input != "s")
    {
        if (input == "?")
        {
            showHelp = !showHelp;
        }
        else
        {
            if (input == "z" && currentItem.ParentId != 0)
            {
                currentItem = data.Find(item => item.Id == currentItem.ParentId && item.File == currentItem.File);
            }
            else
            {
                if (newItem is not null)
                {
                    // Text hinzufügen und speichern
                    newItem.Name = input;
                    data.Add(newItem);
                    saveDataItem();
                    newItem = null;
                }
                else
                {

                    if (modifyItem is not null)
                    {
                        // Änderung speichern
                        modifyItem.Name = input;
                        saveDataItem();
                        modifyItem = null;
                    }
                    else
                    {
                        if (input == "+")
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
                                    saveDataItem();
                                    dependenceItem = null;
                                }
                            }
                            else
                            {
                                if (input == "*")
                                {
                                    // Ändern
                                    modifyItem = currentItem;
                                }
                                else
                                {
                                    if (input == "<" || input == ">")
                                    {
                                        // Ausschneiden
                                        if (cutItem is null)
                                        {
                                            cutItem = currentItem;
                                            currentItem = data.Find(item => item.Id == currentItem.ParentId && item.File == currentItem.File);
                                            cutItem.ParentId = 0;
                                        }
                                        else
                                        {
                                            // Ausgeschnittenen Punkt einfügen
                                            saveDataItem();
                                            cutItem = null;
                                        }
                                    }
                                    else
                                    {
                                        if (input == "-")
                                        {
                                            // Löschen
                                            cutItem = currentItem;
                                            data.Remove(currentItem);
                                            saveDataItem();
                                            cutItem = null;
                                            currentItem = data.Find(item => item.Id == currentItem.ParentId && item.File == currentItem.File);
                                        }
                                        else
                                        {
                                            if (input == ".")
                                            {
                                                // Erledigt/Unerledigt
                                                currentItem.Done = !currentItem.Done;
                                                saveDataItem();
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
                                                        saveDataItem();
                                                        // eine Ebene zurückgehen
                                                        currentItem = data.Find(item => item.Id == currentItem.ParentId && item.File == currentItem.File);
                                                    }
                                                    else
                                                    {
                                                        int number;
                                                        bool isNumber = int.TryParse(input, out number);
                                                        bool found = false;
                                                        found = false;
                                                        foreach (MyClass item in currentList)
                                                        {
                                                            if (isNumber && number == item.Position || input != "" && !isNumber && item.Name.ToLower().Contains(input.ToLower()))
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
                                                                if (depItem != null)
                                                                {
                                                                    if (number == -depItem.Position)
                                                                    {
                                                                        // Abhängigkeit/Verknüpfung löschen, wenn die Nummer mit minus eingegeben wurde
                                                                        currentItem.DependenceIds.Remove(id);
                                                                        break;
                                                                    }
                                                                    if (number == depItem.Position || depItem.Name.ToLower().Contains(input.ToLower()))
                                                                    {
                                                                        currentItem = depItem;
                                                                        found = true;
                                                                        break;
                                                                    }
                                                                }
                                                            }
                                                        }
                                                        if (!found)
                                                        {
                                                            foreach (var item in data)
                                                            {
                                                                if (item.DependenceIds.Contains(currentItem.Id))
                                                                {
                                                                    if (input == item.Position.ToString())
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
                                                            // in allen Daten suchen
                                                            MyClass foundItem = searchItem(input);

                                                            if (foundItem is not null)
                                                                currentItem = foundItem;
                                                        }
                                                        lastSearch = input;
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
        }
        if (currentItem != null)
        {
            currentList = data.Where(item => matchList(item, currentItem)).ToList();
        }
        showList();
    }

}

// die passenden Unterpunkte zuordnen (je nach Joson-Datei)
bool matchList(MyClass item, MyClass parentItem)
{
    return isInfoForUser(item, parentItem) && (item.ParentId == parentItem.Id && item.File == parentItem.File || item.File == parentItem.Name && item.ParentId == 1);
}

// Prüfen, dass bei Terminen nur eigene Punkte reinkommen
bool isInfoForUser(MyClass item, MyClass parentItem)
{
    bool infoForUser = true;
    MyClass parentParent = getParentItem(parentItem);
    if (parentParent != null && parentParent.Name == "Termine")
    {
        infoForUser = (item.UserName == "" || item.UserName == currentUserName);
    }
    return infoForUser;
}

MyClass getParentItem(MyClass item)
{
    return data.Find(i => i.Id == item.ParentId && i.File == item.File || i.Name == item.File && item.ParentId == 1);
}

// Suche im ganzen Baum 
MyClass searchItem(string search)
{
    MyClass item = null;
    int counter = 1;
    foreach (MyClass child in data)
    {
        if (search == "" || foundAllWordsInItem(child, search)) // nach mehr als einem Wort suchen
        {
            if (counter == searchCounter)
            {
                item = child;
                break;
            }
            counter += 1;
        }
    }
    return item;
}

string itemPath(MyClass item)
{
    string path = "";
    while (item != null && item.ParentId != 0)
    {
        item = data.FirstOrDefault(i => i.Id == item.ParentId && i.File == item.File);
        if (item != null)
        {
            path += item.Name;
        }
    }
    return path;
}

bool foundAllWordsInItem(MyClass item, string search)
{
    bool found = false;
    if (isSynonymInText(item.Name, search))
    {
        found = true;
    }
    else
    {
        // auch in Anhängigkeiten suchen
        string depString = "";
        foreach (MyClass child in dependenceList(item))
        {
            depString += child.Name;
        }

        foreach (string word in search.Split(" "))
        {
            if (isSynonymInText(item.Name, word))
            {
                found = true;
                string path = itemPath(item);
                foreach (string nextWord in search.Split(" "))
                {
                    if (nextWord != word && !isSynonymInText(item.Name + path + depString, nextWord))
                    {
                        found = false;
                    }
                }
            }
        }
    }
    return found;
}

bool isSynonymInText(string text, string word)
{
    bool found = false;
    if (text.ToLower().Contains(word.ToLower()))
    {
        found = true;
    }
    else
    {
        MyClass synItem = data.Find(item => item.Name == "Synonyme");
        if (synItem != null)
        {
            synItem = data.Find(item => item.Name.Contains(word) && item.ParentId == synItem.Id);
            if (synItem != null)
            {
                foreach (var synWord in synItem.Name.Split(","))
                {
                    if (text.ToLower().Contains(synWord.ToLower()))
                    {
                        found = true;
                    }
                }
            }
        }
    }
    return found;
}

string getItemPath(MyClass item)
{
    string path = "";
    string help = "";
    MyClass parentItem = getParentItem(item);
    while (parentItem != null && parentItem.ParentId > 1)
    {
        path += help + parentItem.Name;
        parentItem = getParentItem(parentItem);
        help = " < ";
    }
    return path;
}

List<MyClass> dependenceList(MyClass item)
{
    List<MyClass> list = new();
    foreach (int id in item.DependenceIds)
    {
        MyClass depItem = data.Find(item => item.Id == id);
        if (depItem != null)
        {
            list.Add(depItem);
        }
    }
    return list;
}

// Daten lesen
void readData()
{
    filePathArray = new();
    filePathArray.Add(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + @"\SDWorkflow\data.json"); 
    
    try
    {
        if (data.Count == 0)
        {
            executeCommand(Path.GetDirectoryName(filePathArray[0]));
        }
        json = File.ReadAllText(filePathArray[0]);
        data = JsonConvert.DeserializeObject<List<MyClass>>(json);
        foreach (var item in data)
        {
            item.File = filePathArray[0];
        }
    }
    catch (System.Exception)
    {
        data = new();
        data.Add(new MyClass() { Id = 1, Name = "SDWorkflow", File = filePathArray[0] });
    }
    tempData = new();
    foreach (var file in data.Where(i => i.ParentId == 1 && i.Name.EndsWith(".json")))
    {
        filePathArray.Add(file.Name);

        json = File.ReadAllText(file.Name);
        foreach (var item in JsonConvert.DeserializeObject<List<MyClass>>(json))
        {
            item.File = file.Name;
            tempData.Add(item);
        }
    }
    foreach (var item in tempData)
    {
        data.Add(item);
    }

    // Absteigend nach Datum sortieren
    data = data.OrderByDescending(item => item.TimeStamp).ToList(); 
}

void saveDataItem()
{
    MyClass item = null;

    if (File.Exists(currentItem.File))
    {
        json = File.ReadAllText(currentItem.File);
        tempData = JsonConvert.DeserializeObject<List<MyClass>>(json);
    }
    else
    {
        tempData = data.ToList();
    }

    if (newItem != null)
    {
        item = newItem;
        item.Id = tempData.Max(item => item.Id) + 1;
        item.File = currentItem.File;
        item.ParentId = currentItem.Id;
        item.UserName = currentUserName;
        tempData.Add(item);
    }
    if (modifyItem != null)
    {
        item = tempData.Find(i => i.Id == modifyItem.Id);
        item.Name = modifyItem.Name;
        item.DependenceIds = modifyItem.DependenceIds;
    }
    if (cutItem != null)
    {
        item = tempData.Find(i => i.Id == cutItem.Id);
        item.ParentId = currentItem.Id;

        if (data.Find(i => i.Id == currentItem.Id) == null)
        {
            // Löschen
            item = tempData.Find(i => i.Id == currentItem.Id);
            tempData.Remove(item);
        }
    }
    if (item == null)
    {
        item = tempData.Find(i => i.Id == currentItem.Id);
        item.Name = currentItem.Name;
        // Erledigt setzen
        item.Done = currentItem.Done;
        // Abhängigkeiten speichern
        item.DependenceIds = currentItem.DependenceIds;
    }
    item.TimeStamp = DateTime.Now;

    json = JsonConvert.SerializeObject(tempData.ToList(), Formatting.Indented);
    File.WriteAllText(currentItem.File, json);
    File.WriteAllText(currentItem.File + fileNameToDay, json);

    readData();
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
    if (itemName.Contains(@":\") && !itemName.Contains(@".json"))
    {
        // Datei öffnen
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
    catch (System.Exception)
    {
        return 0;
    }
}
static int GetIso8601WeekOfYear(DateTime date)
{
    // Get the Calendar instance associated with the specified culture.
    Calendar calendar = CultureInfo.InvariantCulture.Calendar;

    // Determine the week of the year using ISO 8601 rules.
    return calendar.GetWeekOfYear(date, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday);
}

static void ReadDataFromAcceessDb()
{
    string connectionString = @"Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=r:\sd\mdb\kontroll\kontdat.mdb;";

    using (OdbcConnection connection = new OdbcConnection(connectionString))
    {
        connection.Open();

        string query = "SELECT * FROM MitarbeiterKtl where len(Text) > 0;";
        OdbcCommand command = new OdbcCommand(query, connection);

        using (OdbcDataReader reader = command.ExecuteReader())
        {
            while (reader.Read())
            {
                int id = (int)reader["SatzId"];
                string text = (string)reader["Text"];
                DateTime date = (DateTime)reader["Datum"];

                Console.WriteLine($"ID: {id}, Name: {text}, Date: {date}");
            }
        }
    }

    Console.ReadLine();
}


class MyClass
{
    public int Id { get; set; }
    public string File { get; set; } = "";
    public string Name { get; set; } = "";
    public int ParentId { get; set; }
    public int Position { get; set; }
    public List<int> DependenceIds { get; set; } = new List<int>();
    public bool Done { get; set; } = false;
    public string UserName { get; set; } = "";
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
