using Newtonsoft.Json;
using System;

MyClass currentItem;
MyClass newItem;
MyClass modifyItem;
MyClass cutItem;
MyClass dependenceItem;
List<MyClass> currentList;
string json;
List<MyClass> data;
string? input = "s";
string filePath = @"S:\SDWorkflow\data.json";

newItem = null;
modifyItem = null;
cutItem = null;
dependenceItem = null;

while (input == "s")
{
    try
    {
        json = File.ReadAllText(filePath);
        data = JsonConvert.DeserializeObject<List<MyClass>>(json);
    }
    catch (System.Exception)
    {
        data = new();
        data.Add(new MyClass() { Id = 1, Name = "SDWorkflow" });
    }

    currentItem = data[0];
    currentList = data.Where(item => item.ParentId == 1).ToList();

    showList();

    saveData();
}

void saveData()
{
    json = JsonConvert.SerializeObject(data, Formatting.Indented);
    File.WriteAllText(filePath, json);
}

void showList()
{
    try
    {
        Console.Clear();
    }
    catch (System.Exception) {}

    Console.WriteLine(currentItem.Name + " (" + currentItem.TimeStamp + ")");
    Console.WriteLine();

    if (currentItem.Name.StartsWith("Termine"))
    {
        currentList = currentList.OrderBy(i => dateFromWord(i.Name)).ToList();
    }
    else
    {
        currentList = currentList.OrderBy(i => i.TimeStamp).ToList();
    }

    DateTime dateFromWord(string dateString) {
        DateTime dateTime = DateTime.Today;
        if (dateString.ToLower() == "heute")
        {
            dateTime = DateTime.Today;
        }
        if (dateString.ToLower() == "morgen")
        {
            dateTime = DateTime.Today.AddDays(1);
        }
        if (dateString.EndsWith("."))
        {
            string format = "d.M.";   
            string fullDateString = dateString + DateTime.Today.Year.ToString();
            DateTime.TryParseExact(fullDateString, format, null, System.Globalization.DateTimeStyles.None, out dateTime);
        }
        return dateTime;
    }

    // Punkte anzeigen
    int i = 0;
    foreach (MyClass item in currentList)
    {
        i += 1;
        item.Position = i;

        // Pluszeigen setzen, wenn Unterpunkte, Abhängikeiten oder Verwendungen im Unterpunkt sind
        string plus = "";
        if (data.Where(child => child.ParentId == item.Id).ToList().Count > 0 || item.DependenceIds.Count > 0)
        {
            plus = " +";
        }
        if (plus == "")
        {
            foreach (var depItem in data)
            {
                if (depItem.DependenceIds.Contains(item.Id))
                {
                    plus = " +";
                    break;
                }
            }
        }

        Console.WriteLine(item.Position + " " + item.Name + plus);
    }

    // Abhängigkeiten anzeigen
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
    foreach (var item in data)
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

    Console.WriteLine();
    Console.WriteLine("h Hinzufügen");
    if (currentItem.Id != 1)
    {
        Console.WriteLine("ä Ändern");
        Console.WriteLine("l Löschen/Ausschneiden");
        if (dependenceItem is null)
        {
            Console.WriteLine("a Abhängigkeit hinzufügen");
        }
        else 
        {
            Console.WriteLine("a Abhängigkeit setzen");
        }
    }
    if (cutItem is not null)
    {
        Console.WriteLine("e Einfügen des ausgeschnittenen Punktes");
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
    else {
        Console.WriteLine("Suchen, Nummer oder Befehlsbuchstaben eingeben:");
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
                newItem = new()
                {
                    Id = data.Max(item => item.Id) + 1,
                    Name = input,
                    ParentId = currentItem.Id,
                    TimeStamp = DateTime.Now
                };
                data.Add(newItem);
                saveData();
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
                        if (input == "a")
                        {
                            // Abhängigkeit
                            if (dependenceItem is null)
                            {
                                dependenceItem = currentItem;
                            }
                            else 
                            {
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
                                if (input == "l")
                                {
                                    // Löschen/Ausschneiden
                                    cutItem = currentItem;
                                    currentItem = data.Find(item => item.Id == currentItem.ParentId);
                                    cutItem.ParentId = 0;
                                }
                                else
                                {
                                    if (input == "e")
                                    {
                                        // Einfügen
                                        cutItem.ParentId = currentItem.Id;
                                        cutItem.TimeStamp = DateTime.Now;
                                        cutItem = null;
                                        saveData();
                                    }
                                    else
                                    {
                                        bool found = false;
                                        foreach (MyClass item in currentList)
                                        {
                                            if (input == item.Position.ToString() || item.Name.ToLower().Contains(input.ToLower()))
                                            {
                                                currentItem = item;
                                                found = true;
                                                break;
                                            }
                                        }
                                        if (!found)
                                        {
                                            // Abhängigkeit suchen/aufrufen
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

                                            if (foundItem != null) {
                                                currentItem = foundItem;
                                            }
                                            else {
                                                Console.WriteLine("Nicht gefunden, weiter mit beliebiger Taste.");
                                                string dummy = Console.ReadLine();
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

class MyClass
{
    public int Id { get; set; }
    public string Name { get; set; } = "";
    public int ParentId { get; set; }
    public int Position { get; set; }
    public List<int> DependenceIds { get; set; } = new List<int>();    
    public DateTime TimeStamp { get; set; } = DateTime.Now;
}
