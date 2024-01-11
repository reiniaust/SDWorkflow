using Newtonsoft.Json;
using System;

MyClass currentItem;
MyClass newItem;
MyClass modifyItem;
MyClass cutItem;
List<MyClass> currentList;
string json;
List<MyClass> data;
string? input = "r";
string filePath = @"S:\SDWorkflow\data.json";

while (input == "r")
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
    newItem = null;
    modifyItem = null;
    cutItem = null;

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
    Console.Clear();
    Console.WriteLine(currentItem.Name + " (" + currentItem.TimeStamp + ")");
    Console.WriteLine();
    int i = 0;
    foreach (MyClass item in currentList.OrderBy(i => i.TimeStamp).ToList())
    {
        i += 1;
        item.Position = i;
        string plus = "";
        if (data.Where(child => child.ParentId == item.Id).ToList().Count > 0)
        {
            plus = " +";
        }
        Console.WriteLine(item.Position + " " + item.Name + plus);
    }

    Console.WriteLine();
    Console.WriteLine("h Hinzufügen");
    if (currentItem.Id != 1)
    {
        Console.WriteLine("ä Ändern");
        Console.WriteLine("l Löschen/Ausschneiden");
    }
    if (cutItem is not null)
    {
        Console.WriteLine("e Einfügen");
    }
    if (currentItem.Id != 1)
    {
        Console.WriteLine("z Zurück");
    }
    if (currentItem.Id != 1 || cutItem is not null)
    {
        Console.WriteLine("r Refresh");
    }
    Console.WriteLine("x Ende");
    Console.WriteLine();

    if (input == "h" || input == "ä")
    {
        // Ändern
        Console.WriteLine("Neuer Titel:");
    }

    input = Console.ReadLine();

    if (input != "x" && input != "r")
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
                                        MyClass foundItem = FoundChildItem(currentList, input);

                                        // Neuen Punkt anlegen/hinzufügen
                                        if (foundItem is not null)
                                            currentItem = foundItem;
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
    public DateTime TimeStamp { get; set; } = DateTime.Now;
}


