using Newtonsoft.Json;
using System;

MyClass currentItem;
MyClass cutItem;
List<MyClass> currentList;
string json;
List<MyClass> data;
string? input = "r";

while (input == "r")
{
    json = File.ReadAllText("list.json");
    data = JsonConvert.DeserializeObject<List<MyClass>>(json);

    currentItem = data[0];
    currentList = data.Where(item => item.ParentId== 1).ToList();
    cutItem = null;

    showList();
    
    saveData(); 
}


void saveData()
{
    json = JsonConvert.SerializeObject(data, Formatting.Indented);
    File.WriteAllText("list.json", json);
}


void showList()
{
    Console.Clear();
    Console.WriteLine(currentItem.Name+ " (" + currentItem.TimeStamp + ")");
    Console.WriteLine();
    int i = 0;
    foreach (MyClass item in currentList.OrderBy(i => i.TimeStamp).ToList())
    {
        i += 1;
        item.Position = i;
        Console.WriteLine(item.Position + " " + item.Name);
    }

    Console.WriteLine();
    if (currentItem.Id != 1)
    {
        Console.WriteLine("r Refresh");
        Console.WriteLine("z Zurück");
        if (cutItem is null)
        {
            Console.WriteLine("x Ausschneiden");
        }
    }
    if (cutItem is not null)
    {
        Console.WriteLine("v Einfügen");
    }
    Console.WriteLine("e Ende");
    Console.WriteLine();

    input = Console.ReadLine();

    if (input != "e" && input != "r")
    {
        if (input == "z" && currentItem.ParentId != 0)
        {
            currentItem = data.Find(item => item.Id == currentItem.ParentId);
        }
        else
        {
            if (input == "x")
            {
                cutItem = currentItem;
                currentItem = data.Find(item => item.Id == currentItem.ParentId);
                cutItem.ParentId = 0;
            }
            else
            {
                if (input == "v")
                {
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

                        if (foundItem is null)
                        {
                            currentItem = new()
                            {
                                Id = data.Max(item => item.Id) + 1,
                                Name = input,
                                ParentId = currentItem.Id,
                                TimeStamp = DateTime.Now
                            };
                            data.Add(currentItem);
                            saveData();
                        }
                        else
                        {
                            currentItem= foundItem;
                        }
                    }
                }
            }
        }
        currentList = data.Where(item => item.ParentId == currentItem.Id).ToList();
        showList();
    }

    MyClass FoundChildItem(List<MyClass> list, string search) {
        MyClass item = null;
        foreach (MyClass child in list) {
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

