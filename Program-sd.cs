using Newtonsoft.Json;
using System;

MyClass currentItem;
MyClass cutItem = null;
List<MyClass> currentList;
string json;
json = File.ReadAllText("list.json");
List<MyClass> data = JsonConvert.DeserializeObject<List<MyClass>>(json);

currentItem = data[0];
currentList = data.Where(item => item.ParentId== 1).ToList();

showList();


json = JsonConvert.SerializeObject(data, Formatting.Indented);

File.WriteAllText("list.json", json);


void showList()
{
    Console.Clear();
    Console.WriteLine(currentItem.Name);
    Console.WriteLine();
    int i = 0;
    foreach (MyClass item in currentList)
    {
        i += 1;
        item.Position = i;
        Console.WriteLine(item.Position + " " + item.Name);
    }

    Console.WriteLine();
    if (currentItem.Id != 1)
    {
        Console.WriteLine("z Zurück");
        if (cutItem is null && currentItem.ParentId > 1)
        {
            Console.WriteLine("x Ausschneiden");
        }
        if (cutItem is not null)
        {
            Console.WriteLine("v Einfügen");
        }
    }
    Console.WriteLine("e Ende");

    string? input = Console.ReadLine();

    if (input != "e")
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
                    cutItem = null;
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
                        currentItem = FoundChildItem(currentList, input);

                        if (currentItem is null)
                        {
                            currentItem = new()
                            {
                                Id = data.Max(item => item.Id) + 1,
                                Name = input,
                                ParentId = currentItem.Id
                            };
                            data.Add(currentItem);
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
}

