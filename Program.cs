using Newtonsoft.Json;
using System;
using System.Diagnostics;
using System.Runtime.InteropServices;

MyClass currentItem;
MyClass newItem;
MyClass modifyItem;
MyClass cutItem;
MyClass dependenceItem;
List<MyClass> currentList;
string json;
List<MyClass> data;
string? input = "r";
string filePath = @"S:\SDWorkflow\";
string fileName = "data.json";
string fileNameToDay = "data" + DateTime.Today.ToString().Split(" ")[0].Replace(".","") + ".json";
string varName = "";
string varValue = "";


//HelperClass.SetForeground("Outlook");
executeCommand("subst s: \"g:\\Meine Ablage\"");


while (input == "r")
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

    currentItem = data[0];
    currentList = data.Where(item => item.ParentId == 1).ToList();
    newItem = null;
    modifyItem = null;
    cutItem = null;
    dependenceItem = null;

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
    Console.Clear();

    if (input == "")
    {
        setVariableOrExecuteCommand(currentItem.Name);

        // Text-Datei öffnen
        if (currentItem.Name.Contains(".txt"))
        {
            try
            {
                Process.Start("notepad.exe", currentItem.Name);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.ReadLine();
            }
        }
    }


    // Daten anzeigen
    {
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
        Console.WriteLine("e Einfügen");
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

void setVariableOrExecuteCommand(string itemName)
{
    // Varibale (Platzhalter lesen)
    if (itemName.StartsWith("["))
    {
        varName = itemName.Split("=")[0];
        varValue = itemName.Split("=")[1];
    }

    // Befehl ausführen
    if (itemName.StartsWith("Befehl:"))
    {
        string cmd = itemName.Split("Befehl:")[1];
        if (varName != "")
        {
            cmd = cmd.Replace(varName, varValue);
        }
        executeCommand(cmd);
    }
}

void executeCommand(string command)
{
    //string command = "subst s: \"g:\\Meine Ablage\""; // Replace with the CMD command you want to run
    Console.WriteLine(command);

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

class MyClass
{
    public int Id { get; set; }
    public string Name { get; set; } = "";
    public int ParentId { get; set; }
    public int Position { get; set; }
    public List<int> DependenceIds { get; set; } = new List<int>();
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
