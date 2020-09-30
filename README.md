# WordReport
Generate report from Word document

### Get it
```
PM> Install-Package Mapster
```

### Step 1: Provide word template file
NOTE: template is [Scriban](https://github.com/lunet-io/scriban), please see usage there.  
NOTE2: You need to open selection pane to get or set image name to replace

![image](https://user-images.githubusercontent.com/5763993/94629582-6fbb9080-02ed-11eb-8d04-d17fadcf6e64.png)

### Step 2: Provide data models & images

```csharp
var data = new
{
    teacher = "Ben",
    author = "John Doe",
    students = new[]
    {
        new {name = "Foo", age = 15},
        new {name = "Bar", age = 16},
    }
};
var images = new Dictionary<string, byte[]>
{
    ["signature_pic"] = File.ReadAllBytes("signature.png")
};
```

### Step 3: Load template (NOTE: you can reuse template to generate multiple output)

```csharp
var reporter = WordTemplate.FromFile("Template.docx");
```

### Step 4: Generate output

```csharp
var mem = new MemoryStream();
reporter.Render(mem, data, images);
File.WriteAllBytes("Output.docx", mem.ToArray());
```
![image](https://user-images.githubusercontent.com/5763993/94629843-1ef86780-02ee-11eb-9654-93bcdf1595bc.png)
