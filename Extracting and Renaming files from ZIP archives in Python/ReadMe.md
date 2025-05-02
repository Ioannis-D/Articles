<h1 style="text-align: center;">Extracting and Renaming files from ZIP archives in Python</h1>

This article was originally published in [Medium](https://medium.com/@i.doganos/extracting-and-renaming-files-from-zip-archives-in-python-3c015ec21280).

<div style="text-align: center; id="dot">. . .</div>

If you just need the final code:

```
from zipfile import ZipFile
import os # This is imported to create the user path "~" in Linux

rename_dict = {
    "file1.csv" : "rock.csv",
    "file2.csv": "jazz.csv",
    "file3.csv": "pop.csv"
}

file = "./file.zip"
extract_directory = os.path.expanduser("~/Documents") # You can use the absolute path, for example "C:Users\User\Documents"

with ZipFile(file, 'r') as filezip: # Open the ZIP file in read mode
    for file in filezip.infolist(): # Iterate over the metadata (including filenames)
        filename = file.filename # # Store the original filename in a variable (as the .extract function needs the original filename)
        file.filename = rename_dict[filename] # Change the filename in memory (temporary change, does not alter the zip file itself but it extracts the file with this name)
        filezip.extract(filename, path=extract_directory, pwd=b'filezip101') # Extract each file with a different name to the specified directory. The pwd is the password of the zip file, must be given with as bytes (b'password')
```

<div style="text-align: center; id="dot">. . .</div>

### 1. Introduction and files creation

I love automation. I have noticed that most employees spend a considerable portion of their workday, repeating tasks that they could be automated.
Reducing that repetitive time is one of my main goals at the company I’m currently at.
I have automated a variety of tasks but today, I encountered a process that I hadn’t faced before.
When I had to export the files from a ZIP archive (with a Python script), I exported them with their original name.
But, this time, I needed to extract the archives with specific names.
This is when I realized that, although there is a [Stackoverflow answer](https://stackoverflow.com/a/56362289), I couldn’t find any explanatory article about the process. So, I decided to write one, showing how I did it with the [zipfile library](https://docs.python.org/3/library/zipfile.html).

<div style="text-align: center; id="dot">. . .</div>

For the purposes of this article I created a ZIP file (*file.zip* ) containing three .csv files (*file1.csv* , *file2.csv* , *file3.csv* ). I asked [Le Chat](https://chat.mistral.ai/chat/4c44b6a8-09b9-4ba6-bb8b-a085f044a2f6) to create their content and I zipped them with the password `zipfile101`.For the purposes of this article I created a ZIP file (file.zip) containing three .csv files (file1.csv, file2.csv, file3.csv). I asked Le Chat to create their content and I zipped them with the password zipfile101.

<div style="text-align: center; id="dot">. . .</div>

# 2. A brief introduction in the zipfile library


Before I start with the process per se, I want to cover the basics of the `zipfile` module, which is the library included in the CPython for handling ZIP
archives. I believe reading and understanding the documentation is essential, as it can help in writing more efficient code and in comprehending the module’s capabilities, limitations, and potential use cases. As Matthias Endler put it in his must-read article “[The Best Programmers I Know](https://endler.dev/2025/best-programmers/)”, excellent programmers “*Read the Reference* ”. The purpose of my article isn’t to deep dive into the `zipfile` library. I will only cover the parts that are necessary to:

* read a ZIP archive,
* retrieve its metadata
* extract files (even if the zip is password protected)**with a different name** .

First of all, let’s keep these limitations in mind:

▹*Decompression may fail due to (…) ***unsupported compression method / decryption****. I have faced this multiple times, as some of the files that I receive cannot be decrypted with *zipfile* . This can happen, for example, when the password is hashed. In such cases, different python libraries can be used.

And, most importantly:

▹*Not knowing the default extraction behaviors can cause unexpected decompression results (…).*

The last one scares a bit, but there is no reason to panic. When you understand what each option does, you’ll be an expert in extracting ZIP files. Nevertheless, I mention it because it’s wise to know and understand how the written code behaves, elsewhere unwanted and irreversible actions can be produced.

## 2.1 The *zipfile.ZipFile* class — Reading the ZIP file

From the documentation: “*The class for ****reading*** and writing ZIP files*”. In other words, this is the main class we’ll use to read the ZIP archive and these are its objects:

```
zipfile.ZipFile(
  file, 
  mode='r', 
  compression=ZIP_STORED, 
  allowZip64=True, 
  compresslevel=None, 
  *, 
  strict_timestamps=True, 
  metadata_encoding=None
)
```

I won’t enter into specific details. To read the ZIP, we need to specify the file’s `path` and the `mode='r'` (which stands for ‘reading’ and is the default). The `metadata_encoding` can also be used to read the file, but for the majority of the cases it can be left as *None*.

It supports the `with` statement, so to read it we can do:

```
from zipfile import ZipFile

file = "./file.zip"

with ZipFile(file, 'r') as filezip: 
    # The rest of the code
```

With this, we can inspect and extract the ZIP file’s metadata, which includes the name of the files it contains. Keep in mind that this does not require a password, since the **metadata of ZIP archives is not encrypted**.

Having opened and read the file, the next step is to work with its content.

## 2.2 Explaining the infolist(), getinfo() and namelist()

I will start by explaining the `namelist()` method. From the documentation: “*Return a list of archive members by name.”*

Running the following command we can see that it returns a list of filenames:

```
from zipfile import ZipFile

file = "./file.zip"

with ZipFile(file, 'r') as filezip:
    print(f"Content of namelist(): {filezip.namelist()}")
    print(f"Type of namelist(): {filezip.namelist()}")
```

```
Content of namelist(): ['file1.csv', 'file2.csv', 'file3.csv']
Type of namelist(): <class 'list'>
```

That’s basically it. This method does exactly that: it retrieves the filenames inside the ZIP archive. However, keep in mind that if the ZIP file contains directories, `namelist()` will return the directory names, not the names of the files inside them. For example, if I compress *file.zip* into a new ZIP, the `namelist()` will return:
 ```
file = "./file2.zip"

with ZipFile(file, 'r') as filezip:
    print(f"Content of namelist(): {filezip.namelist()}")
    print(f"Type of namelist(): {filezip.namelist()}")
```
```
Content of namelist(): ['file.zip']
type of namelist(): <class 'list'>
```
In my case that I want to change the filenames, this object isn’t helpful as it doesn’t allow any modifications.

On the other hand, we can retrieve and modify the *ZipInfo* attributes using the `getinfo(name)` and `infolist()` objects. Both return information (as *ZipInfo* objects) about the contents of the ZIP archive, but the main difference is that `infolist()` returns a *list* with the info of **all files** while `getinfo(name)` returns the information of **a specific file**. In practice:

```
with ZipFile(file, 'r') as filezip:
    print(f"Content of infolist(): {filezip.infolist()}")
    print(f"Type of infolist(): {type(filezip.infolist())}")
```
```
Content of infolist(): [<ZipInfo filename='file1.csv' compress_type=99 external_attr=0x20 file_size=186 compress_size=178>, <ZipInfo filename='file2.csv' compress_type=99 external_attr=0x20 file_size=197 compress_size=185>, <ZipInfo filename='file3.csv' compress_type=99 external_attr=0x20 file_size=137 compress_size=145>]
Type of infolist(): <class 'list'>
```
We have a list of *ZipInfo* objects which we can iterate through:

```
with ZipFile(file, 'r') as filezip:
    for file in filezip.infolist():
        print(file)
        print(type(file))
```
```
<ZipInfo filename='file1.csv' compress_type=99 external_attr=0x20 file_size=186 compress_size=178>
<class 'zipfile.ZipInfo'>
<ZipInfo filename='file2.csv' compress_type=99 external_attr=0x20 file_size=197 compress_size=185>
<class 'zipfile.ZipInfo'>
<ZipInfo filename='file3.csv' compress_type=99 external_attr=0x20 file_size=137 compress_size=145>
<class 'zipfile.ZipInfo'>
```
The `getinfo(name)` object returns the same kind of metadata for a specific file (or more accurately, a specific *member* of the ZIP archive):

```
with ZipFile(file, 'r') as filezip:
    print(f"Content of getinfo(name): {filezip.getinfo('file1.csv')}")
    print(f"Type of getinfo(name): {type(filezip.getinfo('file1.csv'))}")
```
```
Content of getinfo(name): <ZipInfo filename='file1.csv' compress_type=99 external_attr=0x20 file_size=186 compress_size=178>
Type of getinfo(name): <class 'zipfile.ZipInfo'>
```
The name can also **include a path within the archive**. For example, if *file1.csv* is stored inside the directory *directory1*, we can retrieve its information by:

```
with ZipFile(file, 'r') as filezip:
    print(filezip.getinfo('directory1/file1.csv')"
```
Whether we are using `infolist()` or `getinfo(name)`, we get a *ZipInfo* object:

```
<ZipInfo filename='file1.csv' compress_type=99 external_attr=0x20 file_size=186compress_size=178
```
To extract its filename we can access the `.filename` attribute:

```
with ZipFile(file, 'r') as filezip:
    for file in filezip.infolist():
        print(file)
        print(file.filename)
        print("\n")
```
```
<ZipInfo filename='file1.csv' compress_type=99 external_attr=0x20 file_size=186 compress_size=178>
file1.csv

<ZipInfo filename='file2.csv' compress_type=99 external_attr=0x20 file_size=197 compress_size=185>
file2.csv

<ZipInfo filename='file3.csv' compress_type=99 external_attr=0x20 file_size=137 compress_size=145>
file3.csv
```
Or, for a specific file:

```
with ZipFile(file, 'r') as filezip:
    print(filezip.getinfo('file1.csv').filename)
```
```
file1.csv
```
## 2.3 Changing the filenames (well, sort of


We can modify the **filename property** of the *ZipInfo* class, by directly assigning the new name. In practice:

```
with ZipFile(file, 'r') as filezip:
    for file in filezip.infolist():
        print(file.filename)
        file.filename = "changed_name.csv"
        print(file.filename)
        print("\n")
```
```
file1.csv
changed_name.csv

file2.csv
changed_name.csv

file3.csv
changed_name.csv
```
Here is the tricky part: What has changed **is not the real name of the file but the member of the class object**.

> In
> other words, this change only happens in memory. What you’re modifying
> is the filename attribute of the ZipInfo object, not the actual file
> inside the ZIP archive.

In my case, to rename each file with a specific filename I created a dictionary:

```
rename_dict = {
    "file1.csv" : "rock.csv",
    "file2.csv": "jazz.csv",
    "file3.csv": "pop.csv"
}

with ZipFile(file, 'r') as filezip:
    print(filezip.infolist().filename)
    for file in filezip.infolist():
        print(file.filename)
        filename = file.filename # This is saved for the extraction that I explain in the next section
        file.filename = rename_dict[filename]
        print(file.filename)
        print("\n")
```
```
file1.csv
rock.csv


file2.csv
jazz.csv


file3.csv
pop.csv
```
Again, remember: these are temporary changes to the *ZipInfo* objects and don’t rename the files inside the ZIP unless you create a new ZIP archive and write the files. What I care for is to **extract** them with a different name.

<div style="text-align: center; id="dot">. . .</div>

Let’s start with the `member`. A filename can be passed to the `extract` method to extract a specific file. In the previous example, the filename attribute for *file1.csv* was modified to *rock.csv* . So, should we pass the new name?

This is one more tricky aspect. **No**, you have to use the **original name** (*file1.csv)*. Why? Do you remember when I said that the change isn’t done directly in the file but in temporary in the memory? That’s why. When the extraction functions are called, it internally fetches the list of filenames **again**.

Nevertheless, **the extracted file will be saved with the new name passed in** `file.filename`. Confusing right? That is why I am writing this, as it took a while to figure this out.

Continuing with the `path` argument, it specifies the directory to extract the files. If no argument is given, then the files will be extracted in the **same directory where the python script is running**, not the directory of the ZIP file.

Finally, the `pwd` attribute is for the password. This is another tricky aspect to keep in min: the password **has to be provided as bytes** , not as a string. The good news is that the conversion is straightforward. In Python, putting a *b* before the string converts it into bytes.

Putting all together:

```
from zipfile import ZipFile
import os # This is imported to create the user path "~" in Linux

rename_dict = {
    "file1.csv" : "rock.csv",
    "file2.csv": "jazz.csv",
    "file3.csv": "pop.csv"
}

file = "./file.zip"
extract_directory = os.path.expanduser("~/Documents") # You can use the absolute path, for example "C:Users\User\Documents"

with ZipFile(file, 'r') as filezip: # Open the ZIP file in read mode
    for file in filezip.infolist(): # Iterate over the metadata (including filenames)
        filename = file.filename # Store the original filename in a variable (as the .extract function needs the original filename)
        file.filename = rename_dict[filename] # Change the filename in memory (temporary change, does not alter the zip file itself but it extracts the file with this name)
        filezip.extract(filename, path=extract_directory, pwd=b'filezip101') # Extract each file with a different name to the specified directory. The pwd is the password of the zip file, must be given with as bytes (b'password')
```

# 3. Conclusion

I hope you found this article informative. While there are other libraries available for unzipping ZIP files, the *zipfile* module is already included in Python, is simple to use, and often gets the job done. The only confusing aspect can be the temporary change of the filename, but now you know how to handle it!
