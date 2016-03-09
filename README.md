# VBad
VBA Obuscation Tools combined with an MS office document generator.

DISCLAIMER: This is only for testing purposes and can only be used where strict consent has been given. Do not use this for illegal purposes, period. 

Please read the LICENSE under readme/LICENSE for the licensing of VBad.

![alt tag](https://raw.githubusercontent.com/Pepitoh/VBad/master/Example/example_ok.PNG)

# Features
VBad is a tool that allows you to obfuscate in many diffrent way pieces of VBA code and integrated directly into MS Office document. You would be able to : 
* Encrypt all String present in your VBA code;
* Encrypt data from your python Script in VBA code (domain name or path for example);
* Randomize each functions or variables' names that you want;
* Chose Encryption method and how and where encryption keys are stored;
* Generate as many unique MS Office u(with different randomize in the VBA) as you want using a filename list and a document Template;
* Enable autodestruction of encryption Keys feature once the VBA has been trigger once; 

#prerequisites
* Office (Excel/Word) for generated final doc (tested only on Office 2010)
* Python 2.7 
* win32com

#How to use 
1. First of all, you need to markdown your orignal VBA to indicate the script what you want to obfuscate/randomize or not :

* All VBA strings are encrypted by default. Moreover, you can exclude encryption of one string by adding an exclude mark ([!!]) at the end of the string. Example :
```vbs
String_Encrypted = "This string will be encrypted"
String_Not_Encrypted = "This string will NOT be encrypted[!!]"
````
* Mark [rdm::x] before a variable or function name will randomize it with a x chars string, Example :
```vbs
Function [rdm::10]Test()  '=> Test() will become randomized with a 10 characters string
[rdm::4]String_1 = "Test"  '=> String_1 wil lbecome randomized with a 4 characters string
``` 
* Mark [var::var_name] will included the string string_to_hide('var_name') from const.py in a encrypted way in the VBA. With that, you can generate string from your python file and include it directly in your VBA (DGA codding for example).
```vbs
Path_to_save_exe = [var::path] '=> string_to_hide("path") will be encrypted and put in the final VBA
``` 

2. Git clone and customize config.py to fit your need, you have to indicate at least : 
```python
template_file = r"C:\tmp\Vbad\Example\Template\template.doc" # The path of the template Office document you want to use to generate your files
filename_list = r"C:\tmp\Vbad\Example\Lists\filename_list.txt" #The path to the file that contain a list of different filename you want to use for your generated files
path_gen_files = r"C:\tmp\Vbad\Example\Results" # Path where your generated Office document will be saved
original_vba_file = r"C:\tmp\Vbad\Example\Orignal_VBA\original_vba.vbs" # The orignal VBA file you want to include, randomize and obfuscate in your malicious documents
trigger_function_name =  "Test" # Function that you want to auto_trigger (in your original_vba_file)
string_to_hide = {"domain_name":"http://www.test.com", "path_to_save":r"C:\tmp\toto"} #Strings that you want to add in your 
```

#Example 
In Example folder, you will find an already marked vba file, a template.doc, a list of 3 filename. You can use it and adapt it as you need.

#TODO : 
* Other encryption methods
* Other key hiding methods 
* .xls generation
* .docx and .xlsx generation

Feel free to contribute :-)

Pepitoh.
