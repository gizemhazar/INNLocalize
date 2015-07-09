# INNLocalize
Application aims create three platform string files (android,ios,windows phone) from an excel file.For this purpose you can use console or gui version.(for console version change localizeGui branch to master branch.)
#Overview
INNLocalize is a java application for convert an excel file to different formats(For android .properties extension, 
for iOS .string extension, for windows phone .resx extension). Giving excel file must consists of key and value
pair (One column must represents keys and the other one must values). Excel supported read formats:
  - Microsoft Excel 97(-2007) (BIFF8) .xls 
  - Microsoft Excel Template (.xlt)
  - Microsoft Excel XML (2007+) (OOXML) .xlsx file formats.
  
Application produces three files from every sheet in excel and imports giving sheet name as these different formats files name and locates them which is related platform's folder. 
#Running on Console
Working on CMD or Terminal, firstly need to produce runnable jar file. After adding source code and external jars to eclipse 
(while this project was creating , external jars used in eclipse so these jars were added to git for achieve jars very easy.) 
Then right click your project and click export ->Runnable jar file , expand the "Java" folder and double click the "Runnable JAR file" option. 
After created jar file type on command prompt
```sh
$ java -jar <jar file path>
```
eg ```$ java -jar /Users/username/Desktop/test.jar```

After runned jar file, drop excel file, which you want to convert, on Terminal (or CMD).
```sh
    drop excel file you want to read 
   /Users/username/desktop/Book1.xlsx 
```
then add destination path
```sh
   enter destination path
   /Users/username/desktop
```

Finally all these steps , files which they're in different formats (.properties,.string,.resx) will be located under relational 
platform folder with their zip files.
