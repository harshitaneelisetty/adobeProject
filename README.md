# adobeProject
Extracting information from PDF invoices using Adobe PDF Services Extract API

First I generated API keys.
Next i downloded -> pip install pdfservices-sdk  
adobeAPImain.py is the main python file
To run in Visual studio code : python adobeAPImain.py
In code all files one by one and stored information in ExtractData.csv file
I created seperate functions to extract each coloumn data 
I tried my best to complete the project.

problems I faced during project:

The problem is adobe extracted different api structure for some pdf, that's why I am getting wrong data in 5 to 6 pdf files, not all the data is wrong only some columns like email, username, because sometime adobe only read .co .co m .c . c om, different structure like this should be hotmail.com not hotma and il.com under that, it affected the whole structure of the pdf,  I recognised that the wrong data happen because of different api structure, I tried in several ways but that hard to automate, there is no pattern to automate. At last i some how managed by carefully observing how data is going and i done some indexing.
