Microbot Name - EmailClassificationMicrobot

Microbot Description:-
This microbot uses Python scripts in backgeound to either create a dataset as a JSON File for Email Classification or classify email against the probability of matched keywords and phrases present in JSON file. It has following Input adn Output parameters.

Input: 

CSVFilePath - Path of CSV file that is to be created while creating dataset for email classification. Bot will split mails into sequence of body of replies that is present in the email chain. these content will be copied into CSV file to create dataset

JSONFilePath - THis works as a dataset for classificaion of email. Bot will read the csv file and create a list of keywords and phrases. It will create nodes in JSON file.

TrainorTest - Pass "Train"  to run robot as a Training dataset or pass "Test" for classifying email.

EMailFolder - Name of Folder under Inbox from which emails are to be processed.

PythonFilePath - Python script file path in which actual code is written for creating dataset or classification. Sampel files are provided in the microbot zip file.

PythonEXEPath - Exe path of installed Python version. Here i have used Python 3.6.4


Output:-

EmailType - Type of email after classification.

Message - Some message regarding execution of process

ErrDesc  - Error desc from catch

ErrCode- Error Code from catch
Below is the list of error codes


204 - No email found
205 - Error creating Dataset
206- Error classifying email
400 - Error in getting position of string in email
401 - Error in classifying email
402 -Error in Creating dataset
404 - File not found
420 - Python execution failure
444 - Error code for execute
503 - Outlook error

In case of successful execution
1. Message: "Inside Execute() : Execution Completed Successfully.";

In case of exception/failure:
1. ErrDesc: Some description of error alongwith exception message
2. ErrCode: some error code as given above
3. Message: "Inside Execute() Exception: Error While executing Request. " + ex.Message;


Dlls used:
EmailClassificationMicrobot 	(Main BOT solution DLL)
SE.Core.Automation.Interfaces
SE.Core.Automation.Models.Common
System
Microsoft.Office.Interop.Outlook
System.IO
System.Diagnostics
Anaconda Python 3.6.4

Product Description:
BOT starts by checking if it has to run as a training set or a test set. This will be decided by the TrainorTest field. Once decided it will start reading emails from the desired email folder. 

For creating dataset JSON file, it will start reading email chain and will extract thee body of each email from the complete email chain. It will copy these body text in the CSV file path provided as input. Once this file is created, it will invoke a run_python() method to execute the bigram.py file. This python script will create phrases and keywords for the Incident type of emails. Here i have considered two tpes of email - For creation of new incident and other for creating followu in Service Now. In Bigram.py file in order to create phrases list for specific type we need to change the respective node in the code. Currently it will be creating dataset for New Incident type of email. Once this script will run completely BOT will send a messahe as Dataset created successfully.

For classification of email, BOT will send each body parts from the chain of reply email to the run_Python() method. run_Python will execute testSet.py script which in turn call Classify() method from EmailClassification.py script in Python.Based on the largest probability of classification, Final type of email will be identified and BOT will update the EMailType field.

Target System:  windows system








