# -*- coding: utf-8 -*-
"""
Created on Mon Feb 26 17:33:13 2018

@author: anuradha.agrawal
"""


import sys

temp = []
newIncedent = 0
followUp = 0
none=0
MailType="None"
MailClassification = "None"
from EmailClassification import Classify

JSONFilePath = sys.argv[1]
CSVFilePath = sys.argv[2]
EmailContent = sys.argv[3:]
#text = ["Observing high latency issue" , "please raise a docket and provide resolution","Please assist to open ticket"]

#print (EmailContent)
for line in EmailContent:
#    print (line)
    MailType = Classify(JSONFilePath,CSVFilePath,line)
    if (MailType=="NewIncident"):
        newIncedent+=1
        print("NI")
    elif (MailType=="FollowUp"):
        followUp+=1
        print("FW")
    else:
        none+=1
        print("N")
if(newIncedent>followUp and newIncedent> none):
    MailClassification="NewIncedent"
    print("NI1")
elif (followUp>newIncedent and followUp > none):
    MailClassification="FollowUp"
    print("FW1")
elif (followUp==newIncedent):
    MailClassification="None"
    print("N1")
else:
    MailClassification="None"
print (MailClassification)
    

    
#    print (sentiment(line),flush=True)
#print (sentiment("Observing high latency issue , please raise a docket and provide resolution"))    

#text1 = "Observing Packet Drop"
#text2 = "Packet"
#    
#print text2.lower() in text1.lower()
#if (text2.lower() in text1.lower()):
#    print "exist"
    