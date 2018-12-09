# -*- coding: utf-8 -*-
"""
Created on Wed Feb 21 13:53:14 2018

@author: Anuradha.Agrawal
"""
import nltk
import csv
import numpy as np
import json

import operator
from nltk.corpus import stopwords


 
  
def sentiment(JSONFilePath,CSVFilePath,text):
#    print(JSONFilePath)
#    print(CSVFilePath)
#    print(text)
    with open(JSONFilePath) as data_file:    
        data = json.load(data_file)
    
    Key_dict={}
    Phrase_dict={}
    incProb=0
    fwProb = 0
   
    NIPhrase=[]
    for inc in data['Incident']:
       NIPhrase.append( inc['Phrases'])
       Phrase_dict[inc['Phrases']]=0
       for k in inc['Keys']:
            Key_dict[k]=0
        
    
    FWPhrase=[]
    for fw in data['FollowUp']:
       FWPhrase.append( fw['Phrases'])
       Phrase_dict[fw['Phrases']]=0 
       for k in fw['Keys']:
           Key_dict[k]=0 
           
#    print (NIPhrase)       
#    print (FWPhrase)
    classification ="None"
  
    for item in NIPhrase:
#                print item
                item_word = nltk.word_tokenize(item)
                text_sent = nltk.sent_tokenize(text)
#                print text_sent
                for sentence in text_sent:
                    #                print sentence
                    new_count=0
                    sent_words = nltk.word_tokenize(sentence)
                    for word in sent_words:
                        #                print word
                        for w in item_word:
                            if(word.lower() == w.lower()):
                                #                print "inside  NI"
                                new_count+=1
                Phrase_dict[item]=new_count
            
    for item in FWPhrase:
                #                print item
                item_word = nltk.word_tokenize(item)
                text_sent = nltk.sent_tokenize(text)
                #                print text_sent
                for sentence in text_sent:
                    #                print sentence
                    followup_count=0
                    sent_words = nltk.word_tokenize(sentence)
                    for word in sent_words:
                        #                print word
                        for w in item_word:
                            if(word.lower() == w.lower()):
                                #                print "inside  NI"
                                followup_count+=1
                Phrase_dict[item]=followup_count
    
   
    #max count of phrase and keys
#    for key, value in Phrase_dict.iteritems():
        #                print key
        #                print value
#    max_key = max(Phrase_dict.iteritems(), key=operator.itemgetter(1))[0]
    
    for keyp, value in Phrase_dict.items():
        if keyp in  NIPhrase:
            #                print "NI"
            for ph in data['Incident']:            
                if(ph['Phrases']==keyp):
                    for key in ph['Keys']:
                        #                print key
                        new_count=0
                        for sentence in text_sent:
                            #                print sentence                        
                            if(key.lower() in sentence.lower()):
                                new_count+=1
                        Key_dict[key]=new_count              
                        incProb+=new_count
        if keyp in FWPhrase:
            #                print "FW"
            for ph in data['FollowUp']:
                if(ph['Phrases']==keyp):
                    for key in ph['Keys']:
                        #                print key
                        followup_count=0
                        for sentence in text_sent:
                            #                print sentence                        
                            if(key.lower() in sentence.lower()):
                                followup_count+=1
                        Key_dict[key]=followup_count
                        fwProb+=followup_count
       
    #                print Key_dict
    #                print new_count
    #                print followup_count
    
   
    
    
    if(incProb<fwProb):
#        print(incProb)
#        print(fwProb)
        classification = "FollowUp"
    elif(incProb>fwProb):
#        print(incProb)
#        print(fwProb)
        classification = "NewIncident"
    else:
#        print(incProb)
#        print(fwProb)
        classification = "None"
    return classification

#print (sentiment('D:\\test.json ','D:\Mail.csv',"Observing high latency issue , please raise a docket and provide resolution"))


#Freq Dist and Normalization of text
text = "Observing high latency issue , please raise a docket and provide resolution.Please resolve issue asap."
words = nltk.word_tokenize(text)
fd = nltk.FreqDist(words)
#print (fd["Observing"])

# NLTK's default English stopwords
default_stopwords = set(nltk.corpus.stopwords.words('english'))
#print (default_stopwords)

# We're adding some on our own - could be done inline like this...
# custom_stopwords = set((u'–', u'dass', u'mehr'))
# ... but let's read them from a file instead (one stopword per line, UTF-8)


all_stopwords = default_stopwords 

# Remove single-character tokens (mostly punctuation)
words = [word for word in words if len(word) > 1]



# Lowercase all words (default_stopwords are lowercase too)
words = [word.lower() for word in words]

# Stemming words seems to make matters worse, disabled
# stemmer = nltk.stem.snowball.SnowballStemmer('german')
# words = [stemmer.stem(word) for word in words]

# Remove stopwords
words = [word for word in words if word not in all_stopwords]
#print (words)

# Calculate frequency distribution
fdist = nltk.FreqDist(words)

 



            
                    