# -*- coding: utf-8 -*-
"""
Created on Wed Apr 11 16:22:05 2018

@author: Anuradha.Agrawal
"""

import nltk
import json
from nltk.corpus import stopwords
import sys
import os

JSONFilePath = sys.argv[1]
CSVFilePath = sys.argv[2]
EmailContent = sys.argv[3:]

print(JSONFilePath)
print(CSVFilePath)
print(EmailContent)


if os.path.exists(JSONFilePath) and os.path.exists(CSVFilePath):
    with open(JSONFilePath) as data_file:
        data = json.load(data_file)
    emmaWords = list(nltk.corpus.gutenberg.words(CSVFilePath))
    lowerwords = [word.lower() for word in emmaWords if word.isalpha() ]
#print(emmaWords)
    sw = stopwords.words('english')
    filtered_sentence = [w for w in lowerwords if not w in sw] # Remving stop words

    emmaBigrams = list(nltk.ngrams(filtered_sentence, 2)) # create bigrams (n-grams of 2) from our emmaTokens
    emmabigramsFreqs = nltk.FreqDist(emmaBigrams) # determine frequency of Bi-grams
#for words, count in emmabigramsFreqs.most_common(20): # for the 15 most common Bi-grams
#    print(count, " ".join(list(words))) # show the count and the create a string from the tuple
#print(emmaBigrams[:5]) # convert to a list and show the first 5 entries
    
    emmaTrigrams = list(nltk.ngrams(filtered_sentence, 3)) # create bigrams (n-grams of 3) from our emmaTokens
    emmaTrigramsFreqs = nltk.FreqDist(emmaTrigrams) # determine frequency of Tri-grams
#for words, count in emmaTrigramsFreqs.most_common(20): # for the 15 most common Tri-grams
#    print(count, " ".join(list(words))) # show the count and the create a string from the tuple
#print(emmaBigrams[:5]) # convert to a list and show the first 5 entries

    emma5grams = list(nltk.ngrams(filtered_sentence, 5)) # create five-grams
    emma5gramsFreqs = nltk.FreqDist(emma5grams) # determine frequency of five-grams
#for words, count in emma5gramsFreqs.most_common(20): # for the 15 most common five-grams
#    print(count, " ".join(list(words))) # show the count and the create a string from the tuple

    for fivegram in emma5grams:
        W=[]
        for bigram in emmaBigrams:      
#        print(fivegram)
            F="%s %s %s %s %s" % fivegram
            B="%s %s" % bigram
#        print(F)
#        print(B)
            if(B in F):
                W.append(B)
#            print(W)
        data['Incident'].append({  
                'Phrases': F,
                'Keys': W
                })

    with open(JSONFilePath, 'w') as outfile:  
        json.dump(data, outfile)
    outfile.close()
else:
    print("Files Not Found")         
    
            
            
            
        