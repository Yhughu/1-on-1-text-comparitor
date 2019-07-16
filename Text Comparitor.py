import docx2txt
import glob
import os
import keyboard
import rouge
import re
import collections
import string
import nltk
import itertools
from nltk.tokenize import sent_tokenize, word_tokenize

liststringsN = []
liststringsP = []
listwcN = []
listwcP = []
aantalfiles = 0
folder_path = 'C:/Users/*' #Here you can enter the path in which your files are placed

def inlezenn ():  #Leest alle N(ieuws) files in
    global aantalfiles
    for filename in glob.glob(os.path.join(folder_path, '*N.docx')):
        liststringsN.append(docx2txt.process(filename))
        print(filename)
        aantalfiles = aantalfiles + 1

def inlezenp (): #Leest alle P(ers) files in
    for filename in glob.glob(os.path.join(folder_path, '*P.docx')):
        liststringsP.append(docx2txt.process(filename))
        print(filename)

def printfiles ():
    print("Tekst N ", end="", flush=True)
    print(liststringsN)
    print("Tekst P ", end="", flush=True)
    print(liststringsP)

def filecount(tempn, tempp): #Count het aantal woorden voor elke in gelezen N file.
    wordsn = word_tokenize(tempn)
    wordsp = word_tokenize(tempp)
    wordsn = [a.lower() for a in wordsn]
    wordsp = [a.lower() for a in wordsp]
    wordsn = [x for x in wordsn if not re.fullmatch('[' + string.punctuation + ']+', x)]
    wordsp = [x for x in wordsp if not re.fullmatch('[' + string.punctuation + ']+', x)]
    del wordsn[400:3000]
    #print("Wordcount overlap: ", end="", flush=True)
    print(round(((len(wordsn)/ len(wordsp)) * 100), 2), end="", flush=True)
    print(" ")

def wordcountoverlap (): #count de wordcountoverlap tussen text N1 en P1, N2 en P2, enz.
    for valueN, valueP in zip(listwcN, listwcP):
        print("Wordcount overlap: ", end="", flush=True)
        print(round(((valueN / valueP) * 100), 2), end="", flush=True)
        print("%")

def wordcomparison (tempn, tempp): #compares the words 1 on 1
    common_word = 0
    wordsn = word_tokenize(tempn)
    wordsp = word_tokenize(tempp)
    wordsn = [a.lower() for a in wordsn]
    wordsp = [a.lower() for a in wordsp]
    wordsn = [x for x in wordsn if not re.fullmatch('[' + string.punctuation + ']+', x)]
    wordsp = [x for x in wordsp if not re.fullmatch('[' + string.punctuation + ']+', x)]
    del wordsn[400:3000]
    #print(len(wordsn))
    #print(wordsn) #Displays the entire list of words
    for x in wordsn:
        a = 0
        for y in wordsp:
            if x == y:
                if a == 0:
                    common_word += 1
                    a += 1
                    break
    #print("Wordoverlap: ", end="", flush=True)
    print(round((common_word / (len(wordsn))) * 100), end="", flush=True)
    print("%")

def wordtest (tempn, tempp): #test if words are in text b and not text a
    nieuwewoorden = 0
    wordsexag = []
    wordsn = word_tokenize(tempn)
    wordsp = word_tokenize(tempp)
    wordsn = [a.lower() for a in wordsn]
    wordsp = [a.lower() for a in wordsp]
    wordsn = [x for x in wordsn if not re.fullmatch('[' + string.punctuation + ']+', x)]
    wordsp = [x for x in wordsp if not re.fullmatch('[' + string.punctuation + ']+', x)]
    for x in wordsn:
        a = 0
        for y in wordsp:
            if x == y:
                if a == 0:
                    a += 1
                    break
        if a == 0:
            nieuwewoorden += 1
            wordsexag.append(x)
            a += 1
    print(nieuwewoorden)
    print(wordsexag)

def prepare_results(p, r, f):
    return '\t{}:\t{}: {:5.2f}\t{}: {:5.2f}\t{}: {:5.2f}'.format(float, 'P', 100.0 * p, 'R', 100.0 * r, 'F1', 100.0 * f)


def rougescore(tempn, tempp, teller): #Performs the Rouge-L metric
    for aggregator in ['Avg', 'Best', 'Individual']:
        #print('Evaluation with {}'.format(aggregator))
        apply_avg = aggregator == 'Avg'
        apply_best = aggregator == 'Best'

        evaluator = rouge.Rouge(metrics=['rouge-l'],
                               max_n=4,
                               limit_length=True,
                               length_limit=100,
                               length_limit_type='words',
                               apply_avg=apply_avg,
                               apply_best=apply_best,
                               alpha=0.5, # Default F1_score
                               weight_factor=1.2,
                               stemming=True)


        hypothesis_1 = tempn
        references_1 = tempp

        all_hypothesis = [hypothesis_1]
        all_references = [references_1]

        scores = evaluator.get_scores(all_hypothesis, all_references)

        for metric, results in sorted(scores.items(), key=lambda x: x[0]):
            if not apply_avg and not apply_best: # value is a type of list as we evaluate each summary vs each reference
                for hypothesis_id, results_per_ref in enumerate(results):
                    nb_references = len(results_per_ref['p'])
                    for reference_id in range(nb_references):
                        print("Hypothesis #", teller, " & Reference #", teller, ": ".format(hypothesis_id, reference_id))
                        print(prepare_results(results_per_ref['p'][reference_id], results_per_ref['r'][reference_id], results_per_ref['f'][reference_id]))


def lcs(s1, s2): #does the LCS algorithm
    tokens1 = word_tokenize(s1)
    tokens2 = word_tokenize(s2)
    tokens1 = [a.lower() for a in tokens1]
    tokens2 = [a.lower() for a in tokens2]
    tokens1 = [x for x in tokens1 if not re.fullmatch('[' + string.punctuation + ']+', x)]
    tokens2 = [x for x in tokens2 if not re.fullmatch('[' + string.punctuation + ']+', x)]
    del tokens1[400:3000]
    cache = collections.defaultdict(dict)
    for i in range(-1, len(tokens1)):
        for j in range(-1, len(tokens2)):
            if i == -1 or j == -1:
                cache[i][j] = 0
            else:
                if tokens1[i] == tokens2[j]:
                    cache[i][j] = cache[i - 1][j - 1] + 1
                else:
                    cache[i][j] = max(cache[i - 1][j], cache[i][j - 1])
    x = cache[len(tokens1) - 1][len(tokens2) - 1]
    print(round((x / (len(tokens1))) * 100), end="", flush=True)
    print("%")




    #print("Wordoverlap: ", end="", flush=True)
    #print(round((common_word / (len(wordsn))) * 100), end="", flush=True)
    #print("%")

def main ():
    inlezenn()
    inlezenp()
    menu()

def print_menu():  ## Your menu design here
    print(30 * "-", "MENU", 30 * "-")
    print("1. Get a percentage based on text length.")
    print("2. Get a percentage based on actual wordoverlap.")
    print("3. Get the rougescore of the text.")
    print("4. Execute the LCS algorithm.")
    print("5. Find words in news text that aren't in press text.")
    print("6. Exit.")
    print(67 * "-")

def menu():
    loop = True
    print_menu()  ## Displays menu
    while loop:  ## While loop which will keep going until loop = False
        if keyboard.is_pressed('1'): #Text length function
            print("Menu 1 has been selected")
            teller = 0
            while teller != aantalfiles:
                #print(teller)
                filecount(liststringsN[teller], liststringsP[teller])
                teller += 1
            menu()
        elif keyboard.is_pressed('2'): #Text overlap function
            print("Menu 2 has been selected")
            teller = 0  # aanmaken variabele.
            while teller != aantalfiles:
                #print(teller)
                wordcomparison(liststringsN[teller], liststringsP[teller])
                teller += 1
            menu()
        elif keyboard.is_pressed('3'):  # rougescore
            print("Menu 3 has been selected")
            teller = 0  # aanmaken variabele.
            while teller != aantalfiles:
                rougescore(liststringsN[teller], liststringsP[teller], teller)
                teller += 1
            menu()
        elif keyboard.is_pressed('4'):
            print("Menu 4 has been selected")
            teller = 0  # aanmaken variabele.
            while teller != aantalfiles:
                lcs(liststringsN[teller], liststringsP[teller])
                teller += 1
            menu()
        elif keyboard.is_pressed('5'): #longestsentence
            print("Menu 5 has been selected")
            teller = 0  # aanmaken variabele.
            while teller != aantalfiles:
                wordtest(liststringsN[teller], liststringsP[teller])
                teller += 1
            menu()
        elif keyboard.is_pressed('6'):
            print("Program has quit")
            break
            ## You can add your code or functions here
main()
