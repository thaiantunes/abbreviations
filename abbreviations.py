#!/usr/bin/python

import re
import docx2txt
import easygui
import csv

def main():

    regex = r"[A-Z]{2,}"

    #dictionary of commom abbreviations that don't need to be added to the table 
    boring_words = {"ID", "LOG", "STAR", "MOON"} 

    #dictionary of abbreviations and their meaning
    fun_words  = {"TBD":"To be defined", "KOM":"Kick Off Meeting", "PO":"Purchase Order"}

    try:
        filename = easygui.fileopenbox()
        text = docx2txt.process(filename)

        abbreviation_list = re.findall(regex, text)
        abbreviation_list = [i for i in abbreviation_list if i not in boring_words]
        abbreviation_list.sort()

        abbreviations = {}
            
        for i in abbreviation_list:
            if i in fun_words:
                abbreviations[i] = fun_words[i]
            else:
                abbreviations[i] = "###"

        with open("abbreviations_table.csv", "w", newline='') as tb:
            writer = csv.writer(tb)
            writer.writerow(["Abbreviation", "Meaning"])
            for i in abbreviations:
                writer.writerow([i,abbreviations[i]])

    except:
        print("Please close the .docx file")

if __name__ == "__main__":
    main()