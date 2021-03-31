'''
Parse returned (docx format) adjudication forms 
Flag missing grades
Create .json dictionary of results
Create pdf summary of overall pass/fail marks, for committee use
Create csv summary of individual piece and overall marks, for board use

*** NB : move all extraneous juror files out of directory containing completed forms before running
*** NB : close all docx files to be parsed before running

THINGS TO BE EDITED EACH YEAR:
-- "Parameters"
'''

import copy
import glob
import json
import os
import random
import subprocess

from docx import Document
from docx.enum.text import WD_BREAK
from docxcompose.composer import Composer
#from docx2pdf import convert

##### Parameters #####
######################

examyear = '2021'

jurors = ['cortez','dzuris','ellis','harwood','hunsberger','lehrer','lens']
candidates = ['1','2','3','4','5','6','7','8']
voting = ['cortez','dzuris','ellis','harwood','hunsberger']
conflict = {'3':['lens']}
numrequired = 5

labelstr = 'final'

##### Helper functions #####
############################

def insensitive_glob(pattern):

    # return all filenames containing case-insensitive version of pattern
    
    def either(c):
        return '[%s%s]' % (c.lower(), c.upper()) if c.isalpha() else c
    
    return glob.glob(''.join(map(either, pattern)))


def get_candnumber(adjform):

    # extract candidate number from adjform Document

    return adjform.paragraphs[3].text.split('\t')[0].split('.')[-1].strip()


def get_pieces(adjform):

    # extract piece names from adjform Document

    piecelines = [ind-1 for ind,par in enumerate(adjform.paragraphs) if par.text == 'piece\t\t\t\t\t\t\tcomposer']

    return [adjform.paragraphs[ind].text.split('\t')[0] for ind in piecelines]


def get_grades(adjform):

    # extract grades from adjform Document for individual pieces

    gradelines = [par.text for par in adjform.paragraphs if par.text[:9] == 'Candidate']

    grades = []
    for line in gradelines:
        gradestr = line.split('\t')[-1].split('Over')[0].strip('_')

        try:
            thisgrade = float(gradestr)
        except:
            # check for fractions with "/" divider
            if '/' in gradestr:
                if ' ' in gradestr:
                    wholegrade = gradestr.split()[0]
                    if wholegrade.isnumeric():
                        fracgrade = gradestr.split()[1].strip().split('/')
                        if len(fracgrade) == 2 and fracgrade[0].isnumeric() and fracgrade[1].isnumeric():
                            thisgrade = int(wholegrade) + float(fracgrade[0])/float(fracgrade[1])
                        else:
                            thisgrade = float(wholegrade)
                    else:
                        thisgrade = 0
                else:
                    gradeparts = gradestr.split('/')
                    wholegrade = gradeparts[0][:-1]
                    fracgrade = [gradeparts[0][-1],gradeparts[1]]
                    if wholegrade.isnumeric():
                        if fracgrade[0].isnumeric() and fracgrade[1].isnumeric():
                            thisgrade = int(wholegrade) + float(fracgrade[0])/float(fracgrade[1])
                        else:
                            thisgrade = float(wholegrade)
                    else:
                        thisgrade = 0
            else:
                # check for "+","-" in grade
                trygrade = gradestr.replace('+','').replace('-','')
                if trygrade.isnumeric():
                    if '+' in gradestr:
                        offset = 0.3
                    elif '-' in gradestr:
                        offset = -0.3
                    thisgrade = float(trygrade)+offset
                else:
                    thisgrade = 0
                
        grades.append(thisgrade)

    return grades
                

def get_req(adjform):

    # extract req/non-req info from adjform Document
    # 1 = required , 0 = non-required

    reqlines = [par.text for par in adjform.paragraphs if par.text[:8] == 'Required']

    reqlist = []
    for line in reqlines:
        reqstr = line.split('\t')[1].strip()
        if reqstr == 'Yes':
            reqlist.append(1)
        elif reqstr == 'No':
            reqlist.append(0)
        else:
            reqlist.append(-1)

    return reqlist
        

def get_overall(adjform):

    # extract overall pass/fail from adjform Document

    overall = ''
    
    try:
        overallline = adjform.tables[0].rows[1].cells[1].paragraphs[1]
    except:
        try:
            overallline = adjform.tables[0].rows[1].cells[1].paragraphs[0]
        except:
            print("can't get line with overall mark")

    # if there are no formatting details in this line, just check text
    if len(overallline.runs) == 1 or len(overallline.runs) == 2:
        if 'do not pass' in overallline.text:
            overall = 'fail'
        elif 'pass' in overallline.text:
            overall = 'pass'
        else:
            overall = ""
            
    # but if there are formatting details the meaning might be different from that of the text string
    elif len(overallline.runs) >= 3:
        if 'do not' in overallline.runs[1].text:
            if overallline.runs[1].font.bold:
                overall = 'fail'
            if overallline.runs[1].font.underline:
                overall = 'fail'
            if overallline.runs[1].font.strike:
                overall = 'pass'
        else:
            overall = 'pass'
            
    else:
        overall = "can't parse"

    return overall
    

def record_grades(thisdict,pieces,grades):

    # record grades for this juror in results dictionary
    
    for ind in range(len(pieces)):

        piece = pieces[ind]
        grade = grades[ind]
        
        if grade == 0:
            print('cand '+candidate+'/'+juror+'/'+piece+' : missing grade')
                
        if piece not in thisdict.keys():
            thisdict[piece] = {juror:grades[ind]}
        else:
            thisdict[piece][juror] = grades[ind]

    return thisdict


def record_overall(thisdict,overall,req_outcomes):

    # record overall grade, check for consistency

    if overall == 'pass':
        thisdict['pass'].append(juror)
        if any(prod>0 and prod<3 for prod in req_outcomes):
            print('cand '+candidate+'/'+juror+' : overall pass, failed req')
    elif overall == 'fail':
        thisdict['fail'].append(juror)
        if all(prod==0 or prod>=3 for prod in req_outcomes):
            print('cand '+candidate+'/'+juror+' : overall fail, all req passed')
    else:
        print('cand '+candidate+'/'+juror+' : overall mark "'+overall+'"')

    return thisdict
    

def make_jurorsummary(results,jurors,voting,conflicts):

    # construct summary of overall grades for committee reference

    # find nonvoting jurors, set up random choice of one for each candidate
    altjurors = [x for x in jurors if x not in voting]
    if altjurors:
        altchoices = random.Random(int(examyear)).choices(range(len(altjurors)),[1]*len(altjurors),k=len(candidates))

    votingsummary = {}
    
    for candidate in candidates:

        # check which jurors marked this candidate
        thisjurors = results[candidate]['pass'] + results[candidate]['fail']
        thisaltjurors = [x for x in altjurors if x in thisjurors]

        # check for juror recusals
        if candidate in conflict.keys():
            recuse = conflict[candidate]
        else:
            recuse = []

        # determine voting jurors for this candidate by picking altjuror if any recusals
        if len(recuse) > len(thisjurors)-numrequired:
            
            # there are a lot of recusals or AWOL voting jurors, just use all jurors possible
            print('cand '+candidate+' : not enough jurors')
            thisvoting = [x for x in thisjurors if x not in recuse]
            
        elif set(voting).intersection(set(recuse)):
            
            # there are recusals and there are enough alternates
            if len(thisaltjurors) == 1:
                thisvoting = [x for x in voting if x not in recuse and x in thisjurors] + thisaltjurors
                if len(thisvoting) < numrequired:
                    print('cand '+candidate+' : not enough jurors')

            elif len(thisaltjurors) > 1:
                thisvoting = [x for x in voting if x not in recuse and x in thisjurors] + [thisaltjurors[altchoices[0]]]
                if len(thisvoting) < numrequired:
                    thisvoting.append(altjurors[1+altchoices[0] % 2])
                    
            else:
                thisvoting = [x for x in voting if x not in recuse and x in thisjurors]
                
            altchoices = altchoices[1:]

        else:
            
            # in effect there are no recusals
            thisvoting = [juror for juror in voting if juror in thisjurors]

            # but if there are not enough voting jurors, still have to use alternates
            if numrequired != len(thisvoting):
                okaltjurors = [x for x in thisaltjurors if x not in recuse]
                if numrequired-len(thisvoting) >= len(okaltjurors):
                    thisvoting += okaltjurors
                else:
                    thisvoting += okaltjurors[altchoices[0]]

            if len(thisvoting) < numrequired:
                print('cand '+candidate+' : not enough jurors')

            altchoices = altchoices[1:]
            

        # store vote tallies
        numpass = len([x for x in results[candidate]['pass'] if x in thisvoting])
        numfail = len([x for x in results[candidate]['fail'] if x in thisvoting])
        votingsummary[candidate] = {'pass/fail':[numpass,numfail],'voting':thisvoting}

        if abs(numpass-numfail) <= 1 or 'prelim' not in labelstr:
            make_candidate_pdf(candidate,thisjurors)
            
    with open(examyear+'votingsummary.json','w') as fh:
        json.dump(votingsummary,fh,indent=4,sort_keys=True)


    # create latex document with grade summary
    preamble = [
        '\\documentclass[10pt]{article}\n', \
        '\n', \
        '\\usepackage{parskip,array}\n', \
        '\\usepackage[scaled=.9]{helvet}\n', \
        '\\usepackage[T1]{fontenc}\n', \
        '\n', \
        '\\addtolength{\\topmargin}{-.9in}\n', \
        '\\addtolength{\\oddsidemargin}{0in}\n', \
        '\\addtolength{\\oddsidemargin}{-1in}\n', \
        '\\addtolength{\\textwidth}{2in}\n', \
        '\\addtolength{\\textheight}{1.7in}\n', \
        '\n', \
        '\\renewcommand\\familydefault{\\sfdefault}\n', \
        '\n', \
        '\\begin{document}\n', \
        '\n', \
        '\\begin{LARGE}\n', \
        '\\noindent {\\bf '+examyear+' '+labelstr+' recording stage results}\\bigskip\\hfill\n', \
        '\\end{LARGE}\n', \
        '\n', \
        '\\begin{large}\n', \
        ]

    votetable = ['\\begin{tabular}{cc}\n', \
                 '{\\bf Candidate}& {\\bf pass/fail} \\makebox[0in][l]{(voting jurors only)}\\smallskip\\\\\n']

    for candidate in candidates:

        thisvote = votingsummary[candidate]
        votetable += [candidate+'& '+str(thisvote['pass/fail'][0])+'/'+str(thisvote['pass/fail'][1])+'\\\\\n']
        
    votetable += ['\\end{tabular}\n\n']

    votelines = preamble + votetable + ['\\end{large}\n','\\end{document}\n']
    
    votefile = examyear+'_'+labelstr+'_recording_summary.tex'
    with open(votefile,'w') as fh:
        for line in votelines:
            _ = fh.write(line)

    subprocess.run(['pdflatex',votefile])


def make_candidate_pdf(candidate,thisjurors):

    # create pdf of all adjudications for juror review

    composer = ''
    for juror in thisjurors:
        
        filelist = insensitive_glob('*'+juror+'*docx')
        thisformname = [form for form in filelist if 'candidate'+candidate in form]

        if not thisformname:
            print('cand '+candidate+'/'+juror+' : strange name for form?')
            thisformname = ''
            
        else:
            thisform = Document(thisformname[0])
            numpar = len(thisform.paragraphs)
            for ind in range(numpar):
                par = thisform.paragraphs[ind]
                if 'Juror' in par.text and 'Signature' in par.text:
                    run = par.add_run()
                    run.add_break(WD_BREAK.PAGE)
                if ind == numpar-1:
                    if len(par.runs) != 0:
                        run = par.add_run()
                        run.add_break(WD_BREAK.PAGE)
            if not composer:
                composer = Composer(thisform)
            else:
                composer.append(thisform)

    composer.save(examyear+'_candidate'+candidate+'_all.docx')
    #convert(examyear+'_candidate'+candidate+'_all.docx')
    

def make_boardsummary(results):

    # create csv for board convenience

    csvlist = ['candidate,piece,'+','.join(jurors)+',,overall,'+','.join(jurors)]
    #csvlist = ['candidate,piece,'+','.join(jurors)+',range,avg,,overall,'+','.join(jurors)]

    for candidate in candidates:

        overallstr = ''

        pieces = [key for key in results[candidate].keys() if key not in ['pass','fail']]
        for piece in pieces:
            
            thisline = candidate + ',' + piece.replace(',','') + ','
            
            for juror in jurors:
                if juror in results[candidate][piece].keys():
                    thisline += str(results[candidate][piece][juror]) + ','
                else:
                    thisline += '0,'
                    
            #thispiecegrades = [v for k,v in results[candidate][piece].items() if v != 0]
            #if thispiecegrades:
            #    graderange = max(thispiecegrades) - min(thispiecegrades)
            #    gradeavg = sum(thispiecegrades)/len(thispiecegrades)
            #else:
            #    graderange = 0
            #    gradeavg = 0
            #    
            #thisline += str(graderange) + ',' + '{:4.2f}'.format(gradeavg) + ','*(len(jurors)+2)
            thisline += ','*(len(jurors)+1)
            csvlist.append(thisline)

        for juror in jurors:
            if juror in results[candidate]['pass']:
                overallstr += ',P'
            elif juror in results[candidate]['fail']:
                overallstr += ',F'
            else:
                overallstr += ','

        csvlist[-1] = csvlist[-1].replace(','*(len(jurors)+2),',,'+overallstr)    
        
    with open(examyear+'_'+labelstr+'_exam_grade_summary.csv','w') as fh:
        for line in csvlist:
            _ = fh.write(line+'\n')
        

        
                
    
############################

results = {}
for candidate in candidates:
    results[candidate] = {'pass':[],'fail':[]}

for juror in jurors:

    print(juror+' : ')

    # get list of this juror's forms
    filelist = insensitive_glob('*'+juror+'*docx')
    filelist = [x for x in filelist if x and 'prelim' not in x]

    for filename in filelist:

        # extract grades from adjudication form
        adjform = Document(filename)

        candidate = get_candnumber(adjform)
        print(candidate,end=' ')
        pieces = get_pieces(adjform)
        reqlist = get_req(adjform)
        grades = get_grades(adjform)
        overall = get_overall(adjform)

        thisdict = copy.deepcopy(results[candidate])

        # store individual piece grades in results
        thisdict = record_grades(thisdict,pieces,grades)
        
        # check for overall pass/fail and consistency
        req_outcomes = [grades[ind]*reqlist[ind] for ind in range(len(pieces))]
        thisdict = record_overall(thisdict,overall,req_outcomes)

        results[candidate] = thisdict
        
    print('')

# create summaries of overall scores
make_jurorsummary(results,jurors,voting,conflict)

#if 'prelim' not in labelstr:
make_boardsummary(results)

with open('results'+examyear+'.json','w') as fh:
    json.dump(results,fh,indent=4,sort_keys=True)
