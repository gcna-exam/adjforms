# adjforms
creating and parsing adjudication forms for the recording stage of the GCNA exam  
***NB : current code assumes old numerical grading system***  

***
### create_adjforms.py
Original version fall 2020 by M. Pan  
Given tab-separated-variable output with candidates' program info from Google form,  
* create Word (.docx) adjudication forms for each candidate and juror  
* create required and repertoire piece forms for each juror  
* create PDF program listing of all candidates' recordings

#### Dependencies  
* Python 3 (used with Python 3.7) with packages docx, docxcompose, mailmerge (, datetime, math, subprocess)
* A LaTeX installation, including pdflatex
* Word templates adjform.docx (replace with adjform_pf.docx for pass/fail grading), overallform.docx, repertoirepieceform.docx, requiredpieceform.docx in the same directory as python script
* Assumes a \*nix-like OS (I believe this is easily generalized)  

#### How to use
* Edit `PARAMETERS` section with data for current year:  
    * Change `examyear` to year for this exam cycle  
    * Change `jurors` to list of current jurors  
    * Change `tsvfile` to name of file containing Google form program info, supplied by candidates  
* If desired, edit list `forms_to_make` in `FORMS TO GENERATE` section to contain only the forms desired:  
    * 'adj' : adjudication forms  
    * 'rep' : repertoire piece form  
    * 'req' : required piece form  
    * 'prog' : program listing  
* `python3 create_adjforms.py`

***
### parse_adjforms.py
Original version winter 2021 by M. Pan  
Given completed adjudication forms from jurors,
* extract grades and overall pass/fail decision from each set of forms, print alerts for missing grades  
* create .csv file containing all grades
* create PDF summary of voting jurors' decisions for all candidates  
* create JSON summary of lists of voting jurors and overall pass/fail results for all candidates
* create JSON summary of all grades

#### Dependencies
* Python 3 (used with Python 3.7) with packages docx, docxcompose (, copy, glob, json, os, random, subprocess)
* A Latex installation, including pdflatex
* Assumes a \*nix-like OS (I believe this is easily generalized)

#### How to use
* Edit `Parameters` section with data for current year:
     * Change `examyear` to year for this exam cycle
     * Change `jurors` to list of jurors who returned forms this year
     * Change `voting` to list of designated voting jurors
     * Change `conflict` to contain all juror recusals; keys are candidate numbers with recusals and values are list of jurors recused for that candidate
     * Change `labelstr` as needed to 'prelim' (for results before juror discussion) or 'final' (for final results)  
* Remove from current working directory all juror forms that are previous versions or otherwise should not be used; check that all latest-version juror forms are in current working directory
* Check that no (latest-version) juror forms are open in Word
* `python3 parse_adjforms.py`
* Keep an eye on the standard output; any missing grades or missing forms will be flagged there

##### Note : how are voting jurors selected in case of (a) recusal(s)?
* If the recusal(s) is(are) of alternate jurors only, do nothing
* If there are at most five non-recused jurors for this candidate, keep all of them; if there are fewer than five, print a warning
* If there are four non-recused voting jurors and two non-recused alternate jurors, pick an alternate juror at random to act as voting juror for this candidate only
