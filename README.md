# qrexam

## python scripts for online marking of exams

The purpose of these scripts is to format exams so that they can be marked electronically 
without requiring the people marking the exam to be together in one room.

The scripts were developed at Heidelberg university to facilitate the marking of the 
physics 1 and 2 exams (with ~400 written exams) during Corona times. The first version of 
the code was written by *Udo Friess* and is contained here in the directory v1. The scripts 
are released under GPL3.

List of authors for the current version:

Udo Friess          udo.friess@iup.physik.uni-heidelberg.de

Robert W. Schmidt   rschmidt@uni-heidelberg.de

Please check required_packages.txt for required moduls/packages.

**Usage**:

The idea is to scan the exams after they are written. The exams are numbered, which is 
encoded in a QR code. In Heidelberg the exams are usually distributed in the writing halls, 
so that it is essentially random which student gets which exam/number. This set of scripts 
(different from the ones in the directory v1) requires them to be scanned in separate piles 
sorted by problem number. This was done so that the additional pages at the end of each 
exam can be attributed to the correct person. In order to create the final pdf for the 
students again, an xlsx file connecting the student names and the exam numbers is required. 
In principle the student names only need to be given on the first page as all pages are 
connected by the exam number. Thus, marking can be done 'blindly' without knowing the name.

This sequence of scripts lets you

1. Create exams pdfs to be delievered to the printer.

2. Create a data base from all scanned exams.

3. Use this data base to sort the scanned exams by problem number and by grouping them into 
   'batches'.

4. Re-group the pdfs from point (3) to create the final pdfs for each
student that can then be uploaded, e.g., to an exam inspection tool.

**Detailed comments:**

- The script numbering has been chosen to be reminiscent of the original version contained 
  in 'v1'.

- 01_QRcode_on_exam_single.py: This script will create *NumKlausuren* versions of the pdf 
  file demo_klausur.pdf. The exams will be put into the directory 'tray'. The script 
  assumes that there is one page per problem plus a title page and three additional pages 
  at the end. (We will scan the reverse sides of each page, but we do not print anything 
  there). The exams will be numbered continuously, with the exception of a few 
  'superstitional' numbers contained in *avoid_numbers*. On the title page the exam number 
  is also included below the name field.

  The script assumes *NumAufgaben* problems. There is an empty space on the bottom of each 
  page for the QR code.

- 04_Scan_QR_Codes_single.py: This script will scan all scanned pdfs in the directory 
  'scans'. Note that the scans have to be sorted by problem number. All file names have to 
  start with a code aXX which encodes the problem number, e.g. a02_57851854854252.pdf The 
  script will generate a pandas data base which contains the location of all pages and 
  their orientation (rotated/flipped).

- 05_Check_Teilnehmer.py: This script allows you a plausibility check whether all submitted 
  exams have all the required problems. If one problem is missing (see script 09), it can 
  be indidcated in the scritp (special case, best to avoid for starters).

- 06_sort_scans_by_problem.py: This script will create the problem folders in the directory 
  'exams_sorted'. The number of problems is contained in 'num_aufgaben'. The problems pdfs 
  are divided into 'batches' with a maximum number 'batch_size'. In fact, the batches can 
  be even smaller if certain printed exams were not actually used (or were omitted for 
  reasons of 'superstition'). If 'batch_size' is known,# it is easy to find a certain exam 
  number in the files. (e.g. exam 35 in batch 2 if batch_size is 20).

- The normal mode in Heidelberg then was to mark those files and drop the marked pdfs in a 
  directory 'exams_marked' with the same structure as exams_sorted.

- 07_collect_exams.py: This script will use the pandas data bases created in previous steps 
  and the files in 'exams_marked' to create the actual student exam pdfs in exams_out. It 
  require an xlsx file 'results.xlsx' which relate the Names and the exam numbers. For 
  this, we use the columns 'ID', 'Matrikelnummer', 'Name', 'Vorname', 'Nummer'. (can be 
  modified of course for other contexts, it is however necessary to connect exam number and 
  student name somehow).

- 09_add_pages.py [optional]: This script can be used to add marked problems that were 
  marked on paper and only scanned afterwards (amounts to a special service to the people 
  marking if they feel more comfortable with paper). It essentially modifies the data base 
  created by '06_sort_scans_by_problem.py' so that the relevant problem number is added to 
  the data base. Handle with care.
