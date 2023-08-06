@echo off

echo Processing the Transcipt
python -W "ignore" cle\\Process_Transcript.py

echo Generating the Tutors Assignments
python cle\\Generate_Tutors_Assignment.py

echo Process Done!
