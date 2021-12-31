# QuestionAsker

This is a personal project that I made over the summer of 2020 to aid my mathematics studies. The Question Asker program is a Python flashcard GUI implemented using Tkinter. The user must create an Excel spreadsheet like that found within the repository, the first column with the question and the second with the answer (Latex can be typed inside `$$` characters as normal). The program will count the number of times a question is asked and answered correctly, these values will be saved in the following two columns. 

NOTE: The spreadsheet cannot be opened at the same time as the GUI is in use, otherwise saving of the spreadsheet will not be possible. 

---

# Instructions for Use

Firstly, ensure that `os.chdir(...)` on line x points to the correct directory (that which contains the spreadsheet). Then ensure that loadfilename correctly contains the name of the spreadsheet (Week 2 Summary.xlsx) by default. When the program is first run, you will see the main menu.

![First Page](assets/StartPage.png)

By pressing the 'Study' button you will progress to the second page where you can select sheets from the spreadsheet to work on, you can select as many as desired. 



