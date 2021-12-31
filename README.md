# QuestionAsker

This is a personal project that I made over the summer of 2020 to aid my mathematics studies. The Question Asker program is a Python flashcard GUI implemented using Tkinter. The user must create an Excel spreadsheet like that found within the repository, the first column with the question and the second with the answer (Latex can be typed inside `$$` characters as normal). The program will count the number of times a question is asked and answered correctly, these values will be saved in the following two columns. 

NOTE: The spreadsheet cannot be opened at the same time as the GUI is in use, otherwise saving of the spreadsheet will not be possible. 

---

# Instructions for Use

Firstly, ensure that `os.chdir(...)` on line x points to the correct directory (that which contains the spreadsheet). Then ensure that loadfilename correctly contains the name of the spreadsheet (Notes.xlsx) by default. When the program is first run, you will see the main menu.

<p align="center" width="100%">
    <img width="40%" src="https://user-images.githubusercontent.com/77517061/147829153-951f6f8b-fa30-4d96-91fa-bc81ecd0eae7.png">
</p>

By pressing the 'Study' button you will progress to the second page where you can select sheets from the spreadsheet to work on, you can select as many as desired. 

<p align="center" width="100%">
    <img width="40%" src="https://user-images.githubusercontent.com/77517061/147829237-7265c226-b43e-4241-b4e5-55167121c962.png">
</p>

After pressing 'Go!', you will be presented with a question in blue and a space to type an answer.

<p align="center" width="100%">
    <img width="40%" src="https://user-images.githubusercontent.com/77517061/147829288-9919cf8d-dd95-4e40-81d5-7d5e36073995.png">
</p>

The 'Okay!' button will take you to the answer page, where the correct answer is displayed (with Latex formatting) and your answer is displayed below. Check buttons are available to select if your response was correct, this information will also be saved into the spreadsheet. The 'Next Question!' button repeats the process and 'Finish!' Takes us to a Recap screen then back to the main menu.

The main menu also includes a 'Try Me Text' button that allows you to check your latex as you create your spreadsheet during revision. This can be useful as some latex commands are not correctly displayed by matplotlib. The Try Me Text page is similar to the Question and Answer pages as shown below. 

<p align="center" width="100%">
    <img width="45%" src="https://user-images.githubusercontent.com/77517061/147829411-0f30cea9-7d5c-40ad-bd94-a5144964f769.png">
</p>

<p align="center" width="100%">
    <img width="55%" src="https://user-images.githubusercontent.com/77517061/147829417-5cc450c7-a370-4543-895a-9ffc27f1d5ab.png">
</p>








