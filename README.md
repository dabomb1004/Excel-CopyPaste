# Excel-CopyPaste
This was created for anyone who has to copy and paste a bunch of crap from one place to another. This uses openpyxl and borrows from many contributors but makes it a bit more user friendly and organized. If you want to improve it and abstract it more by all means do so. 

# How to Use 

Example:
Say you want to copy some cell from sheet A and paste to cells in sheet B. Well this is how you go about it. 

Create an instance of copysheet(). 
The copysheet takes three arguments: copysheet("file path of file you want to copy from", "file path of file you want to paste to", "file path you want to save to")

So in this scenario, we want to copy from excel file A and paste to excel file B and we'll save into an entirely new file called excel file C: 

CopySheet = copysheet("C://Users//Copy//Puppets//A.xlsx", "C://Users//Copy//Pastries//B.xlsx", "C://Users//Copy//Pastries//C.xlsx")

Then call the copyPaste method on the newly created instance. 
copyPaste takes the following arguments: copyPaste(Sheet Tab of file you want to copy, column letter start , column # start, column letter end, column # end, 
                                                  (Sheet Tab of file you want to paste, column letter start, column # start, column letter end, column # end)

So in this case, I want to copy the sheet called "pinocchio which is in file A and copy cells B25:F40 and paste to file B to a sheet called "jiminy" in cells B36:F51
CopySheet.copyPaste("pinocchio", "B", 25, "F",40,"jiminy", "B", 36,"F", 51)

The final step is to save using the save() method 

Copysheet.save()

# Additional notes

Now say I wanted to grab other tables within a different sheet in file A and paste it to a sheet tab in file B. 
I don't have to create a new instance. I can just call copyPaste() method again: 

CopySheet.copyPaste("A different tab sheet to paste to", "B", 20, "F",40,"A different tab sheet to paste to", "B", 36,"F", 51)

CopySheet.save()

If however, you wanted to copy and paste from different files than A and B but continue to add onto file C, you would create a new instance but make file C as the paste 
argument (2nd argument) in copysheet() object. So here is what you would do: 

CopySheet2 = copysheet("C://Users//Copy//Puppets//X.xlsx", "C://Users//Copy//Pastries//C.xlsx", "C://Users//Copy//Pastries//C.xlsx")
CopySheet2.copyPaste("A different tab sheet to paste to", "B", 20, "F",40,"A different tab sheet to paste to", "B", 36,"F", 51)

CopySheet2.save()

As you can see above, we are copying from X.xlsx (which is the first argument in copysheet()) and then we are pasting into C.xlsx, the file we saved all our paste
changes in file B to. So essentially what this is doing is adding on top of the changes you already made and saved into file C. 

Below are the arguments that copyPaste takes copyPaste("Name of Sheet Tab to Copy", "Column Start", Row # Start,"Column End", Row # End, "Name of Sheet Tab to Paste", "Column Start", Row # Start, "Column End", Row # End)

Hope this helps. Also, I originally had a sql function that could connect and query into a database and used a custom function to turn it into a pandas dataframe before 
converting it into a list to paste into excel cells. But since there are different ways of connecting to DB, I left it blank for others to customize it for their use 
cases. Anyways hope this saves someone some time. Cheers! 
