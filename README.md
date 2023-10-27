# VBA-challenge

Before beginning this project, I quickly ran through the approach to the logic at the end of a tutoring session with Simon Rennocks.
He helped me to think through the process for the first portion and I used that to then go back and start writing code. 

I ran into some trouble calculating the Stock Total Volume. Using google, I shifted that variable to be a variant and the code worked.
Dim Stock_Volume_Total As Variant
I did the same for help to find how to format column J and found this line of code: 
'Column Formatting
ws.Range("J:J").NumberFormat = "0.00"

Another tutor, Matthew Werth, assisted me in the last portion of the project, specifically this code:
'Check "Greatest Increase" Values
        If Percentage_Change > Greatest_Increase Then
            Greatest_Increase = Percentage_Change
He also assisted me in updating my code to run through all worksheets within the file.
Together we checked online what code would adjust the column size and applied this line of code.
ws.Range("I:Q").EntireColumn.AutoFit
