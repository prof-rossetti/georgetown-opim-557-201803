# Project 1 Checkpoints

This document provides a progression of sequential checkpoints to help focus your development efforts.

It assumes you have already configured your desired user interface.

## Checkpoint 1: Capture and Display Inputs

  1. Capture user inputs and display them in a message box.
  2. After demonstrating your ability to display raw input values in a message box, format currency and percentage values as applicable. Don't worry if at this point you are writing the same formatting-related snippet of code in multiple places. The next step is about simplifying, or "refactoring" the code to remove such duplication.
  3. You are encouraged but not required to define one or more custom functions to perform the numeric formatting. Hint: you will need to pass the number as a parameter when invoking the function(s).

## Checkpoint 2: Display Outputs

  1. After displaying inputs, also display the outputs. Don't worry if you haven't performed the necessary calculations yet. Use hard-coded example values for now. Your objective is to get the display right. Format currency and percentage values as applicable.

## Checkpoint 3: Perform Calculations

  1. Calculate the savings balance for the end of one year, and display it (see Checkpoint 2) as if it were the final balance.
  2. Forget about calculating anything. See if you can loop through each year between the customer's current age and their desired retirement age. Optionally produce a message box to display each age value. This is a temporary check to make sure you are looping properly. Make sure to increment the age to avoid getting stuck in an infinite loop!
  3. Modify the code inside your loop to calculate the final savings balance, and display it (see Checkpoint 2).
  4. Finally, modify the code inside your loop to calculate the remaining output values, and display them (see Checkpoint 2). Hint: you may need to declare additional variables outside the loop's scope.

## Checkpoint 4: Validate Inputs

Perform this checkpoint only if your interface allows users to enter invalid input values by mistake.

  1. After capturing each input, detect its datatype and optionally display the datatype in a message box (loosely referred to here as a "validation message box").
  2. For each numeric input, detect whether or not it falls within an acceptable range, and display the determination (e.g. "valid" vs "invalid") in a validation message box.
  3. Instead of displaying validation message boxes in all cases for all inputs, only display a validation message box if a given value is invalid.
  4. For each invalid input, instead of displaying a message box, programmatically exit the sub-procedure after displaying a friendly error message box to the user with instructions on how to fix the problem.
