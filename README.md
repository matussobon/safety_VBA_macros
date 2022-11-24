# FMEA MS Excel macros library

- This library contains MS Excel(VBA) macros essential for speeding up the FMEA table creation process. 
- These macros are uploaded in the form of .bas file but it is possible to open them as .txt file and copy the content into MS Excel *Developer module*. 
- The individual steps for this method are: 
   1. Download .bas file from this repository into your computer
   2. Right-click on the file
   3. Click on **Open/Open with...** and choose Notepad (or any other text editor of your choice)
   4. Highlight the whole code (or the part you need) and open MS Excel 
   5. In the top panel choose *Developer* tab (If its not there, you might need to enable it: [How to enable Developer tab](https://support.microsoft.com/en-us/office/show-the-developer-tab-e1192344-5e56-4d45-931b-e5fd9bea2d45))
   6. To open the Visual Basic editor/enviroment click on the *Visual Basic* icon in the *Developer* tab
   7. In the Visual Basic editor right-click on the white space in the *Project* window (this windows should be on the top left side of your screen), choose *Insert* and then *Module*
   8. Now you've created a module where you can copy the macro

### Macros included in this repository:
- [x] capacitors.bas
- [x] diodes.bas
- [x] inductors.bas
- [x] microcircuits.bas
- [x] microcircuits_specials.bas
- [x] resitors.bas
- [x] transistors.bas

### The basic structure of the macros: 
- Each macro consists of multiple sub functions with the same working principle
- These functions are then called in the *main* function to create a working macro
- The function used throughout the macros are called: *Split*, *DoesSheetExists*, *CopyCellValueToNewSheet*, *Transpose_new*, *ParameterLookUp*, *StrFind*,
  - The basic structure of the *main* function looks like this:
    1. *CopyCellValueToNewSheet*
    2. *Split*
    3. *Transpose_new*
    4. *ParameterLookUp*
    5. *StrFind*
   - The content of these functions is in principle the same for each macro, the only functions different between the macros are *ParameterLookUp* and *StrFind* 

### The working principle of each function:
- Functions *Split*, *Transpose_new* are used for formatting the working sheet, these functions are essential for the proper function of a macro and they **SHOULD NOT BE EDITED IN ANY WAY**
- Function *CopyCellValueToNewSheet* copies cell values from active sheet into the new sheet to create FMEA table
   - this functions uses subfunction *DoesSheetExists* to check whether the destination sheet already exists, if it does not the function will create one
- Function *ParameterLookUp* copies component failure rate and part description into the working sheet based on the parameters in the B and C column from the source sheet(Capacitors/Inductors/etc.)
- Function *StrFind* calculates the failure rate based on the component type and mode and outputs the result into the destination sheets and copies it into clipboard






