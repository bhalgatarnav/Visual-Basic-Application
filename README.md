# Visual Basic for Applications
VBA is a ackend software for microsoft applications.
The VBA Editor offers several windows designed for developing and debugging macros, each serving a distinct purpose:

1. **Project Window**:

   **Purpose:** This window presents all components within your VBA project, including modules, forms, class modules, and references to external libraries. It facilitates management of these components, allowing addition of new ones and easy navigation between them.

   **Importance:** The Project window provides a structural overview of your entire project, simplifying organization and access to different code segments.

2. **Property Inspector**:

   **Purpose:** The Property Inspector displays properties and their corresponding values for the currently selected object, be it a form, control, or module. It permits modification of these properties to customize appearance and behavior of objects.

   **Importance:** This window enables fine-tuning of functionality and aesthetics of forms and controls without direct code manipulation.

3. **Code Editor**:

   **Purpose:** This serves as the workspace for writing and editing VBA code. It offers features like syntax highlighting, auto-completion, and debugging tools to aid in code creation and troubleshooting.

   **Importance:** The Code Editor is the nucleus of the VBA development environment, where instructions are crafted to automate tasks and extend functionalities within Office applications.

4. **Immediate Window**:

   **Purpose:** The Immediate Window allows direct execution of VBA commands and immediate viewing of their outcomes. It's beneficial for testing code snippets, evaluating expressions, and debugging macros.

   **Importance:** Offering a swift and interactive means to test and debug code, the Immediate window obviates the need to run the entire macro or set breakpoints.

5. **Locals Window**:

   **Purpose:** This window exhibits values of all variables and objects within the current scope during code execution. It aids in comprehending data changes throughout the macro and identifying potential issues.

   **Importance:** As a valuable debugging tool, the Locals window permits inspection of variable and object states at specific code execution points.

* All code is written as subroutines in the Macro. any code you want to run in the macro should be enclosed within `sub` and `End Sub`.

## Variables Declarations
` Dim Variable As VariableType`
`integer` `long` `double` `string` `boolean`

* `CDbl()` - convert to type double

* `CInt()` - convert to type integer

* `CLng()` - convert to type long

* `CStr()` - convert to type string

* `CBool()` - convert to type boolean

## Performing Calculations:

Precedence
`^` Power
`+ve or -ve` Positive/Negative
`* or /` Multiplication/Division
`%`  Modulus
`+ or -` Addition/Subtraction

## Complex Mathematical Operations.
Trigonometric: `sin()`, `cos()`, `tan()`
Logarithmic: `log()`
Exponential: `exp()`
Absolute Value: `abs ()`
Square Root: `sqr()`
Rounding: `round()`

## Accessing the value of the worksheet
`var = ActiveSheet.Cells(Rows,Columns).Value`
* ActiveSheet tells the editor to access the value of the current open worksheet.
The **index** for the Rows and Columns starts with 1.

## Saving and Opening the Excel Files.
The excel should be saved in the excel Macro-Enabled Workbook. (.xlsm)

## Printing Values

* `Debug.Print x` would print the value of x in the **immediate window**.
* `MsgBox x` would print the value of x in the **Message Box ( Alert )** Pop up Alert.


## Using the Conditional Operators and statements.
* Value1 `=` Value2 (Equals to)
* Value1 `<>` Value2 (Not Equals to)
* Value1 `<` Value2 (Lessthan)
* Value1 `<=` Value2 (Less than Equal to)
* Value1 `>` Value2 (Greater than)
* Value1 `>=` Value2 (Greater than Equal to)

* Value1 `And` Value2
* Value1 `Or` Value2
* `Not`(Value1)

The operators used above work for numbers, single characters and strings. the ASCII values are used to campare these values and then finding the corresponding value. Two strings are said to be equal if they have the same characters in the same pattern. You can determine the order of strings by using the above operators.

` "Hear" < "Here" `
The result will be True as hear is precent to here in the alphabetical order.

## If else Ladder.

* Syntax:
```
If condition Then
 conditional code
ElseIf condition Then
 conditional code
Else 
 else code
End If
```

# Loops
* Conditional Based Reptition: **More Versatile** The identified actions are continued untill some event occurs that causes the algorithm to stop repeating.

* Counter Based Repetition: The identified actions are repeated for a specific number of times determined prior to starting the repeatition.


## For.... Next Loop

**First Method**
```
For var = start to End Step Change
   ....code to repeat....
Next
```
**Second method**
```
For Each Item In Array
   ....code to repeat....
Next
```


## Initialization of Array

```
Dim ArrayName(N) as VariableType
Dim 2DArray(R, C) as VariableType

ReDim ArrayName(N) as VariableType
ReDim Preserve ArrayName(N) as VariableType
```



## Functions in VBA

```Function FunctionName(Parameter) As Type
   function code
   FunctionName = value to return
   
   end Function```