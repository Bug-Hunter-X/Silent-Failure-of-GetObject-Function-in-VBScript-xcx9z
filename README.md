# Silent Failure of GetObject Function in VBScript

This repository demonstrates a common yet easily overlooked issue in VBScript: the silent failure of the `GetObject` function when the specified object is not found.  The problem stems from the use of `On Error Resume Next`, which suppresses error reporting, making it difficult to detect and handle failures gracefully.

## Problem

The `GetObject` function in VBScript is frequently used to access COM objects.  However, if the specified object does not exist, `GetObject` returns `Nothing` without generating an error. If `On Error Resume Next` is used to handle potential errors, the script continues execution without any indication that `GetObject` failed.

## Solution

The improved version explicitly checks the return value of `GetObject` and provides more informative error handling.  Instead of relying on `On Error Resume Next`, the code directly checks if the returned object is `Nothing`, providing clear feedback to the user.

## Usage

1.  Clone the repository.
2.  Open the `bug.vbs` and `bugSolution.vbs` files.
3.  Run each script separately and observe the difference in error handling.