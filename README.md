This repository demonstrates a subtle bug in VBScript's `IsEmpty` function. The `IsEmpty` function, intended to check for uninitialized variables or empty variants, incorrectly identifies an empty string as empty.  This leads to unexpected behavior in functions expecting different handling for truly empty parameters versus empty strings. The `bug.vbs` file illustrates the problem, while `bugSolution.vbs` offers a corrected approach.

**Problem:**
The `IsEmpty` function returns `True` for an empty string (""), leading to potential logic errors when a function expects to differentiate between an uninitialized variable and an empty string.

**Solution:**
The solution uses the `Len` function to explicitly check the length of the string parameter before proceeding. This reliably distinguishes between an empty string and an uninitialized variant or other empty values.