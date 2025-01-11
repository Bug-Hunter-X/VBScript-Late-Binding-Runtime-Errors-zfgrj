# VBScript Late Binding Runtime Errors

This repository demonstrates a common error in VBScript programming: runtime errors caused by late binding.  Late binding, while flexible, can lead to unexpected crashes if the objects or methods being accessed don't exist at runtime.  The example shows how to mitigate this by properly handling potential errors.

## Bug Description

VBScript's late binding allows you to work with objects without explicit type declarations.  This flexibility comes at the cost of runtime errors if the object or a specific method isn't available. This is often encountered when working with COM objects like Excel, where the application might not be installed, or when accessing external libraries.

## Solution

The solution addresses this by incorporating error handling using the `Err` object.  The code checks the `Err.Number` property after attempting to create or access the object. If an error has occurred, the program gracefully handles it instead of crashing.