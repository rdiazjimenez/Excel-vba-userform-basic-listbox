# excel-vba-userform-basic-listbox-demo
This is a simple demo on how to perform basic operations like Create, Read, Update and Delete information stored in an Excel table with a userform, a listbox and some command buttons.


To use it just download the xlsm file to your computer, enable macros, and click the "Show UserForm" button.

To check the code, press ALT + F11

File contains:
oUserForm = The actual userform
mCode = Procedures that perform basic operations (Create, Read, Update and Delete) and a couple of routines to load and clear the textboxes inside the userform.

Code is commented and descriptive in order to make it easy to understand.

Some basic programming stuff missing (on purpose in order to keep the demo simple):
- Error handling
- Code and procedures optimization / simplification.


Here are a couple of tips to have in mind when you design a solution that includes loading Excel data into a Listbox in a UserForm:

1) Try to store the Excel information into an structured Excel Table (Visit this link to learn more: https://support.office.com/en-ie/article/create-and-format-tables-e81aa349-b006-4f8a-9806-5af9df0ac664)

2) Use Option Explicit at the top of each module in VBA so you have more control of the variables you use and create (Visit this link to learn more: https://docs.microsoft.com/en-us/dotnet/visual-basic/language-reference/statements/option-explicit-statement)

3) Use descriptive names for controls inside a UserForm (Visit this link to learn more: https://rtmccormick.com/2015/11/23/vba-control-naming-conventions/)
