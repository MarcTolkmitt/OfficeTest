# <u>Project: OfficeTest</u>



This demo program is made for my other project 'DbaseFrame'. 

I want to use the Interop-coms for office. Being practical small to write the needed script i found out about some problems in handling these com classes.

December 2024 they changed the debugger- he runs now in lazy mode. From the beginning i wasn't pleased about that, because an added 'using ...' can give you an enormous amount of errors. 

In this example you have the UI done in WPF. Trying to use the com-classes for office shows that Office must be scripted with a visual language from Microsoft. It starts with the UI's partial class 'MainWindow' derived from 'Window'. Both the Excel and the Word com classes have a 'Window' too, leading to an amount of errors. This problem made me drop the 'using ... '-statement for the interop classes. 

I even think you can't change that decision for any other projects because of the debugger being lazy. Better stay secure and working than ignoring possible mistakes.

On interest you should watch the project 'DbaseFrame'.

