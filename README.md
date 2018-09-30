# RSWordXML
Open source library for building word document. Suitable for .Net web apps, API's and web jobs. Built using .Net framework 4.6.1   

Welcome to the RSWordXML wiki!

Everything works as Stream here.

This DLL is used to build word documents asynchronously with your Application. Used .Net framework 4.6.1 to build this package. In future, I'll implement parallel processing using pre-processor directives and other elements.

With the latest build RSWordXML 1.0.5.1 in nuget,user can insert contents in between the existing word document.

Here are the steps to accomplish it,

Added the DLL from nuget using the above given URL.
invoke ExistingDocumentEdit function in class RSWord with parameters.
KeyText parameter should be the same which is given in your template document stream.
Example: if you template word document stream contains AddMyContent string as text in fourth page, then if "AddMyContent" passed as keyinput. The input paragraphs will be rendered in that location.
I'll come up with updates.
