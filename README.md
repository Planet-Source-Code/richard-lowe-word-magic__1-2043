<div align="center">

## Word Magic


</div>

### Description

This program allows simple desktop access the the Microsoft Word spelling and thesaurus engine using OLE Automation.

You can Spell Check, Produce Anangrams, use the Thesaurus and look up the meaning of words. THIS IS A COMPLETE WORKING APPLICATION
 
### More Info
 
Must have Microsoft Word Installed


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Richard Lowe](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/richard-lowe.md)
**Level**          |Unknown
**User Rating**    |5.0 (30 globes from 6 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Complete Applications](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/complete-applications__1-27.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/richard-lowe-word-magic__1-2043/archive/master.zip)





### Source Code

```
'===========================================================================
'Start a new project
'add a ComboBox named cboInput
'add a ListBox named lstDisplay
'add a Command Button named cmdHelp caption Help
'add a Command Button named cmdExit caption Exit
'add 4 Command Buttons (command array) named
'cmdAction(0)	caption Spelling
'cmdAction(1)	caption Wildcard
'cmdAction(2)	caption Anagarm
'cmdAction(3)	Caption Lookup
'In the Project/References menu option tick the reference for
'Microsoft Word 8.0 Object Library
'===========================================================================
'paste the following code
Option Explicit
'============================================================
'== Author : Richard Lowe
'== Date : June 99
'== Contact : riklowe@hotmail.com
'============================================================
'== Desciption
'==
'== This program enable quick and easy desktop access to
'== the Microsoft Word spelling and thesaurus engine.
'==
'============================================================
'== Version History
'============================================================
'== 1.0 06-Jun-99 RL Initial Release. Spelling Only
'== 1.1 07-Jun-99 RL Added Widcard, Anagram and Lookup
'== 1.2 08-Jun-99 RL Added Help
'============================================================
'------------------------------------------------------------
'Define constants
'------------------------------------------------------------
Const HeightLimit = 5000
Const WidthLimit = 5640
'------------------------------------------------------------
'Dimension variables
'------------------------------------------------------------
Dim objMsWord As Word.Application
Dim SugList As SpellingSuggestions
Dim sug As SpellingSuggestion
Dim synInfo As SynonymInfo
Dim synList As Variant
Dim AntList As Variant
Private Sub cmdAction_Click(Index As Integer)
'------------------------------------------------------------
' dimension local variables
'------------------------------------------------------------
Dim strTemp As String
Dim blnRet As Boolean
Dim iCount As Integer
'------------------------------------------------------------
' Asign an error handler
'------------------------------------------------------------
On Error GoTo eh_Trap:
'------------------------------------------------------------
' If cboInput has changed, add it as an entry to the list
'------------------------------------------------------------
 If cboInput.List(0) <> cboInput Then
  cboInput.AddItem cboInput, 0
 End If
'------------------------------------------------------------
'Assign the objMsWord object reference to the Word application
'------------------------------------------------------------
 Set objMsWord = New Word.Application
'------------------------------------------------------------
'Due to a bug, you have to open a file to use GetSpellingSuggestions
'This is documented in Q169545 on microsoft knowledge base
'------------------------------------------------------------
 objMsWord.WordBasic.FileNew  'open a doc
 objMsWord.Visible = False  'hide the doc
'------------------------------------------------------------
' clear display area
'------------------------------------------------------------
 lstDisplay.Clear
'------------------------------------------------------------
' select which button has been pressed
'------------------------------------------------------------
 Select Case Index
 Case 0
'------------------------------------------------------------
'Spelling
'------------------------------------------------------------
  blnRet = objMsWord.CheckSpelling(cboInput)
'------------------------------------------------------------
'if incorrectly spelt, check for suggestions. Iterate and display
'------------------------------------------------------------
  If blnRet = True Then
   lstDisplay.AddItem "OK"
  Else
   Set SugList = objMsWord.GetSpellingSuggestions(cboInput, _
   SuggestionMode:=wdSpelling)
   If SugList.Count = 0 Then
    lstDisplay.AddItem "No suggestions"
   Else
    For Each sug In SugList
     lstDisplay.AddItem sug.Name
    Next sug
   End If
  End If
 Case 1
'------------------------------------------------------------
'WildCard
'------------------------------------------------------------
  Set SugList = objMsWord.Application.GetSpellingSuggestions(cboInput, _
  SuggestionMode:=wdWildcard)
'------------------------------------------------------------
'If entries found, Iterate and display
'------------------------------------------------------------
  If SugList.Count = 0 Then
   lstDisplay.AddItem "No suggestions"
  Else
   For Each sug In SugList
    lstDisplay.AddItem sug.Name
   Next sug
  End If
 Case 2
'------------------------------------------------------------
'Anagram
'------------------------------------------------------------
  Set SugList = objMsWord.GetSpellingSuggestions(cboInput, _
  SuggestionMode:=wdAnagram)
'------------------------------------------------------------
'If entries found, Iterate and display
'------------------------------------------------------------
  If SugList.Count = 0 Then
   lstDisplay.AddItem "No suggestions"
  Else
   For Each sug In SugList
    lstDisplay.AddItem sug.Name
   Next sug
  End If
 Case 3
'------------------------------------------------------------
'Lookup
'------------------------------------------------------------
'------------------------------------------------------------
'Assign the synInfo object reference to the Word Synonym Information
'------------------------------------------------------------
  Set synInfo = objMsWord.SynonymInfo(cboInput)
  lstDisplay.AddItem "--- MEANING ---"
'------------------------------------------------------------
'If entries found, Iterate and display
'------------------------------------------------------------
  If synInfo.MeaningCount >= 2 Then
   synList = synInfo.MeaningList
   For iCount = 1 To UBound(synList)
    lstDisplay.AddItem synList(iCount)
   Next iCount
  Else
   lstDisplay.AddItem "None"
  End If
  lstDisplay.AddItem "--- SYNONYM ---"
'------------------------------------------------------------
'If entries found, Iterate and display
'------------------------------------------------------------
  If synInfo.MeaningCount >= 2 Then
   synList = synInfo.SynonymList(2)
   For iCount = 1 To UBound(synList)
    lstDisplay.AddItem synList(iCount)
   Next iCount
  Else
   lstDisplay.AddItem "None"
  End If
  Set synInfo = Nothing
 End Select
'------------------------------------------------------------
'Clean exit point
'------------------------------------------------------------
eh_exit:
 objMsWord.Quit
 Set objMsWord = Nothing
 cboInput.SetFocus
Exit Sub
'------------------------------------------------------------
'Error Handler
'------------------------------------------------------------
eh_Trap:
 lstDisplay.AddItem Err & vbTab & Error$
 Resume eh_exit:
End Sub
Private Sub cmdExit_Click()
 Unload Me
End Sub
Private Sub cmdHelp_Click()
'------------------------------------------------------------
'Display help information in the viewing area
'------------------------------------------------------------
 lstDisplay.Clear
 lstDisplay.AddItem "Spelling "
 lstDisplay.AddItem "Enter a word into the box above, press 'Spelling'"
 lstDisplay.AddItem "Correctly spelt words will display 'OK'"
 lstDisplay.AddItem "Incorrectly spelt words will display a list of "
 lstDisplay.AddItem "choices that most closely match the word"
 lstDisplay.AddItem " "
 lstDisplay.AddItem "Wildcard "
 lstDisplay.AddItem "Enter a word into the box above, press 'Wildcard'"
 lstDisplay.AddItem "Use a ? to indicate an unkown letter"
 lstDisplay.AddItem "Use a * to indicate muliple unkown letters"
 lstDisplay.AddItem "Examples (?) - Cl?se, Un?no?n "
 lstDisplay.AddItem "Examples (*) - Cl*, C*e"
 lstDisplay.AddItem " "
 lstDisplay.AddItem "Anangram "
 lstDisplay.AddItem "Enter a word into the box above, press 'Anagram'"
 lstDisplay.AddItem "The program will find all words in the "
 lstDisplay.AddItem "dictionary containing those letters "
 lstDisplay.AddItem " "
 lstDisplay.AddItem "Lookup "
 lstDisplay.AddItem "Enter a word into the box above, press 'Lookup'"
 lstDisplay.AddItem "The program will find the meaning and synonym "
 lstDisplay.AddItem "for the word from the dictionary "
 lstDisplay.AddItem " "
 lstDisplay.AddItem "General "
 lstDisplay.AddItem "Double click on an entry in this list box"
 lstDisplay.AddItem "and it will be transfered to the box above."
 lstDisplay.AddItem "Use the up and down arrows on the keyboard "
 lstDisplay.AddItem "or select the arrow at the right hand side "
 lstDisplay.AddItem "of the above box, to scroll through all of "
 lstDisplay.AddItem "the word you have entered."
 lstDisplay.AddItem ""
 lstDisplay.AddItem "Please e-mail any comments / suggestions to"
 lstDisplay.AddItem "me - It's great to get feedback."
 lstDisplay.AddItem "My e-mail address is riklowe@hotmail.com"
 lstDisplay.AddItem ""
End Sub
Private Sub Form_Load()
 cboInput.Clear
End Sub
Private Sub Form_Resize()
'------------------------------------------------------------
'Do not let the screen size get to small, so that the button
'are always visible
'------------------------------------------------------------
 Select Case Me.WindowState
 Case vbNormal
  If Me.Height < HeightLimit Then
   Me.Height = HeightLimit
  End If
  lstDisplay.Height = Me.Height - 1000
  Me.Width = WidthLimit
 Case Else
 End Select
End Sub
Private Sub lstDisplay_DblClick()
'------------------------------------------------------------
'Move entry from listbox into combo box
'------------------------------------------------------------
 cboInput.AddItem lstDisplay, 0
 cboInput.ListIndex = 0
 lstDisplay.Clear
 cboInput.SetFocus
End Sub
```

