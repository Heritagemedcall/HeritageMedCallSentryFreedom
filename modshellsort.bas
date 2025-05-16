Attribute VB_Name = "modShellSort"
Option Explicit

Option Base 1

Public Function ShellSortElapsedTime(c As Collection, ByVal Descending As Boolean) As Collection

        Dim Hold               As Long  ' specific to sort
        Dim Gap                As Long  ' specific to sort
        Dim i                  As Long  ' specific to sort
        Dim Min                As Long  ' always 1
        Dim Max                As Long  ' same as number of objects in collection
        Dim Compare            As Long  ' the current value
        Dim IndexValue         As Long  ' the current index
        Dim TestValue          As Long  ' to test against
        Dim j                  As Long
        Dim sx()               As Long  ' array for sorting of indexes

        Dim hist               As cAlarmHistory

        Dim newcoll            As Collection  ' the new collection

        Dim DoEventsCounter    As Long

10      Set newcoll = New Collection

20      ReDim sx(c.Count, 2) As Long  ' we need a count number of elements with 2 columns

30      For j = 1 To c.Count
40        Set hist = c(j)
50        sx(j, 1) = hist.ElapsedTime
60        sx(j, 2) = j
70      Next

80      Min = 1
90      Max = c.Count
100     Gap = Min

110     Do                           ' figureb optimum gap
120       Gap = 3 * Gap + 1
130     Loop Until Gap > Max

        Dim t
140     t = Timer

150     Do
160       Gap = Gap \ 3
170       For i = Gap + Min To Max
180         If DoEventsCounter > 200 Then
190           DoEventsCounter = 0
200           DoEvents
210         End If
220         DoEventsCounter = DoEventsCounter + 1
230         Compare = sx(i, 1)
240         IndexValue = sx(i, 2)
250         Hold = i

260         If Descending Then
270           TestValue = sx(Hold - Gap, 1)
280           Do While TestValue < Compare
290             If DoEventsCounter > 200 Then
300               DoEventsCounter = 0
310               DoEvents
320             End If
330             DoEventsCounter = DoEventsCounter + 1
                ' swap the value and the index
340             sx(Hold, 1) = sx(Hold - Gap, 1)  ' swap real values
350             sx(Hold, 2) = sx(Hold - Gap, 2)  ' swap indexes
360             Hold = Hold - Gap
370             If Hold < Min + Gap Then
380               Exit Do
390             End If
400             TestValue = sx(Hold - Gap, 1)
410           Loop
420         Else  ' ascending
430           TestValue = sx(Hold - Gap, 1)
440           Do While TestValue > Compare
450             If DoEventsCounter > 200 Then
460               DoEventsCounter = 0
470               DoEvents
480             End If
490             DoEventsCounter = DoEventsCounter + 1
                ' swap the value and the index
500             sx(Hold, 1) = sx(Hold - Gap, 1)
510             sx(Hold, 2) = sx(Hold - Gap, 2)
520             Hold = Hold - Gap
530             If Hold < Min + Gap Then
540               Exit Do
550             End If
560             TestValue = sx(Hold - Gap, 1)
570           Loop
580         End If

590         sx(Hold, 1) = Compare
600         sx(Hold, 2) = IndexValue
610       Next i

620     Loop Until Gap = 1
        'Debug.Print "time " & Timer - t
630     For j = 1 To Max
          ' Debug.Print sx(j, 1), sx(j, 2)
640       newcoll.Add c(sx(j, 2))
650     Next
660     Set c = Nothing
670     Set ShellSortElapsedTime = newcoll


End Function

