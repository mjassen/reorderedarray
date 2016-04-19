'TheReorderedArray Function -- for an element of the array, specify the element's current(old) position and its desired(new) position. The function puts that element into the new position and all the other elements become appropriately reordered.

Function reorderarray(OldPos, NewPos, SourceArray)

Dim LengthOfOurArray
Dim DestinationArray()


LengthOfOurArray = UBound(SourceArray) 
   
redim DestinationArray(LengthOfOurArray)


DestinationArray(NewPos) = SourceArray(OldPos)
   
   
Dim SourceCursor
SourceCursor = 0
Dim DestinationCursor
DestinationCursor = 0

for LoopCursor=0 to LengthOfOurArray

If (SourceCursor <> OldPos) And (DestinationCursor <> NewPos) Then
DestinationArray(DestinationCursor) = SourceArray(SourceCursor)
End If

If SourceCursor <> OldPos Then
	DestinationCursor = DestinationCursor + 1
End If

If DestinationCursor <> NewPos Then
	SourceCursor = SourceCursor + 1
End If

next	   


reorderarray = DestinationArray


End Function









   



dim days(5)
   ' Referring elements by index
   days(0) = "mon"
   days(1) = "tue"
   days(2) = "wed"
   days(3) = "thu"
   days(4) = "fri"

   
   
   
   
   
   dim messagey
   
      for i=0 to 5-1
      messagey = messagey & "Element " & i+1 & " = " & days(i) & vbCrLf
   next   
   
   
   
   
'lets call TheReorderedArray -- tell it to put the 0th-position element into the 1st-position.
TheReorderedArray = reorderarray(0,1,days)





      
   messagey = messagey & vbCrLf & vbCrLf
	  
   for i=0 to 5-1
      messagey = messagey & "Element " & i+1 & " = " & TheReorderedArray(i) & vbCrLf
   next


   
MsgBox messagey




