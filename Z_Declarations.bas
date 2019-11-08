Attribute VB_Name = "Z_Declarations"
Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
Public prevColl(0 To 4) As Dictionary 'din moveloop
Public randnum As Integer 'din skelmove
Public tempAnim As Range 'din skelmove

'ce e mai jos (pana la spatiu) e din moveloop
'Public pozVerf As Range
Public tEmp As Range
Public isRight As Boolean
Public isLeft As Boolean
Public isDown As Boolean
Public isUp As Boolean
Public isBlocked As Boolean
Public prevnumX As Integer
Public prevnumY As Integer

Public skeletal1 As Range
Public skeletal2 As Range
Public skeletal3 As Range
Public skeletal4 As Range
Public skeletal5 As Range
Public skeletal6 As Range
Public skeletal7 As Range

Public fantomas As Range
Public fantomas2 As Range
Public fantomas3 As Range

Public anim1 As Range
'facut sa lase dara pe alta pagina
Public anim2 As Range

Public pahar1 As Range
Public pahar2 As Range
Public pahar3 As Range

Public para1 As Range
Public para2 As Range
Public para3 As Range

Public beer1 As Range
Public beer2 As Range
Public beer3 As Range

Public cola1 As Range
Public cola2 As Range
Public cola3 As Range

Public cig1 As Range
Public cig2 As Range
Public cig3 As Range

Public ochi1 As Range
Public ochi2 As Range

Public mobs(1 To 15) As Range

Public anim As Range

Public waitfortalk As Integer
Public waitformouth As Integer
Public waitforeyes As Integer
Public waitformove As Integer
Public noMobs As Integer
Public u As Integer
