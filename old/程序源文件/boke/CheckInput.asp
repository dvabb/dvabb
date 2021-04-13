<%
Function CheckText(str)
	Dim Chk
	Dim Re,TempStr
	Chk = False
	If Not IsNull(Str) Then
		If Instr(Str,"=")>0 or Instr(Str,"%")>0 or Instr(Str,"?")>0 or Instr(Str,"&")>0 or Instr(Str,";")>0 or Instr(Str,",")>0 or Instr(Str,"'")>0 or Instr(Str,",")>0 or Instr(Str,chr(34))>0 or Instr(Str,chr(9))>0 or Instr(Str,"$")>0 or Instr(Str,"|")>0 Then
			Chk = False
		Else
			Chk = True
		End If
		Set Re=New RegExp
			Re.IgnoreCase =True
			Re.Global=True
			Re.Pattern="(\s)"
			TempStr = Re.Replace(Str,"")
			TempStr = Replace(TempStr,chr(32),"")
			TempStr = Replace(TempStr,"