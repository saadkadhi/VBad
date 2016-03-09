Function [rdm::15]Test_Function(ByVal [rdm::10]FileToTest As String) As Boolean
   FileExists = (Dir(FileToTest) <> "")
End Function

Public Function [rdm::15]Test()
	[rdm::8]kikou = "AZERTY1234"
	[rdm::15]test3 = "HELLO_WORLD"
	[rdm::5]Excluded_string = "ExcludedString[!!]"
	MsgBox test3
	[rdm::12]domain_list = [var::domain_name]
	[rdm::4]path = [var::path_to_save] 
	msgbox domain_list
	msgbox path
	
End Function


