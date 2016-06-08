'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "ACTIONS - MAIN MENU.vbs"
start_time = timer 

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN	   'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF use_master_branch = TRUE THEN			   'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		Else											'Everyone else should use the release branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/RELEASE/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		End if
		SET req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a FuncLib_URL
		req.open "GET", FuncLib_URL, FALSE							'Attempts to open the FuncLib_URL
		req.send													'Sends request
		IF req.Status = 200 THEN									'200 means great success
			Set fso = CreateObject("Scripting.FileSystemObject")	'Creates an FSO
			Execute req.responseText								'Executes the script code
		ELSE														'Error message
			critical_error_msgbox = MsgBox ("Something has gone wrong. The Functions Library code stored on GitHub was not able to be reached." & vbNewLine & vbNewLine &_
                                            "FuncLib URL: " & FuncLib_URL & vbNewLine & vbNewLine &_
                                            "The script has stopped. Please check your Internet connection. Consult a scripts administrator with any questions.", _
                                            vbOKonly + vbCritical, "BlueZone Scripts Critical Error")
            StopScript
		END IF
	ELSE
		FuncLib_URL = "C:\BZS-FuncLib\MASTER FUNCTIONS LIBRARY.vbs"
		Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
		Set fso_command = run_another_script_fso.OpenTextFile(FuncLib_URL)
		text_from_the_other_script = fso_command.ReadAll
		fso_command.Close
		Execute text_from_the_other_script
	END IF
END IF
'END FUNCTIONS LIBRARY BLOCK================================================================================================

'LOADING LIST OF SCRIPTS FROM GITHUB REPOSITORY===========================================================================
'IF run_locally = FALSE or run_locally = "" THEN	   'If the scripts are set to run locally, it skips this and uses an FSO below.
''	IF use_master_branch = TRUE THEN			   'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
''		script_list_URL = "https://raw.githubusercontent.com/MN-Script-Team/DHS-MAXIS-Scripts/master/COMPLETE%20LIST%20OF%20SCRIPTS.vbs"
''	Else											'Everyone else should use the release branch.
''		script_list_URL = "https://raw.githubusercontent.com/MN-Script-Team/DHS-MAXIS-Scripts/RELEASE/COMPLETE%20LIST%20OF%20SCRIPTS.vbs"
''	End if
'
'	'<<<<TEMPORARY JUST FOR TESTING
'	script_list_URL = "https://raw.githubusercontent.com/MN-Script-Team/DHS-MAXIS-Scripts/VKC-list-of-scripts-vbs-file/COMPLETE%20LIST%20OF%20SCRIPTS.vbs"
'	
'	SET req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a script_list_URL
'	req.open "GET", script_list_URL, FALSE							'Attempts to open the script_list_URL
'	req.send													'Sends request
'	IF req.Status = 200 THEN									'200 means great success
'		Set fso = CreateObject("Scripting.FileSystemObject")	'Creates an FSO
'		Execute req.responseText								'Executes the script code
'	ELSE														'Error message
'		critical_error_msgbox = MsgBox ("Something has gone wrong. The script list code stored on GitHub was not able to be reached." & vbNewLine & vbNewLine &_
'                                        "Script list URL: " & script_list_URL & vbNewLine & vbNewLine &_
'                                        "The script has stopped. Please check your Internet connection. Consult a scripts administrator with any questions.", _
'                                        vbOKonly + vbCritical, "BlueZone Scripts Critical Error")
'        StopScript
'	END IF
'ELSE
	script_list_URL = "C:\DHS-MAXIS-Scripts\COMPLETE LIST OF SCRIPTS.vbs"
	Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
	Set fso_command = run_another_script_fso.OpenTextFile(script_list_URL)
	text_from_the_other_script = fso_command.ReadAll
	fso_command.Close
	Execute text_from_the_other_script
'END IF
'END FUNCTIONS LIBRARY BLOCK================================================================================================

class subcat
	public subcat_name
	public subcat_button
End class

'CUSTOM FUNCTIONS===========================================================================================================
Function declare_main_menu_dialog(script_category)

	'Runs through each script in the array and generates a list of subcategories based on the category located in the function
	For current_script = 0 to ubound(script_array)
		If ucase(script_array(current_script).category) = ucase(script_category) then																								'If the script in the array is of the correct category (ACTIONS/NOTES/ETC)...
			For each listed_subcategory in script_array(current_script).subcategory																									'...then iterate through each listed subcategory, and...
				If listed_subcategory <> "" and InStr(subcategory_list, ucase(listed_subcategory)) = 0 then subcategory_list = subcategory_list & "|" & ucase(listed_subcategory)	'...if the listed subcategory isn't blank and isn't already in the list, then add it to our handy-dandy list.	
			Next
		End if
	Next
	
	subcategory_list = split(subcategory_list, "|")
	
	For i = 0 to ubound(subcategory_list)
		ReDim Preserve subcategory_array(i)
		set subcategory_array(i) = new subcat
		If subcategory_list(i) = "" then subcategory_list(i) = "MAIN"
		subcategory_array(i).subcat_name = subcategory_list(i)
	Next
	
	BeginDialog dialog1, 0, 0, 600, 340, script_category & " scripts main menu dialog"
	 	Text 5, 5, 435, 10, script_category & " scripts main menu: select the script to run from the choices below."
	  	ButtonGroup ButtonPressed
		
		subcat_button_position = 5
		
		For i = 0 to ubound(subcategory_array)
		
			
		
			'Displays the button and text description-----------------------------------------------------------------------------------------------------------------------------
			'FUNCTION		HORIZ. ITEM POSITION	VERT. ITEM POSITION		ITEM WIDTH	ITEM HEIGHT		ITEM TEXT/LABEL										BUTTON VARIABLE
			PushButton 		subcat_button_position, 20, 					50, 		15, 			subcategory_array(i).subcat_name, 					subcat_button_placeholder
			
			subcategory_array(i).subcat_button = subcat_button_placeholder	'The .button property won't carry through the function. This allows it to escape the function. Thanks VBScript.
			subcat_button_position = subcat_button_position + 50
			subcat_button_placeholder = subcat_button_placeholder + 1
		Next
		
		'' 	PushButton 445, 10, 65, 10, "SIR instructions", 	SIR_instructions_button
		'This starts here, but it shouldn't end here :)
		vert_button_position = 50

		For current_script = 0 to ubound(script_array)
			If ucase(script_array(current_script).category) = ucase(script_category) then
			
				'<<<<<<RIGHT HERE IT SHOULD ITERATE THROUGH SUBCATEGORIES AND BUTTONS PRESSED TO DETERMINE WHAT THE CURRENTLY DISPLAYED SUBCATEGORY SHOULD BE, THEN ONLY DISPLAY SCRIPTS THAT MATCH THAT CRITERIA
			
				SIR_button_placeholder = button_placeholder + 1	'We always want this to be one more than the button_placeholder
			
				'Displays the button and text description-----------------------------------------------------------------------------------------------------------------------------
				'FUNCTION		HORIZ. ITEM POSITION	VERT. ITEM POSITION		ITEM WIDTH	ITEM HEIGHT		ITEM TEXT/LABEL										BUTTON VARIABLE
				PushButton 		5, 						vert_button_position, 	10, 		10, 			"?", 												SIR_button_placeholder
				PushButton 		18,						vert_button_position, 	120, 		10, 			script_array(current_script).script_name, 			button_placeholder
				Text 			120 + 23, 				vert_button_position, 	500, 		10, 			"--- " & script_array(current_script).description
				'----------
				vert_button_position = vert_button_position + 15	'Needs to increment the vert_button_position by 15px (used by both the text and buttons)
				'----------
				script_array(current_script).button = button_placeholder	'The .button property won't carry through the function. This allows it to escape the function. Thanks VBScript.
				script_array(current_script).SIR_instructions_button = SIR_button_placeholder	'The .button property won't carry through the function. This allows it to escape the function. Thanks VBScript.
				button_placeholder = button_placeholder + 2
			End if
		next

		CancelButton 540, 320, 50, 15
	EndDialog
End function
'END CUSTOM FUNCTIONS=======================================================================================================

'Starting these with a very high number, higher than the normal possible amount of buttons.
'	We're doing this because we want to assign a value to each button pressed, and we want
'	that value to change with each button. The button_placeholder will be placed in the .button
'	property for each script item. This allows it to both escape the Function and resize
'	near infinitely. We use dummy numbers for the other selector buttons for much the same reason,
'	to force the value of ButtonPressed to hold in near infinite iterations.
button_placeholder 			= 24601
subcat_button_placeholder 	= 1701

'Displays the dialog
Do
	If ButtonPressed = "" then 
		call declare_main_menu_dialog("Actions")
		ready_to_exit_loop = false
	End if

	dialog

	If ButtonPressed = 0 then stopscript
	MsgBox ButtonPressed 

	'Runs through each script in the array... if the selected script (buttonpressed) is in the array, it'll run_from_GitHub
	For i = 0 to ubound(script_array)
		If ButtonPressed = script_array(i).SIR_instructions_button then call open_URL_in_browser(script_array(i).SIR_instructions_URL)
	Next    

	'Runs through each script in the array... if the selected script (buttonpressed) is in the array, it'll run_from_GitHub
	For i = 0 to ubound(script_array)
		If ButtonPressed = script_array(i).button then 
			ready_to_exit_loop = true		'Doing this just in case a stopscript or script_end_procedure is missing from the script in question
			call run_from_GitHub(script_array(i).script_URL)
		End if
	Next    
	

Loop until ready_to_exit_loop = true



stopscript
