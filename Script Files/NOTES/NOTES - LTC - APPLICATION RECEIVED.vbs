'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTES - LTC - APPLICATION RECEIVED.vbs"
start_time = timer

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN		'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF default_directory = "C:\DHS-MAXIS-Scripts\Script Files\" OR default_directory = "" THEN						
			'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		ELSEIF beta_agency = "" or beta_agency = True then							'If you're a beta agency, you should probably use the beta branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/BETA/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		Else																		'Everyone else should use the release branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/RELEASE/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		End if
		SET req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a FuncLib_URL
		req.open "GET", FuncLib_URL, FALSE							'Attempts to open the FuncLib_URL
		req.send													'Sends request
		IF req.Status = 200 THEN									'200 means great success
			Set fso = CreateObject("Scripting.FileSystemObject")	'Creates an FSO
			Execute req.responseText								'Executes the script code
		ELSE														'Error message, tells user to try to reach github.com, otherwise instructs to contact Veronica with details (and stops script).
			MsgBox 	"Something has gone wrong. The code stored on GitHub was not able to be reached." & vbCr &_ 
					vbCr & _
					"Before contacting Veronica Cary, please check to make sure you can load the main page at www.GitHub.com." & vbCr &_
					vbCr & _
					"If you can reach GitHub.com, but this script still does not work, ask an alpha user to contact Veronica Cary and provide the following information:" & vbCr &_
					vbTab & "- The name of the script you are running." & vbCr &_
					vbTab & "- Whether or not the script is ""erroring out"" for any other users." & vbCr &_
					vbTab & "- The name and email for an employee from your IT department," & vbCr & _
					vbTab & vbTab & "responsible for network issues." & vbCr &_
					vbTab & "- The URL indicated below (a screenshot should suffice)." & vbCr &_
					vbCr & _
					"Veronica will work with your IT department to try and solve this issue, if needed." & vbCr &_ 
					vbCr &_
					"URL: " & FuncLib_URL
					script_end_procedure("Script ended due to error connecting to GitHub.")
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

'DIALOGS----------------------------------------------------------------------------------------------------
BeginDialog case_number_dialog, 0, 0, 141, 80, "Case number dialog"
  EditBox 65, 10, 65, 15, case_number
  EditBox 65, 30, 30, 15, MAXIS_footer_month
  EditBox 100, 30, 30, 15, MAXIS_footer_year
  ButtonGroup ButtonPressed
    OkButton 25, 55, 50, 15
    CancelButton 80, 55, 50, 15
  Text 10, 30, 50, 15, "Footer month:"
  Text 10, 10, 50, 15, "Case number: "
EndDialog


BeginDialog LTC_app_recd_dialog, 0, 0, 286, 415, "LTC application received dialog"
  EditBox 75, 35, 65, 15, appl_date
  EditBox 75, 55, 65, 15, appl_type
  CheckBox 150, 45, 105, 10, "A transfer has been reported", transfer_reported_check
  CheckBox 150, 60, 140, 10, "Spousal allocation has been requested", spousal_allocation_check
  EditBox 160, 75, 120, 15, forms_needed
  EditBox 30, 95, 30, 15, CFR
  EditBox 110, 95, 170, 15, HH_comp
  EditBox 65, 115, 215, 15, pre_FACI_ADDR
  DropListBox 65, 135, 215, 15, "Select one..."+chr(9)+"Age 65 or older"+chr(9)+"Adult without children"+chr(9)+"Blind/disabled"+chr(9)+"Child under 21"+chr(9)+"Parent/Caretaker"+chr(9)+"Pregnant", basis_of_elig_droplist
  EditBox 35, 155, 245, 15, FACI
  EditBox 60, 175, 220, 15, retro_request
  EditBox 35, 195, 245, 15, AREP
  EditBox 60, 215, 220, 15, SWKR
  EditBox 60, 235, 220, 15, INSA
  EditBox 60, 255, 220, 15, adult_signatures
  EditBox 50, 275, 230, 15, veteran_info
  EditBox 50, 295, 230, 15, LTCC
  EditBox 55, 315, 225, 15, actions_taken
  CheckBox 5, 345, 220, 10, "Check here to have the script update PND2 to show client delay.", update_PND2_check
  CheckBox 5, 360, 280, 10, "Check here to have the script create a TIKL to deny at the 45 day mark (NON-DISA).", TIKL_45_day_check
  CheckBox 5, 375, 265, 10, "Check here to have the script create a TIKL to deny at the 60 day mark (DISA).", TIKL_60_day_check
  EditBox 90, 395, 80, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 175, 395, 50, 15
    CancelButton 230, 395, 50, 15
    PushButton 180, 25, 45, 10, "next panel", next_panel_button
    PushButton 230, 25, 45, 10, "next memb", next_memb_button
    PushButton 5, 160, 25, 10, "FACI:", FACI_button
    PushButton 5, 200, 25, 10, "AREP:", AREP_button
    PushButton 25, 220, 30, 10, "SWKR:", SWKR_button
    PushButton 5, 240, 25, 10, "INSA/", INSA_button
    PushButton 30, 240, 25, 10, "MEDI:", MEDI_button
    PushButton 15, 15, 25, 10, "TYPE", TYPE_button
    PushButton 40, 15, 25, 10, "PROG", PROG_button
    PushButton 65, 15, 25, 10, "HCRE", HCRE_button
    PushButton 90, 15, 25, 10, "REVW", REVW_button
    PushButton 115, 15, 25, 10, "MEMB", MEMB_button
  Text 5, 100, 20, 10, "CFR:"
  Text 70, 100, 40, 10, "HH Comp:"
  Text 5, 120, 60, 10, "Pre FACI address:"
  Text 5, 140, 60, 10, "Basis of eligibilty:"
  Text 5, 180, 55, 10, "Retro requested:"
  Text 5, 220, 20, 10, "PHN/"
  Text 5, 260, 55, 10, "Adult signatures:"
  Text 5, 300, 40, 10, "LTCC info:"
  Text 5, 320, 50, 10, "Actions taken:"
  Text 30, 400, 60, 10, "Worker signature:"
  Text 5, 40, 55, 10, "Application date:"
  Text 5, 80, 150, 10, "Forms needed? 1503, 3543, 3050, 5181, AA:"
  GroupBox 10, 5, 135, 25, "General STAT navigation:"
  GroupBox 175, 5, 105, 35, "STAT-based navigation"
  Text 5, 60, 65, 10, "Appl type received:"
  Text 5, 280, 45, 10, "Veteran info:"
  GroupBox 0, 335, 285, 55, "Actions"
  ButtonGroup ButtonPressed
    PushButton 230, 15, 45, 10, "prev. memb", prev_memb_button
    PushButton 180, 15, 45, 10, "prev. panel", prev_panel_button
EndDialog


'VARIABLES TO DECLARE----------------------------------------------------------------------------------------------------
HH_memb_row = 05


'THE SCRIPT----------------------------------------------------------------------------------------------------
'Connecting to BlueZone
EMConnect ""

'Searching for case number and footer month/year	
call MAXIS_case_number_finder(case_number)
Call MAXIS_footer_finder(MAXIS_footer_month, MAXIS_footer_year)

'Showing the case number dialog, transmits to check for MAXIS.
Do
  Dialog case_number_dialog
  cancel_confirmation
  If case_number = "" then MsgBox "You must type a case number."
Loop until case_number <> ""

'Now it checks to make sure MAXIS is running on this screen.
Call check_for_MAXIS(True)

'Navigating to STAT/HCRE so we can grab the app date
call navigate_to_MAXIS_screen("stat", "hcre")

'Creating a custom dialog for determining who the HH members are
call HH_member_custom_dialog(HH_member_array)

'Grabs autofill info from STAT
call autofill_editbox_from_MAXIS(HH_member_array, "HCRE-retro", retro_request)
call autofill_editbox_from_MAXIS(HH_member_array, "HCRE", appl_date)
call autofill_editbox_from_MAXIS(HH_member_array, "AREP", AREP)
call autofill_editbox_from_MAXIS(HH_member_array, "FACI", FACI)
call autofill_editbox_from_MAXIS(HH_member_array, "INSA", INSA)
call autofill_editbox_from_MAXIS(HH_member_array, "MEDI", MEDI)
call autofill_editbox_from_MAXIS(HH_member_array, "MEMB", HH_comp)
call autofill_editbox_from_MAXIS(HH_member_array, "SWKR", SWKR)

'Now, because INSA and MEDI will go on the same variable, we're going to add INSA to MEDI. To separate them in the case note, we have to add a semicolon (assuming both have data).
If INSA <> "" and MEDI <> "" then
  INSA = INSA & "; " & MEDI
Else
  INSA = INSA & MEDI
End if

'The main dialog
Do
	Do
		Do
			Do
				Do
					Dialog LTC_app_recd_dialog
					cancel_confirmation
					If buttonpressed <> -1 then Call MAXIS_dialog_navigation
				Loop until buttonpressed = -1 or buttonpressed = 0
				If actions_taken = "" then MsgBox "You must fill in the actions taken section. Please try again."
			Loop until actions_taken <> ""
			If worker_signature = "" then MsgBox "You must sign your case note!"
		Loop until worker_signature <> ""
		If basis_of_elig_droplist = "Select one..." THEN MsgBox "You must select the client's MA basis of eligibility."
	Loop until basis_of_elig_droplist <> "Select one..."
	IF TIKL_45_day_check = 1 and TIKL_60_day_check = 1 then MsgBox "You must choose to TIKL out for 45 or 60 days, not both."
LOOP until (TIKL_45_day_check = 1 AND TIKL_60_day_check = 0) or (TIKL_45_day_check = 0 AND TIKL_60_day_check = 0) OR (TIKL_45_day_check = 0 AND TIKL_60_day_check = 1)

'UPDATING PND2----------------------------------------------------------------------------------------------------
If update_PND2_check = 1 then 
	call navigate_to_screen("rept", "pnd2")
	EMGetCursor PND2_row, PND2_col
	EMReadScreen PND2_SNAP_status_check, 1, PND2_row, 62
	If PND2_SNAP_status_check = "P" then EMWriteScreen "C", PND2_row, 62
	EMReadScreen PND2_HC_status_check, 1, PND2_row, 65
	If PND2_HC_status_check = "P" then
		EMWriteScreen "x", PND2_row, 3
		transmit
		person_delay_row = 7
		Do
			EMReadScreen person_delay_check, 1, person_delay_row, 39
			If person_delay_check <> " " then EMWriteScreen "c", person_delay_row, 39
			person_delay_row = person_delay_row + 2
		Loop until person_delay_check = " " or person_delay_row > 20
		PF3
	End if
	PF3
	EMReadScreen PND2_check, 4, 2, 52
	If PND2_check = "PND2" then
		MsgBox "PND2 might not have been updated for client delay. There may have been a MAXIS error. Check this manually after case noting."
		PF10
		client_delay_check = 0
	End if
End if

'THE TIKL's----------------------------------------------------------------------------------------------------
If TIKL_45_day_check = 1 then
	call navigate_to_MAXIS_screen("dail", "writ")
	call create_MAXIS_friendly_date(appl_date, 45, 5, 18) 
	EMSetCursor 9, 3
	Call write_variable_in_TIKL("HC pending 45 days. Evaluate for possible denial. If any members are elderly/disabled, allow an additional 15 days and reTIKL out.")
	transmit
	PF3
End if

If TIKL_60_day_check = 1 then
	call navigate_to_MAXIS_screen("dail", "writ")
	call create_MAXIS_friendly_date(appl_date, 60, 5, 18) 
	EMSetCursor 9, 3
	Call write_variable_in_TIKL("HC pending 60 days. Evaluate for possible denial. If any members are elderly/disabled, allow an additional 15 days and reTIKL out.")
	transmit
	PF3
End if


'THE CASE NOTE----------------------------------------------------------------------------------------------------
call start_a_blank_CASE_NOTE
'Writing the case note
Call write_variable_in_CASE_NOTE("***LTC intake***")
If appl_date <> "" then call write_bullet_and_variable_in_CASE_NOTE("Application date", appl_date)
If appl_type <> "" then call write_bullet_and_variable_in_CASE_NOTE("Application type received", appl_type)
If forms_needed <> "" then call write_bullet_and_variable_in_CASE_NOTE("Forms Needed", forms_needed)
If HH_comp <> "" then call write_bullet_and_variable_in_CASE_NOTE("HH comp", HH_comp)
If CFR <> "" then call write_bullet_and_variable_in_CASE_NOTE("CFR", CFR)
If pre_FACI_ADDR <> "" then call write_bullet_and_variable_in_CASE_NOTE("Pre FACI address", pre_FACI_ADDR)
If basis_of_elig_droplist <> "" then call write_bullet_and_variable_in_CASE_NOTE("Basis of eligibility", basis_of_elig_droplist)
If FACI <> "" then call write_bullet_and_variable_in_CASE_NOTE("FACI", FACI)
If retro_request <> "" then call write_bullet_and_variable_in_CASE_NOTE("Retro request", retro_request)
If AREP <> "" then call write_bullet_and_variable_in_CASE_NOTE("AREP", AREP)
If SWKR <> "" then call write_bullet_and_variable_in_CASE_NOTE("PHN/SWKR", SWKR)
If INSA <> "" then call write_bullet_and_variable_in_CASE_NOTE("INSA/MEDI", INSA)
If adult_signatures <> "" then call write_bullet_and_variable_in_CASE_NOTE("Adult signatures", adult_signatures)
If LTCC <> "" then call write_bullet_and_variable_in_CASE_NOTE("LTCC info", LTCC)
IF veteran_info <> "" then call write_bullet_and_variable_in_CASE_NOTE("Veteran information", veteran_info)
call write_bullet_and_variable_in_CASE_NOTE("Actions taken", actions_taken)
If transfer_reported_check = 1 THEN call write_variable_in_CASE_NOTE("* A transfer has been reported.")
IF spousal_allocation_check = 1 THEN Call write_variable_in_CASE_NOTE("* Spousal allocation has been requested.")
If update_PND2_check = 1 THEN Call write_variable_in_CASE_NOTE("* PND2 updated to show client delay.")
IF TIKL_45_day_check = 1 Then call write_variable_in_CASE_NOTE("* Set TIKL for 45 days to recheck case.")
IF TIKL_60_day_check = 1 Then call write_variable_in_CASE_NOTE("* Set TIKL for 60 days to recheck case.")
call write_variable_in_CASE_NOTE("---")
call write_variable_in_CASE_NOTE(worker_signature)

script_end_procedure("")