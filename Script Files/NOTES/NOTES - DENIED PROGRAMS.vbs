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

'THE SCRIPT----------------------------------------------------------------------------------------------------
'SCRIPT CONNECTS, THEN FINDS THE CASE NUMBER
EMConnect ""
Call MAXIS_case_number_finder(MAXIS_case_number)

Call check_for_MAXIS(True)

'THE DIALOG----------------------------------------------------------------------------------------------------
'This dialog uses a dialog_shrink_amt variable, along with an if...then which is decided by the global variable case_noting_intake_dates.
BeginDialog denied_dialog, 0, 0, 401, 385 - dialog_shrink_amt, "Denied progs dialog"
  EditBox 65, 5, 55, 15, MAXIS_case_number
  EditBox 185, 5, 55, 15, application_date
  CheckBox 60, 25, 35, 10, "SNAP", SNAP_check
  CheckBox 145, 25, 25, 10, "HC", HC_check
  CheckBox 230, 25, 35, 10, "Cash", cash_check
  CheckBox 315, 25, 40, 10, "Emer", emer_check
  EditBox 60, 40, 55, 15, SNAP_denial_date
  EditBox 145, 40, 55, 15, HC_denial_date
  EditBox 230, 40, 55, 15, cash_denial_date
  EditBox 315, 40, 55, 15, emer_denial_date
  CheckBox 60, 60, 60, 10, "Missing Verifs", missing_verifs_SNAP_checkbox
  CheckBox 145, 60, 60, 10, "Missings Verifs", missing_verifs_HC_checkbox
  CheckBox 230, 60, 60, 10, "Missing Verifs", missing_verifs_CASH_checkbox
  CheckBox 315, 60, 60, 10, "Missing Verifs", missing_verifs_EMER_checkbox
  CheckBox 60, 75, 65, 10, "Denied on Pnd2", denied_pnd2_SNAP_checkbox
  CheckBox 230, 75, 65, 10, "Denied on Pnd2", denied_pnd2_CASH_checkbox
  CheckBox 315, 75, 65, 10, "Denied on Pnd2", denied_pnd2_EMER_checkbox
  CheckBox 60, 90, 75, 10, "Withdrawn on Pnd2", withdraw_pnd2_SNAP_checkbox
  CheckBox 145, 90, 75, 10, "Withdrawn on Pact", withdraw_pact_HC_checkbox
  CheckBox 230, 90, 75, 10, "Withdrawn on Pnd2", withdraw_pnd2_CASH_checkbox
  CheckBox 315, 90, 75, 10, "Withdrawn on Pnd2", withdraw_pnd2_EMER_checkbox
  EditBox 65, 105, 330, 15, reason_for_denial
  EditBox 140, 125, 255, 15, verifs_needed
  Text 30, 145, 350, 25, "Check here to have the script add the verifs needed to denial notices. This will list the contents of the above box on the client denial notice. List each of the specific mandatory verifications that were used for the denial."
  CheckBox 15, 140, 10, 25, "", edit_notice_check
  EditBox 50, 170, 345, 15, other_notes
  'If case_noting_intake_dates = True then
    CheckBox 15, 200, 360, 10, "Check here if requested proofs were not provided, interview was completed (if applicable) and this case pended", requested_proofs_not_provided_check
    CheckBox 15, 225, 365, 10, "Denied SNAP for self-declaration of income over 165% FPG (hold for 30 days, with an add'l 30 for proration)", self_declaration_of_income_over_165_FPG
    CheckBox 15, 245, 130, 10, "Client is disabled (60 day HC period)", disabled_client_check
    CheckBox 15, 260, 305, 10, "Check here if there are any programs still open/pending (doesn't become intake again yet)", open_prog_check
    EditBox 105, 275, 235, 15, open_progs
    CheckBox 15, 290, 330, 10, "Check here if there are any HH members still open on HC (won't require a HCAPP to add a member)", HH_membs_on_HC_check
    EditBox 105, 305, 235, 15, HH_membs_on_HC
    GroupBox 5, 190, 390, 140, "Important items that affect the intake date/documentation:"
    Text 40, 210, 300, 10, " the full 30 day period (or 45/60 days for HC). Applies a 30 day reinstate period."
    Text 35, 275, 70, 10, "If so, list them here:"
    Text 35, 310, 70, 10, "If so, list them here:"
  'Else
    'EditBox 165, 190, 200, 15, open_progs
    'EditBox 190, 210, 200, 15, HH_membs_on_HC
    'Text 5, 195, 150, 10, "If there are any open programs, list them here: "
    'Text 5, 215, 175, 10, "If there are any HH membs open on HC, list them here: "
  'End if
  CheckBox 5, 335 - dialog_shrink_amt, 65, 10, "Updated MMIS?", updated_MMIS_check
  CheckBox 80, 335 - dialog_shrink_amt, 155, 10, "Check here if you sent a NOMI to this client.", NOMI_check
  CheckBox 245, 335 - dialog_shrink_amt, 95, 10, "WCOM added to notice?", WCOM_check
  CheckBox 30, 350 - dialog_shrink_amt, 125, 10, "Check here to TIKL to send to CLS.", TIKL_check
  EditBox 75, 365 - dialog_shrink_amt, 70, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 265, 365 - dialog_shrink_amt, 50, 15
    CancelButton 320, 365 - dialog_shrink_amt, 50, 15
    PushButton 250, 5, 145, 15, "Autofill previous denied progs script info", autofill_previous_info_button
    PushButton 345, 335 - dialog_shrink_amt, 50, 10, "SPEC/WCOM", SPEC_WCOM_button
  Text 5, 25, 50, 10, "Denied Progs: "
  Text 5, 10, 50, 10, "Case number:"
  Text 125, 10, 55, 10, "Application date:"
  Text 5, 110, 55, 10, "Other Reasons: "
  Text 5, 130, 130, 10, "Verifs/docs/apps needed (if applicable):"
  Text 5, 175, 45, 10, "Other notes:"
  Text 5, 45, 45, 10, "Denial Date: "
  Text 5, 60, 40, 10, "Reasons:"
  Text 5, 370 - dialog_shrink_amt, 65, 10, "Worker signature: "
EndDialog

Dialog denied_dialog
cancel_confirmation

If self_declaration_of_income_over_165_FPG = 1 THEN	
	call navigate_to_MAXIS_screen("STAT", "PROG")
	EMReadScreen int_date, 8, 10, 55
          	int_date = replace(int_date, " ", "/")
	call navigate_to_MAXIS_screen("ELIG", "FS")
	transmit
	EMWriteScreen "x", 15, 4
	transmit
	EMReadScreen reported_income, 10, 9, 30
	reported_income = trim(reported_income)
	EMReadScreen max_gross_income, 10, 15, 67
	max_gross_income = trim(max_gross_income)	
End if



'NOW IT CASE NOTES THE DATA.
call start_a_blank_case_note
Call write_variable_in_case_note("----Denied " & progs_denied & "----")
call write_bullet_and_variable_in_case_note("SNAP denial date", SNAP_denial_date)
call write_bullet_and_variable_in_case_note("HC denial date", HC_denial_date)
call write_bullet_and_variable_in_case_note("cash denial date", cash_denial_date)
call write_bullet_and_variable_in_case_note("Emer denial date", emer_denial_date)
call write_bullet_and_variable_in_case_note("Application date", application_date)
call write_bullet_and_variable_in_case_note("Reason for denial", reason_for_denial)
call write_bullet_and_variable_in_case_note("Coding for denial", coded_denial)
call write_bullet_and_variable_in_case_note("Verifs needed", verifs_needed)
	'adding case note portion to cover Self Declaration of Over Income Policy
	If self_declaration_of_income_over_165_FPG = 1 THEN
		call write_variable_in_case_note("---")
		call write_variable_in_case_note("   ***Self Declaration of Over Income Policy for SNAP***")	
		call write_variable_in_case_note("* Date of Interview: " & int_date)
		call write_variable_in_case_note("* Client's Stated Total Income: $" & reported_income)
		call write_variable_in_case_note("* Max Gross Income 165% of FPG: $" & max_gross_income)
		call write_variable_in_case_note("* Denial Reason: Client stated their income is greater than 165% of FPG")
	End If
If updated_MMIS_check = 1 then call write_variable_in_case_note("* Updated MMIS.")
If disabled_client_check = 1 then call write_variable_in_case_note("* Client is disabled.")
If WCOM_check = 1 then call write_variable_in_case_note("* Added WCOM to notice.")
If NOMI_check = 1 then call write_variable_in_case_note("* Sent NOMI to client.")
If case_noting_intake_dates = True then
	call write_variable_in_case_note("---")
	If HC_check = 1 then call write_bullet_and_variable_in_case_note("Last HC REIN date", HC_last_REIN_date)
	If SNAP_check = 1 then call write_bullet_and_variable_in_case_note("Last SNAP REIN date", SNAP_last_REIN_date)
	If cash_check = 1 then call write_bullet_and_variable_in_case_note("Last cash REIN date", cash_last_REIN_date)
	If emer_check = 1 then call write_bullet_and_variable_in_case_note("Last emer REIN date", emer_last_REIN_date)
	If open_prog_check = 1 or HH_membs_on_HC_check = 1 then 
		If open_progs <> "" then call write_bullet_and_variable_in_case_note("Open programs", open_progs)
		If HH_membs_on_HC <> "" then call write_bullet_and_variable_in_case_note("HH members remaining on HC", HH_membs_on_HC)
	Else
		call write_variable_in_case_note("* All programs denied. Case becomes intake again on " & intake_date & ".")
	End if
Else
	If open_progs <> "" then call write_bullet_and_variable_in_case_note("Open programs", open_progs)
	If HH_membs_on_HC <> "" then call write_bullet_and_variable_in_case_note("HH members remaining on HC", HH_membs_on_HC)
End if
call write_bullet_and_variable_in_case_note("Other notes", other_notes)
call write_variable_in_case_note("---")
call write_variable_in_case_note(worker_signature)


script_end_procedure
