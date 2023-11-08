# openai-email-digest
Create a digest of outlook emails using openai api


# Instructions

## Set up Poetry Environment

poetry install

## Build The EXE

- poetry run python build.py
- EXE can be found in dist\Email Summariser.exe
- pin EXE to task bar

## Create local folder to store JSON email files

e.g. create a folder at C:\Emails

## Add VBA to Outlook

- Enable the developer part of the ribbon: Right-click on ribbon > Customise Ribbon > Check Developer
- On Develop Tab Click "Visual Basic"
- Right Click on Modules > Import File > Import vba/modEmails.bas
- If applicable change M_STR_EMAIL_FOLDER to your email folder (default is C:\Emails).
- Manually run the export once by clicking into RunSaveEmailsToJsonFiles and pressing F5
- To automate export Double-click on "ThisOutlookSession" and paste in the following code

	Option Explicit

	Private Sub Application_ItemLoad(ByVal Item As Object)

		RunSaveEmailsToJsonFiles

	End Sub

	Private Sub Application_Reminder(ByVal Item As Object)
		
		RunSaveEmailsToJsonFiles
		
	End Sub

## Running the EXE
- Click on the Email Summariser Icon in your task bar
- Click generate summary





