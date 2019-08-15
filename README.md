# MailTemplate

The project allow you to compose email out of templates with placeholder fields inside Outlook.

## How to install

Use the files available in the Release section to install the Add-in or the VBA Macro.

### Add-in

1. Extract the files where you want to keep them.

	By default windows has a Add-In folder: `%APPDATA%\Microsoft\AddIns`.

2. Run the `Setup.exe`.

The Template selector will be available in the Home Ribbon of Outlook.

### VBA Macro

1. Input `[ALT]+[F11]`
2. Right-click on the root element in the Project sidebar (normally named `Project1`).
3. Click "Import File..."
4. Navigate where all the VBA files downloaded are located.
5. Select a file and click "Open".
6. Repeat 2 to 5 for each files.
7. Save the project.

Then the Macro "Project1.ShowTemplateSelector" is installed, but you'll need add it to your Ribbon or the quick access bar to use it.

## How to use create a template email

1. Use "New Email" feature to compose the template email.

	Indicate the variable informations in the template as follow: `{:Variable Name}`.

	Also, if you have a signature, remove it since Outlook automatically add the signature to the template.

2. Click on "File" > "Save As"
3. The save type must be: "Outlook Template (*.oft)"; and must be located in the folder: `%APPDATA%\Microsoft\Templates`.

## How to compose an email from a template

1. Be sure your template is created.
2. Open the template selector
	* For Add-in: Click the "Select" button in the "Template" group of the "Home" tab.
	* For Macro: Run the Macro named : ShowTemplateSelector
3. Select your template from the dropdown.
4. Fill the fields in the "Variable" group.
5. Click "Apply".
6. A new message will be displayed with the content of the template with fields replaced.
7. Continue as you would normally do with a regular email.
