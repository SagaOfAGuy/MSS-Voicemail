# AutoVoicemail
Autovoicemail is a Powershell script that automates VCU MSS voicemail closing duties.  

## Dependencies
- This script requires the use of PSSlack to send images to Slack using powershell. 
- Navigate to this project's folder and clone the PSSlack repository using this command: ```git clone https://github.com/RamblingCookieMonster/PSSlack.git```

### Requirements
- Make sure Powershell execution policy allows for the execution of Powershell scripts on system. 
- Make sure that Chrome browser is installed on system as the script uses Chrome for browser automation. 
- Have a Slack account

### Creating a Slack App
- Go to the Slack [page]("https://api.slack.com/apps") for creating a Slack App
- Assign Slack App a desired Slack workspace and name. For example, you can name the app "AutoVoicemail", and make the Workspace "MSS-Classroom Support"


### Generating Slack Bot User Token
- After creating the app, you should be on the **Basic Information** page. Under **Add Features and Functionality**, click the **Permissions** option
- Verify that you are on the **OAUTH & Permissions** page
- Under **Scopes**, add the OAUTH Scopes **files:write** and **channels:join** under **Bot Token Scopes**
- Scroll back to the top of the **OAUTH & Permissions** page, and click **Install to Workspace**
- Next, allow the Slack App the necessary permissions by clicking **allow**.
- On the page, there should be a string that starts with **xoxb-**. This is your Bot User token. Save this string in a safe location

### Invite Slack App to Channel
- Open up the Slack application on your computer
- Create a separate slack channel for Voicemail Screenshot validation. For example, **#voicemail**
- Navigate to the **#voicemail** channel, and in the message box at the bottom, type ```/invite @AutoVoicemail```, with **AutoVoicemail** being the name of the slack app. 


### Configuration
```powershell
# Navigate to this project folder
cd \Path\To\Project

# Inside the AutoVoicemail folder, find and edit SendSlack.ps1
notepad SendSlack.ps1

# Set the filename to the full file location of Image.png. For example, C:\Users\$USER\Documents\AutoVoicemail\Image.png
$filepath="C:\Users\$USER\Documents\AutoVoicemail\Image.png"

# Change the following variables within SendSlack.ps1 to your Bot User Token and Channel ID
$channel_name="YourChannelName"
$bot__user_token="xoxb-my-token"

# Save changes to SendSlack.ps1
```

### Credentials
```powershell 
# Find and edit voicemail.ps1
notepad voicemail.ps1

# Change the variables below per your mailbox and security code (Lines 46 & 47)
$mailbox_code="YourMailboxCodeGoesHere"; # Put mailbox code in this variable or in separate file
$sec_code="YourSecurityCodeGoesHere";

# Find and change the variable below to provide script with web voicemail URL (Line 21) 
$url='YourWebVoicemailURL'


# Save changes to voicemail.ps1
```

### Usage
```powershell
# Navigate to this project folder
cd \Path\To\Project

# Run the voicemail script first to login to voicemail utility and screenshot proof of voicemail change. It will also run SendSlack.ps1 after screenshot is taken. 
.\voicemail.ps1
```
## License
[MIT](https://choosealicense.com/licenses/mit/)
