# Import necessary Slack Modules from PSSlack for Powershell 
$import_path = $pwd.Path + "\PSSlack-master\PSSlack\PSSlack.psd1"; 
Import-Module $import_path; 

# Variables for sending screenshot via PSSlack module
$filepath = ""
$channel_name=""
$bot_user_token = ""

# Load Bot User token 
Set-PSSlackConfig -Token $bot_user_token; 

# Send screenshot to Slack 
Send-SlackFile -Channel $channel_name -Path $filepath -title "Voicemail Screenshot"; 