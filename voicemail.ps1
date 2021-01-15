# Credit goes to https://universecitiz3n.tech/selenium/Selenium-Powershell/ for setup 
# https://powershellone.wordpress.com/2015/02/12/waiting-for-elements-presence-with-selenium-web-driver/ 
# Webdriver.Support.dll can be downloaded from https://github.com/AdyKalra/WAF-WebAutomationFramework/blob/master/WAF/Packages/Selenium.Support/3.4.0/lib/net40/WebDriver.Support.dll


# Load Selenium DLL files 
$PathToFolder = "$pwd\"
[System.Reflection.Assembly]::LoadFrom("{0}\WebDriver.dll" -f $PathToFolder);
[System.Reflection.Assembly]::LoadFrom("{0}\WebDriver.Support.dll" -f $PathToFolder);
if ($env:Path -notcontains ";$PathToFolder" ) {
    $env:Path += ";$PathToFolder"
}


# Configure Chrome Browser
$ChromeOptions = New-Object OpenQA.Selenium.Chrome.ChromeOptions
$ChromeOptions.AddArgument('start-maximized')
$ChromeOptions.AcceptInsecureCertificates = $True
$ChromeDriver = New-Object OpenQA.Selenium.Chrome.ChromeDriver($ChromeOptions)
$webDriverWait = New-Object OpenQA.Selenium.Support.UI.WebDriverWait($ChromeDriver, 30); 
$url=""; 

# Function for typing in strings to form elements
function type_string($form_element, $string) {
    for($index = 0; $index -lt $string.Length; $index++) {
        $form_element.SendKeys($string.Substring($index,1)); 
        $random_num = Get-Random -Minimum 0.001 -Maximum 0.08; 
        start-sleep -Seconds $random_num; 
    }

}
# Login to Mailbox 
function Login() { 
    # Start Chrome, and get form elements

    # Define this URL 
    $ChromeDriver.Url = $url;
    $ChromeDriver.SwitchTo().Frame("wpm"); # This gets the frame of page; need to switch to get access to HTML elements
    
    $username = $webDriverWait.Until([OpenQA.Selenium.Support.UI.ExpectedConditions]::ElementIsVisible( [OpenQA.Selenium.by]::Id("mailbox"))); 
    $password = $webDriverWait.Until([OpenQA.Selenium.Support.UI.ExpectedConditions]::ElementIsVisible( [OpenQA.Selenium.by]::Id("password"))); 
    $button =   $webDriverWait.Until([OpenQA.Selenium.Support.UI.ExpectedConditions]::ElementIsVisible( [OpenQA.Selenium.by]::ClassName("submit"))); 


    # Type in mailbox and code
    $mailbox_code=""; # Put mailbox code in this variable or in separate file
    $sec_code="";     # Put password in this variable or in separate file 

    type_string $username $mailbox_code; 
    type_string $password $sec_code; 
    start-sleep -Seconds 2; 
    $button.Click(); 

    # Login to voicemail 
    Write-Host "Logged in!"
    start-sleep 5; 
    # Find and click personal settings
} 
# Enable out of office voicemail box 
function EnableOutOfOffice() { 
    
    $personal_settings = $webDriverWait.Until([OpenQA.Selenium.Support.UI.ExpectedConditions]::ElementIsVisible( [OpenQA.Selenium.by]::Id("menu_item_personal")));
    $personal_settings.Click(); 
    # Enable out of office option  
    Sleep 10
    $out_of_office =  @"  
    var office = document.getElementsByTagName("input");
    office[3].click();  
"@;

    $ChromeDriver.ExecuteScript($out_of_office);      
}
# Take screenshot of enabled voicemail 
function Screenshot() { 
    # JavasCript code for displaying date for further validation 
    $script =  @"  
    var d = new Date(); 
    var month = d.getMonth()+1; 
    var day = d.getDate(); 
    var year = d.getFullYear();  
    var userInfo = document.getElementsByClassName("userInfo"); 
    userInfo[0].innerText = month + '-' + day + '-' + year; 
"@;
 
    # Click OK button to approve of out of office option 
    $ok_button = $webDriverWait.Until([OpenQA.Selenium.Support.UI.ExpectedConditions]::ElementIsVisible( [OpenQA.Selenium.by]::Id("ok"))); 
    $ok_button.Click();

    # Execute JavaScript code to put date on page
    $ChromeDriver.ExecuteScript($script); 

    # Stuck on Screenshot. Need to get this working. Javascript not executing properly on page.  
    $base64_string = $ChromeDriver.GetScreenshot().AsBase64EncodedString; 
    $img = [Drawing.Bitmap]::FromStream([IO.MemoryStream][Convert]::FromBase64String($base64_string)); 
    $img.Save($pwd.Path+"\Image.png"); 
}
function Logout() { 
    # Logout of voicemail. Define the window.location below 
    $logout = @"
    window.location=$url+"/main.php?&module=home&page=messages&main_page_logout=""; 
"@
    $ChromeDriver.ExecuteScript($logout); 
    write-host "Logged out!"; 

    # Kill chromedriver process  
    Stop-Process -Name chromedriver
}
# Send Image file to slack channel 
function SendToSlack() { 
    .\SendSlack.ps1
}

function main() { 
    Login
    EnableOutOfOffice
    Screenshot
    Logout
    SendToSlack    
    exit 1 
} 
main
