Clear-Host
#Set Varibles
$CompanyName = "Company"
$SigSource = "\\net_share\Signatures"

#Connect to ad and set Signatures Path
$AppData=(Get-Item env:appdata).value
$SigPath = "\Microsoft\Signatures"
$LocalSignaturePath = $AppData+$SigPath
$fullPath = $LocalSignaturePath+"\"+$CompanyName+".docx"
$fullPathHTM = $LocalSignaturePath+"\"+$CompanyName+".htm"

#Get User information
$UserName = $env:username
$Filter = "(&(objectCategory=User)(samAccountName=$UserName))"
$Searcher = New-Object System.DirectoryServices.DirectorySearcher
$Searcher.Filter = $Filter
$ADUserPath = $Searcher.FindOne()
$ADUser = $ADUserPath.GetDirectoryEntry()
$ADDisplayName = $ADUser.DisplayName
$ADEmailAddress = $ADUser.mail
$ADTitle = $ADUser.title
$ADTelePhoneNumber = $ADUser.TelephoneNumber
$ADDepartment = $ADUser.Department
$CompanyRegPath = "HKCU:\Software\"+$CompanyName

#Check reg
function isRegistryValue {
    param (
        [parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]$Path,
        [parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]$ValueName
    )
    try {Get-ItemProperty -Path $Path | Select-Object -ExpandProperty $ValueName -ErrorAction Stop | Out-Null
        return $true }
    catch {
        return $false}
}


#Delete Folder
Function DELETE_Folder { 
If (Test-Path -Path $LocalSignaturePath) {
Remove-Item -Path $LocalSignaturePath -Recurse
}}

#Delete REG
Function DELETE_REG {
If (Test-Path -Path $CompanyRegPath) {
Remove-Item -Path $CompanyRegPath -Recurse
}}

#Create REG PATH
Function CREATE_REG {
if (Test-Path $CompanyRegPath)
{}
else
{New-Item -path "HKCU:\Software" -name $CompanyName}

if (Test-Path $CompanyRegPath"\OutlookSignatureSettings")
{}
else
{New-Item -path $CompanyRegPath -name "OutlookSignatureSettings"}
}

#Set main signatures
Function Main_Sign {
if (isRegistryValue -Path "HKCU:\Software\Microsoft\Office\16.0\Outlook\Setup" -ValueName ImportPRF) {
    $currentState = (Get-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Outlook\Setup" -Name $key.Value_Name).$( $key.Value_Name )
    if (($currentState -eq $key.Value) -or (([string]::IsNullOrEmpty($key.Value)) -and ([string]::IsNullOrEmpty($currentState))))
    {}
    else {
        Remove-ItemProperty -Path HKCU:\Software\Microsoft\Office\16.0\Outlook\Setup -Name First-Run -Force -ErrorAction SilentlyContinue -Verbose
        New-ItemProperty HKCU:'\Software\Microsoft\Office\16.0\Common\MailSettings' -Name 'ReplySignature' -Value $CompanyName -PropertyType 'String' -Force
        New-ItemProperty HKCU:'\Software\Microsoft\Office\16.0\Common\MailSettings' -Name 'NewSignature' -Value $CompanyName -PropertyType 'String' -Force
}
}
else {
    Remove-ItemProperty -Path HKCU:\Software\Microsoft\Office\16.0\Outlook\Setup -Name First-Run -Force -ErrorAction SilentlyContinue -Verbose
    New-ItemProperty HKCU:'\Software\Microsoft\Office\16.0\Common\MailSettings' -Name 'ReplySignature' -Value $CompanyName -PropertyType 'String' -Force
    New-ItemProperty HKCU:'\Software\Microsoft\Office\16.0\Common\MailSettings' -Name 'NewSignature' -Value $CompanyName -PropertyType 'String' -Force
}
}

#Set data to reg
Function REG_DATA {
Set-ItemProperty $CompanyRegPath"\OutlookSignatureSettings" -name SignatureSourceFiles -Value $SigSource
New-ItemProperty $CompanyRegPath"\OutlookSignatureSettings" -name Title -PropertyType String -Value $ADTitle
New-ItemProperty $CompanyRegPath"\OutlookSignatureSettings" -name EmailAddress -PropertyType String -Value $ADEmailAddress
New-ItemProperty $CompanyRegPath"\OutlookSignatureSettings" -name Department -PropertyType String -Value $ADDepartment
New-ItemProperty $CompanyRegPath"\OutlookSignatureSettings" -name TelePhoneNumber -PropertyType String -Value $ADTelePhoneNumber
}

#Create signatures
Function Create_SIGN { 

Copy-Item -Path $SigSource $AppData"\Microsoft" -Recurse -Force

#Set Enviroment for New-Object -com word.application
$ReplaceAll = 2
$FindContinue = 1
$MatchCase = $True
$MatchWholeWord = $True
$MatchWildcards = $False
$MatchSoundsLike = $False
$MatchAllWordForms = $False
$Forward = $True
$Wrap = $FindContinue
$Format = $False

#Start applet
$MSWord = New-Object -com word.application
#$MSWord.Visible = $True
$MSWord.Documents.Open($fullPath)

$FindText = "DisplayName"
$ReplaceText = $ADDisplayName.ToString()
$MSWord.Selection.Find.Execute($FindText, $MatchCase, $MatchWholeWord, $MatchWildcards, $MatchSoundsLike, $MatchAllWordForms, $Forward, $Wrap, $Format, $ReplaceText, $ReplaceAll )

$FindText = "Title"
$ReplaceText = $ADTitle.ToString()
$MSWord.Selection.Find.Execute($FindText, $MatchCase, $MatchWholeWord, $MatchWildcards, $MatchSoundsLike, $MatchAllWordForms, $Forward, $Wrap, $Format, $ReplaceText, $ReplaceAll )

$FindText = "Telephone"
$ReplaceText = $ADTelePhoneNumber.ToString()
$MSWord.Selection.Find.Execute($FindText, $MatchCase, $MatchWholeWord, $MatchWildcards, $MatchSoundsLike, $MatchAllWordForms, $Forward, $Wrap, $Format, $ReplaceText, $ReplaceAll )

$FindText = "Department"
$ReplaceText = $ADDepartment.ToString()
$MSWord.Selection.Find.Execute($FindText, $MatchCase, $MatchWholeWord, $MatchWildcards, $MatchSoundsLike, $MatchAllWordForms, $Forward, $Wrap, $Format, $ReplaceText, $ReplaceAll )

$MSWord.Selection.Find.Execute("Email")
$MSWord.ActiveDocument.Hyperlinks.Add($MSWord.Selection.Range, "mailto:"+$ADEmailAddress.ToString(), $missing, $missing, $ADEmailAddress.ToString())

#Save chenged documents
$MSWord.ActiveDocument.Save()
#Save to HTM
$saveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], "wdFormatHTML");
[ref]$BrowserLevel = "microsoft.office.interop.word.WdBrowserLevel" -as [type]
$MSWord.ActiveDocument.WebOptions.OrganizeInFolder = $true
$MSWord.ActiveDocument.WebOptions.UseLongFileNames = $true
$MSWord.ActiveDocument.WebOptions.BrowserLevel = $BrowserLevel::wdBrowserLevelMicrosoftInternetExplorer6
$path = $LocalSignaturePath+"\"+$CompanyName+".htm"
$MSWord.ActiveDocument.saveas([ref]$path, [ref]$saveFormat)

#Save to RTF
$saveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], "wdFormatRTF");
$path = $LocalSignaturePath+"\"+$CompanyName+".rtf"
$MSWord.ActiveDocument.SaveAs([ref] $path, [ref]$saveFormat)

#Save to TXT
$saveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], "wdFormatText");

#Needn't 
#$path = $LocalSignaturePath+"\"+$CompanyName+".rtf"
#$MSWord.ActiveDocument.SaveAs([ref] $path, [ref]$saveFormat)

$path = $LocalSignaturePath+"\"+$CompanyName+".txt"
$MSWord.ActiveDocument.SaveAs([ref] $path, [ref]$SaveFormat)
$MSWord.ActiveDocument.Close()
$MSWord.Quit()
}

#Check all
if ($ADTitle -eq ($Title = (Get-ItemProperty $CompanyRegPath"\OutlookSignatureSettings").Title))
    {if ($ADDepartment -eq ($Department = (Get-ItemProperty $CompanyRegPath"\OutlookSignatureSettings").Department))
        {if ($ADTelePhoneNumber -eq ($TelePhoneNumber = (Get-ItemProperty $CompanyRegPath"\OutlookSignatureSettings").TelePhoneNumber))
            {if (Test-Path -PathType Leaf -Path $fullPathHTM)
                {Main_Sign
                }
            else{
            DELETE_Folder
            DELETE_REG
            CREATE_REG
            REG_DATA
            Create_SIGN
            Main_Sign}}
        else {
        DELETE_Folder    
        DELETE_REG
        CREATE_REG
        REG_DATA
        Create_SIGN
        Main_Sign}}
    else {
    DELETE_Folder
    DELETE_REG
    CREATE_REG
    REG_DATA
    Create_SIGN
    Main_Sign}}
else {
DELETE_Folder
DELETE_REG
CREATE_REG
REG_DATA
Create_SIGN
Main_Sign}
Exit
