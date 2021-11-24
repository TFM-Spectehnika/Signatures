Clear-Host
#������ ���� � ������
$CompanyName = "TFM-Spectehnika"
$DomainName = "krut.ru"
$SigSource = "\\srv-db1\Config.1C$\Signatures"
$ForceSignatureNew = 1 #����� �������������, ������ ������� �� ��������� ��� ����� �����. 0 = no force, 1 = force
$ForceSignatureReplyForward = 1 #����� �������������, ������ ������� �� ��������� ��� �������/��������� �����. 0 = no force, 1 = force

#���������� ����������
$AppData=(Get-Item env:appdata).value
$SigPath = "\Microsoft\Signatures"
$LocalSignaturePath = $AppData+$SigPath
$RemoteSignaturePathFull = $SigSource+"\"+$CompanyName+".docx"
$fullPath = $LocalSignaturePath+"\"+$CompanyName+".docx"
$fullPathHTM = $LocalSignaturePath+"\"+$CompanyName+".htm"

#������ ���������� ������������
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

#������� �������� �������
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


#������� �������� �����
Function DELETE_PATH { 
If (Test-Path -Path $LocalSignaturePath) {
Remove-Item -Path $LocalSignaturePath -Recurse
}}

#������� ��� �������� �������
Function DELETE_REG {
If (Test-Path -Path $CompanyRegPath) {
Remove-Item �Path $CompanyRegPath �Recurse
}}

#������� �������� ������������� ����� ������� � �������� � ������ ����������
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

#������� ������� ������ ������� �������� �� ���������
Function Main_Sign {
if (isRegistryValue -Path "HKCU:\Software\Microsoft\Office\16.0\Outlook\Setup" -ValueName ImportPRF) {
    $currentState = (Get-ItemProperty �Path "HKCU:\Software\Microsoft\Office\16.0\Outlook\Setup" -Name $key.Value_Name).$( $key.Value_Name )
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

#���������� ����������� ����� �������
Function REG_DATA {
$SigVersion = (gci $RemoteSignaturePathFull).LastWriteTime #������ ����� �������� ������� �������
$ForcedSignatureNew = (Get-ItemProperty $CompanyRegPath"\OutlookSignatureSettings").ForcedSignatureNew
$ForcedSignatureReplyForward = (Get-ItemProperty $CompanyRegPath"\OutlookSignatureSettings").ForcedSignatureReplyForward
$SignatureVersion = (Get-ItemProperty $CompanyRegPath"\OutlookSignatureSettings").SignatureVersion
Set-ItemProperty $CompanyRegPath"\OutlookSignatureSettings" -name SignatureSourceFiles -Value $SigSource
New-ItemProperty $CompanyRegPath"\OutlookSignatureSettings" -name Title -PropertyType String -Value $ADTitle
New-ItemProperty $CompanyRegPath"\OutlookSignatureSettings" -name EmailAddress -PropertyType String -Value $ADEmailAddress
New-ItemProperty $CompanyRegPath"\OutlookSignatureSettings" -name Department -PropertyType String -Value $ADDepartment
New-ItemProperty $CompanyRegPath"\OutlookSignatureSettings" -name TelePhoneNumber -PropertyType String -Value $ADTelePhoneNumber
$SignatureSourceFiles = (Get-ItemProperty $CompanyRegPath"\OutlookSignatureSettings").SignatureSourceFiles
}

#������� �������� �������
Function Create_SIGN { 

Copy-Item -Path $SigSource $AppData"\Microsoft" -Recurse -Force

#������ ���������� ��� New-Object -com word.application
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

#�������� ��������
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

#��������� ���������� ��������
$MSWord.ActiveDocument.Save()
#������� HTM
$saveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], "wdFormatHTML");
[ref]$BrowserLevel = "microsoft.office.interop.word.WdBrowserLevel" -as [type]
$MSWord.ActiveDocument.WebOptions.OrganizeInFolder = $true
$MSWord.ActiveDocument.WebOptions.UseLongFileNames = $true
$MSWord.ActiveDocument.WebOptions.BrowserLevel = $BrowserLevel::wdBrowserLevelMicrosoftInternetExplorer6
$path = $LocalSignaturePath+"\"+$CompanyName+".htm"
$MSWord.ActiveDocument.saveas([ref]$path, [ref]$saveFormat)

#������� RTF
$saveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], "wdFormatRTF");
$path = $LocalSignaturePath+"\"+$CompanyName+".rtf"
$MSWord.ActiveDocument.SaveAs([ref] $path, [ref]$saveFormat)

#������� TXT
$saveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], "wdFormatText");

#�� ������
#$path = $LocalSignaturePath+"\"+$CompanyName+".rtf"
#$MSWord.ActiveDocument.SaveAs([ref] $path, [ref]$saveFormat)

$path = $LocalSignaturePath+"\"+$CompanyName+".txt"
$MSWord.ActiveDocument.SaveAs([ref] $path, [ref]$SaveFormat)
$MSWord.ActiveDocument.Close()
$MSWord.Quit()

#}
}

#�������� ������� �� ��������� ���������� � ������������
if ($ADTitle -eq ($Title = (Get-ItemProperty $CompanyRegPath"\OutlookSignatureSettings").Title))
    {if ($ADDepartment -eq ($Department = (Get-ItemProperty $CompanyRegPath"\OutlookSignatureSettings").Department))
        {if ($ADTelePhoneNumber -eq ($TelePhoneNumber = (Get-ItemProperty $CompanyRegPath"\OutlookSignatureSettings").TelePhoneNumber))
            {if (Test-Path -PathType Leaf -Path $fullPathHTM)
                {Write-Host "1"
                Main_Sign
                }
            else{DELETE_REG
            CREATE_REG
            Write-Host "2"
            REG_DATA
            Create_SIGN
            Main_Sign}}
        else {DELETE_REG
        CREATE_REG
        Write-Host "3"
        REG_DATA
        Create_SIGN
        Main_Sign}}
     else {DELETE_REG
     CREATE_REG
     Write-Host "4"
     REG_DATA
     Create_SIGN
     Main_Sign}}
else {DELETE_REG
CREATE_REG
Write-Host "5"
REG_DATA
Create_SIGN
Main_Sign}




#Stamp registry-values for OutlookSignatureSettings if they doesn`t match the initial script variables. Note that these will apply after the second script run when changes are made in the Custom variables-section.
#����� ������� ����, ��� ������ �������, ����� ����� �� ������ - �������� ���� ������� ��� �������������, � ��� ����� �� ����� �������� ������� ��������� �������. ������� ��� ��� ���� ��������� �������� �������������.
if ($ForcedSignatureNew -eq $ForceSignatureNew){}
else
{Set-ItemProperty $CompanyRegPath"\OutlookSignatureSettings" -name ForcedSignatureNew -Value $ForceSignatureNew}

if ($ForcedSignatureReplyForward -eq $ForceSignatureReplyForward){}
else
{Set-ItemProperty $CompanyRegPath"\OutlookSignatureSettings" -name ForcedSignatureReplyForward -Value $ForceSignatureReplyForward}

if ($SignatureVersion -eq $SigVersion){}
else
{Set-ItemProperty $CompanyRegPath"\OutlookSignatureSettings" -name SignatureVersion -Value $SigVersion}
exit