<#
.SYNOPSIS
    This script will get the top or latest tweets from the past day for a hashtag or a specific user

.DESCRIPTION 
    This script will get the top or latest tweets from the past day for a hashtag or a specific user

.PARAMETER Hashtag
    Filter tweets by hashtag

.PARAMETER User
    Filter tweets by user

.PARAMETER ResultType
    Return recent, popular or mixed tweets

.PARAMETER Recipients
    Comma separated list of email recipients

.Example
    Get-Tweets.ps1 -Hashtag ConfigMgr -ResutType Recent -Recipients "me@mail.com"

    Will retrieve a list of the latest teets with the hashtag #ConfigMgr and mail them to me

.Example
    Get-Tweets.ps1 -User bachirica -ResutType Recent -Recipients "me@mail.com"

    Will retrieve a list of the latest teets with the hashtag #ConfigMgr and mail them to me

.NOTES
    - Author: Bernardo Achirica
	- Email : 
	- CreationDate: 2018.05.04
	- LastModifiedDate: 2018.06.23
	- Version: 1.2
    - History:
        - 1.2: Added the option to retreive tweets for a specifit user.
#>

#region Parameters

Param (
    [Parameter(
        ParameterSetName = "Hashtag",
        HelpMessage="Hashtag tu be used as filter",
        Mandatory = $true
    )]
    [String]$Hashtag,

    [Parameter(
        ParameterSetName = "User",
        Mandatory = $true
    )]
    [string]$User,

    [Parameter(
        Mandatory = $true,
        HelpMessage="Result type (recent, popular, mixed)"
    )]
    [ValidateSet(
        'Recent','Popular','Mixed'
    )]
    [String]$ResultType,

    [Parameter(
        Mandatory = $true,
        HelpMessage="Comma separated list of email recipients ""mail1@mail.com"", ""mail2@mail.com"""
    )]
    [Array]$Recipients
)

#endregion

#region Variables

# SMTP Server configuration
$SMTPServer = "smtp.gmail.com"
$SMTPUserName = "xxxxx"
$SMTPPassword = "xxxxx"
$SMTPFromAddress = "xxxxx"

# Twitter API Configuration
# To get an API Key, etc go to: https://apps.twitter.com/
$TwitterAPIKey = "xxxxx"
$TwitterAPISecret = "xxxxx"
$TwitterAccessToken = "xxxxx"
$TwitterAccessTokenSecret = "xxxxx"

#endregion

#region MainScript

Import-Module "C:\Users\Bernardo\OneDrive\Tech\Scripts\InvokeTwitterAPIs\InvokeTwitterAPIs.psm1"

$secpasswd = ConvertTo-SecureString $SMTPPassword -AsPlainText -Force
$SMTPCredential = New-Object System.Management.Automation.PSCredential ($SMTPUserName, $secpasswd)

$Date = ($(Get-Date)).ToString("yyyy-MM-dd")
$OAuth = @{'ApiKey' = $TwitterAPIKey; 'ApiSecret' = $TwitterAPISecret; 'AccessToken' = $TwitterAccessToken; 'AccessTokenSecret' = $TwitterAccessTokenSecret}

$Results = @()

if ($Hashtag) {
    $Tweets = Invoke-TwitterRestMethod -RestVerb GET -ResourceURL "https://api.twitter.com/1.1/search/tweets.json" -Parameters @{"q" = "#$Hashtag AND -filter:retweets AND -filter:replies"; "result_type" = "$ResultType"; "count" = 100; "tweet_mode" = "extended"; "until" = $Date } -OAuthSettings $OAuth

    foreach ($Tweet in $Tweets.statuses) {
        $TwYear = ($Tweet.created_at).Substring(26,4)
        $TwMonth = ($Tweet.created_at).Substring(4,3)
        $TwDay = ($Tweet.created_at).Substring(8,2)
        $TwTime = ($Tweet.created_at).Substring(11,8)

        $TwDateTime = Get-Date "$TwYear-$TwMonth-$TwDay $TwTime"

        if ($TwDateTime -gt $(Get-Date).AddHours(-24)) {
            $Results += [pscustomobject]@{
                TwUser = $Tweet.user.name
                TwScreenName = $Tweet.user.screen_name
                TwRetweets = $Tweet.retweet_count
                TwFullText = $Tweet.full_text
                TwID = $Tweet.id
            }
        }
    }
    $Results = $Results | Sort-Object -Property TwRetweets -Descending
    $MailSubject = "Twitter Monitor: #$($Hashtag) - $($Date)"
}

if ($User) {
    $Tweets = Invoke-TwitterRestMethod -RestVerb GET -ResourceURL "https://api.twitter.com/1.1/statuses/user_timeline.json" -Parameters @{"screen_name" = "$User"; "count" = 100; "tweet_mode" = "extended"; "until" = $Date } -OAuthSettings $OAuth

    foreach ($Tweet in $Tweets){
        $TwYear = ($Tweet.created_at).Substring(26,4)
        $TwMonth = ($Tweet.created_at).Substring(4,3)
        $TwDay = ($Tweet.created_at).Substring(8,2)
        $TwTime = ($Tweet.created_at).Substring(11,8)

        $TwDateTime = Get-Date "$TwYear-$TwMonth-$TwDay $TwTime"

        if ($TwDateTime -gt $(Get-Date).AddHours(-24)) {
            $Results += [pscustomobject]@{
                TwUser = $Tweet.user.name
                TwScreenName = $Tweet.user.screen_name
                TwRetweets = $Tweet.retweet_count
                TwFullText = $Tweet.full_text
                TwID = $Tweet.id
            }
        }
    }
    $Results = $Results | Sort-Object -Property TwID -Descending
    $MailSubject = "Twitter Monitor: @$($User) - $($Date)"
}

$HTMLExit = '<html><head><meta http-equiv="Content-Type" content="text/html; charset=Windows-1252">'
$HTMLExit += '<style>table { font-family: "Segoe UI", Frutiger, "Frutiger Linotype", "Dejavu Sans", "Helvetica Neue", Arial, sans-serif; width:100%; } table, th, td { border: 1px solid black; border-collapse: collapse; } th, td { padding: 15px; text-align: left; } table#t01 tr:nth-child(even) { background-color: #D4E7FF; } table#t01 tr:nth-child(odd) { background-color: #fff; } table#t01 th { background-color: #2A4B9C; color: white; }</style>'
$HTMLExit += '</head><body><table id="t01"><tbody><tr><th>Author</th><th>Retweets</th><th>Tweet</th></tr>'

foreach ($Result in $Results) {
    $HTMLExit += '<tr>'
    $HTMLExit += '<td style="text-align:center;">' + $Result.TwUser + '<br>(<a href="http://twitter.com/' + $Result.TwScreenName + '">@' + $Result.TwScreenName + '</a>)</td>'
    $HTMLExit += '<td style="text-align:center;"><a href="http://twitter.com/' + $Result.TwScreenName + '/status/' + $Result.TwID + '">' + $Result.TwRetweets + '</a></td>'
    $HTMLExit += '<td>' + $Result.TwFullText + '</td>'
    $HTMLExit += '</tr>'
}

$HTMLExit += '</tbody></table></body></html>'

$HTMLExit | Out-File test.html

Send-MailMessage -From: $SMTPFromAddress -To $SMTPFromAddress -Bcc $Recipients -Subject $MailSubject -BodyAsHtml $HTMLExit -SmtpServer $SMTPServer -UseSsl -Credential $SMTPCredential

#endregion

