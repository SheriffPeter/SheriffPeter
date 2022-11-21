<details>
<summary>Draw a picture in console</summary>

```powershell
function Draw-Picture
{
	Param (
		[ValidateScript({ Test-Path -Path $_ -PathType Leaf })]
		[string]$ImageFile,
		[ValidateSet('FillTerminal', 'Fit')]
		[string]$Fit = 'Fit'
	)

	$null   = [Reflection.Assembly]::LoadWithPartialName('System.Drawing')
	$SrcImg = [System.Drawing.Image]::FromFile($ImageFile)

	switch ($Fit) {
		'FillTerminal' {
			[int]$newWidth  = $host.UI.rawui.WindowSize.Width
			[int]$newHeight = $host.UI.rawui.WindowSize.Height * 2
			$SrcImg = $SrcImg.GetThumbnailImage($newWidth, $newHeight, $null, 2)
		}
		'Fit' {
			[float]$imgRatio = $SrcImg.Width / $SrcImg.Height
			[float]$conRatio = $host.UI.rawui.WindowSize.Width / ($host.UI.rawui.WindowSize.Height * 2)

			if ([Math]::Abs(1 - $imgRatio) -gt [Math]::Abs(1 - $conRatio)) {
				if ($imgRatio -lt 1) {
					[int]$newHeight = $host.UI.RawUI.WindowSize.Height * 2
					[int]$newWidth  = $newHeight * $imgRatio
				} else {
					[int]$newWidth  = $host.UI.rawui.WindowSize.Width
					[int]$newHeight = $newWidth / $imgRatio
				}
			} else {
				if ($conRatio -lt 1) {
					[int]$newWidth  = $host.UI.rawui.WindowSize.Width
					[int]$newHeight = $newWidth / $imgRatio
				} else {
					[int]$newHeight = $host.UI.RawUI.WindowSize.Height * 2
					[int]$newWidth  = $newHeight * $imgRatio
				}
			}

			$SrcImg = $SrcImg.GetThumbnailImage($newWidth, $newHeight, $null, 2)
		}
	}

	$pixelStrings = for ($i = 0; $i -lt $SrcImg.Height; $i += 2) {
		if( $i -ne 0 ) { "`n" }
		for ($j = 0; $j -lt $SrcImg.Width; $j++) {
			$back = $SrcImg.GetPixel($j, $i)
			if ($i -ge $SrcImg.Height - 1) {
				$foreVT = "" 
			} else {
				$fore   = $SrcImg.GetPixel($j, $i + 1)
				$foreVT = "$([char]27)[38;2;$($fore.R);$($fore.G);$($fore.B)m"
			}

			$backVT = "$([char]27)[48;2;$($back.R);$($back.G);$($back.B)m"

			"$backVT$foreVT▄$([char]27)[0m"
		}
	}

	Write-Host ([string]::Join('', $pixelStrings))
}
```

Examples :

![image](uploads/980fe634e6c2384fb644aabfcc8cfc30/image.png)

![image](uploads/0e9d27813a87b041ba953fc6c4b0b9c0/image.png)

</details>

<details>
<summary>Get Outlook appointments</summary>

```powershell
function Get-Appointments
{
	param([System.DateTime]$date = [System.DateTime]::Now)

	$outlook = New-Object -ComObject Outlook.Application
	$mapi = $outlook.GetNamespace('MAPI')

	$filter = "[MessageClass]='IPM.Appointment' AND [Start]>='$($date.ToShortDateString())' AND [End]<'$($date.AddDays(1).ToShortDateString())'"

	$appointments = $mapi.GetDefaultFolder(9<#Folder Calendar#>).Items
	$appointments.IncludeRecurrences = $true
	$appointments.Sort("[Start]")

	$appointments.Restrict($filter) | ForEach-Object {
		$requiredAttendees = $_.RequiredAttendees.Trim(";").Split(";") | ForEach-Object { $_.Trim() }

		if( [string]::IsNullOrEmpty($_.OptionalAttendees) )
		{
			$optionalAttendees = @()
		}
		else
		{
			$optionalAttendees = $_.OptionalAttendees.Trim(";").Split(";") | ForEach-Object { $_.Trim() }
		}

		$attendees = foreach($attendee in (($requiredAttendees + $optionalAttendees) | Sort-Object))
		{
			if( $_.Organizer -eq $attendee )
			{
				"$([char]0x1b)[42m${attendee}$([char]0x1b)[0m"
			}
			elseif( $attendee -in $optionalAttendees )
			{
				"$([char]0x1b)[32m${attendee}$([char]0x1b)[0m"
			}
			else
			{
				"$([char]0x1b)[92m${attendee}$([char]0x1b)[0m"
			}
		}

		if( $_.ResponseStatus -eq 2 )
		{
			$subject = "$([char]0x1b)[34m$($_.Subject.Trim())$([char]0x1b)[0m"
		}
		else
		{
			$subject = "$([char]0x1b)[44m$($_.Subject.Trim())$([char]0x1b)[0m"
		}

		"${subject}`n`tfrom $([char]0x1b)[96m$($_.Start.ToString('HH:mm'))$([char]0x1b)[0m to $([char]0x1b)[96m$($_.End.ToString('HH:mm'))$([char]0x1b)[0m`n`twith $($attendees -join `",`n`t     `")"
	}
}
```

Example :

![image](uploads/0cb27992682ae767a1bfe400ebcd9b5e/image.png)

```powershell
function Get-WeekAppointments
{
	$date = [System.DateTime]::Now
	$date = $date.AddDays(1-[int]$date.DayOfWeek)
	foreach($i in 1..5)
	{
		"╭―$('―' * $date.ToLongDateString().Length)―╮"
		
		"│ $($date.ToLongDateString()) │"
		
		"╰―$('―' * $date.ToLongDateString().Length)―╯"
		Get-Appointments $date
		$date = $date.AddDays(1)
	}
}
```

Example:

![image](uploads/341c7b71292d9a291c690a572f6ce8f4/image.png)

</details>

<details>
<summary>Create a console shortcut key to set clipboard with current working directory</summary>

```powershell
Set-PSReadLineKeyHandler -Key ctrl+p {
	Set-Clipboard (Convert-Path $pwd)
}
```

</details>

<details>
<summary>Create a console shortcut key to explore current working directory</summary>

```powershell
Set-PSReadLineKeyHandler -Key ctrl+e {
	explorer.exe (Convert-Path $pwd)
}
```

</details>

<details>
<summary>Create a console shortcut key to launch web site from clipboard content</summary>

```powershell
Set-PSReadLineKeyHandler -Key ctrl+l {
	$content = Get-Clipboard
	if( $content )
	{
		$content = $content.Trim()
		if( ($content -as [System.URI]).AbsoluteURI )
		{
			Start-Process $content
		}
		elseif( $content -match '^SA[0-9]+$' )
		{
			Start-Process "https://wlib-si.cm-cic.fr?mnc=SARA&ref=${content}"
		}
		elseif( $content -match '^RU[0-9]+$' )
		{
			Start-Process "https://wlib-si.cm-cic.fr?mnc=RUBIS&ref=${content}"
		}
	}
}
```

Example:

Copy `SA0000048173153` and press CTRL+L then SARA website opens.

![image](uploads/dd48873823c11f006865cfb72c4371b4/image.png)

</details>

<details>
<summary>Format XML to pretty colored tree</summary>

```powershell
function Format-XML ([xml]$xml, $indent=2)
{
	function print-xmlelement([System.Xml.XmlNode]$node, [string]$path, [bool]$isLastNode, [string]$indentString, [System.Text.StringBuilder]$sb)
	{
		$newPath = $path + "/"
		if( -not [string]::IsNullOrEmpty($node.prefix) )
		{
			$newPath += "$([char]0x1b)[95m$($node.Prefix)$([char]0x1b)[0m:"
		}
		$newPath += "$([char]0x1b)[93m$($node.LocalName)$([char]0x1b)[0m"

		[void]$sb.Append($indentString)
	
		$newIndentString = $indentString

		if( -not [string]::IsNullOrEmpty($path) )
		{
			if( $isLastNode )
			{
				$newIndentString += " "
			}
			else
			{
				$newIndentString += "│"
			}
			$newIndentString += " "*$indent
	
			if( $isLastNode )
			{
				[void]$sb.Append("└")
			}
			else
			{
				[void]$sb.Append("├")
			}
			[void]$sb.Append("─"*$indent)
		}

		switch( $node.NodeType )
		{
			"Text"
			{ [void]$sb.Append("$([char]0x1b)[31m$($node.Value)$([char]0x1b)[0m") }
			"CDATA"
			{ [void]$sb.Append("$([char]0x1b)[31m$($node.Value)$([char]0x1b)[0m") }
			"Comment"
			{ [void]$sb.Append("$([char]0x1b)[90m$($node.Value)$([char]0x1b)[0m") }
			"Element"
			{ [void]$sb.Append($newPath) }
			default
			{ [void]$sb.Append($node.NodeType.ToString()) }
		}
		[void]$sb.Append("`n")

		$attributeIndentString = $newIndentString
		if( $node.ChildNodes.Count )
		{
			$attributeIndentString += "│"
		}
		else
		{
			$attributeIndentString += " "
		}
		$attributeIndentString += " "*$indent

		#GetNamespaceOfPrefix
		#GetPrefixOfNamespace

		[int]$attibutePadding = 0
		foreach($attribute in $node.Attributes)
		{
			[int]$attibutePadding = [math]::Max($attibutePadding, $attribute.Name.Length)
		}
		foreach($attribute in $node.Attributes)
		{
			[void]$sb.Append($attributeIndentString)
			[void]$sb.Append("$([char]0x1b)[96m$($attribute.Name.PadRight($attibutePadding))$([char]0x1b)[0m")
			[void]$sb.Append(" : ")
			[void]$sb.Append("$([char]0x1b)[92m$($attribute.'#text')$([char]0x1b)[0m")
			[void]$sb.Append("`n")
		}

		foreach($childNode in $node.ChildNodes)
		{
			print-xmlelement $childNode $newPath ($childNode -eq $node.ChildNodes[$node.ChildNodes.Count -1]) $newIndentString $sb
		}
	}

	function print-xmlroot([xml]$xml, [System.Text.StringBuilder]$sb)
	{
		print-xmlelement $xml.DocumentElement "" $true "" $sb
	}

	$sb = [System.Text.StringBuilder]::new()
	print-xmlroot $xml $sb
	$sb.ToString()
}
```

Example:

![image](uploads/13fd5d58998d96385670f00f82776f23/image.png)

</details>

