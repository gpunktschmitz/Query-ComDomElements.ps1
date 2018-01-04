function Query-ComDomElements([string]$Query, $Dom = $false, $Property = $false) {
	if($Dom -eq $false -or $Dom -eq $null) {
		throw "DOM not valid or set"
	}
	
	$SecondQuery = $false
	if($Query.IndexOf('>') -ne -1) {
		$SecondQuery = $Query.Substring($Query.IndexOf('>')+1).Trim()
		$Query = $Query.Substring(0,$Query.IndexOf('>')).Trim()
	}
	
	switch($Query.Substring(0,1)) {
		'.' {
			if($Property -ne $false -and $SecondQuery -eq $false) {
				$result = [System.__ComObject].InvokeMember("getElementsByClassName",[System.Reflection.BindingFlags]::InvokeMethod, $null, $dom, $Query.Substring(1)) | select $Property
			} else {
				$result = [System.__ComObject].InvokeMember("getElementsByClassName",[System.Reflection.BindingFlags]::InvokeMethod, $null, $dom, $Query.Substring(1))
			}
			Break
		}
		'#' {
			if($Property -ne $false -and $SecondQuery -eq $false) {
				$result = [System.__ComObject].InvokeMember("getElementById",[System.Reflection.BindingFlags]::InvokeMethod, $null, $dom, $Query.Substring(1)) | select $Property
			} else {
				$result = [System.__ComObject].InvokeMember("getElementById",[System.Reflection.BindingFlags]::InvokeMethod, $null, $dom, $Query.Substring(1))
			}
			Break
		}
		default {
			if($Property -ne $false -and $SecondQuery -eq $false) {
				$result = [System.__ComObject].InvokeMember("getElementsByTagName",[System.Reflection.BindingFlags]::InvokeMethod, $null, $dom, $Query) | select $Property
			} else {
				$result = [System.__ComObject].InvokeMember("getElementsByTagName",[System.Reflection.BindingFlags]::InvokeMethod, $null, $dom, $Query)
			}
			Break
		}
	}
	
	if($SecondQuery -eq $false) {
		if($Property -ne $false) {
			return $result.$Property
		} else {
			return $result
		}
	} else {
		return Query-ComDomElements -Query $SecondQuery -Dom $result -Property $Property
	}
}

