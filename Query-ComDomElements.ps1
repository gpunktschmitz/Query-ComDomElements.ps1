function Query-ComDomElements([string]$Filter, $Dom = $false, $Property = $false) {
	if($Dom -eq $false -or $Dom -eq $null) {
		throw "DOM not valid or set"
	}
	
	$SecondFilter = $false
	if($Filter.IndexOf('>') -ne -1) {
		$SecondFilter = $Filter.Substring($Filter.IndexOf('>')+1).Trim()
		$Filter = $Filter.Substring(0,$Filter.IndexOf('>')).Trim()
	}
	
	switch($Filter.Substring(0,1)) {
		'.' {
			if($Property -ne $false -and $SecondFilter -eq $false) {
				$result = [System.__ComObject].InvokeMember("getElementsByClassName",[System.Reflection.BindingFlags]::InvokeMethod, $null, $dom, $Filter.Substring(1)) | select $Property
			} else {
				$result = [System.__ComObject].InvokeMember("getElementsByClassName",[System.Reflection.BindingFlags]::InvokeMethod, $null, $dom, $Filter.Substring(1))
			}
			Break
		}
		'#' {
			if($Property -ne $false -and $SecondFilter -eq $false) {
				$result = [System.__ComObject].InvokeMember("getElementById",[System.Reflection.BindingFlags]::InvokeMethod, $null, $dom, $Filter.Substring(1)) | select $Property
			} else {
				$result = [System.__ComObject].InvokeMember("getElementById",[System.Reflection.BindingFlags]::InvokeMethod, $null, $dom, $Filter.Substring(1))
			}
			Break
		}
		default {
			if($Property -ne $false -and $SecondFilter -eq $false) {
				$result = [System.__ComObject].InvokeMember("getElementsByTagName",[System.Reflection.BindingFlags]::InvokeMethod, $null, $dom, $Filter) | select $Property
			} else {
				$result = [System.__ComObject].InvokeMember("getElementsByTagName",[System.Reflection.BindingFlags]::InvokeMethod, $null, $dom, $Filter)
			}
			Break
		}
	}
	
	if($SecondFilter -eq $false) {
		if($Property -ne $false) {
			return $result.$Property
		} else {
			return $result
		}
	} else {
		return Query-ComDomElements -Filter $SecondFilter -Dom $result -Property $Property
	}
}

