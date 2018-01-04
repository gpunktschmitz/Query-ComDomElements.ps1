<#
    Query-DomElements.ps1 - Yu Yagi (miriyagi@miriyagi.jp)
    Seek elements in a DOM objectreturned from Internet Explorer COM object using jQuery-like selectors
    *** Usage ***
    For example, you can get a list of Yahoo! headlines as follows (unless they change the HTML structure):
    [object] $ie = New-Object -ComObject InternetExplorer.Application
    $ie.Visible = $false
    $ie.Navigate("http://www.yahoo.co.jp/")
    while ($ie.Busy) { Start-Sleep 1 }
    Query-DomElements -Dom $ie.document.documentElement -Query "#topicsfb > .topicsindex > .emphasis > li" | Select-Object innerText
    *** Restrictions ***
    Currently, only following selectors and their combinations are supported:
    - '#id'
    - '.class'
    - 'element'
    - 'parent child'
    - 'parent > child'
    *** License ***
    Dual-licensed under CC0 1.0 (http://creativecommons.org/publicdomain/zero/1.0/)
    and NYSL 0.92 (http://www.kmonos.net/nysl/NYSL.TXT)
#>

function Query-DomElements() {
    param (
        [parameter(Mandatory = $true, HelpMessage = "DOM document element object")]
        [object] $Dom,
        [parameter(Mandatory = $true, HelpMessage = "Selector query")]
        [string] $Query
    )

    Write-Debug "Element = $($Dom.nodeName), Query = '$Query'"
    if ($Dom -eq $null) {
        throw "DOM is null"
    }

    # Ignores any selectors before ID specification
    if ($Query.LastIndexOf(" #") -ne -1) {
        Write-Debug "Original: '$Query'"
        $Query = $Query.Substring($Query.LastIndexOf(" #") + 1)
        Write-Debug "Modified: '$Query'"
    }

    # Parse the target element name
    [bool] $childOnly = $false
    do {
        [string] $target, $Query = ($Query -split " ", 2)
        if ([String]::IsNullOrEmpty($target)) {
            return $Dom
        } elseif ($target -eq ">") {
            $childOnly = $true
        }
    } while ($target -eq ">")

    # End of search
    if ([String]::IsNullOrEmpty($target)) {
        return $Dom;
    }
    
    if ($childOnly) {
        # When ">" is specified, filter child nodes matching the given selector.
        $tags = @([Linq.Enumerable]::Range(0, $Dom.childNodes.length) | %{ $Dom.childNodes.item($_) } | ?{
            return -not ($target.StartsWith(".") -and (($_.className -split " ") -inotcontains $target.SubString(1)))
              -and -not ($target.StartsWith("#") -and ($_.id -ne $target.SubString(1)))
              -and -not (!$target.StartsWith(".") -and !$target.StartsWith("#") -and $_.nodeName -ne $target)
        })
    } else {
        # When ">" is not specified, find elements using getElement functions.
        if ($target.StartsWith(".")) {
            $tags = $Dom.getElementsByClassName($target.SubString(1))
            $tags = @([Linq.Enumerable]::Range(0, $tags.length) | %{ $tags.item($_) })
        } elseif ($target.StartsWith("#")) {
            $tags = @($Dom.document.getElementById($target.SubString(1)))
        } else {
            $tags = $Dom.getElementsByTagName($target)
            $tags = @([Linq.Enumerable]::Range(0, $tags.length) | %{ $tags.item($_) })
        }
    }
    
    # No matching nodes were found
    if ($tags -eq $null) {
        return
    }

    # Find elements recursively with the remaining query
    return $tags | %{ Query-DomElements -Dom $_ -Query $Query }
}
