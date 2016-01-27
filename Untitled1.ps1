if($($z[$5][9]) -match ","){$sugg = $($z[$5][9]).Substring(0, $($z[$5][9]).IndexOf(','))}else{$sugg = $($z[$5][9])}
            Write-Host "$($z[$5][2]) / $($z[$5][9]) / $($z[$5][7]) / $($z[$5][8]) / $($z[$5][10])" -ForegroundColor Cyan
            Write-Host "Suggested Part name: '$sugg'"
            $part = Read-Host "Enter name"
            if($part -eq ""){$part = $sugg}
            $drawing = Read-Host "Enter Drawing number for $part"