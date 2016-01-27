

$bookmark = @(@("145028-0043",300,$false),@("14599",800,$false),@("145025-9999",1300,$false),@("145029-1254",900,$false))




    for($5 = 0; $5 -lt $bookmark.count; $5++){



        Write-Progress -activity “$($bookmark[$5][0])” -id 43 -PercentComplete ((($5+1) / $bookmark.count)*100)
        
        for($b = 100; $b -le $bookmark[$5][1]; $b = $b + 100){

            Write-Progress -activity “Print” -id 44 -ParentId 43 -PercentComplete ((($b-50) / ($bookmark[$5][1]))*100) -status "$("{0:D4}" -f $b)/$($bookmark[$5][1])" -CurrentOperation "Job Card"
        start-sleep -m 500
                        Write-Progress -activity “Print” -id 44 -ParentId 43 -PercentComplete (($b / $bookmark[$5][1])*100) -status "$("{0:D4}" -f $b)/$($bookmark[$5][1])" -CurrentOperation "Drawing"
                        start-sleep -m 200
        }




    }
    Write-Progress -activity “Print” -id 44 -status "x" -Completed
    Write-progress -activity "x” -id 43 -status "x"  -Completed