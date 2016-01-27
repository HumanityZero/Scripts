$tree= $null
for($i=1; $i -le 97; $i++){

    $tree += "=IF(part_$i!C24="
    $tree += '"y","hardened",IF('
    $tree += "part_$i!C26="
    $tree += '"y","black oxid finish",'
    $tree += "IF(part_$i!C28="
    $tree += '"y","nitrided",'
    $tree += "IF(part_$i!C30="
    $tree += '"y","coated","soft"))))'
    $tree += "`n"


}
$tree | clip