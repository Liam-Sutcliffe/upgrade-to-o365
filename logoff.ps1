quser | Select-Object -Skip 1 | ForEach-Object {
    $id = ($_ -split ' +')[-5]
    logoff $id
}
