$users = Get-ChildItem C:\Users

foreach ($user in $users)
{
    $folder = "C:\users\" + $user + "\AppData\Roaming\DocsCorp\pdfDocs compareDocs\config" 
    Remove-Item -path $folder -Recurse -Force -ErrorAction silentlycontinue
}
