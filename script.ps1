Add-Type -AssemblyName System.Windows.Forms

#検索開始パスの指定
while ($true) {
    $filePath = Read-Host "開始パスを指定してください(例 D:\)"
    if((Test-Path -Path $filePath)){
        break
    }
    Write-Output "パスが間違っています"
}

#拡張子の指定
$fileExtension = Read-Host "拡張子を指定してください(例 .png .json)"
$fileExtension = $fileExtension -replace "\.",""

#ダイアログの設定
$saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
$saveFileDialog.Filter = "CSV Files (*.csv)|*.csv"
$saveFileDialog.Title = "CSVを保存する場所を選択してください"
$saveFileDialog.DefaultExt = "csv"
$saveFileDialog.FileName = $fileExtension+"出力"
$saveFileDialog.InitialDirectory = [Environment]::GetFolderPath('Desktop') # デスクトップを初期フォルダに設定

#ShowDialogの戻り値がOKだったらtrue
if($saveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK){
    
    Write-Output "実行します"
    #csvファイルを作成、保存
    Get-ChildItem -Path $filePath -Filter "*.$fileExtension" -Recurse | ForEach-Object{
        New-Object psobject -Property @{
            Name = $_.Name
            Path = $_.FullName
        }
        
    } | Export-Csv -NoTypeInformation -Encoding utf8 -Path $saveFileDialog.FileName

    
    Write-Output "CSVファイルが保存されました"
}else{
    Write-Output "保存がキャンセルされました"
}
Write-Output "処理を終了します"
Start-Sleep -Seconds 3