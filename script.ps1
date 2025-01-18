Add-Type -AssemblyName System.Windows.Forms

#�����J�n�p�X�̎w��
while ($true) {
    $filePath = Read-Host "�J�n�p�X���w�肵�Ă�������(�� D:\)"
    if((Test-Path -Path $filePath)){
        break
    }
    Write-Output "�p�X���Ԉ���Ă��܂�"
}

#�g���q�̎w��
$fileExtension = Read-Host "�g���q���w�肵�Ă�������(�� .png .json)"
$fileExtension = $fileExtension -replace "\.",""

#�_�C�A���O�̐ݒ�
$saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
$saveFileDialog.Filter = "CSV Files (*.csv)|*.csv"
$saveFileDialog.Title = "CSV��ۑ�����ꏊ��I�����Ă�������"
$saveFileDialog.DefaultExt = "csv"
$saveFileDialog.FileName = $fileExtension+"�o��"
$saveFileDialog.InitialDirectory = [Environment]::GetFolderPath('Desktop') # �f�X�N�g�b�v�������t�H���_�ɐݒ�

#ShowDialog�̖߂�l��OK��������true
if($saveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK){
    
    Write-Output "���s���܂�"
    #csv�t�@�C�����쐬�A�ۑ�
    Get-ChildItem -Path $filePath -Filter "*.$fileExtension" -Recurse | ForEach-Object{
        New-Object psobject -Property @{
            Name = $_.Name
            Path = $_.FullName
        }
        
    } | Export-Csv -NoTypeInformation -Encoding utf8 -Path $saveFileDialog.FileName

    
    Write-Output "CSV�t�@�C�����ۑ�����܂���"
}else{
    Write-Output "�ۑ����L�����Z������܂���"
}
Write-Output "�������I�����܂�"
Start-Sleep -Seconds 3