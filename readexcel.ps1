# ���i�ƌ��C�̏o�͐�̃t�@�C���I�[�v��
$qualificationFile = "shikaku.csv"
$trainingFile = "kensyu.csv"

# �ǂݍ��݊J�n�s
$qualificationRowNumber = 4
$trainingRowNumber = 4

$qualificationSheet = "AAA"
$trainingSheet = "BBB"

$shikakWriter = New-Object System.IO.StreamWriter($qualificationFile, $false, [System.Text.Encoding]::GetEncoding("sjis"))
$traiinigWriter = New-Object System.IO.StreamWriter($trainingFile, $false, [System.Text.Encoding]::GetEncoding("sjis"))
# �w�b�_�[�ǉ�
$shikakWriter.WriteLine("�Ј��ԍ�,�N,���i��")
$traiinigWriter.WriteLine("�Ј��ԍ�,�N,���C��")

# Excel�t�@�C���̈ꗗ���擾
$files = Get-Item *.xlsx

# Excel�t�@�C�����`�F�b�N��
foreach($file in $files){
  Write-Host($file)

  # �t�@�C��������Ј��ԍ��擾
  $employeeId =  $file.Name.Substring(0, 6);

  # Excel�t�@�C���̗L���`�F�b�N
  if (!(Test-Path $file)) {
    Write-Host $file not found.
    exit 1
  }
  
  try {
      $file = (Get-ChildItem $file).FullName
      $excel = New-Object -ComObject Excel.Application
      $excel.Visible = $false    # <- Excel��\��
      $book = $excel.Workbooks.Open($file)

      # ���i���̃t�@�C���o��
      $sheet1 = $excel.Worksheets.Item($qualificationSheet)   # <- ���i���̃V�[�g��I��
      $line = $qualificationRowNumber

      $item1 = $sheet1.Cells.Item($line,1).Text
      $item2 = $sheet1.Cells.Item($line,2).Text

      while ( ($item1 -ne "") -and ($item2 -ne "")) {
          # Write-Host "$line,$employeeId,$item1,$item2"    # <- �擾��������\��
          $shikakWriter.WriteLine("$employeeId,$item1,$item2") # <- �擾����������������

          $line++
          $item1 = $sheet1.Cells.Item($line,1).Text
          $item2 = $sheet1.Cells.Item($line,2).Text
      }

      # ���C���̃t�@�C���o��
      $sheet2 = $excel.Worksheets.Item($trainingSheet)   # <- ���C���̃V�[�g��I��
      $line = $trainingRowNumber

      $item1 = $sheet2.Cells.Item($line,1).Text
      $item2 = $sheet2.Cells.Item($line,2).Text

      while ( ($item1 -ne "") -and ($item2 -ne "")) {
          # Write-Host "$line,$employeeId,$item1,$item2"    # <- �擾��������\��
          $traiinigWriter.WriteLine("$employeeId,$item1,$item2") # <- �擾����������������

          $line++
          $item1 = $sheet2.Cells.Item($line,1).Text
          $item2 = $sheet2.Cells.Item($line,2).Text
      }

  } catch [Exception] {
      $exp = $error[0].ToString()
      Write-Host $exp
  } finally {
      # Excel�N���[�Y
      $excel.Quit()
      # �ϐ��̔j��
      $book = $null
      $excel = $null
      [GC]::Collect()
  }

  # �p�����^�`�F�b�N
  if ($file -eq $null) {
    $thie = $MyInvocation.MyCommand.Name
    "Usage: $this <Excel file>"
    exit 1
  }
}

# �t�@�C���N���[�Y
$shikakWriter.Close()
$traiinigWriter.Close()
