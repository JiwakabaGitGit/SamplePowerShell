# 資格と研修の出力先のファイルオープン
$qualificationFile = "shikaku.csv"
$trainingFile = "kensyu.csv"

# 読み込み開始行
$qualificationRowNumber = 4
$trainingRowNumber = 4

$qualificationSheet = "AAA"
$trainingSheet = "BBB"

$shikakWriter = New-Object System.IO.StreamWriter($qualificationFile, $false, [System.Text.Encoding]::GetEncoding("sjis"))
$traiinigWriter = New-Object System.IO.StreamWriter($trainingFile, $false, [System.Text.Encoding]::GetEncoding("sjis"))
# ヘッダー追加
$shikakWriter.WriteLine("社員番号,年,資格名")
$traiinigWriter.WriteLine("社員番号,年,研修名")

# Excelファイルの一覧を取得
$files = Get-Item *.xlsx

# Excelファイルをチェックし
foreach($file in $files){
  Write-Host($file)

  # ファイル名から社員番号取得
  $employeeId =  $file.Name.Substring(0, 6);

  # Excelファイルの有無チェック
  if (!(Test-Path $file)) {
    Write-Host $file not found.
    exit 1
  }
  
  try {
      $file = (Get-ChildItem $file).FullName
      $excel = New-Object -ComObject Excel.Application
      $excel.Visible = $false    # <- Excel非表示
      $book = $excel.Workbooks.Open($file)

      # 資格情報のファイル出力
      $sheet1 = $excel.Worksheets.Item($qualificationSheet)   # <- 資格情報のシートを選択
      $line = $qualificationRowNumber

      $item1 = $sheet1.Cells.Item($line,1).Text
      $item2 = $sheet1.Cells.Item($line,2).Text

      while ( ($item1 -ne "") -and ($item2 -ne "")) {
          # Write-Host "$line,$employeeId,$item1,$item2"    # <- 取得した情報を表示
          $shikakWriter.WriteLine("$employeeId,$item1,$item2") # <- 取得した情報を書き込み

          $line++
          $item1 = $sheet1.Cells.Item($line,1).Text
          $item2 = $sheet1.Cells.Item($line,2).Text
      }

      # 研修情報のファイル出力
      $sheet2 = $excel.Worksheets.Item($trainingSheet)   # <- 研修情報のシートを選択
      $line = $trainingRowNumber

      $item1 = $sheet2.Cells.Item($line,1).Text
      $item2 = $sheet2.Cells.Item($line,2).Text

      while ( ($item1 -ne "") -and ($item2 -ne "")) {
          # Write-Host "$line,$employeeId,$item1,$item2"    # <- 取得した情報を表示
          $traiinigWriter.WriteLine("$employeeId,$item1,$item2") # <- 取得した情報を書き込み

          $line++
          $item1 = $sheet2.Cells.Item($line,1).Text
          $item2 = $sheet2.Cells.Item($line,2).Text
      }

  } catch [Exception] {
      $exp = $error[0].ToString()
      Write-Host $exp
  } finally {
      # Excelクローズ
      $excel.Quit()
      # 変数の破棄
      $book = $null
      $excel = $null
      [GC]::Collect()
  }

  # パラメタチェック
  if ($file -eq $null) {
    $thie = $MyInvocation.MyCommand.Name
    "Usage: $this <Excel file>"
    exit 1
  }
}

# ファイルクローズ
$shikakWriter.Close()
$traiinigWriter.Close()
