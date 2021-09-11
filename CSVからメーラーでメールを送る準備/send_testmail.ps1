
#タイトルとアドレスからメールを生成する

function MAKE_MAIL($name,$address) {
 
    # 変数定義
    $Subject = "テストメール" + ($name)
    $MailTo = $address
    $Body = @()
    $CrLf = "`r`n"

    echo $Subject

    $MailBody = "これはテストです"

    # mailto のリンク形式の文字列を生成
    $EncodeTxt = 'mailto:' + $MailTo + `
                    '?subject=' + ($Subject) + `
                    '&body=' + ($MailBody) 
                
    Start-Process $EncodeTxt


}
 
#現在ディレクトリ取得
$CurrentDir = Split-Path $MyInvocation.MyCommand.Path
echo $CurrentDir

#CSVファイルパス
$filePath = ($CurrentDir)+"\サンプル.csv"
echo $filePath

#CSV読み込み
$products = Import-Csv $filePath  -Encoding UTF8
$products | Format-Table

#ぐるぐる回す
Import-Csv $filePath | ForEach-Object {

  MAKE_MAIL $_.number $_.address

}

