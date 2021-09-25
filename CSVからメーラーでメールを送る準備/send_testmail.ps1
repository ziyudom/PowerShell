
#タイトルとアドレスからメールを生成する
$BR = "%0D%0A" #改行コード
#現在ディレクトリ取得
$CurrentDir = Split-Path $MyInvocation.MyCommand.Path

#メールbody用のテキストファイル取得
$txtFilePath = ($CurrentDir)+"\BODYテンプレート.txt"
#送信リストのCSVファイルパス
$filePath = ($CurrentDir)+"\サンプル.csv"

function MAKE_MAIL($title,$address,$bodyTxt) {
 
    $Subject = $title
    $MailTo = $address
    $MailBody = $bodyTxt

    # mailto のリンク形式の文字列を生成
    $EncodeTxt = 'mailto:' + $MailTo + `
                    '?subject=' + ($Subject) + `
                    '&body=' + ($MailBody)
                
    Start-Process $EncodeTxt
}

#テキストファイルを読み込む
function READ_TEXT($txtFilePath){

    $f = (Get-Content -Encoding UTF8 $txtFilePath) -as [string[]]

    #返却用テキスト
    $rtnText = ""
    #一行ごとに改行コードつける
    foreach ($l in $f) {
      $rtnText = $rtnText + $l + $BR
    }
    
    return $rtnText
}
 
#テキストファイルから読み込む
$tempBodyText = READ_TEXT $txtFilePath

#CSV読み込み(UTF8) ※EXCELで作成してメモ帳で名前を付けて保存からUTF8に変換する
$products = Import-Csv $filePath  -Encoding UTF8
$products | Format-Table

#CSV分ぐるぐる回す
Import-Csv $filePath | ForEach-Object {

  #テキスト差し替え
  $bodyText = $tempBodyText.Replace("{{text}}",$_.text)
  #メールを作成する
  MAKE_MAIL $_.title $_.address $bodyText

}

