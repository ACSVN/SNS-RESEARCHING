$strPath = (Convert-Path ..)

#$path = "C:\SNS-TWITTER\images\"                     # 出力先
$Path = $strPath + "\images\"
$file = "Cap.png"             # {0:000} 0番目の引数を3桁の0埋め数字とする。
Add-Type -A System.Windows.Forms      # このセッションに型を追加。
[Windows.Forms.Screen]::PrimaryScreen.Bounds |                   # ディスプレイサイズを取得し、
    % {[Drawing.Bitmap]::New($_.Width, $_.Height)} |             # そのサイズのBitmapデータ生成。
    sv b                                                         # Set-Variable 変数への代入。bにBitmapオブジェクトを代入。

[Drawing.Graphics]::FromImage($b).CopyFromScreen(0, 0, $b.Size)  # 描画オブジェクトで画面キャプチャして、Bitmapデータへ反映。

md -f $path                           # フォルダ作成。-f:force 既に存在する場合エラーを抑止。

1..1000 |                             # 1～1000まで
    % { $path + $file -f $_ }  |      # 繰り返し、フォーマット。c:\tmp\Cap_001.pngなどに変換。
    ? { (Test-Path $_) -eq $false } | # その内、ファイルが存在しないパスのみを取得。
    select -F 1 |                     # 最初の1件のみにフィルタして、
    % { $b.Save($_) }                 # そのパスにBitmapデータをファイル保存。