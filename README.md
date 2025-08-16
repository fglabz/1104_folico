# A002_folico

## 要件定義
次のスクリプトを作ろう。
目的：エクセルのアクティブセルの内容と同じ名前のフォルダを指定の場所（C:\Users\ken1r\projects\work）内につくり、決まったアイコンを設定する。スクリプトはVBAとPythonの協働とする

VBA部分
１．マクロを実行するとアクティブセルの内容を外部入力としてPythonを呼び出して処理を渡す（"C:\Users\ken1r\projects\work\A002_folico\src\set_folder_icon.py"）
２．CLI画面は「文字列」をPythonに渡したことを表示して２秒後に終了

Python部分
１．受け取った文字列にフォルダの名前に含められない文字列が含まれていれば、その旨を報告し、即時終了。
２．work内に同名のフォルダがあるかチェックし、存在していればその旨を報告し、当該フォルダをエクスプローラーで開いたのち、プログラム終了。
３．フォルダを作成、その中に　/src フォルダ、 /archive フォルダ　READEME.md　を作成する。
４．desktop.iniを作成し、（"C:\Users\ken1r\projects\work\A002_folico\Benjigarner-Rise-Folder-Summer-generic.256.ico"）アイコンを設定
５．エクスプローラーでフォルダを開いて終了

VBA と Python　を書いて