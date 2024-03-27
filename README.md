W 案件用ツール  
[Wiki](https://nanairoware.nanairo-inc.jp/member/tasks/wiki/project/FE-%E3%83%AF%E3%82%A4%E3%82%BA210101?id=237)

# 概要
* ExtractVariableFromExcel は、`Excel`と 検索文字列を入力すると、検索文字列が使用されているセルの値を抽出してくれるツールです。


# 使い方
* 実行ファイルである、`ExtractVariableFromExcel.exe` に`Excel`ファイルを`Drag & Drop`してください。
* 読み込める拡張子は、`.xlsx, .xlsm, .xltx, .xltm` です。
* ファイルを読み込むと、検索文字列の入力を促されるので、入力して `Enter` を押してください。


# 仕様
* 複数ファイルを読み込めます
* 読み込める拡張子以外のファイルを引数として渡すと、そのファイルはスキップされます。
* 抽出結果は、`Search.xlsx`として、デスクトップに出力されます。
