# PowerPoint-note-counter

PowerPoint のノートに記載した文字数をカウントする PowerShell スクリプト。

## TL;DR

### できること

`powerPointNoteCounter.ps1`（PowerShell スクリプト）を実行すると、PowerPoint ファイル（pptx または ppt）のノートに記載した内容を読み取り、ノートの全文字から取り消し線の文字を引いた数を標準出力する。さらに、PowerShell スクリプトがあるディレクトリ内に`note.txt`（ノートの全内容を出力したテキストファイル）と`strikes.txt`（ノートで取り消し線の内容を出力したテキストファイル）を作成する。

### 実行方法

PowerShell の実行のやり方がわからない場合は、[PowerShell を実行する方法](#powershell-を実行する方法)を参照。

PowerShell スクリプトの実行により標準入力で PowerPoint ファイル（pptx または ppt）のパスを入力するよう求められるので、ファイルの絶対パスまたは PowerShell スクリプトからみた相対パスを入力する（もしくは、実行方法によってはファイルのドラッグ＆ドロップでも可）する必要がある。

## PowerShell スクリプトの詳細

PowerPoint ファイルを ZIP 変換して解凍することで PoewrPoint ファイル内のノートに相当するファイルを取り出すことができる。そのファイルに対して、正規表現パターン検索でノートに記載した内容と取り消し線の内容を取得し、その文字数をカウント後に標準出力する。また、取得した文字もそれぞれ`note.txt`と`strikes.txt`に出力する。

## Appendix

### PowerShell を実行する方法

1. [ファイルの詳細ページ（このリンクを Ctrl キーを押しながらクリック。ファイルのダウンロードはされません）](https://github.com/sasaiyu/PowerPoint-note-counter/blob/main/powerPointNoteCounter.ps1)を開き、Code の内容を選択してクリップボードにコピーしてください。なお、[このリンクをクリックするとファイルをダウンロード](https://github.com/sasaiyu/PowerPoint-note-counter/raw/refs/heads/main/powerPointNoteCounter.ps1)することもできます。

1. メモ帳にクリップボードの内容をペーストして、ファイル名を`powerPointNoteCounter.ps1`にして、**エンコードを ASCII**にして保存してください（任意の名前でも問題ありませんが、以降の手順でファイル名を置き換えてください）。なお、**ファイルの拡張子は ps1 であること**に気をつけてください。

1.`powerPointNoteCounter.ps1`のファイルを右クリックして、PowerShell で実行を選択してください。選択してもウィンドウが開かない場合は、[こちら](#powershell-スクリプトを実行をしてもウィンドウが開かない)

1. PowerShell のウィンドウが開くので、PowerPoint ファイルのパスを入力するか、ファイルをドラッグ＆ドロップしてください。なお、パスは絶対パス（C:\から始まるパス）または、`powerPointNoteCounter.ps1`があるフォルダからみたパス（例えば、同じフォルダにある場合はファイル名のみ）を入力してください。

1. PowerShell の処理が実行されて、文字数が表示されれば完了です。確認後に Enter を押すか、ウィンドウを閉じれば終了します。

### トラベルシューティング

#### PowerShell スクリプトを実行をしてもウィンドウが開かない

スクリプトの実行方法に問題がないか確認してください。特に、[PowerShell スクリプトを右クリックして、PowerShell で実行](#powershell-を実行する方法)しても解消しない場合は **PowerShell の実行ポリシーが有効になっていない可能性があります。以下の手順により有効化することができます。**

以下の内容を確認してください。

- WIndows10 の場合

Windows の設定を開き、更新とセキュリティ>開発者向け>開発者モードをオンにする。

- Windows11 の場合

Windows10 の場合を参照して設定してください。それでも実行できない場合は、署名されていない PowerShell を実行できないようセキュリティポリシーが設定されている可能性があります。Windows の設定を開き設定の検索に PowerShell と入力したあとに、検索結果の署名なしでローカル PowerShell スクリプトを実行できるようにするをクリックし、設定をオンにしてください。

#### Windows メモ帳にコピーするときの注意点

Windows メモ帳で対応している文字コードは、UTF-8 か ASCII であるため、直接 SJIS の文字をペーストすることはできない。つまり、[GitHub](https://github.com/sasaiyu/PowerPoint-note-counter/blob/main/powerPointNoteCounter.ps1)にある Copy raw file ボタンをクリックしてコピーしても`powerPointNoteCounter.ps1`の文字コードは SJIS のため文字化けしてしまう。そのため、Code を選択して UTF8 で開いたメモ帳に （UTF8 で）貼り付けたあとに、ASCII で保存することで、SJIS として開くことができる。なお、ASCII で開いたメモ帳に SJIS を貼り付けると文字化けするので、最初は UTF8 でメモ帳を開く必要がある。
