Attribute VB_Name = "Com"
Option Explicit

Function IsDirectory(ByRef FilePath As String, ByRef FileSystemObject As Object) As Boolean
  
  '引数   FilePath          対象ファイルが存在するディレクトリパス。
  '       FileSystemObject  ファイルシステムオブジェクト。
  '戻り値 指定したディレクトリが存在する場合、Trueを返す。
  '       指定したディレクトリが存在しない場合、Falseを返す。
  '処理概要 FilePathが空文字列または、存在しないディレクトリの場合Falseを返す。
  
  If Len(FilePath) = 0 Or Dir(FilePath, vbDirectory) = "" Then
   IsDirectory = False
  Else
   IsDirectory = True
  End If

End Function

Function GetNumberOfFiles(ByRef FilePath As String, ByRef FileSystemObject As Object) As Integer
  
  '引数   FilePath          対象ファイルが存在するディレクトリパス。
  '       FileSystemObject  ファイルシステムオブジェクト。
  '戻り値 指定したディレクトリに存在するファイル数を整数値で返す。
  '処理概要 FilePathに指定したディレクトリに存在する標準ファイル数を取得する。
 
 GetNumberOfFiles = FileSystemObject.GetFolder(FilePath).Files.Count

End Function

Function IsInputTextBox(ByRef Message As String, ByRef FilePath As String) As Boolean

  '引数   Message           InputTextBox出力時ヘッダーメッセージ
  '       FilePath          空文字変数FilePath。呼び出し元から参照渡しする。
  '戻り値 キャンセルボタンが押下された場合、Falseを返す。
  '       1文字以上入力された場合、Trueを返す。
  '処理概要 InputTextBoxが表示後、ユーザがキャンセルボタンを押下したか判定する。

  FilePath = InputBox(Message)
  'キャンセルボタン押下を判定する。
  If StrPtr(FilePath) = 0 Then
   IsInputTextBox = False
  Else
   IsInputTextBox = True
  End If

End Function

Function DelSpace(ByRef Text As String) As String

 '引数   Text              スペース混在チェック対象文字列
 '戻り値 半角全角スペースを取り除いた文字列を返す。
 '処理概要 引数Textに全角,半角スペースが含まれる場合空文字列に変換する。
 
 Text = Replace(Text, " ", "")
 Text = Replace(Text, "　", "")
 DelSpace = Text
 
End Function

Sub GetFileAttribute(ByRef FileSystemObject As Object, ByRef FilePath As String)

 '引数   FileSystemObject   ファイルシステムオブジェクト。
 '       FilePath           対象ファイルが存在するディレクトリパス。
 '戻り値
 '処理概要 対象ディレクトリ内に存在するファイルの属性を取得する。

 '変数宣言
 Dim Buf  As Object

 'フォルダ内ブックを順に取得
 'Attributes       :ファイルの属性を取得または設定します。
 'DateCreated      :ファイルの作成日時を取得します。
 'DateLastAccessed :最後にアクセスした日時を取得します。
 'DateLastModified :最後に更新された日時を取得します。
 'Drive            :指定したファイルが存在するドライブ文字（「C:」「D:」など）を取得します。
 'Name             :指定したファイルの名前を取得または設定します。
 'ParentFolder     :指定したファイルが格納されているフォルダ（Folder                                                                                                                                                      オブジェクト）を取得します。
 'Path             :ファイルのパスを取得します。
 'ShortName        :8.3形式のファイル名を取得します。
 'ShortPath        :8.3形式のパスを取得します。
 'Size             :ファイルの容量をバイト単位で取得します。
 'Type             :ファイルの種類をあらわす文字列を取得します。
 
 
 For Each Buf In FileSystemObject.GetFolder(FilePath).Files
 MsgBox "名前：" & Buf.Name & vbCrLf & _
        "サイズ：" & Buf.Size & vbCrLf & _
        "Attributes:" & Buf.Attributes & vbCrLf
 Next
End Sub

 '引数   FilePath           対象ファイルが存在するディレクトリパス。
 '処理概要 指定されたテキストファイルから１行ずつ読み込む処理をする。

Sub ReadTextFile(ByRef FilePath As String)
    Dim Buf As String
    Open FilePath For Input As #1
        Do Until EOF(1)
            Line Input #1, Buf
            IsWord (Buf)
        Loop
    Close #1
End Sub

 '引数
 '処理概要 設定したパラメータ数に応じて、指定行列値を読み込みワード配列を作成する。
 
Sub MakeWordList()
  '変数宣言
  Dim NumberOfParameters As Integer
  Dim Count As Integer
  Dim Word As Variant
  
  'WordLstに設定するパラメータ数を決める。
  NumberOfParameters = InputBox("パラメータ数を決定してください。")
  
  ReDim Preserve WordList(NumberOfParameters)
  ThisWorkbook.Activate
  For Count = 0 To NumberOfParameters - 1
    WordList(Count) = ThisWorkbook.Sheets(1).Cells(Count + 1, 1)
  Next Count
End Sub

 '引数   Text    テキストファイルから読み込んだ1行分の文字列
 '処理概要  テキストファイルから読み込んだ文字列に、MakeWordListのワードが含まれているか判定する。
 '使い方    ①呼び出し元でPublic WordList() As Stringを宣言する。

Sub IsWord(ByRef Text As String)
 Dim Word As Variant
 For Each Word In WordList
 If InStr(Text, Word) >= 1 And Word <> "" Then
  Debug.Print Text
  Debug.Print Len(Text)
 End If
 Next Word
End Sub

