Option Explicit
' 作成から一定期間経ったものを自動的に削除する
' 2017/10/17 yo16


' ---------------------------
' 設定
' ---------------------------

' ログファイルのフォルダ
Dim logDir : logDir = "H:\コールセンター共有\911_ツール\一時フォルダを整理\生存期間1ヶ月"


' 起動時の引数から、下記の情報を取得
Dim objArgs : Set objArgs = WScript.Arguments
If( objArgs.Count < 3 )Then
	MsgBox _
		"起動引数が必要です。" & vbCrLf & vbCrLf & _
		"1.生存期間日数" & vbCrLf & _
		"2.削除対象フォルダのフルパス" & vbCrLf & _
		"3.削除移動フラグ [ 1:削除 | 2:移動 ]" & vbCrLf & _
		"4.移動先フォルダ(移動時のoption)", _
		vbOkOnly + vbCritical, _
		"一定期間後に削除"
	WScript.Quit
End If
If( (objArgs(2) = 2) And (objArgs.Count < 4) )Then
	MsgBox _
		"移動を指示したときは第４引数が必要です。" & vbCrLf & vbCrLf & _
		"1.生存期間日数" & vbCrLf & _
		"2.削除対象フォルダのフルパス" & vbCrLf & _
		"3.削除移動フラグ [ 1:削除 | 2:移動 ]" & vbCrLf & _
		"4.移動先フォルダ(移動時のoption)", _
		vbOkOnly + vbCritical, _
		"一定期間後に削除"
	WScript.Quit
End If
' 生存日数
Dim lifeSpan : lifeSpan = Int(objArgs(0))
' 処理対象フォルダのフルパス
Dim targetDir : targetDir = objArgs(1)
' 削除移動フラグ [ 1:削除 | 2:移動 ]
Dim deleteFlag : deleteFlag = objArgs(2)
' 移動の場合の移動先フォルダ
Dim movetoDir : movetoDir = ""
If( objArgs.Count >= 4 )Then
	movetoDir = objArgs(3)
	If( Right(movetoDir, 1) <> "\" )Then
		movetoDir = movetoDir & "\"
	End If
End If

'msgbox "lifeSpan:"&lifeSpan & vbCrLf & _
'	"targetDir:"& targetDir & vbCrLf & _
'	"deleteFlag:"& deleteFlag & vbCrLf & _
'	"movetoDir:"& movetoDir


' ---------------------------
' 処理開始
' ---------------------------
Dim objFs
Set objFs = CreateObject("Scripting.FileSystemObject")
Dim objFolder
Set objFolder = objFs.GetFolder(targetDir)

' 今日の日にちを取得
Dim dtToday
dtToday = Date

' すべてのサブフォルダに対して更新日を確認して、条件に一致するなら削除する
Dim colSubfolders
Set colSubfolders = objFolder.SubFolders
Dim objSubfolder
For Each objSubfolder in colSubfolders
	If( DateDiff( "d", objSubfolder.DateLastModified, dtToday ) > lifeSpan )Then
		' 生存期間を超える日数の差がある
		' → 削除 or 移動
		If( deleteFlag = 1 )Then
			objFs.DeleteFolder objSubfolder.Path
		Else
'msgbox objSubfolder.Path &"/" &movetoDir
			objFs.MoveFolder objSubfolder.Path,movetoDir
		End If
	End If
Next
Set objSubFolder = Nothing
Set colSubfolders = Nothing

' すべてのファイルに対して更新日を確認して、条件に一致するなら削除する
Dim colFiles
Set colFiles = objFolder.Files
Dim objFile
For Each objFile in colFiles
	If( objFile.Name <> WScript.ScriptName )Then	' 自分自身を除く
		If( DateDiff( "d", objFile.DateLastModified, dtToday ) > lifeSpan )Then
			' 生存期間を超える日数の差がある
			' → 削除 or 移動
			If( deleteFlag = 1 )Then
				objFs.DeleteFile objFile.Path
			Else
'msgbox objFile.Path & "/" & movetoDir
				objFs.MoveFile objFile.Path, movetoDir
			End If
		End If
	End If
Next
Set objFile = Nothing
Set colFiles = Nothing


Set objFolder = Nothing



' めんどくさいからべた書き
'objFs.CopyFile ".\★★このフォルダは１週間くらいで勝手に消します","H:\コールセンター共有\999_一時フォルダ\★★このフォルダは１週間くらいで勝手に消します"
'objFs.CopyFile ".\★★消されたら困るファイルは個人へ","H:\コールセンター共有\999_一時フォルダ\★★消されたら困るファイルは個人へ"
Set objFs = Nothing

msgbox "end"
