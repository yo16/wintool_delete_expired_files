Option Explicit
' �쐬��������Ԍo�������̂������I�ɍ폜����
' 2017/10/17 yo16


' ---------------------------
' �ݒ�
' ---------------------------

' ���O�t�@�C���̃t�H���_
Dim logDir : logDir = "H:\�R�[���Z���^�[���L\911_�c�[��\�ꎞ�t�H���_�𐮗�\��������1����"


' �N�����̈�������A���L�̏����擾
Dim objArgs : Set objArgs = WScript.Arguments
If( objArgs.Count < 3 )Then
	MsgBox _
		"�N���������K�v�ł��B" & vbCrLf & vbCrLf & _
		"1.�������ԓ���" & vbCrLf & _
		"2.�폜�Ώۃt�H���_�̃t���p�X" & vbCrLf & _
		"3.�폜�ړ��t���O [ 1:�폜 | 2:�ړ� ]" & vbCrLf & _
		"4.�ړ���t�H���_(�ړ�����option)", _
		vbOkOnly + vbCritical, _
		"�����Ԍ�ɍ폜"
	WScript.Quit
End If
If( (objArgs(2) = 2) And (objArgs.Count < 4) )Then
	MsgBox _
		"�ړ����w�������Ƃ��͑�S�������K�v�ł��B" & vbCrLf & vbCrLf & _
		"1.�������ԓ���" & vbCrLf & _
		"2.�폜�Ώۃt�H���_�̃t���p�X" & vbCrLf & _
		"3.�폜�ړ��t���O [ 1:�폜 | 2:�ړ� ]" & vbCrLf & _
		"4.�ړ���t�H���_(�ړ�����option)", _
		vbOkOnly + vbCritical, _
		"�����Ԍ�ɍ폜"
	WScript.Quit
End If
' ��������
Dim lifeSpan : lifeSpan = Int(objArgs(0))
' �����Ώۃt�H���_�̃t���p�X
Dim targetDir : targetDir = objArgs(1)
' �폜�ړ��t���O [ 1:�폜 | 2:�ړ� ]
Dim deleteFlag : deleteFlag = objArgs(2)
' �ړ��̏ꍇ�̈ړ���t�H���_
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
' �����J�n
' ---------------------------
Dim objFs
Set objFs = CreateObject("Scripting.FileSystemObject")
Dim objFolder
Set objFolder = objFs.GetFolder(targetDir)

' �����̓��ɂ����擾
Dim dtToday
dtToday = Date

' ���ׂẴT�u�t�H���_�ɑ΂��čX�V�����m�F���āA�����Ɉ�v����Ȃ�폜����
Dim colSubfolders
Set colSubfolders = objFolder.SubFolders
Dim objSubfolder
For Each objSubfolder in colSubfolders
	If( DateDiff( "d", objSubfolder.DateLastModified, dtToday ) > lifeSpan )Then
		' �������Ԃ𒴂�������̍�������
		' �� �폜 or �ړ�
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

' ���ׂẴt�@�C���ɑ΂��čX�V�����m�F���āA�����Ɉ�v����Ȃ�폜����
Dim colFiles
Set colFiles = objFolder.Files
Dim objFile
For Each objFile in colFiles
	If( objFile.Name <> WScript.ScriptName )Then	' �������g������
		If( DateDiff( "d", objFile.DateLastModified, dtToday ) > lifeSpan )Then
			' �������Ԃ𒴂�������̍�������
			' �� �폜 or �ړ�
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



' �߂�ǂ���������ׂ�����
'objFs.CopyFile ".\�������̃t�H���_�͂P�T�Ԃ��炢�ŏ���ɏ����܂�","H:\�R�[���Z���^�[���L\999_�ꎞ�t�H���_\�������̃t�H���_�͂P�T�Ԃ��炢�ŏ���ɏ����܂�"
'objFs.CopyFile ".\���������ꂽ�獢��t�@�C���͌l��","H:\�R�[���Z���^�[���L\999_�ꎞ�t�H���_\���������ꂽ�獢��t�@�C���͌l��"
Set objFs = Nothing

msgbox "end"
