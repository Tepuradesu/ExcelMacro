Attribute VB_Name = "Com"
Option Explicit

Function IsDirectory(ByRef FilePath As String, ByRef FileSystemObject As Object) As Boolean
  
  '����   FilePath          �Ώۃt�@�C�������݂���f�B���N�g���p�X�B
  '       FileSystemObject  �t�@�C���V�X�e���I�u�W�F�N�g�B
  '�߂�l �w�肵���f�B���N�g�������݂���ꍇ�ATrue��Ԃ��B
  '       �w�肵���f�B���N�g�������݂��Ȃ��ꍇ�AFalse��Ԃ��B
  '�����T�v FilePath���󕶎���܂��́A���݂��Ȃ��f�B���N�g���̏ꍇFalse��Ԃ��B
  
  If Len(FilePath) = 0 Or Dir(FilePath, vbDirectory) = "" Then
   IsDirectory = False
  Else
   IsDirectory = True
  End If

End Function

Function GetNumberOfFiles(ByRef FilePath As String, ByRef FileSystemObject As Object) As Integer
  
  '����   FilePath          �Ώۃt�@�C�������݂���f�B���N�g���p�X�B
  '       FileSystemObject  �t�@�C���V�X�e���I�u�W�F�N�g�B
  '�߂�l �w�肵���f�B���N�g���ɑ��݂���t�@�C�����𐮐��l�ŕԂ��B
  '�����T�v FilePath�Ɏw�肵���f�B���N�g���ɑ��݂���W���t�@�C�������擾����B
 
 GetNumberOfFiles = FileSystemObject.GetFolder(FilePath).Files.Count

End Function

Function IsInputTextBox(ByRef Message As String, ByRef FilePath As String) As Boolean

  '����   Message           InputTextBox�o�͎��w�b�_�[���b�Z�[�W
  '       FilePath          �󕶎��ϐ�FilePath�B�Ăяo��������Q�Ɠn������B
  '�߂�l �L�����Z���{�^�����������ꂽ�ꍇ�AFalse��Ԃ��B
  '       1�����ȏ���͂��ꂽ�ꍇ�ATrue��Ԃ��B
  '�����T�v InputTextBox���\����A���[�U���L�����Z���{�^�����������������肷��B

  FilePath = InputBox(Message)
  '�L�����Z���{�^�������𔻒肷��B
  If StrPtr(FilePath) = 0 Then
   IsInputTextBox = False
  Else
   IsInputTextBox = True
  End If

End Function

Function DelSpace(ByRef Text As String) As String

 '����   Text              �X�y�[�X���݃`�F�b�N�Ώە�����
 '�߂�l ���p�S�p�X�y�[�X����菜�����������Ԃ��B
 '�����T�v ����Text�ɑS�p,���p�X�y�[�X���܂܂��ꍇ�󕶎���ɕϊ�����B
 
 Text = Replace(Text, " ", "")
 Text = Replace(Text, "�@", "")
 DelSpace = Text
 
End Function

Sub GetFileAttribute(ByRef FileSystemObject As Object, ByRef FilePath As String)

 '����   FileSystemObject   �t�@�C���V�X�e���I�u�W�F�N�g�B
 '       FilePath           �Ώۃt�@�C�������݂���f�B���N�g���p�X�B
 '�߂�l
 '�����T�v �Ώۃf�B���N�g�����ɑ��݂���t�@�C���̑������擾����B

 '�ϐ��錾
 Dim Buf  As Object

 '�t�H���_���u�b�N�����Ɏ擾
 'Attributes       :�t�@�C���̑������擾�܂��͐ݒ肵�܂��B
 'DateCreated      :�t�@�C���̍쐬�������擾���܂��B
 'DateLastAccessed :�Ō�ɃA�N�Z�X�����������擾���܂��B
 'DateLastModified :�Ō�ɍX�V���ꂽ�������擾���܂��B
 'Drive            :�w�肵���t�@�C�������݂���h���C�u�����i�uC:�v�uD:�v�Ȃǁj���擾���܂��B
 'Name             :�w�肵���t�@�C���̖��O���擾�܂��͐ݒ肵�܂��B
 'ParentFolder     :�w�肵���t�@�C�����i�[����Ă���t�H���_�iFolder                                                                                                                                                      �I�u�W�F�N�g�j���擾���܂��B
 'Path             :�t�@�C���̃p�X���擾���܂��B
 'ShortName        :8.3�`���̃t�@�C�������擾���܂��B
 'ShortPath        :8.3�`���̃p�X���擾���܂��B
 'Size             :�t�@�C���̗e�ʂ��o�C�g�P�ʂŎ擾���܂��B
 'Type             :�t�@�C���̎�ނ�����킷��������擾���܂��B
 
 
 For Each Buf In FileSystemObject.GetFolder(FilePath).Files
 MsgBox "���O�F" & Buf.Name & vbCrLf & _
        "�T�C�Y�F" & Buf.Size & vbCrLf & _
        "Attributes:" & Buf.Attributes & vbCrLf
 Next
End Sub

 '����   FilePath           �Ώۃt�@�C�������݂���f�B���N�g���p�X�B
 '�����T�v �w�肳�ꂽ�e�L�X�g�t�@�C������P�s���ǂݍ��ޏ���������B

Sub ReadTextFile(ByRef FilePath As String)
    Dim Buf As String
    Open FilePath For Input As #1
        Do Until EOF(1)
            Line Input #1, Buf
            IsWord (Buf)
        Loop
    Close #1
End Sub

 '����
 '�����T�v �ݒ肵���p�����[�^���ɉ����āA�w��s��l��ǂݍ��݃��[�h�z����쐬����B
 
Sub MakeWordList()
  '�ϐ��錾
  Dim NumberOfParameters As Integer
  Dim Count As Integer
  Dim Word As Variant
  
  'WordLst�ɐݒ肷��p�����[�^�������߂�B
  NumberOfParameters = InputBox("�p�����[�^�������肵�Ă��������B")
  
  ReDim Preserve WordList(NumberOfParameters)
  ThisWorkbook.Activate
  For Count = 0 To NumberOfParameters - 1
    WordList(Count) = ThisWorkbook.Sheets(1).Cells(Count + 1, 1)
  Next Count
End Sub

 '����   Text    �e�L�X�g�t�@�C������ǂݍ���1�s���̕�����
 '�����T�v  �e�L�X�g�t�@�C������ǂݍ��񂾕�����ɁAMakeWordList�̃��[�h���܂܂�Ă��邩���肷��B
 '�g����    �@�Ăяo������Public WordList() As String��錾����B

Sub IsWord(ByRef Text As String)
 Dim Word As Variant
 For Each Word In WordList
 If InStr(Text, Word) >= 1 And Word <> "" Then
  Debug.Print Text
  Debug.Print Len(Text)
 End If
 Next Word
End Sub

