Attribute VB_Name = "mod_Utility"

Option Explicit
'Module: Utility
'Description:�q�Υ\��u��A�@���}��������eExcel�M�צs���ϥΡC

Private cApp As cApplication

Public Sub Start()
    Set cApp = New cApplication
End Sub

Public Function DetectCSVEncoding(ByVal filePath As String) As Long
    ' �ϥΤG�i��Ū���˥��i����R
    Dim stream As Object
    Dim bomBytes As Variant      ' ���� .Read �����G
    Dim sampleBytes() As Byte    ' �Ω�Ū���˥�
    Dim defaultEncoding As Long
    Dim detectedEncoding As Long
    Dim i As Long, byteValue As Integer, byteCount As Integer
    Dim isLikelyUTF8 As Boolean
    Dim bytesRead As Long
    
    defaultEncoding = 950 ' �w�]�� Big5
    detectedEncoding = defaultEncoding
    
    On Error GoTo DetectionError
    
    Set stream = CreateObject("ADODB.Stream")
    With stream
        .Type = 1 ' adTypeBinary - �l�רϥΤG�i��Ҧ�
        .Open
        .LoadFromFile filePath
        
        ' --- �B�J 1: �ˬd UTF-8 BOM ---
        If .Size >= 3 Then
            .Position = 0
            bomBytes = .Read(3) ' �{�b�i�H���T��ȵ� Variant
            
            ' �ˬd Variant ���}�C������
            If bomBytes(0) = &HEF And bomBytes(1) = &HBB And bomBytes(2) = &HBF Then
                detectedEncoding = 65001 ' UTF-8 with BOM
                Debug.Print "�˴��� UTF-8 �s�X (BOM)"
                GoTo Cleanup ' �w�T�w�A����M�z�B�J
            End If
        End If
        
        ' --- ����N�X�O������ ---
        ' --- �B�J 2: �p�G�S�� BOM�AŪ���G�i��˥��i��ҵo�����R ---
        .Position = 0 ' ���m��m
        If .Size > 0 Then
            ' Ū���e 4KB �Υ������e�@���˥�
            Dim sampleSize As Long
            sampleSize = WorksheetFunction.Min(4096, .Size)
            sampleBytes = .Read(sampleSize) ' Ū���G�i��ƾڨ�Ʋ�
            bytesRead = UBound(sampleBytes) + 1 ' ���Ū�����줸�ռ� (+1 �]���O 0-based)
        Else
            bytesRead = 0 ' ���ɮ�
        End If
        
        .Close ' Ū���˥��Y�i����
    End With
    Set stream = Nothing ' ���� stream ����

    ' --- �B�J 3: ���R�G�i��˥����e (�ҵo��) ---
    ' [�O������]
    If bytesRead > 0 Then
        isLikelyUTF8 = False
        byteCount = 0 ' �Ω�l�� UTF-8 �h�줸�էǦC
        
        For i = 0 To bytesRead - 1 ' �M��Ū�����줸�ռƲ� (0-based)
            byteValue = sampleBytes(i) ' �������o�줸�խ� (0-255)
            
            If byteCount = 0 Then ' �ˬd�O�_���h�줸�էǦC���_�l�줸��
                If byteValue >= &H80 Then ' �D ASCII �r��
                    If byteValue >= &HC2 And byteValue <= &HDF Then ' UTF-8 ���줸�էǦC�_�l (C2-DF)
                        byteCount = 1
                    ElseIf byteValue >= &HE0 And byteValue <= &HEF Then ' UTF-8 �T�줸�էǦC�_�l (E0-EF)
                        byteCount = 2
                    ElseIf byteValue >= &HF0 And byteValue <= &HF4 Then ' UTF-8 �|�줸�էǦC�_�l (F0-F4)
                        byteCount = 3
                    Else
                        ' �o�{�L�Ī��_�l�줸�� (�i�ण�O UTF-8)
                        isLikelyUTF8 = False
                        Debug.Print "�o�{�L�Ī� UTF-8 �_�l�줸��: " & Hex(byteValue) & " at position " & i & "�A�ɦV�D UTF-8"
                        Exit For ' ���A�~���ˬd
                    End If
                End If
            Else ' �ˬd�O�_�����Ī�����줸�� (80-BF)
                If byteValue >= &H80 And byteValue <= &HBF Then
                    byteCount = byteCount - 1 ' ���Ӥ@�ӫ���줸��
                    If byteCount = 0 Then
                        isLikelyUTF8 = True ' �ܤ֧��@�ӧ��㪺�h�줸�էǦC
                    End If
                Else
                    ' �o�{�L�Ī�����줸�� (�֩w���O UTF-8)
                    isLikelyUTF8 = False
                    Debug.Print "�o�{�L�Ī� UTF-8 ����줸��: " & Hex(byteValue) & " at position " & i & "�A�T�w�D UTF-8"
                    Exit For ' ���A�~���ˬd
                End If
            End If
        Next i
        
        ' �B�~�ˬd�G�p�G byteCount �b�����ɤ��� 0�A��ܧǦC������A�i�ण�O UTF-8
        If byteCount <> 0 Then
             isLikelyUTF8 = False
             Debug.Print "UTF-8 �ǦC�b�˥������B������A�ɦV�D UTF-8"
        End If
        
        ' �ھڤ��R���G�P�_
        If isLikelyUTF8 Then
             ' �p�G�b�˥����o�{�F���Ī� UTF-8 �h�줸�ռҦ��B�ǦC����
             detectedEncoding = 65001 ' UTF-8 without BOM
             Debug.Print "�ҵo���˴��G�ɦV UTF-8 �s�X (�L BOM)"
        Else
             ' �p�G�˥����S���o�{���㪺 UTF-8 �Ҧ��A�Ϊ̵o�{�F�L��/������Ҧ�
             detectedEncoding = defaultEncoding ' �O���w�] Big5
             Debug.Print "�ҵo���˴��G���o�{���T UTF-8 �Ҧ��A�ϥιw�]�s�X: " & defaultEncoding
        End If
        
    Else
        ' ���ɮשΥu�]�t ASCII�A�ϥιw�]�s�X
        Debug.Print "�ɮ׬��ũΥu�t ASCII�A�ϥιw�]�s�X: " & defaultEncoding
    End If

Cleanup:
    DetectCSVEncoding = detectedEncoding ' ��^�̲��˴����G
    Set stream = Nothing ' �T�O����
    On Error GoTo 0 ' ��_���`���~�B�z
    Exit Function

DetectionError:
    Debug.Print "Ū���ɮװ����s�X�ɵo�Ϳ��~: " & Err.Description & "�A�ϥιw�]�s�X: " & defaultEncoding
    DetectCSVEncoding = defaultEncoding ' �X���ɪ�^�w�]��
    Set stream = Nothing
    On Error GoTo 0
End Function
