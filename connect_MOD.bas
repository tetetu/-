Attribute VB_Name = "connect_MOD"
Option Compare Database
Option Explicit

Private CN As ADODB.Connection

Public Sub connect()
Dim openFl As Boolean
 
    '�I�[�v���t���O�I�t
    openFl = False
 
    If CN Is Nothing Then
        '�I�u�W�F�N�g�����݂��Ȃ��ꍇ
 
        '�I�u�W�F�N�g����
        Set CN = New ADODB.Connection
 
        '�I�[�v���t���O�I��
        openFl = True
    Else
        '�I�u�W�F�N�g�����݂���ꍇ
 
        '�R�l�N�V������Ԃ��N���[�Y�̏ꍇ
        If CN.State = adStateClosed Then
 
            '�I�[�v���t���O�I��
            openFl = True
        End If
    End If
 
    '�I�[�v���t���O���I���̏ꍇ
    If openFl = True Then
 
        Dim constr As String
 
'    OLE DB
    Const MYPROVIDERE = "Provider=SQLOLEDB;"
    Const MYSERVER = "Data Source=DESKTOP-NOL35GE;"                '�T�[�o�[,�|�[�g
    Const MYNINSYO = "Trusted_connection=no;"                    'Windows�F�؂̏ꍇ�i�u"Integrated Security=SSPI;"�v�ł��悢�j
    Const MYDATABASE = "Initial Catalog=Reward;"                '�ڑ�����f�[�^�x�[�X��
'    constr = MYPROVIDERE & MYSERVER & MYNINSYO & MYDATABASE       'Windows�F�؂̏ꍇ
    Const USER = "User Id=sb;"                               'SQL Server�F�؂̏ꍇ�̂ݎw��
    Const PSWD = "Password=3469;"                               'SQL Server�F�؂̏ꍇ�̂ݎw��
    constr = MYPROVIDERE & MYSERVER & MYDATABASE & USER & PSWD    'SQL Server�F�؂̏ꍇ
    Debug.Print constr

    
    

    Dim rs As New ADODB.Recordset
  
    CN.ConnectionString = constr
    CN.Open
    End If

       

   

    
End Sub

' �f�[�^�x�[�X�ւ̐ڑ�����������
Public Sub disconnect()
    CN.Close
    Set CN = Nothing
End Sub

' ������SQL�������s���AADODB.Recordset��Ԃ�
Public Function execute(sql As String) As ADODB.Recordset
    Dim rs As New ADODB.Recordset

    ' �^�C���A�E�g�ݒ� (20��)
    CN.CommandTimeout = 60 * 20

    ' �������ꂽ�s�����������b�Z�[�W�����ʃZ�b�g�̈ꕔ�Ƃ��ĕԂ���Ȃ��悤�ɂ���
    CN.execute ("SET NOCOUNT ON")

    ' �x�����b�Z�[�W�����ʃZ�b�g�̈ꕔ�Ƃ��ĕԂ���Ȃ��悤�ɂ���
    CN.execute ("SET ANSI_WARNINGS OFF")

    ' �I�[�o�[�t���[�����0���Z���ɂ�NULL��Ԃ�
    CN.execute ("SET ARITHABORT OFF")


    rs.Open sql, CN, adOpenStatic, adLockBatchOptimistic

    Do
        ' ���R�[�h�̑��삪�ł���I�u�W�F�N�g�Ⴕ���͎���RecordSet���Ƃꂸ�A�R�l�N�V��������ɂȂ����ꍇ�I��
        If rs.State = adStateOpen Or rs.ActiveConnection Is Nothing Then
            Exit Do
        End If
        Set rs = rs.NextRecordset()
    Loop

    Set execute = rs

    ' �ݒ�OFF
    CN.execute ("SET NOCOUNT OFF")
    CN.execute ("SET ANSI_WARNINGS ON")
    CN.execute ("SET ARITHABORT ON")
End Function

' �g�����U�N�V�������J�n����
Public Sub BeginTransaction()
    CN.BeginTrans
End Sub

' �g�����U�N�V�������R�~�b�g����
Public Sub CommitTransaction()
    CN.CommitTrans
End Sub

' �g�����U�N�V���������[���o�b�N����
Public Sub RollbackTransaction()
    CN.RollbackTrans
End Sub



