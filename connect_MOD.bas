Attribute VB_Name = "connect_MOD"
Option Compare Database
Option Explicit

Private CN As ADODB.Connection

Public Sub connect()
Dim openFl As Boolean
 
    'オープンフラグオフ
    openFl = False
 
    If CN Is Nothing Then
        'オブジェクトが存在しない場合
 
        'オブジェクト生成
        Set CN = New ADODB.Connection
 
        'オープンフラグオン
        openFl = True
    Else
        'オブジェクトが存在する場合
 
        'コネクション状態がクローズの場合
        If CN.State = adStateClosed Then
 
            'オープンフラグオン
            openFl = True
        End If
    End If
 
    'オープンフラグがオンの場合
    If openFl = True Then
 
        Dim constr As String
 
'    OLE DB
    Const MYPROVIDERE = "Provider=SQLOLEDB;"
    Const MYSERVER = "Data Source=DESKTOP-NOL35GE;"                'サーバー,ポート
    Const MYNINSYO = "Trusted_connection=no;"                    'Windows認証の場合（「"Integrated Security=SSPI;"」でもよい）
    Const MYDATABASE = "Initial Catalog=Reward;"                '接続するデータベース名
'    constr = MYPROVIDERE & MYSERVER & MYNINSYO & MYDATABASE       'Windows認証の場合
    Const USER = "User Id=sb;"                               'SQL Server認証の場合のみ指定
    Const PSWD = "Password=3469;"                               'SQL Server認証の場合のみ指定
    constr = MYPROVIDERE & MYSERVER & MYDATABASE & USER & PSWD    'SQL Server認証の場合
    Debug.Print constr

    
    

    Dim rs As New ADODB.Recordset
  
    CN.ConnectionString = constr
    CN.Open
    End If

       

   

    
End Sub

' データベースへの接続を解除する
Public Sub disconnect()
    CN.Close
    Set CN = Nothing
End Sub

' 引数のSQL文を実行し、ADODB.Recordsetを返す
Public Function execute(sql As String) As ADODB.Recordset
    Dim rs As New ADODB.Recordset

    ' タイムアウト設定 (20分)
    CN.CommandTimeout = 60 * 20

    ' 処理された行数を示すメッセージが結果セットの一部として返されないようにする
    CN.execute ("SET NOCOUNT ON")

    ' 警告メッセージが結果セットの一部として返されないようにする
    CN.execute ("SET ANSI_WARNINGS OFF")

    ' オーバーフローおよび0除算時にはNULLを返す
    CN.execute ("SET ARITHABORT OFF")


    rs.Open sql, CN, adOpenStatic, adLockBatchOptimistic

    Do
        ' レコードの操作ができるオブジェクト若しくは次のRecordSetがとれず、コネクションが空になった場合終了
        If rs.State = adStateOpen Or rs.ActiveConnection Is Nothing Then
            Exit Do
        End If
        Set rs = rs.NextRecordset()
    Loop

    Set execute = rs

    ' 設定OFF
    CN.execute ("SET NOCOUNT OFF")
    CN.execute ("SET ANSI_WARNINGS ON")
    CN.execute ("SET ARITHABORT ON")
End Function

' トランザクションを開始する
Public Sub BeginTransaction()
    CN.BeginTrans
End Sub

' トランザクションをコミットする
Public Sub CommitTransaction()
    CN.CommitTrans
End Sub

' トランザクションをロールバックする
Public Sub RollbackTransaction()
    CN.RollbackTrans
End Sub



