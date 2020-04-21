Imports System.IO
Imports System.IO.File
Imports System.Text
Imports System.Security.AccessControl
Imports System.Security.Principal


Public Class Class1

    Public Sub Done()
        Console.WriteLine("Start A")

        Me.DoExclusiveProcess()

        Console.WriteLine("End A")
    End Sub



    Private Sub DoExclusiveProcess()
        Dim dateStr As String = Date.Now.ToString("yyyyMMddhhmmssfff")

        '作成するシステムミューテックスの名前。グローバルシステムミューテックスにする。
        Dim mutexName As String = "Global\" & My.Application.Info.AssemblyName
        'Mutexインスタンス
        Dim mutex As Threading.Mutex = Nothing
        'MutexインスタンスがWaitOneメソッドでシグナルを受信したかを示すフラグ
        Dim signaled As Boolean = False

        Try
            '================================================================================
            'Mutexインスタンス取得
            '================================================================================
            Try
                '既存のシステムミューテックスを開く
                mutex = Threading.Mutex.OpenExisting(mutexName, MutexRights.Synchronize Or MutexRights.Modify)
                'mutex = Threading.Mutex.OpenExisting(mutexName)

            Catch ex As Threading.WaitHandleCannotBeOpenedException
                '指定した名前のシステムミューテックスが存在しない場合、新規に作成する

                'MutexSecurityオブジェクトを既定値(全ユーザーの全アクセスを拒否)で作成
                Dim mutexSecurity As New MutexSecurity()

                '全ユーザに一致するSIDを取得
                Dim sid As New SecurityIdentifier(WellKnownSidType.WorldSid, Nothing)

                '上記のsidのユーザに以下のアクセス権を許可:
                '・MutexRights.Synchronize ... WaitOneメソッドでMutexを待機できるようにするため
                '・MutexRights.Modify      ... ReleaseMutexメソッドでMutexの所有権を解放できるようにするため
                Dim mutexAccessRule As New MutexAccessRule(sid, MutexRights.Synchronize Or MutexRights.Modify, _
                                                           AccessControlType.Allow)

                'MutexSecurityオブジェクトに上記のアクセス制御ルールを追加
                mutexSecurity.AddAccessRule(mutexAccessRule)

                '上記のアクセス制御セキュリティを適用したシステムミューテックスのインスタンスを作成(※初期所有権は付与しない)
                Dim createdNew As Boolean
                mutex = New Threading.Mutex(False, mutexName, createdNew, mutexSecurity)
                'mutex = New Threading.Mutex(False, mutexName)
            End Try


            '================================================================================
            'Mutexの所有権を取得
            '================================================================================
            '現在のインスタンスがシグナルを受け取るまでか、指定の時間(ms)が経過するまではブロック。
            'タイムアウトしたら処理終了。
            '※タイムアウトしMutexの所有権を取得していない状態で処理を進めるとプロセス間排他にならない。
            signaled = mutex.WaitOne(90000)
            If signaled = False Then
                'タイムアウトのためMutexの所有権を取得できなかった
                Console.WriteLine("タイムアウト.")
                Return
            End If
            '▼▼▼▼▼▼▼▼▼▼
            'プロセス間排他処理開始

        Catch ex As Exception
            Console.WriteLine("予期しないエラー:" & Environment.NewLine & ex.Message)
            Throw
        End Try


        Try
            For i As Integer = 1 To 10
                Console.WriteLine("Start A ... " & i.ToString("D2"))
                Threading.Thread.Sleep(5000)

                Dim fileName As String = "A_" & dateStr & "_" & i.ToString("D2") & ".txt"

                Try
                    Using fs As _
                            New FileStream(".\" & fileName, FileMode.Create, FileAccess.ReadWrite, FileShare.ReadWrite)
                        If fs IsNot Nothing Then
                            fs.Close()
                        End If
                    End Using
                Catch ex As Exception
                    Console.WriteLine(ex.ToString())
                End Try

                Console.WriteLine("End A ... " & i.ToString("D2"))
            Next

        Finally
            '================================================================================
            'Mutexの所有権を解放
            '※WaitOneメソッドでMutexの所有権取得したら、必ずReleaseMutexメソッドでMutexの所有権を解放する。
            '================================================================================
            If mutex IsNot Nothing AndAlso signaled = True Then
                Try
                    mutex.ReleaseMutex()
                    signaled = False
                    'プロセス間排他処理終了
                    '▲▲▲▲▲▲▲▲▲▲
                Catch ex As Exception
                    '呼び出し元のスレッドがミューテックスを所有していない場合。
                    '処理終了せずログ出力のみでOK。
                    Console.WriteLine("予期しないエラー:" & Environment.NewLine & ex.Message)
                End Try
            End If
        End Try

    End Sub

End Class
