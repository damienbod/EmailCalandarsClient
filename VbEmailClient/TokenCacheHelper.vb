Imports System.IO
Imports System.Security.Cryptography
Imports Microsoft.Identity.Client

Namespace GraphEmailClient
    Module TokenCacheHelper
        Private ReadOnly CacheFilePath As String = System.Reflection.Assembly.GetExecutingAssembly().Location & ".msalcache.bin"
        Private ReadOnly FileLock As Object = New Object()

        Private Sub BeforeAccessNotification(ByVal args As TokenCacheNotificationArgs)
            SyncLock FileLock
                args.TokenCache.DeserializeMsalV3(If(File.Exists(CacheFilePath), ProtectedData.Unprotect(File.ReadAllBytes(CacheFilePath), Nothing, DataProtectionScope.CurrentUser), Nothing))
            End SyncLock
        End Sub

        Private Sub AfterAccessNotification(ByVal args As TokenCacheNotificationArgs)
            If args.HasStateChanged Then

                SyncLock FileLock
                    File.WriteAllBytes(CacheFilePath, ProtectedData.Protect(args.TokenCache.SerializeMsalV3(), Nothing, DataProtectionScope.CurrentUser))
                End SyncLock
            End If
        End Sub

        Friend Sub EnableSerialization(ByVal tokenCache As ITokenCache)
            tokenCache.SetBeforeAccess(AddressOf BeforeAccessNotification)
            tokenCache.SetAfterAccess(AddressOf AfterAccessNotification)
        End Sub
    End Module
End Namespace
