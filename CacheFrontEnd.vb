Imports System
Imports System.Text
Imports System.Web
Imports System.IO
Imports System.Net
Imports System.Security.Cryptography
Imports System.Configuration.ConfigurationManager
Imports System.Linq
Imports Microsoft.WindowsAzure
Imports Microsoft.WindowsAzure.StorageClient
Imports System.Threading

Public Class CacheFrontEnd
    Implements IHttpModule

    Dim StartDT As DateTime = DateTime.MinValue 'Start time of module
    Dim CacheStartDT As DateTime = DateTime.MinValue 'Start time of cache operations
    Dim CacheEndDT As DateTime = DateTime.MinValue 'End time of cache operations

    Public Sub Init(ByVal app As HttpApplication) Implements IHttpModule.Init
        AddHandler app.BeginRequest, AddressOf Me.OnBeginRequest
        AddHandler app.ResolveRequestCache, AddressOf Me.OnCacheRequest
    End Sub

    Public Sub Dispose() Implements IHttpModule.Dispose

    End Sub

    Public Sub OnBeginRequest(ByVal s As Object, ByVal e As EventArgs)
        'Record the startup time of the module
        StartDT = DateTime.Now


    End Sub

    Public Sub OnCacheRequest(ByVal s As Object, ByVal e As EventArgs)
        Dim app As HttpApplication = CType(s, HttpApplication)

        'Only execute if the request is for a PHP file
        If app.Context.Request.Url.GetLeftPart(UriPartial.Path).EndsWith(".php") Then
            'Construct URL in the same manner as the Project Nami (WordPress) Plugin
            Dim URL As String = ""
            Dim scheme As String = ""
            If app.Context.Request.IsSecureConnection Then
                scheme = "https://"
            Else
                scheme = "http://"
            End If
            URL = scheme & app.Context.Request.ServerVariables("HTTP_HOST") & app.Context.Request.ServerVariables("REQUEST_URI")
            Dim NewURI As New Uri(URL)
            URL = NewURI.GetLeftPart(UriPartial.Query)

            'Determine if the User Agent is available
            If Not IsNothing(app.Context.Request.ServerVariables("HTTP_USER_AGENT")) Then
                'If Project Nami (WordPress) thinks this is a mobile device, salt the URL to generate a different key
                If wp_is_mobile(app.Context.Request.ServerVariables("HTTP_USER_AGENT")) Then
                    URL &= "|mobile"
                End If
            End If


            'Generate key based on the URL via MD5 hash
            Dim MD5Hash As String = getMD5Hash(URL)

            'Check cookies and abort if either user is logged in or the Project Nami (WordPress) Plugin has set a commenter cookie on this user for this page
            If Not IsNothing(app.Context.Request.Cookies) Then
                For Each ThisCookie As String In app.Context.Request.Cookies.AllKeys
                    If Not IsNothing(ThisCookie) Then
                        If ThisCookie.ToLower.Contains("wordpress_logged_in") Or ThisCookie.ToLower.Contains("comment_post_key_" & MD5Hash.ToLower) Then
                            Exit Sub
                        End If
                    End If
                Next
            End If

            'Record the start time of cache operations
            CacheStartDT = DateTime.Now

            'Randomized cleanup function
            If CInt(Math.Ceiling(Rnd() * 200)) = 42 Then
                System.Threading.ThreadPool.QueueUserWorkItem(AddressOf ContainerCleanUp)
            End If

            'Set up connection to the cache
            Dim ThisStorageAccount As CloudStorageAccount = CloudStorageAccount.Parse("DefaultEndpointsProtocol=http;AccountName=" & System.Configuration.ConfigurationManager.AppSettings("ProjectNamiBlobCache.StorageAccount") & ";AccountKey=" & System.Configuration.ConfigurationManager.AppSettings("ProjectNamiBlobCache.StorageKey"))
            Dim ThisBlobClient As CloudBlobClient = ThisStorageAccount.CreateCloudBlobClient
            Dim ThisContainer As CloudBlobContainer = ThisBlobClient.GetContainerReference(System.Configuration.ConfigurationManager.AppSettings("ProjectNamiBlobCache.StorageContainer"))
            Dim ThisBlob As CloudBlob

            'Attempt to access the blob
            Try
                ThisBlob = ThisContainer.GetBlobReference(MD5Hash)
            Catch ex As Exception
                Exit Sub
            End Try

            'Fetch metadata for the blob.  If fails, the blob is not present
            Try
                ThisBlob.FetchAttributes()
            Catch ex As Exception
                Exit Sub
            End Try

            'Check the TTL of the blob and delete it if it has expired
            If ThisBlob.Properties.LastModifiedUtc.AddSeconds(ThisBlob.Metadata("Projectnamicacheduration")) < DateTime.UtcNow Then 'Cache has expired
                ThisBlob.Delete()
                Exit Sub
            Else
                'Determine if Proactive mode is enabled
                If Not IsNothing(System.Configuration.ConfigurationManager.AppSettings("ProjectNamiBlobCache.Proactive")) Then
                    If System.Configuration.ConfigurationManager.AppSettings("ProjectNamiBlobCache.Proactive") = "1" Then
                        'Determine if the blob will expire within the next 20% of its total TTL
                        If ThisBlob.Properties.LastModifiedUtc.AddSeconds(ThisBlob.Metadata("Projectnamicacheduration")) < DateTime.UtcNow.AddSeconds(ThisBlob.Metadata("Projectnamicacheduration") * 0.2) Then 'Extend the cache duration and let the current request through
                            'Update blob metadata to reset the LastModifiedUtc and allow this request through in an attempt to reseed the cache
                            ThisBlob.Metadata("Projectnamicacheduration") = ThisBlob.Metadata("Projectnamicacheduration") - 1
                            ThisBlob.SetMetadata()
                            Exit Sub
                        End If
                    End If
                End If

                'If we've gotten this far, then we both have something to serve from cache and need to serve it, so get it from blob storage
                Dim CacheString As String = ThisBlob.DownloadText()

                'If the blob is empty, delete it
                If CacheString.Trim.Length = 0 Then
                    ThisBlob.Delete()
                    Exit Sub
                End If

                'Record the end time of cache operations
                CacheEndDT = DateTime.Now

                'Determine if Debug mode is enabled
                If Not IsNothing(System.Configuration.ConfigurationManager.AppSettings("ProjectNamiBlobCache.Debug")) Then
                    If System.Configuration.ConfigurationManager.AppSettings("ProjectNamiBlobCache.Debug") = "1" Then
                        'Calculate the milliseconds spent until cache operations began, and until completion
                        Dim CacheStartTS As TimeSpan = CacheStartDT - StartDT
                        Dim CacheEndTS As TimeSpan = CacheEndDT - StartDT
                        'Insert debug data before the closing HEAD tag
                        CacheString = CacheString.Replace("<" & Chr(47) & "head>", "<!-- CacheStart " & CacheStartTS.TotalMilliseconds & " CacheEnd " & CacheEndTS.TotalMilliseconds & " -->" & vbCrLf & "<!-- Key " & MD5Hash & " ServerVar " & URL & " Rewrite " & app.Context.Request.Url.GetLeftPart(UriPartial.Query) & " -->" & vbCrLf & "<" & Chr(47) & "head>")
                    End If
                End If

                'Set 200 status, MIME type, and write the blob contents to the response
                app.Context.Response.StatusCode = 200
                app.Context.Response.Write(CacheString)
                If app.Context.Request.ServerVariables("REQUEST_URI").ToLower.EndsWith(".xml") Or CacheString.ToLower.StartsWith("<?xml") Then
                    app.Context.Response.ContentType = "application/xml"
                Else
                    app.Context.Response.ContentType = "text/html"
                End If

                'Notify IIS we are done and to abort further operations
                app.CompleteRequest()
            End If
        End If
    End Sub

    Function getMD5Hash(ThisString As String) As String
        Dim md5Obj As New MD5CryptoServiceProvider
        Dim ResultBytes() As Byte = md5Obj.ComputeHash(System.Text.Encoding.ASCII.GetBytes(ThisString))
        Dim StrResult As String = ""

        Dim b As Byte
        For Each b In ResultBytes
            StrResult += b.ToString("x2")
        Next
        Return (StrResult)
    End Function

    Function wp_is_mobile(UserAgent As String) As Boolean
        'List of mobile User Agent lookups from Project Nami (WordPress) vars.php
        Dim MobileAgents() As String = {"Mobile", "Android", "Silk/", "Kindle", "BlackBerry", "Opera Mini", "Opera Mobi"}

        'If the User Agent contains any of the array elements, return True
        If MobileAgents.Any(Function(str) UserAgent.Contains(str)) Then
            Return True
        Else
            Return False
        End If

    End Function

    Sub ContainerCleanUp(state As Object)
        Try
            'Set up connection to the cache
            Dim ThisStorageAccount As CloudStorageAccount = CloudStorageAccount.Parse("DefaultEndpointsProtocol=http;AccountName=" & System.Configuration.ConfigurationManager.AppSettings("ProjectNamiBlobCache.StorageAccount") & ";AccountKey=" & System.Configuration.ConfigurationManager.AppSettings("ProjectNamiBlobCache.StorageKey"))
            Dim ThisBlobClient As CloudBlobClient = ThisStorageAccount.CreateCloudBlobClient
            Dim ThisContainer As CloudBlobContainer = ThisBlobClient.GetContainerReference(System.Configuration.ConfigurationManager.AppSettings("ProjectNamiBlobCache.StorageContainer"))
            Dim LogBlob As CloudBlob

            Try
                LogBlob = ThisContainer.GetBlobReference("LastCleanup")

                'Fetch metadata for the blob.  If fails, the blob is not present
                Try
                    LogBlob.FetchAttributes()
                Catch ex As Exception
                    LogBlob.UploadText("")
                    LogBlob.FetchAttributes()
                End Try

                If LogBlob.Properties.LastModifiedUtc.AddHours(1) < DateTime.UtcNow Then 'Cleanup needed
                    For Each ThisBlob As CloudBlob In ThisContainer.ListBlobs
                        ThisBlob.FetchAttributes()
                        If ThisBlob.Properties.LastModifiedUtc.AddHours(1) < DateTime.UtcNow Then 'Blob is old
                            ThisBlob.Delete()
                        End If
                    Next
                End If

            Catch ex As Exception
                Exit Sub
            End Try

        Catch ex As Exception
            'Do Nothing
        End Try
    End Sub
End Class

