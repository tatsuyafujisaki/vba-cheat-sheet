Option Explicit

Private Function EncodeBase64(ByVal s As String)
    With New DOMDocument60 'Microsoft XML, v6.0
        With .createElement("foo")
            .DataType = "bin.base64"
            .NodeTypedValue = GetBytes(s)
            EncodeBase64 = .Text
        End With
  End With
End Function

Private Function DecodeBase64(ByVal s As String)
    With New DOMDocument60 'Microsoft XML, v6.0
        With .createElement("foo")
            .DataType = "bin.base64"
            .Text = s
            DecodeBase64 = GetString(.NodeTypedValue)
        End With
  End With
End Function

Private Function GetBytes(ByVal s As String) As Byte()
    With New ADODB.Stream 'Microsoft ActiveX Data Object x.x Library
        .Open
        .Type = adTypeText
        .Charset = "_autodetect"
        .WriteText s
        .Position = 0
        .Type = adTypeBinary
        GetBytes = .Read()
        .Close
    End With
End Function

Private Function GetString(bytes() As Byte) As String
    With New ADODB.Stream 'Microsoft ActiveX Data Object x.x Library
        .Open
        .Type = adTypeBinary
        .Write bytes
        .Position = 0
        .Type = adTypeText
        .Charset = "_autodetect"
        GetString = .ReadText()
        .Close
    End With
End Function