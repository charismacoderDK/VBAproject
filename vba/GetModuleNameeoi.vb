Sub SaveVbaScriptsToGithub()


Dim username As String 'Step 01 : 定義變數
Dim repo_name As String
Dim file_name As String
Dim access_token As String
Dim payload As String


Dim xml_obj As MSXML2.XMLHTTP60 'Step 02 : 建立XML物件，HTTP溝通之必要物件。



Dim VBAEditor As VBIDE.VBE 'Step 03 : 建立VBIDE將程式文字化，巨集安全性必須設定關閉。
Dim VBProj As VBIDE.VBProject
Dim VBCodeMod As VBIDE.CodeModule
Dim VBRawCode As String



Set VBAEditor = Application.VBE 'Step 04 : 注意!!要關閉巨集安全性


'Step 5: 抓取PERSONAL.XLSB
Set VBProj = VBAEditor.VBProjects(1)

'Step06 : 抓取模組裡面名為 XXXX的 Model , 我們要上傳的Code。
Set VBCodeMod = VBProj.VBComponents.Item("GetModuleName").CodeModule

'Step07 : 計算 XXXX Model裡面的code 行數。
VBRawCode = VBCodeMod.Lines(startline:=1, Count:=VBCodeMod.CountOfLines)

'Debug.Print VBRawCode


' 將上敘 VBRawCode 之 String 轉換成 Encode模式
RawCodeEncoded = EncodeBase64(text:=VBRawCode)

'Debug.Print "轉換之後的程式碼 : " + RawCodeEncoded

'xml
Set xml_obj = New MSXML2.XMLHTTP60

'串接API
        base_url = "https://api.github.com/repos/"
        repo_name = "VBAproject/"
        username = "charismacoderDK/"
        file_name = "vba/GetModuleName.vb"
        access_token = "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"
' 組成 URL
        full_url = base_url + username + repo_name + "contents/" + file_name + "?ref=main"
        'Debug.Print full_url
        
        xml_obj.Open bstrMethod:="PUT", bstrURL:=full_url, varAsync:=True
        
        'set the headers.
        xml_obj.setRequestHeader bstrHeader:="Accept", bstrvalue:="application/vnd.github.v3+json"
        xml_obj.setRequestHeader bstrHeader:="Authorization", bstrvalue:="token " + access_token
        
        payload = "{""message"": "" This is my message3"", ""content"":"""
        'Debug.Print "第一次 " + payload + vbLf
        payload = payload + Application.Clean(RawCodeEncoded)
        'Debug.Print "第二次 " + payload + vbLf
        payload = payload + """}"
        'Debug.Print "第三:" + vbLf + payload + vbLf
        
        'payload = ThisWorkbook.Sheets("Sheet2").Range("A6").Value
        'payload = Replace(ThisWorkbook.Sheets("Sheet2").Range("B6").Value, vbLf, "")
        Debug.Print payload
        xml_obj.send varBody:=payload
        
        'wait till it is finish.
        
        While xml_obj.readyState <> 4
            DoEvents
        Wend
        
        Debug.Print "Full URL: " + full_url + vbLf
        Debug.Print "STATUS TEXT : " + xml_obj.statusText + vbLf
        Debug.Print "PAYLOAD: " + payload + vbLf
        
    
    





End Sub

Function EncodeBase64(text As String) As String

'定義格式
Dim arrData() As Byte

'1. 使用Document Object Model 將HTML 定義成物件,'避免版本造成錯誤，盡量使用 6.0
'2. 要再查一下Node
Dim objXML As MSXML2.DOMDocument60
Dim objNode As MSXML2.IXMLDOMElement

'文字格式 -> Unicode format  || vbFromUnicode : 將字串從 Unicode 轉換成系統的預設字碼頁。
arrData = StrConv(text, vbFromUnicode)

'設定物件
Set objXML = New MSXML2.DOMDocument60
Set objNode = objXML.createElement("b64") 'tag name


'定義Node 的資料格式
objNode.DataType = "bin.base64"

'指定 Node 值
objNode.nodeTypedValue = arrData

'※這邊非常重要
'必須將空白取代
' ASCII　chr(10)：換行符號 = vblf
'必須測試一下
EncodeBase64 = Replace(objNode.text, vbLf, "")
'EncodeBase64 = objNode.text
'釋放記憶體

Set objNode = Nothing
Set objXML = Nothing

End Function


'Comment 02:
'=====================
'
'HTML與XML最主要的分別乃為前者主要是用來撰寫網頁用的語言，'且該Html語言(標籤)都是全球統一的，您無法自定標籤，只能變更其標籤屬性。'而後者最主要的功能是用來「資料傳遞」用，例如A網站可將要分享出來的資料(如最新訊息或產品資訊…等)，'轉成XML格式讓B網站可以直接讀取及引用，因此使用者可自行定義標籤(tags)名稱及結構，以利引用者辦識結構及資料內容。
'=====================


Sub select123()

Dim myWB      As Workbook
Dim myWS      As Worksheet
Dim myCell    As Range

Set myWB = ThisWorkbook             '*** see NOTE2
Set myWS = myWB.Sheets("Sheet2")
Set myCell = myWS.Range("A6")

A = myCell.Value

Debug.Print A

End Sub

Sub select1234()



A = ThisWorkbook.Sheets("Sheet2").Range("A6").Value


Debug.Print A

End Sub
