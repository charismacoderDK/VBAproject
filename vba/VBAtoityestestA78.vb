Sub SaveVbaScriptsToGithub()
'�`�N�ƶ�
'�ϥΫe�A����JAPI������T�A�j�MToken�C

Dim username As String 'Step 01 : �w�q�ܼ�
Dim repo_name As String
Dim file_name As String
Dim access_token As String
Dim payload As String


Dim xml_obj As MSXML2.XMLHTTP60 'Step 02 : �إ�XML����AHTTP���q�����n����C



Dim VBAEditor As VBIDE.VBE 'Step 03 : �إ�VBIDE�N�{����r�ơA�����w���ʥ����]�w�����C
Dim VBProj As VBIDE.VBProject
Dim VBCodeMod As VBIDE.CodeModule
Dim VBRawCode As String



Set VBAEditor = Application.VBE 'Step 04 : �`�N!!�n���������w����


'Step 5: ���PERSONAL.XLSB
Set VBProj = VBAEditor.VBProjects(1)

'Step06 : ����Ҳո̭��W�� XXXX�� Model , �ڭ̭n�W�Ǫ�Code�C
Set VBCodeMod = VBProj.VBComponents.Item("General_Format").CodeModule

'Step07 : �p�� XXXX Model�̭���code ��ơC
VBRawCode = VBCodeMod.Lines(startline:=1, Count:=VBCodeMod.CountOfLines)

'Debug.Print VBRawCode


' �N�W�� VBRawCode �� String �ഫ�� Encode�Ҧ�
RawCodeEncoded = EncodeBase64(text:=VBRawCode)

Debug.Print "�ഫ���᪺�{���X : " + RawCodeEncoded

'xml
Set xml_obj = New MSXML2.XMLHTTP60

'�걵API �ж�J��T
        base_url = "https://api.github.com/repos/"
        repo_name = "" 'Ex : "VAproject/"
        username = "" 'Ex :  "charismacoderDK/"
        file_name = "" ' Ex : "vba/test4.vb"
        access_token = ""
' �զ� URL
        full_url = base_url + username + repo_name + "contents/" + file_name + "?ref=main"
        Debug.Print full_url
        
        xml_obj.Open bstrMethod:="PUT", bstrURL:=full_url, varAsync:=True
        
        'set the headers.
        xml_obj.setRequestHeader bstrHeader:="Accept", bstrvalue:="application/vnd.github.v3+json"
        xml_obj.setRequestHeader bstrHeader:="Authorization", bstrvalue:="token " + access_token
        
        payload = "{""message"": "" This is my message3"", ""content"":"""
        payload = payload + Application.Clean(RawCodeEncoded)
        payload = payload + """}"
        
        xml_obj.send varBody:=payload
        
        'wait till it is finish.
        
        While xml_obj.readyState <> 4
            DoEvents
        Wend
        
        Debug.Print "Full URL: " + full_url
        Debug.Print "STATUS TEXT : " + xml_obj.statusText
        Debug.Print "PAYLOAD: " + payload
        
    
    





End Sub

Function EncodeBase64(text As String) As String

'�w�q�榡
Dim arrData() As Byte

'1. �ϥ�Document Object Model �NHTML �w�q������,'�קK�����y�����~�A�ɶq�ϥ� 6.0
'2. �n�A�d�@�UNode
Dim objXML As MSXML2.DOMDocument60
Dim objNode As MSXML2.IXMLDOMElement

'��r�榡 -> Unicode format  || vbFromUnicode : �N�r��q Unicode �ഫ���t�Ϊ��w�]�r�X���C
arrData = StrConv(text, vbFromUnicode)

'�]�w����
Set objXML = New MSXML2.DOMDocument60
Set objNode = objXML.createElement("b64") 'tag name


'�w�qNode ����Ʈ榡
objNode.DataType = "bin.base64"

'���w Node ��
objNode.nodeTypedValue = arrData

'���o��D�`���n
'�����N�ťը��N
' ASCII�@chr(10)�G����Ÿ� = vblf
'�������դ@�U
EncodeBase64 = Replace(objNode.text, vbLf, "")
'EncodeBase64 = objNode.text
'����O����

Set objNode = Nothing
Set objXML = Nothing

End Function


'Comment 02:
'=====================
'
'HTML�PXML�̥D�n�����O�D���e�̥D�n�O�ΨӼ��g�����Ϊ��y���A'�B��Html�y��(����)���O���y�Τ@���A�z�L�k�۩w���ҡA�u���ܧ������ݩʡC'�ӫ�̳̥D�n���\��O�Ψӡu��ƶǻ��v�ΡA�ҦpA�����i�N�n���ɥX�Ӫ����(�p�̷s�T���β��~��T�K��)�A'�নXML�榡��B�����i�H����Ū���ΤޥΡA�]���ϥΪ̥i�ۦ�w�q����(tags)�W�٤ε��c�A�H�Q�ޥΪ̿��ѵ��c�θ�Ƥ��e�C
'=====================
'testteeeee