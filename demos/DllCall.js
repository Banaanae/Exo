function NormalizePath(path) {
    var cc = DllCall("GetFullPathName", "str", path, "uint", 0, "ptr", 0, "ptr", 0, "uint")
    var buf = Buffer(cc*2)
    DllCall("GetFullPathName", "str", path, "uint", cc, "ptr", buf, "ptr", 0)
    return StrGet(buf)
}

MsgBox(NormalizePath('./demo.js'))