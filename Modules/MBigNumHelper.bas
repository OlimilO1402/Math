Attribute VB_Name = "MBigNumHelper"
Option Explicit

'#[allow(dead_code)]
'pub fn same_sizer(s1:&String,s2:&String)->(String,String){
'
'    let mut s22=String::from(s2);
'    let mut s11=String::from(s1);
'
'    if s1.len()>=s2.len() {
'        s22=pad(s2,s1.len()-s2.len());
'    }else if s1.len()<=s2.len() {
'        s11=pad(s1,s2.len()-s1.len());
'    }
'
'    (s11,s22)
'
'
'}
Public Sub same_sizer(ByRef s1_inout As String, ByRef s2_inout As String)
    Dim l1 As Long: l1 = Len(s1_inout)
    Dim l2 As Long: l2 = Len(s2_inout)
    If l1 < l2 Then s1_inout = PadLeft(s1_inout, l2)
    If l2 < l1 Then s2_inout = PadLeft(s2_inout, l1)
End Sub

'helper.rs
'#[allow(dead_code)]
'pub fn is_greater_or_equal(s1:&String,s2:&String)->bool{
'
'    let (s11,s22)=same_sizer(s1, s2);
'    let mut is_greater_or_equal=true;
'    for i in 0..s11.len() {
'        if s11.chars().nth(i).unwrap() as i32 == s22.chars().nth(i).unwrap() as i32 {
'            continue
'        }
'        if s11.chars().nth(i).unwrap() as i32 >= s22.chars().nth(i).unwrap() as i32 {
'            is_greater_or_equal=true;
'        }else{
'            is_greater_or_equal=false;
'        }
'        break;
'    }
'    is_greater_or_equal
'
'}
Public Function is_greater_or_equal(s1 As String, s2 As String) As Boolean
    Dim s11 As String: s11 = s1
    Dim s22 As String: s22 = s2
    same_sizer s11, s22
    Dim i As Long
    Dim c1 As Integer, c2 As Integer
    For i = 1 To Len(s11)
        c1 = AscW(Mid(s11, i, 1))
        c2 = AscW(Mid(s22, i, 1))
        If c1 > c2 Then
            Exit For
        ElseIf c1 < c2 Then
            Exit Function
        End If
    Next
    is_greater_or_equal = True
End Function

'pub fn is_greater(s1:&String,s2:&String)->bool{
'    let (s11,s22)=same_sizer(s1, s2);
'    let mut is_greater=true;
'    for i in 0..s11.len() {
'        if s11.chars().nth(i).unwrap() as i32 == s22.chars().nth(i).unwrap() as i32 {
'            continue
'        }
'        if s11.chars().nth(i).unwrap() as i32 > s22.chars().nth(i).unwrap() as i32 {
'            is_greater=true;
'        }else{
'            is_greater=false;
'        }
'        break;
'    }
'    is_greater
'}
Public Function is_greater(s1 As String, s2 As String) As Boolean
    Dim s11 As String: s11 = s1
    Dim s22 As String: s22 = s2
    same_sizer s11, s22
    Dim i As Long
    Dim c1 As Integer, c2 As Integer
    For i = 1 To Len(s11)
        c1 = AscW(Mid(s11, i, 1))
        c2 = AscW(Mid(s22, i, 1))
        If c1 > c2 Then
            Exit For
        Else 'If c1 < c2 Then
            Exit Function
        End If
    Next
    is_greater = True
End Function

'#[allow(dead_code)]
'pub fn left_zero_kill(s:&String) ->String{
'    let mut st = String::from("");
'    let  still_found_zero=true;
'    for _i in 0..s.len() {
'        if s.chars().nth(_i).unwrap() as i32 -48 == 0 && still_found_zero{
'            continue;
'        }else{
'            st.insert_str(0,&s[_i..]);
'            break;
'        }
'    }
'    st
'}
Public Function left_zero_kill(s As String) As String
    Dim st As String: st = Trim(s)
    Dim i As Long, l As Long: l = Len(st)
    Dim c As Integer
    For i = 1 To l
        c = AscW(Mid(st, i, 1)) - 48
        If c <> 0 Then
            left_zero_kill = Mid(st, i, l - i + 1)
            Exit Function
        End If
    Next
End Function
