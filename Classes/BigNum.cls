VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BigNum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'https://github.com/arash-hacker/BIIIIIIIIIIIG.2
Private m_Value As String

Friend Sub New_(NumericValue)
    m_Value = NumericValue
End Sub

Public Property Get Value() As String
    Value = m_Value
End Property
Public Property Let Value(ByVal NumericString As String)
    m_Value = NumericString
End Property

'add.rs
'use super::helper::same_sizer;
'
'pub fn add(s1:&String,s2:&String)->String{
'        let (s11,s22)=same_sizer(s1, s2);
'        let total:String=add_one_to_one(&s11,&s22);
'        //println!("{} !! {}",s11,s22);
'        //println!("{}",total);
'        total
'    }
'
'pub fn add_one_to_one(s1:&String,s2:&String)->String{
'
'    let mut carry=0;
'    let mut result:String=String::from("");
'    //println!("{} >> {}: {} {}",s1,s2,s1.len(),s2.len());
'    for i in (0..s1.len()).rev() {
'
'        //println!("{} ",(s2.chars().nth(i).unwrap() as i32 -48));
'        let car:i32=(s1.chars().nth(i).unwrap() as i32 -48) + (s2.chars().nth(i).unwrap() as i32 -48)+carry;
'        if car>=10 {
'            let mod10 =car%10;
'            let div10 =car/10;
'            result.insert_str(0,&mod10.to_string());
'            carry=div10;
'
'        }else{
'            result.insert_str(0,&car.to_string());
'            carry=0;
'        }
'        //println!("{}",result);
'    }
'    if carry>0{
'        result.insert_str(0,&carry.to_string());
'    }
'    result
'}

'sub.rs
'    use super::helper::same_sizer;
'    use super::helper::is_greater;
'
'    #[allow(dead_code)]
'    pub fn sub_one_to_one(s1:&String,s2:&String)->String{
'        let mut carry=0;
'        let mut result:String=String::from("");
'
'        for i in (0..s1.len()).rev() {
'
'
'            let mut car:i32=(s1.chars().nth(i).unwrap() as i32 -48-carry) - (s2.chars().nth(i).unwrap() as i32 -48);
'
'            if car<0{
'                car=(s1.chars().nth(i).unwrap() as i32 + 10 -48-carry) - (s2.chars().nth(i).unwrap() as i32 -48);
'
'                result.insert_str(0,&car.to_string());
'                carry=1;
'
'            }else{
'                result.insert_str(0,&car.to_string());
'                carry=0;
'            }
'
'            //println!("{}",result);
'        }
'       result
'    }
'    #[allow(dead_code)]
'    pub fn sub(s1:&String,s2:&String)->String{
'        let (s11,s22)=same_sizer(s1, s2);
'        let mut total:String;
'        if is_greater(&s1,&s2) {
'            total=sub_one_to_one(&s11,&s22);
'        }else{
'            total=sub_one_to_one(&s22,&s11);
'        }
'        total
'
'    }

'multiple.rs
'use super::helper::same_sizer;
'use super::add::add;
'
'#[allow(dead_code)]
'pub fn multiple(s1:&String,s2:&String)->String{
'        let (s11,s22)=same_sizer(s1, s2);
'        let total:String=multiple_many_to_many(&s11,&s22);
'        //println!("{} * {} = {} ",s1,s2,total);
'        total
'    }
'
'#[allow(dead_code)]
'pub fn multiple_one_to_many(s1:&String,s2:char)->String{
'    let mut carry=0;
'    let mut result=String::from("");
'    for i in (0..s1.len()).rev() {
'
'        let car:i32=(s1.chars().nth(i).unwrap() as i32 -48) * (s2 as i32 -48)+carry;
'
'        if car>=10 {
'            let mod10 =car%10;
'            let div10 =car/10;
'            result.insert_str(0,&mod10.to_string());
'            carry=div10;
'
'        }else{
'            result.insert_str(0,&car.to_string());
'            carry=0;
'        }
'
'    }
'    if carry>0{
'        result.insert_str(0,&carry.to_string());
'    }
'
'    result
'}
'
'#[allow(dead_code)]
'pub fn multiple_10_pow_x(one_to_many:&String,pow:usize)->String{
'    let mut result=String::from(one_to_many);
'    for _i in 0..pow {
'        result.insert_str(result.len(), "0")
'    }
'    result
'
'}
'
'
'#[allow(dead_code)]
'pub fn multiple_many_to_many(s1:&String,s2:&String)->String{
'
'
'        let mut result:String=String::from("0");
'
'        let mut result_one_row;
'
'        for i in (0..s1.len()).rev() {
'
'            let  one_to_many=multiple_one_to_many(&s1,s2.chars().nth(i).unwrap());
'            result_one_row=multiple_10_pow_x(&one_to_many,s1.len()-i-1);
'            result=add(&result_one_row,&result);
'        }
'        result
'    }

'div.rs
'use super::helper::same_sizer;
'use super::helper::left_zero_kill;
'use super::helper::is_greater;
'use super::helper::is_greater_or_equal;
'use super::add::add;
'use super::sub::sub;
'use super::multiple::multiple;
'
'#[allow(dead_code)]
'pub fn div(s1:&String,s2:&String)->String{
'        let (mut s11, mut s22)=same_sizer(s1, s2);
'        s11=left_zero_kill(&s11);
'        s22=left_zero_kill(&s22);
'       // println!("{} / {}",s11,s22);
'        let total:String=div_many_to_many(&s11,&s22);
'        total
'
'}
'
'#[allow(dead_code)]
'pub fn div_chunks(s1:&String,s2:&String)->(String,String){
'
'
'    let mut divider  =String::from("0");
'    let mut reminder ;
'
'   // println!("{} {}",s1,s2);
'    while is_greater_or_equal(&s1, &multiple(&s2,&divider)) {
'
'        divider=add(&divider,&String::from("1"));
'
'        if is_greater(&divider, &String::from("11") ){
'            panic!("error in divider");
'        }
'    }
'    if is_greater_or_equal(&s1, &multiple(&s2,&divider)){
'        divider=add(&divider,&String::from("1") );
'    }
'
'    divider=left_zero_kill(&sub(&divider, &String::from("1")));
'    reminder=sub(&s1,&multiple(&divider, &s2));
'
'    if divider.len()==0 {
'       return (String::from("0"),reminder)
'    }
'
'    (divider,reminder)
'}
'
'
'
'#[allow(dead_code)]
'pub fn div_many_to_many(divee:&String,s2:&String)->String{
'
'        let mut result:String=String::from("0");
'        let mut s1_copy=String::from(divee);
'        let mut s1=String::from("");
'        let mut space    = 0;
'
'        //println!("| s1 {} | s2 {} |  | | result {} >> s1_copy {} ",s1,s2,result, s1_copy);
'
'        while &s1_copy.len() >= &s2.len() {
'            //is_greater(&s1_copy, &s2) {
'            if s2.len()+space > s1_copy.len() {
'                break;
'            }
'
'            s1=String::from(&s1_copy[0..s2.len()+space]);
'
'         //   println!("new s1 is {}",s1);
'
'            let (divider,reminder)=div_chunks(&s1,&s2);
'
'
'            result.insert_str(result.len(), &divider.to_string() );
'
'            if is_greater_or_equal(&String::from("0"), &reminder){
'              //  println!("{}",">>>>>>>>>>>>>>>");
'                s1_copy=format!("{}{}",
'                    &String::from("0").repeat(s2.len()+space),
'                    &s1_copy[s2.len()+space..]);
'
'            }else{
'
'                s1_copy=format!("{}{}",
'                    &reminder[&reminder.len()-s2.len()-space..],
'                    &s1_copy[s2.len()+space..]);
'
'            }
'
'            space+=1;
'
'            //delta+=1;
'          // println!(" s1 {} | s2 {} | divider {} | reminder {} | result {} >> s1_copy {} space {} ",s1,s2,divider,&reminder[&reminder.len()-s2.len()..],result, s1_copy,space);
'        }
'
'        result
'    }

'lib.rs
'pub mod add;
'pub mod helper;
'pub mod sub;
'pub mod div;
'pub mod multiple;
