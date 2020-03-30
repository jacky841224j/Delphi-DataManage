unit StrUtil;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, StdCtrls, StrUtils,
  DataM, ADODB, IniFiles,Dialogs,Math,DBGrids;

type ChineseDate=
  record
    YY: word;
    MM: word;
    DD: word;
  end;

function AddSQLList(Data_String:String) : String;
function Fix_Size(data_string:string;data_size:integer) : string;
function SPACE(data_size:integer) : string;
function Fix_SizeL(data_string:string;data_size:integer) : string;
function Fix_Num(NumVar:integer;data_size:integer) : string;
function Fix_SNum(data_string:string;data_size:integer) : string;
function Chk_Num(NumVar:String) : Boolean;
function ListSHT(SCR_List:String):String;
function CutLast(S:String):String;
function CutFirst(S:String):String;
function Fix_Float(data_string:string;data_size:integer) : string;
function Fix_MMDD(data_string:string) : string;
Function Rounda(Value: Extended; Len: Word): Extended;
Function Rounda2(Value: Extended; Len: Word): integer;
Function ChkFont(sLine : String) : String;                //抓造字
Function TestDataRange(InputData: String): Boolean;       // 檢查輸入的資料是否數值
Function CheckDate(Y, M, D: String): String;              // 檢查輸入的日期是否正確 (格式為YY,MM,DD)
Function CheckDate1(Y, M, D: String): String;             // 檢查輸入的日期是否正確 (格式為YY,MM,DD)
Function CheckDate2(InputData: String): String;           // 檢查輸入的日期是否正確 (格式為89/1/1)
Function CheckIdCode(InputData: String): String;          // 檢查輸入的身份證編號是否正確
Function EnglishToChenese(EDate: TDateTime): Chinesedate; // 英文日期轉中文日期
Function ChineseToEnglish(Y, M, D: string): TDateTime;    // 中文日期轉英文日期
Function GetLastDate(Y, M: String): Integer;              // 取得該月的最後一日
Function ChangeValue(Input: Real): String;                // 將數值四捨五入
Function Flt_Com(Input: String): String;                  // 濾豆號
Function Flt_SP(Input: String): String;                   // 濾空白
function Flt_Any(Input:string;M:string):string;           // 濾字串M中任何字元
function Ans_Cnt(Ans:String):string;                      //AB=>F
function Flt_Under(INPUT:string) : string;              //濾底線
function Ans_Nec2Scantron(Ans:String):string;          //將答案nec格式轉成scantron格式
function Cn_FN(Path,FileName:string):string;           //組合路徑與檔名
function Check_AllLetter(StrIn:string):Boolean;         //檢查字串是否全為字母組成
function Cut_Comm(var strIn:string):string;         //切豆號前後字串
function check_contain(str1,str2:string):boolean;  //字串二之字元是否有包含字串一所有字元之情況
function CNV_Num(str:string):string;
function chgAnsMap(Ans: string): String;            //答案對應轉換
//add by Carol 20090401 答案卡比對
Function  CheckNo(No:String):Boolean;    //條碼檢核
Function  GetAnswerStr(Data:string;AnsCount:integer):String;//擷取答案字串
Function  GetAns(AnswerStr: String; QNo: Integer): String;      //取出答案字串中某題之答案
Function  ReplaceAns(AnswerStr,Ans:String;QNo:integer):String;     //取代某題之答案
Function  AddZero(Str:String;Leng:integer):String;
function  ReplaceStr(const AText, AFromText, AToText: AnsiString): AnsiString;
Function  ChangeValue2(a: real): Integer;
procedure DeleteTree(path: String); //清空目錄
procedure SQLExec(DataSet:TADOQuery; Str:String);
procedure OpenSQL(DateSet : TADOQuery; SQLstr:string);
//procedure LoadINI(fPath:String);
//procedure SaveINI(fPath, sD, sY, sT, sU, sP, sS, sHasStudentData:String);
function ExtractFilePath(const FileName: string): string;
Function TableExists(TableName : String):Boolean;  //檢查Table 是否存在

//add by Purea 20110106
Function DBGridRecordSize(mColumn: TColumn):Boolean; //記錄最大欄寬，在OnDrawColumnCell中使用
Function DBGridAutoSize(mDBGrid: TDBGrid; mOffset: Integer = 5):Boolean; //自動調整欄寬，在OnActivate中使用
Function Strsplit(input: string; schar: Char; sindex: cardinal):string;  //字串分割

// add shinyu 20130109
function CheckExamQuestionStruc(ExamNo, SubNo: string): Boolean;
function CheckExamQuestionRead(DataPath, ReadAccess: string): Boolean; //add fish 20130416是否已讀卡
function GetColsExits(DateSet : TADOQuery; sTable, sField : string):Boolean;  //Add by Carol 20131115



implementation

uses UMain;


Function TableExists(TableName : String):Boolean;
Var
    Sl : TStrings;
Begin
    Sl := TstringList.Create;
    // 使用ADO ===============================
    DM.ADOConnection1.GetTableNames(Sl,True);
    Result := Sl.IndexOf(TableName)<> -1 ;
    // 使用BDE ===============================
  //  Session.GetTableNames('dbAlias',TableName,False,True,Sl);
  //  Result := Sl.Count>0 ;
    // 擇一使用 ===============================
    Sl.Free;
End;


function GetColsExits(DateSet : TADOQuery; sTable, sField : string):Boolean;  //Add by Carol 20131115
Var
  Str1 : String;
  ii : Integer;
begin
  Str1 := 'Select Top 1 * From '+sTable;
  OpenSQL(DateSet, Str1);

  Result := False;
  for ii := 0 to DateSet.FieldCount - 1 do
  begin
    if DateSet.Fields.Fields[ii].FieldName=sField then
    begin
      Result := True;
      Break;
    end
    else
      Result := False;
  end;


end;


function ExtractFilePath(const FileName: string): string;
var
  I: Integer;
begin
  I := LastDelimiter(PathDelim + DriveDelim, FileName);
  Result := Copy(FileName, 1, I);
end;

{
procedure SaveINI(fPath, sD, sY, sT, sU, sP, sS, sHasStudentData: String);
var
  Ini: TIniFile;
  hWnd: Integer;
  AItem : Boolean;
begin
  fPath := fPath+'_Data';
  AItem := True;
  if not DirectoryExists(fPath) then   //判斷存檔目錄是否存在,不存在則創建檔案
  begin
    ForceDirectories(fPath); //新增目錄
    hWnd := FileCreate(fPath+'\SMCC.ini');
    FileClose(hWnd);
    AItem := False;
  end;
  try
    Ini := TIniFile.Create(fPath+'\SMCC.ini');  //建立物件
    Ini.WriteString('DetailData', 'TermNo',sT);  //寫入INI檔案 (分項,變數名稱,內容值)
    Ini.WriteString('DetailData', 'TYear', sY);
    Ini.WriteString('DetailData', 'Depart', sD);  //寫入INI檔案 (分項,變數名稱,內容值)
    Ini.WriteString('DetailData', 'UserName', sU);
    Ini.WriteString('DetailData', 'PassWord', sP);  //寫入INI檔案 (分項,變數名稱,內容值)
    Ini.WriteString('DetailData', 'Sname', sS);
    Ini.WriteString('DetailData', 'HasStudentData', sHasStudentData);
    if not AItem then
     Ini.WriteString('DetailData', 'Item','');


  finally
    Ini.Free; //釋放物件
  end;
end;
}
{
procedure LoadINI(fPath:String);
var
  aIni:TInifile;
begin
  if not FileExists(fPath+'_Data\SMCC.ini') then   //判斷檔案是否存在
  begin
    showmessage('使用者設定檔遺失！！');
    Exit;
  end;

  try
    // 讀取 *.ini 檔
    //aIni := TIniFile.Create(ChangeFileExt(Application.ExeName, '.ini' ) );
    aIni := TIniFile.Create(fPath+'_Data\SMCC.ini');//路徑+檔名.INI
    //aIni.WriteString('SysConfig2','Test','TTTTT');

    fmMain.sTerm := Copy(Trim(aIni.ReadString('DetailData','TermNo','')),1,1);      //
    fmMain.sTYear := Trim(aIni.ReadString('DetailData','TYear',''));      //
    fmMain.sDept := Trim(aIni.ReadString('DetailData','Depart',''));    //
    fmMain.sUID := Trim(aIni.ReadString('DetailData','UserName',''));    //
    fmMain.sPW := Trim(aIni.ReadString('DetailData','PassWord',''));      //
   // fmMain.sName := Trim(aIni.ReadString('DetailData','Sname',''));      //
   // fmMain.ItemOn := Trim(aIni.ReadString('DetailData','ItemOn',''));      //使否隱藏標準版流程
   // fmMain.sUpdate := Trim(aIni.ReadString('DetailData','sUpdate',''));  // add by ken 20121001 是否顯示資料庫更新功能
   // fmMain.sHasStudentData:= Trim(aIni.ReadString('DetailData','HasStudentData',EmptyStr));
  finally
    aIni.Free;
  end;
end;
}

procedure OpenSQL(DateSet: TADOQuery; SQLstr:string);
begin
  DateSet.Close;
  DateSet.SQL.Clear;
  DateSet.SQL.Add(SQLstr);
  try
    DateSet.Open;
  except
    ShowMessage(SQLStr);
  end;

end;


procedure SQLExec(Dataset: TADOQuery; Str: String);
begin
  Dataset.Close;
  Dataset.SQL.Clear;
  Dataset.SQL.Add(Str);
  Dataset.ExecSQL;
end;

procedure DeleteTree(path: String);
var
  p, FileName:string;
  iFindResult: integer;
  SearchRec: TSearchRec;
begin
  if (Copy(path, Length(path), 1)= '\') then
    p:= Copy(path, 1, Length(path)-1)
  else
    p:= path;

  iFindResult := FindFirst((p+'\*.*'), faAnyFile, SearchRec);
  while iFindResult = 0 do
    begin
      // ListBox1.Items.Add(SearchRec.Name); ＜- 將檔名置入listbox
      FileName:= p + '\' + SearchRec.Name;
      DeleteFile(FileName);
      iFindResult := FindNext(SearchRec);
    end;
  FindClose(SearchRec);
end;

Function ChangeValue2(a: real): Integer;
var b:real;
begin
  b:=frac(a);
  if b>0.85 then
    Result:=trunc(a)+2
  else
    Result:=trunc(a)+1;
end;


function ReplaceStr(const AText, AFromText, AToText: AnsiString): AnsiString;
begin
  Result := AnsiReplaceStr(AText, AFromText, AToText);
end;

Function CheckNo(No:String):Boolean;
var
  s: Integer;
begin
  s:=(StrToInt(Copy(No,1,1))*1+StrToInt(Copy(No,2,1))*3+StrToInt(Copy(No,3,1))*7+
      StrToInt(Copy(No,4,1))*1+StrToInt(Copy(No,5,1))*3+StrToInt(Copy(No,6,1))*7+
      StrToInt(Copy(No,7,1))*1+StrToInt(Copy(No,8,1))*3+StrToInt(Copy(No,9,1))*7) mod 10;
  if (IntToStr(s) = Copy(No, 10, 1)) then
    Result:= True
  else
    Result:= False;
end;

Function GetAnswerStr(Data: String; AnsCount: Integer): String;  //將答案切成與題數相同
var
  SL: TStringList;
  i: Integer;
begin
  Result:= '';
  if (AnsCount > 0) then
    begin
      Result:= Data;
      SL:= TStringList.Create;
      Try
        SL.Clear;
        SL.CommaText:= Data;

        if (AnsCount > SL.Count-1) then
          begin
            for i:= SL.Count to AnsCount do
              begin
                Result:= Result + ',';
              end;
          end
        else if (AnsCount < SL.Count-1) then
          begin
            for i:= SL.Count-2 downto AnsCount do
              begin
                SL.Delete(i);
              end;
            Result:= SL.CommaText;
          end;
      Finally
        SL.Free;
      End;
    end;

end;

Function GetAns(AnswerStr: String; QNo: Integer): String;
var
  SL: TStringList;
begin
  SL:= TStringList.Create;
  Try
    SL.Clear;
    SL.CommaText:= AnswerStr;
    if (QNo > SL.Count-1) then
      Result:= ''
    else
      Result:= SL.Strings[QNo-1];
  Finally
    SL.Free;
  End;


end;

Function ReplaceAns(AnswerStr, Ans: String; QNo: Integer): String;
var
  SL: TStringList;
begin
  SL:= TStringList.Create;
  Try
    SL.Clear;
    SL.CommaText:= AnswerStr;
    if (QNo < 1) or (QNo > SL.Count-1) then
      begin
        Result:= AnswerStr;
      end
    else
      begin
        SL.Strings[QNo-1]:= Ans;
        Result:= SL.CommaText;
      end;
  Finally
    SL.Free;
  End;


end;

Function AddZero(Str: String; Leng: Integer): String;
var
  i: Integer;
begin
  Result:= Str;
  for i:= Length(Str)+1 to Leng do
    Result:= '0' + Result;


end;


function chgAnsMap(Ans: string): String;
begin

        if (Ans = 'A____') or (Ans = '1____') or (Ans = 'A') or (Ans = '1') then
          Result:='A'
        else if (Ans = '_B___') or (Ans = '_2___') or (Ans = 'B') or (Ans = '2') then
          Result:='B'
        else if (Ans = '__C__') or (Ans = '__3__') or (Ans = 'C') or (Ans = '3') then
          Result:='C'
        else if (Ans = '___D_') or (Ans = '___4_') or (Ans = 'D') or (Ans = '4') then
          Result:='D'
        else if (Ans = '____E') or (Ans = '____5') or (Ans = 'E') or (Ans = '5') then
          Result:='E'
        else if (Ans = 'AB___') or (Ans = '12___') or (Ans = 'AB') or (Ans = '12') then
          Result:='F'
        else if (Ans = 'A_C__') or (Ans = '1_3__') or (Ans = 'AC') or (Ans = '13') then
          Result:='G'
        else if (Ans = 'A__D_') or (Ans = '1__4_') or (Ans = 'AD') or (Ans = '14') then
          Result:='H'
        else if (Ans = 'A___E') or (Ans = '1___5') or (Ans = 'AE') or (Ans = '15') then
          Result:='I'
        else if (Ans = '_BC__') or (Ans = '_23__') or (Ans = 'BC') or (Ans = '23') then
          Result:='J'
        else if (Ans = '_B_D_') or (Ans = '_2_4_') or (Ans = 'BD') or (Ans = '24') then
          Result:='K'
        else if (Ans = '_B__E') or (Ans = '_2__5') or (Ans = 'BE') or (Ans = '25') then
          Result:='L'
        else if (Ans = '__CD_') or (Ans = '__34_') or (Ans = 'CD') or (Ans = '34') then
          Result:='M'
        else if (Ans = '__C_E') or (Ans = '__3_5') or (Ans = 'CE') or (Ans = '35') then
          Result:='N'
        else if (Ans = '___DE') or (Ans = '___45') or (Ans = 'DE') or (Ans = '45') then
          Result:='O'
        else if (Ans = 'ABC__') or (Ans = '123__') or (Ans = 'ABC') or (Ans = '123') then
          Result:='P'
        else if (Ans = 'AB_D_') or (Ans = '12_4_') or (Ans = 'ABD') or (Ans = '124') then
          Result:='Q'
        else if (Ans = 'AB__E') or (Ans = '12__5') or (Ans = 'ABE') or (Ans = '125') then
          Result:='R'
        else if (Ans = 'A_CD_') or (Ans = '1_34_') or (Ans = 'ACD') or (Ans = '134') then
          Result:='S'
        else if (Ans = 'A_C_E') or (Ans = '1_3_5') or (Ans = 'ACE') or (Ans = '135') then
          Result:='T'
        else if (Ans = 'A__DE') or (Ans = '1__45') or (Ans = 'ADE') or (Ans = '145') then
          Result:='U'
        else if (Ans = '_BCD_') or (Ans = '_234_') or (Ans = 'BCD') or (Ans = '234') then
          Result:='V'
        else if (Ans = '_BC_E') or (Ans = '_23_5') or (Ans = 'BCE') or (Ans = '235') then
          Result:='W'
        else if (Ans = '_B_DE') or (Ans = '_2_45') or (Ans = 'BDE') or (Ans = '245') then
          Result:='X'
        else if (Ans = '__CDE') or (Ans = '__345') or (Ans = 'CDE') or (Ans = '345') then
          Result:='Y'
        else if (Ans = 'ABCD_') or (Ans = '1234_') or (Ans = 'ABCD') or (Ans = '1234') then
          Result:='Z'
        else if (Ans = 'ABC_E') or (Ans = '123_5') or (Ans = 'ABCE') or (Ans = '1235') then
          Result:='*'
        else if (Ans = 'AB_DE') or (Ans = '12_45') or (Ans = 'ABDE') or (Ans = '1245') then
          Result:='$'
        else if (Ans = 'A_CDE') or (Ans = '1_345') or (Ans = 'ACDE') or (Ans = '1345') then
          Result:='%'
        else if (Ans = '_BCDE') or (Ans = '_2345') or (Ans = 'BCDE') or (Ans = '2345') then
          Result:='='
        else if (Ans = 'ABCDE') or (Ans = '12345') then
          Result:='&'
        else
          Result:=Ans;

end;



function CNV_Num(str:string):string;
var
  x:integer;
begin
  result:='';
  if Chk_Num(copy(str,1,1)) then
  begin
    x:=StrToInt(copy(str,1,1));
    case x of
      0:result:='０';
      1:result:='一';
      2:result:='二';
      3:result:='三';
      4:result:='四';
      5:result:='五';
      6:result:='六';
      7:result:='七';
      8:result:='八';
      9:result:='九';
    end;
  end;
end;

function check_contain(str1,str2:string):boolean;
var
  i:integer;
begin
  result:=True;
  for i := 1 to length(str1) do
  begin
    if pos(str1[i],str2) < 0 then
    begin
      result:=False;
    end;
  end;
end;

function Cut_Comm(var strIn:string):string;
var
  p:integer;
begin
  p:=pos(',',strIn);
  if p>0 then
  begin
   result := copy(strIn,1,p-1);
    strIn := copy(strIn,p+1,length(strIn)-p);
  end
  else
  begin
    result := strIn;
    strIN:='';
  end;
end;

function Check_AllLetter(StrIn:string):Boolean;
var
  i:integer;
begin
  result:=true;
  for i:= 1 to length(UpperCase(StrIn)) do
  begin
    if ((StrIn[i] <= 'A') and (StrIn[i] >= 'Z')) then result := false;
  end;
end;

function Cn_FN(Path,FileName:string):string;           //組合路徑與檔名
begin
  if RightStr(Path,1) <> '\' then Path:=Path+'\';
  result := Path+FileName;
end;

function Ans_Nec2Scantron(Ans:String):string;
var
  p:integer;
begin
  Ans:=Ans+',';
  result:='';
  p:=pos(',',Ans);
  while p > 0 do
  begin
    result:=result+Ans_Cnt(copy(Ans,1,p-1));
    Ans := copy(Ans,p+1,length(Ans)-p);
    p:=pos(',',Ans);
  end;
end;

function Flt_Under(INPUT:string):string;
var
  i:integer;
  OUTPUT:string;
begin
  OUTPUT:='';
  for i := 1 to length(INPUT) do
  begin
    if (INPUT[i] <> '_') and (INPUT[i] <> ' ') then OUTPUT :=OUTPUT+INPUT[i];
  end;
  result:=OUTPUT;
end;

function Ans_Cnt(Ans:String):string;
begin
  if (Ans = 'A____') or (Ans = '1____') or (Flt_Under(Ans) = 'A') or (Ans = '1') then result:='A'
  else if (Ans = 'B____') or (Ans = '_2___') or (Flt_Under(Ans) = 'B') or (Ans = '2') then result:='B'
  else if (Ans = 'C____') or (Ans = '__3__') or (Flt_Under(Ans) = 'C') or (Ans = '3') then result:='C'
  else if (Ans = 'D____') or (Ans = '___4_') or (Flt_Under(Ans) = 'D') or (Ans = '4') then result:='D'
  else if (Ans = 'E____') or (Ans = '____5') or (Flt_Under(Ans) = 'E') or (Ans = '5') then result:='E'
  else if (Ans = 'AB___') or (Ans = '12___') or (Flt_Under(Ans) = 'AB') or (Ans = '12') then result:='F'
  else if (Ans = 'A_C__') or (Ans = '1_3__') or (Flt_Under(Ans) = 'AC') or (Ans = '13') then result:='G'
  else if (Ans = 'A__D_') or (Ans = '1__4_') or (Flt_Under(Ans) = 'AD') or (Ans = '14') then result:='H'
  else if (Ans = 'A___E') or (Ans = '1___5') or (Flt_Under(Ans) = 'AE') or (Ans = '15') then result:='I'
  else if (Ans = '_BC__') or (Ans = '_23__') or (Flt_Under(Ans) = 'BC') or (Ans = '23') then result:='J'
  else if (Ans = '_B_D_') or (Ans = '_2_4_') or (Flt_Under(Ans) = 'BD') or (Ans = '24') then result:='K'
  else if (Ans = '_B__E') or (Ans = '_2__5') or (Flt_Under(Ans) = 'BE') or (Ans = '25') then result:='L'
  else if (Ans = '__CD_') or (Ans = '__34_') or (Flt_Under(Ans) = 'CD') or (Ans = '34') then result:='M'
  else if (Ans = '__C_E') or (Ans = '__3_5') or (Flt_Under(Ans) = 'CE') or (Ans = '35') then result:='N'
  else if (Ans = '___DE') or (Ans = '___45') or (Flt_Under(Ans) = 'DE') or (Ans = '45') then result:='O'
  else if (Ans = 'ABC__') or (Ans = '123__') or (Flt_Under(Ans) = 'ABC') or (Ans = '123') then result:='P'
  else if (Ans = 'AB_D_') or (Ans = '12_4_') or (Flt_Under(Ans) = 'ABD') or (Ans = '124') then result:='Q'
  else if (Ans = 'AB__E') or (Ans = '12__5') or (Flt_Under(Ans) = 'ABE') or (Ans = '125') then result:='R'
  else if (Ans = 'A_CD_') or (Ans = '1_34_') or (Flt_Under(Ans) = 'ACD') or (Ans = '134') then result:='S'
  else if (Ans = 'A_C_E') or (Ans = '1_3_5') or (Flt_Under(Ans) = 'ACE') or (Ans = '135') then result:='T'
  else if (Ans = 'A__DE') or (Ans = '1__45') or (Flt_Under(Ans) = 'ADE') or (Ans = '145') then result:='U'
  else if (Ans = '_BCD_') or (Ans = '_234_') or (Flt_Under(Ans) = 'BCD') or (Ans = '234') then result:='V'
  else if (Ans = '_BC_E') or (Ans = '_23_5') or (Flt_Under(Ans) = 'BCE') or (Ans = '235') then result:='W'
  else if (Ans = '_B_DE') or (Ans = '_2_45') or (Flt_Under(Ans) = 'BDE') or (Ans = '245') then result:='X'
  else if (Ans = '__CDE') or (Ans = '__345') or (Flt_Under(Ans) = 'CDE') or (Ans = '345') then result:='Y'
  else if (Ans = 'ABCD_') or (Ans = '1234_') or (Flt_Under(Ans) = 'ABCD') or (Ans = '1234') then result:='Z'
  else if (Ans = 'ABC_E') or (Ans = '123_5') or (Flt_Under(Ans) = 'ABCE') or (Ans = '1235') then result:='*'
  else if (Ans = 'AB_DE') or (Ans = '12_45') or (Flt_Under(Ans) = 'ABDE') or (Ans = '1245') then result:='$'
  else if (Ans = 'A_CDE') or (Ans = '1_345') or (Flt_Under(Ans) = 'ACDE') or (Ans = '1345') then result:='%'
  else if (Ans = '_BCDE') or (Ans = '_2345') or (Flt_Under(Ans) = 'BCDE') or (Ans = '2345') then result:='='
  else if (Ans = 'ABCDE') or (Ans = '12345') then result:='&'
  else  result:=' ';
end;

function AddSQLList(Data_String:String) : String;
var
  tmp,OutStr:String;
begin
  Data_String:=Trim(Data_String);
  If Data_String[Length(Data_String)] <> ',' then Data_String:=Trim(Data_String)+',';
  tmp:=Trim(copy(Data_String,1,pos(',',Data_String)-1));
  Data_String:=Trim(Copy(Data_String,(pos(',',Data_String)+1),length(Data_String)-pos(',',Data_String)));
  OutStr:='( '''+Tmp+''' ,';
  while Data_String <> '' do
  begin
    tmp:=Trim(copy(Data_String,1,pos(',',Data_String)-1));
    Data_String:=trim(Copy(Data_String,pos(',',Data_String)+1,length(Data_String)-pos(',',Data_String)));
    OutStr:= OutStr+' '''+Tmp+''' ,';
  end;
  OutStr:=CutLast(OutStr)+' )';
  Result:=OutStr;
end;

//抓造字
Function ChkFont(sLine : String) : String;
var
  str1, str2: string;
  k: integer;
begin
  k:=1;
  str2:='';

  while k <> (Length(sLine)+1) do
    begin
      str1:=Copy(sLine,k,1);
      if str1[1] < #127 then
        begin
          str2:=str2+str1[1];
          k:=k+1;
        end;

      if str1[1] > #127 then
        begin
          if (str1[1] > #249) and (str1[1] < #255) then   //FA - FE
            begin
              if (Copy(sLine,k+1,1) > #63) and (Copy(sLine,k+1,1) < #255) then //40 - FE
                begin
                  //str1:='('+Copy(sLine,k,2)+')';
                  str1:='＊';
                  str2:=str2+str1;
                  k:=k+2;
                end
              else
                begin
                  str1:=Copy(sLine,k,2);
                  str2:=str2+str1;
                  k:=k+2;
                end;
            end
          else
            if (str1[1] > #128) and (str1[1] < #161) then  //81 - A0
              begin
                if (Copy(sLine,k+1,1) > #63) and (Copy(sLine,k+1,1) < #255) then //40 - FE
                  begin
                    str1:='＊';
                    str2:=str2+str1;
                    k:=k+2;
                  end
                else
                  begin
                    str1:=Copy(sLine,k,2);
                    str2:=str2+str1;
                    k:=k+2;
                  end;
              end
            else
              if (str1[1] = #198) then //C6
                begin
                  if (Copy(sLine,k+1,1) > #160) and (Copy(sLine,k+1,1) < #255) then // A1 - FE
                    begin
                      str1:='＊';
                      str2:=str2+str1;
                      k:=k+2;
                    end
                  else
                    begin
                      str1:=Copy(sLine,k,2);
                      str2:=str2+str1;
                      k:=k+2;
                    end;
                end
              else
                if (str1[1] > #198) and (str1[1] < #201) then //C7 - C8
                  begin
                    if (Copy(sLine,k+1,1) > #63) and (Copy(sLine,k+1,1) < #255) then // 40 - FE
                      begin
                        str1:='＊';
                        str2:=str2+str1;
                        k:=k+2;
                      end
                    else
                      begin
                        str1:=Copy(sLine,k,2);
                        str2:=str2+str1;
                        k:=k+2;
                      end;
                  end
                else
                  begin
                    str1:=Copy(sLine,k,2);
                    str2:=str2+str1;
                    k:=k+2;
                  end;
        end;    // if str1[1] > #127 then
    end;   // while k <> (Length(sLine)+1) do
    Result:= str2;
end;




// 檢查輸入的資料是否數值
Function TestDataRange(InputData: String): Boolean;
var
  i: word;
begin
  Result:= True;
  for i:= 1 to Length(InputData) do
    if ((copy(InputData, i, 1) < '0' ) or (copy(InputData, i, 1) > '9')) then
      begin
        Result:= False;
        Break;
      end;
end;


// 檢查輸入的日期是否正確 (格式為 YY, MM, DD)
Function CheckDate(Y, M, D: String): String;
var
  Year, Month, Date: integer;
  Mdays: Array[1..12] of smallint;
begin
  Result:= '正確';
  if (Length(Y) = 0) or (Length(M) = 0) or (Length(D) = 0) then
    begin
      if (Length(Y) = 0) and (Length(M) = 0) and (Length(D) = 0) then
        begin
          Result:= '日期輸入空白';
        end
      else
        begin
          Result:= '日期輸入不完整';
        end;
      Exit;
    end;
  if NOT (TestDataRange(Y) and TestDataRange(M) and TestDataRange(D)) then
    begin
      Result:= '日期輸入應為數值型態';
      Exit;
    end;
  Mdays[1]:=31; Mdays[2]:=28; Mdays[3]:=31; Mdays[4]:=30; Mdays[5]:=31; Mdays[6]:=30; Mdays[7]:=31; Mdays[8]:=31; Mdays[9]:=30; Mdays[10]:=31; Mdays[11]:=30; Mdays[12]:=31;
  Year:= StrToInt(Y)+1911;
  Month:= StrToInt(M);
  Date:= StrToInt(D);
  if ((Year mod 400 = 0) or ((Year mod 4 = 0) and (Year mod 100 <> 0))) then
    Mdays[2]:= 29;
  if ((Year < 1) or (Month < 1) or (Month > 12) or (Date < 1) or (Date > Mdays[Month])) then
    Result:= '日期輸入有錯誤';
end;


// 檢查輸入的日期是否正確 (格式為 YY, MM, DD)
Function CheckDate1(Y, M, D: String): String;
var
  Year, Month, Date: word;
  EDate: TDatetime;
begin
  Result:= '正確';
  if (Length(Y) = 0) or (Length(M) = 0) or (Length(D) = 0) then
    if (Length(Y) <> 0) or (Length(M) <> 0) or (Length(D) <> 0) then
      begin
        Result:= '日期輸入不完整';
        Exit;
      end;

  if NOT (TestDataRange(Y) and TestDataRange(M) and TestDataRange(D)) then
    begin
      Result:= '日期輸入應為數值型態';
      Exit;
    end;
// 檢查日期是否符合萬年曆
  Year:= StrToInt(Y)+1911;
  Month:= StrToInt(M);
  Date:= StrToInt(D);
  if Month = 12 then
    begin
      Year:= Year+1;
      Month:= 1;
    end
  else
    Month:= Month+1;
  Edate:= EncodeDate(Year, Month, 1);
  EDate:= EDate-1;
  DecodeDate(EDate, Year, Month, Date);
  if StrToInt(D) > Date then
    Result:= '日期輸入有錯誤';
end;


// 檢查輸入的日期是否正確 (格式為89/1/1)
Function CheckDate2(InputData: String): String;
var
  Y, M, D: String;
  i, count: smallint;
begin
  count:= 0;
  for i:= 1 to Length(InputData) do
    begin
      if (count = 0) and (copy(InputData, i, 1) <> '/') then
        Y:= Y+copy(InputData, i, 1)
      else if (count = 1) and (copy(InputData, i, 1) <> '/') then
        M:= M+copy(InputData, i, 1)
      else if (count = 2) then
        D:= D+copy(InputData, i, 1)
      else if (copy(InputData, i, 1) = '/') then
        Inc(Count);
    end;
  Result:= CheckDate(Y, M, D);
end;


// 檢查輸入的身份證編號是否正確
Function CheckIdCode(InputData: String): String;
var
  i, Sum: smallint;
begin
  if Length(InputData) < 10 then
    begin
      Result:= '身份證編號應為十位數';
      Exit;
    end;
  InputData:= UpperCase(InputData);
  if ((copy(InputData, 1, 1)) < 'A') or ((copy(InputData, 1, 1)) > 'Z') then
    begin
      Result:= '身份證編號首碼應為英文字母';
      Exit;
    end;
  for i:=2 to 10 do
    if ((copy(InputData, i, 1)) < '0') or ((copy(InputData, i, 1)) > '9') or ((copy(InputData, i, 1)) = ' ') then
      begin
        Result:= '身份證編號後九碼應為阿拉伯數字';
        Exit;
      end;
//A=10 B=11 C=12 D=13 E=14 F=15 G=16 H=17 J=18 K=19 L=20 M=21 N=22 P=23 Q=24 R=25 S=26 T=27 U=28 V=29 W=30 X=31 Y=32 Z=33 I=34 O=35
//( N1的十位數 + N1的個位數 ×9 + N2 ×8 + N3 ×7 + N4 ×6 + N5 ×5 + N6 ×4 + N7 ×3 + N8 ×2 + N9 + N10 ) ÷10 = 0
  for i:=10 to 35 do
    if (copy(InputData, 1, 1) = copy('ABCDEFGHJKLMNPQRSTUVXYWZIO', i-9, 1)) then
      break;
  Sum:=(i div 10)+(i mod 10)*9+strTOint(copy(InputData, 10, 1));
  for i:=2 to 9 do
    Sum:=Sum+strTOint(copy(InputData, i, 1))*(10-i);
  if (Sum mod 10 <> 0) then
    begin
      Result:= '身份證編號錯誤';
      Exit;
    end
  else
    Result:= '正確';
end;


// 英文日期轉中文日期
Function EnglishToChenese(EDate: TDateTime): Chinesedate;
var
  Year, Month, Day: word;
begin
  DecodeDate(EDate, Year, Month, Day);
  Result.YY:= Year-1911;
  Result.MM:= Month;
  Result.DD:= Day;
end;


// 中文日期轉英文日期
Function ChineseToEnglish(Y, M, D: string): TDateTime;
begin
  Result:= EncodeDate(StrToInt(Y)+1911, StrTOInt(M), StrToInt(D));
end;


// 取得該月的最後一日
Function GetLastDate(Y, M: String): Integer;
var
  Year, Month, Date: word;
  EDate: TDatetime;
begin
  Year:= StrToInt(Y)+1911;
  Month:= StrToInt(M);
  if Month = 12 then
    begin
      Year:= Year+1;
      Month:= 1;
    end
  else
    Month:= Month+1;
  Edate:= EncodeDate(Year, Month, 1);
  EDate:= EDate-1;
  DecodeDate(EDate, Year, Month, Date);
  Result:= Date;
end;


// 將數值四捨五入
Function ChangeValue(Input: Real): String;
var
  Temp: String;
begin
  Str(Input: 10: 2, Temp);
  Result:= TrimLeft(Temp);
end;

Function Flt_Com(Input: String): String;                // 濾豆號
var
  i:Integer;
  OutStr:String;
begin
  OutStr:='';
  for i := 1 to Length(Input) do
  begin
    if Input[i] <> ',' then OutStr:=OutStr+Input[i];
  end;
  Result:=OutStr;
end;

Function Flt_SP(Input: String): String;                // 濾空白
var
  i:Integer;
  OutStr:String;
begin
  OutStr:='';
  for i := 1 to Length(Input) do
  begin
    if Input[i] <> ' ' then OutStr:=OutStr+Input[i];
  end;
  Result:=OutStr;
end;

function Flt_Any(Input:string;M:string):string;
var
  s,t:string;
begin
  s:='';
  if (length(Input)>0) then
  begin
    t:=copy(Input,1,1);
    if (pos(t,M)<1) then s:=t;
    result := s+Flt_Any(copy(Input,2,Length(Input)-1),M);
  end
  else result:= s;
end;

function Fix_Float(data_string:string;data_size:integer) : string;
var SPACE,tmp_1,tmp_2:string;
    i:integer;
begin
   SPACE:='0';
   if Pos('.',data_string) <> 0 then
   begin
     tmp_1:=copy(data_string,1,Pos('.',data_string));
     tmp_2:=copy(data_string,Pos('.',data_string)+1,data_size);

     for i:=length(tmp_2)+1 to 2 do
        tmp_2:=tmp_2+SPACE;
     result:=tmp_1+tmp_2;
   end
   else
   begin
     result:=data_string+'.00';
   end;
end;

Function Rounda(Value: Extended; Len: Word): Extended;
begin
  Result:= StrToFloat(Format('%.'+IntToStr(Len)+'f',[Value]));
//DEMO: Rounda(123.45,1):=123.5;
end;

Function Rounda2(Value: Extended; Len: Word): integer;
begin
  Result:= StrToint(Format('%.'+IntToStr(Len)+'f',[Value]));
//DEMO: Rounda(123.45,1):=123.5;
end;

function Fix_MMDD(data_string:string) : string;
begin
   if copy(data_string,1,1)=' ' then result:='0'+copy(data_string,2,1)
   else result:=data_string;
end;

function CutFirst(S:String):String;
begin
  result:= copy(S,2,Length(S));
end;

function CutLast(S:String):String;
begin
  result := copy(S,1,Length(S)-1);
end;

function ListSHT(SCR_List:String):String;
type
  TItems=array[1..999] of Integer;
var
  i,j,Item,Curr,Last,head :integer;
  Items:TItems;
  ListCNV, sList : String;
begin
  ListCNV:='';
  SCR_List:= Trim(Scr_List);
  if SCR_List[Length(SCR_List)] = ',' then SCR_List:=COPY(SCR_List,1,Length(SCR_List)-1);
  if SCR_List[1] = ',' then SCR_List:=COPY(SCR_List,2,Length(SCR_List));
  SCR_List:= Trim(Scr_List);
  if Trim(Scr_List)<>''then
  begin
    i:=1;
    while pos(',',SCR_List) > 0 do
    begin
      sList := Trim(copy(SCR_List,1,pos(',',SCR_List)-1));
      if Pos('00',sList)>0 then
      begin
         Delete(sList,1,2);
      end
      else if Pos('0',sList)>0 then
      begin
         Delete(sList,1,1);
      end;
      Item:=StrToInt(sList);
      SCR_List:=Copy(SCR_List,pos(',',SCR_List)+1,Length(SCR_List)-pos(',',SCR_List));
      Items[i]:=Item;
      i:=i+1;
    end;
    Item:=StrToInt(SCR_List);
    Items[i]:=Item;
    Last:=i;
    for i:=1 to last-1 do
    begin
      for j:= 1 to last-1 do
      begin
        if Items[j]> Items[j+1] then
        begin
          Item:=Items[j];
          Items[j]:=Items[j+1];
          Items[j+1]:=Item;
        end;
      end;
    end;
    Curr:=Items[1];
    head:=Items[1];
    ListCNV:=IntToStr(Items[1]);
    for i := 2 to last do
    begin
      if Items[i] <> Curr+1 then
      begin
        ListCNV:=ListCNV+','+IntToStr(Items[i]);
        head:=Items[i];
      end
      else
      begin
        If Items[i] = head+1 then
        begin
          ListCNV:=ListCNV+'~'+IntToStr(Items[i]);
        end
        else
        begin
          ListCNV:=copy(ListCNV,1,Length(ListCNV)-Length(IntToStr(Curr)))+IntToStr(Items[i]);
        end;
      end;
      Curr:=Items[i];
    end;
  end;
  result:=ListCNV;
end;

function Fix_Size(data_string:string;data_size:integer) : string;
var SPACE,tmp:string;
    i:integer;
begin

   SPACE:=' ';
   tmp:=copy(data_string,1,data_size);

   for i:=length(tmp)+1 to data_size do
      tmp:=tmp+SPACE;
   result:=tmp;

end;

function SPACE(data_size:integer) : string;
var STYLE,tmp:string;
    i:integer;
begin

   STYLE:=' ';
   tmp:='';

   for i:=length(tmp)+1 to data_size do
      tmp:=tmp+STYLE;
   result:=tmp;
end;

function Fix_SizeL(data_string:string;data_size:integer) : string;
var SPACE,tmp:string;
    i:integer;
begin

   SPACE:=' ';
   tmp:=copy(data_string,1,data_size);

   for i:=length(tmp)+1 to data_size do
      tmp:=SPACE+tmp;
   result:=tmp;

end;


function Fix_Num(NumVar:integer;data_size:integer) : string;
var SPACE,tmp:string;
    i:integer;
begin

   SPACE:='0';
   tmp:=inttostr(NumVar);
   for i:=length(tmp)+1 to data_size do
      tmp:=SPACE+tmp;
   result:=tmp;

end;

function Fix_SNum(data_string:string;data_size:integer) : string;
var SPACE,tmp:string;
    i,sw:integer;
begin

   if trim(data_string) = '' then Result:='';
   tmp:='';
   SPACE:='0';
   try
     sw:=StrToInt(data_string);
   except
     sw:=0;
   end;
   if sw<>0 then
   begin
     tmp:= copy(IntToStr(sw),1,data_Size);
     for i:=length(tmp)+1 to data_size do
     begin
       tmp:=SPACE+tmp;
     end;
   end;
   result:=tmp;
end;

function Chk_Num(NumVar:String) : Boolean;
var sw:integer;
begin
   Result:=True;
   try
     sw:=StrToInt(NumVar);
   except
     Result:=False;
   end;

end;

function DBGridRecordSize(mColumn: TColumn):   Boolean;
var
 M : Integer;
begin
    Result :=   False;
    if not Assigned(mColumn.Field) then Exit;
    mColumn.Field.Tag := Max(mColumn.Field.Tag,
        TDBGrid(mColumn.Grid).Canvas.TextWidth(mColumn.Field.DisplayText));
    M := mColumn.Field.Tag;
    Result := True;
end;

function DBGridAutoSize(mDBGrid: TDBGrid; mOffset: Integer = 5):   Boolean;
var
    I: Integer;
begin
    Result := False;
    if not Assigned(mDBGrid) then Exit;
    if not Assigned(mDBGrid.DataSource) then Exit;
    if not Assigned(mDBGrid.DataSource.DataSet) then Exit;
    if not mDBGrid.DataSource.DataSet.Active then Exit;

    for I := 0 to mDBGrid.Columns.Count - 1 do
    begin
        if not mDBGrid.Columns[I].Visible then   Continue;
        if Assigned(mDBGrid.Columns[I].Field) then
           mDBGrid.Columns[I].Width := Max(mDBGrid.Columns[I].Field.Tag,
           mDBGrid.Canvas.TextWidth(mDBGrid.Columns[I].Title.Caption))+ mOffset
        else mDBGrid.Columns[I].Width :=
             mDBGrid.Canvas.TextWidth(mDBGrid.Columns[I].Title.Caption)+ mOffset;
        mDBGrid.Refresh;
    end;
   Result := True;
end;

function Strsplit(input: string; schar: Char; sindex: cardinal):string;
var
  ii,kk:cardinal;
begin
  input := input+schar;
  kk:=0;
  for ii := 0 to length(input)-1 do
  begin
    if input[ii+1]=schar then begin
      if sindex=0 then begin
        result:=copy(input,kk+1,ii-kk);
        exit;
      end;
      dec(sindex);
      kk:=ii+1;
    end;
  end;
  result:=copy(input,kk+1,ii-kk+1);
end;

function CheckExamQuestionStruc(ExamNo, SubNo: string): Boolean;
var
  SQLstr: string;
begin
  Result:= False;
  SQLstr:= ' SELECT Count(*) as Cnt FROM Exam_Question'+
           ' WHERE Exam_No = '+  QuotedStr(ExamNo)+
           ' And Sub_No = '+ QuotedStr(SubNo);
  OpenSQL(DM.qryTemp, SQLstr);
  if DM.qryTemp.IsEmpty then
    Exit;
  DM.qryTemp.First;
  Result:= DM.qryTemp.FieldByName('Cnt').AsInteger > 1;
end;

function CheckExamQuestionRead(DataPath, ReadAccess: string): Boolean;
begin
  Result:= False;
  if FileExists(DataPath + '\' + ReadAccess + '.mdb')   then
  Result:= true;
end;
end.
