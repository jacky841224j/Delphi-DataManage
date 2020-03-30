unit PAExcel;

interface
 uses
  Windows,classes, Messages, Dialogs, Variants, SysUtils, Controls,  DB,ADODB
  ,Comobj,Excel2000, Clipbrd;


 procedure PAExcelExport(AQuery : TADOQuery; WorkSheetName, SaveFileName : String);
 procedure PAExcelExport2(IData : TStringList; WorkSheetName, SaveFileName : String);

implementation

{
 參數說明：
  AQuery : 查詢SQL Open後再傳入
  WorkSheetName : 自訂Excel工作表名稱
  SaveFileName : 存檔名稱
}
//SQL資料匯出Excel
procedure PAExcelExport(AQuery : TADOQuery; WorkSheetName, SaveFileName : String);
var
 ExcelApp : Variant;
 StrFormat : OleVariant;
 I : Integer;
 DataStr : String;
 LData : TStringList;
begin
   try
    if AQuery.RecordCount > 0 then
     AQuery.Recordset.MoveFirst;

   except
    ShowMessage('ADOQuery Linker Error');
    Exit;
   end;

   LData := TStringList.Create;
   LData.BeginUpdate;
   With AQuery.Recordset do
   while not Eof do
   begin
    DataStr := '';
    for I := 0 to Fields.Count - 1 do
     DataStr := DataStr + Trim(VarToStr(Fields[I].Value)) + #9;

    System.Delete(DataStr,Length(DataStr), 1);
    LData.Append(DataStr);
    MoveNext;
   end;

   DataStr := '';
   for I := 0 to AQuery.Fields.Count - 1 do
    DataStr := DataStr + AQuery.Fields[I].DisplayName + #9;

   System.Delete(DataStr,Length(DataStr), 1);
   LData.Insert(0,DataStr);
   LData.EndUpdate;



   try
    ExcelApp:= CreateOleObject('Excel.Application'); //Create Excel 物件
   except
     ShowMessage('請檢查電腦是否安裝Excel...');
     LData.Free;
     Exit;
   end;

   if FileExists(SaveFileName) then
    DeleteFile(SaveFileName);

   Excelapp.WorkBooks.Add; //新增工作簿(預設為三個工作表)
   ExcelApp.Visible := False; //不顯示Excel 視窗
   ExcelApp.WorkSheets[1].Activate;
   ExcelApp.WorkSheets[1].Name:= WorkSheetName; //工作表更名
   strFormat := '@';  //@: 儲存格格式改為文字
   ExcelApp.WorkSheets[1].Cells.NumberFormatLocal := strFormat; //設定儲格格式(一定要宣告OleVariant，直接等於'@'無
   Clipboard.Clear;  //先清空剪貼簿
   Clipboard.AsText := LData.Text; //複製資料到剪貼簿
   ExcelApp.Range['A1'].PasteSpecial; //在A1貼上
   Clipboard.Clear;  //用完清空剪貼簿

  // ExcelApp.Range['A1'].Select;
   ExcelApp.Selection.Columns.AutoFit;  //自動調整全部欄寬
   ExcelApp.Range['A1:D1'].ColumnWidth := 3; //設定欄寬(可給小數點)
   ExcelApp.Range['F1:I1'].ColumnWidth := 3; //設定欄寬(可給小數點)
   ExcelApp.Range['K1:N1'].ColumnWidth := 3; //設定欄寬(可給小數點)

   ExcelApp.Range['A1'].WrapText := True; //欄位文字自動換行
   ExcelApp.Range['B1'].WrapText := True; //欄位文字自動換行
   ExcelApp.Range['C1'].WrapText := True; //欄位文字自動換行
   ExcelApp.Range['D1'].WrapText := True; //欄位文字自動換行
   ExcelApp.Range['F1'].WrapText := True; //欄位文字自動換行
   ExcelApp.Range['G1'].WrapText := True; //欄位文字自動換行
   ExcelApp.Range['H1'].WrapText := True; //欄位文字自動換行
   ExcelApp.Range['I1'].WrapText := True; //欄位文字自動換行
   ExcelApp.Range['K1'].WrapText := True; //欄位文字自動換行
   ExcelApp.Range['L1'].WrapText := True; //欄位文字自動換行
   ExcelApp.Range['M1'].WrapText := True; //欄位文字自動換行
   ExcelApp.Range['N1'].WrapText := True; //欄位文字自動換行

   ExcelApp.Selection.CurrentRegion.Select; //只選取連續有值非Null的欄位(有空值的欄位即停止)

   for i := 1 to 4 do
     ExcelApp.Selection.Borders[I].LineStyle := 1; //設定邊框 (1: 實線2: 虛線)

   ExcelApp.ActiveWorkBook.Saved := True; //設定不存檔，若不設定關閉時會出現"是否存檔的對話框"
   ExcelApp.WorkBooks[1].SaveAs(SaveFileName); //存檔
   ExcelApp.WorkBooks.close;  //關閉Excel
   ExcelApp.Quit;             //離開Excel
   ExcelApp:=Unassigned;      //釋放ExcelApp;

   LData.Free;
end;

procedure PAExcelExport2(IData : TStringList; WorkSheetName, SaveFileName : String);
var
 ExcelApp : Variant;
 StrFormat : OleVariant;
 I : Integer;
 DataStr : String;
begin
   try
    ExcelApp:= CreateOleObject('Excel.Application'); //Create Excel 物件
   except
     ShowMessage('請檢查電腦是否安裝Excel...');
     IData.Free;
     Exit;
   end;

   if FileExists(SaveFileName) then
    DeleteFile(SaveFileName);

   Excelapp.WorkBooks.Add; //新增工作簿(預設為三個工作表)
   ExcelApp.Visible := False; //不顯示Excel 視窗
   ExcelApp.WorkSheets[1].Activate;
   ExcelApp.WorkSheets[1].Name:= WorkSheetName; //工作表更名
   strFormat := '@';  //@: 儲存格格式改為文字
   ExcelApp.WorkSheets[1].Cells.NumberFormatLocal := strFormat; //設定儲格格式(一定要宣告OleVariant，直接等於'@'無
   Clipboard.Clear;  //先清空剪貼簿
   Clipboard.AsText := IData.Text; //複製資料到剪貼簿
   ExcelApp.Range['A1'].PasteSpecial; //在A1貼上
   Clipboard.Clear;  //用完清空剪貼簿

  // ExcelApp.Range['A1'].Select;
   ExcelApp.Selection.Columns.AutoFit;  //自動調整全部欄寬
   ExcelApp.Range['A1:K1'].ColumnWidth := 7; //設定欄寬(可給小數點)

   ExcelApp.Range['A1'].WrapText := True; //欄位文字自動換行
   ExcelApp.Range['B1'].WrapText := True; //欄位文字自動換行
   ExcelApp.Range['C1'].WrapText := True; //欄位文字自動換行
   ExcelApp.Range['D1'].WrapText := True; //欄位文字自動換行
   ExcelApp.Range['F1'].WrapText := True; //欄位文字自動換行
   ExcelApp.Range['G1'].WrapText := True; //欄位文字自動換行
   ExcelApp.Range['H1'].WrapText := True; //欄位文字自動換行
   ExcelApp.Range['I1'].WrapText := True; //欄位文字自動換行
   ExcelApp.Range['K1'].WrapText := True; //欄位文字自動換行
   ExcelApp.Range['L1'].WrapText := True; //欄位文字自動換行
   ExcelApp.Range['M1'].WrapText := True; //欄位文字自動換行
   ExcelApp.Range['N1'].WrapText := True; //欄位文字自動換行

   ExcelApp.Selection.CurrentRegion.Select; //只選取連續有值非Null的欄位(有空值的欄位即停止)

   for i := 1 to 4 do
     ExcelApp.Selection.Borders[I].LineStyle := 1; //設定邊框 (1: 實線2: 虛線)

   ExcelApp.ActiveWorkBook.Saved := True; //設定不存檔，若不設定關閉時會出現"是否存檔的對話框"
   ExcelApp.WorkBooks[1].SaveAs(SaveFileName); //存檔
   ExcelApp.WorkBooks.close;  //關閉Excel
   ExcelApp.Quit;             //離開Excel
   ExcelApp:=Unassigned;      //釋放ExcelApp;

   IData.Clear;
   IData.Free;
end;


end.
