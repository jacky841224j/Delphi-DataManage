unit PAExcel;

interface
 uses
  Windows,classes, Messages, Dialogs, Variants, SysUtils, Controls,  DB,ADODB
  ,Comobj,Excel2000, Clipbrd;


 procedure PAExcelExport(AQuery : TADOQuery; WorkSheetName, SaveFileName : String);
 procedure PAExcelExport2(IData : TStringList; WorkSheetName, SaveFileName : String);

implementation

{
 �Ѽƻ����G
  AQuery : �d��SQL Open��A�ǤJ
  WorkSheetName : �ۭqExcel�u�@��W��
  SaveFileName : �s�ɦW��
}
//SQL��ƶץXExcel
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
    ExcelApp:= CreateOleObject('Excel.Application'); //Create Excel ����
   except
     ShowMessage('���ˬd�q���O�_�w��Excel...');
     LData.Free;
     Exit;
   end;

   if FileExists(SaveFileName) then
    DeleteFile(SaveFileName);

   Excelapp.WorkBooks.Add; //�s�W�u�@ï(�w�]���T�Ӥu�@��)
   ExcelApp.Visible := False; //�����Excel ����
   ExcelApp.WorkSheets[1].Activate;
   ExcelApp.WorkSheets[1].Name:= WorkSheetName; //�u�@���W
   strFormat := '@';  //@: �x�s��榡�אּ��r
   ExcelApp.WorkSheets[1].Cells.NumberFormatLocal := strFormat; //�]�w�x��榡(�@�w�n�ŧiOleVariant�A��������'@'�L
   Clipboard.Clear;  //���M�ŰŶKï
   Clipboard.AsText := LData.Text; //�ƻs��ƨ�ŶKï
   ExcelApp.Range['A1'].PasteSpecial; //�bA1�K�W
   Clipboard.Clear;  //�Χ��M�ŰŶKï

  // ExcelApp.Range['A1'].Select;
   ExcelApp.Selection.Columns.AutoFit;  //�۰ʽվ������e
   ExcelApp.Range['A1:D1'].ColumnWidth := 3; //�]�w��e(�i���p���I)
   ExcelApp.Range['F1:I1'].ColumnWidth := 3; //�]�w��e(�i���p���I)
   ExcelApp.Range['K1:N1'].ColumnWidth := 3; //�]�w��e(�i���p���I)

   ExcelApp.Range['A1'].WrapText := True; //����r�۰ʴ���
   ExcelApp.Range['B1'].WrapText := True; //����r�۰ʴ���
   ExcelApp.Range['C1'].WrapText := True; //����r�۰ʴ���
   ExcelApp.Range['D1'].WrapText := True; //����r�۰ʴ���
   ExcelApp.Range['F1'].WrapText := True; //����r�۰ʴ���
   ExcelApp.Range['G1'].WrapText := True; //����r�۰ʴ���
   ExcelApp.Range['H1'].WrapText := True; //����r�۰ʴ���
   ExcelApp.Range['I1'].WrapText := True; //����r�۰ʴ���
   ExcelApp.Range['K1'].WrapText := True; //����r�۰ʴ���
   ExcelApp.Range['L1'].WrapText := True; //����r�۰ʴ���
   ExcelApp.Range['M1'].WrapText := True; //����r�۰ʴ���
   ExcelApp.Range['N1'].WrapText := True; //����r�۰ʴ���

   ExcelApp.Selection.CurrentRegion.Select; //�u����s�򦳭ȫDNull�����(���ŭȪ����Y����)

   for i := 1 to 4 do
     ExcelApp.Selection.Borders[I].LineStyle := 1; //�]�w��� (1: ��u2: ��u)

   ExcelApp.ActiveWorkBook.Saved := True; //�]�w���s�ɡA�Y���]�w�����ɷ|�X�{"�O�_�s�ɪ���ܮ�"
   ExcelApp.WorkBooks[1].SaveAs(SaveFileName); //�s��
   ExcelApp.WorkBooks.close;  //����Excel
   ExcelApp.Quit;             //���}Excel
   ExcelApp:=Unassigned;      //����ExcelApp;

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
    ExcelApp:= CreateOleObject('Excel.Application'); //Create Excel ����
   except
     ShowMessage('���ˬd�q���O�_�w��Excel...');
     IData.Free;
     Exit;
   end;

   if FileExists(SaveFileName) then
    DeleteFile(SaveFileName);

   Excelapp.WorkBooks.Add; //�s�W�u�@ï(�w�]���T�Ӥu�@��)
   ExcelApp.Visible := False; //�����Excel ����
   ExcelApp.WorkSheets[1].Activate;
   ExcelApp.WorkSheets[1].Name:= WorkSheetName; //�u�@���W
   strFormat := '@';  //@: �x�s��榡�אּ��r
   ExcelApp.WorkSheets[1].Cells.NumberFormatLocal := strFormat; //�]�w�x��榡(�@�w�n�ŧiOleVariant�A��������'@'�L
   Clipboard.Clear;  //���M�ŰŶKï
   Clipboard.AsText := IData.Text; //�ƻs��ƨ�ŶKï
   ExcelApp.Range['A1'].PasteSpecial; //�bA1�K�W
   Clipboard.Clear;  //�Χ��M�ŰŶKï

  // ExcelApp.Range['A1'].Select;
   ExcelApp.Selection.Columns.AutoFit;  //�۰ʽվ������e
   ExcelApp.Range['A1:K1'].ColumnWidth := 7; //�]�w��e(�i���p���I)

   ExcelApp.Range['A1'].WrapText := True; //����r�۰ʴ���
   ExcelApp.Range['B1'].WrapText := True; //����r�۰ʴ���
   ExcelApp.Range['C1'].WrapText := True; //����r�۰ʴ���
   ExcelApp.Range['D1'].WrapText := True; //����r�۰ʴ���
   ExcelApp.Range['F1'].WrapText := True; //����r�۰ʴ���
   ExcelApp.Range['G1'].WrapText := True; //����r�۰ʴ���
   ExcelApp.Range['H1'].WrapText := True; //����r�۰ʴ���
   ExcelApp.Range['I1'].WrapText := True; //����r�۰ʴ���
   ExcelApp.Range['K1'].WrapText := True; //����r�۰ʴ���
   ExcelApp.Range['L1'].WrapText := True; //����r�۰ʴ���
   ExcelApp.Range['M1'].WrapText := True; //����r�۰ʴ���
   ExcelApp.Range['N1'].WrapText := True; //����r�۰ʴ���

   ExcelApp.Selection.CurrentRegion.Select; //�u����s�򦳭ȫDNull�����(���ŭȪ����Y����)

   for i := 1 to 4 do
     ExcelApp.Selection.Borders[I].LineStyle := 1; //�]�w��� (1: ��u2: ��u)

   ExcelApp.ActiveWorkBook.Saved := True; //�]�w���s�ɡA�Y���]�w�����ɷ|�X�{"�O�_�s�ɪ���ܮ�"
   ExcelApp.WorkBooks[1].SaveAs(SaveFileName); //�s��
   ExcelApp.WorkBooks.close;  //����Excel
   ExcelApp.Quit;             //���}Excel
   ExcelApp:=Unassigned;      //����ExcelApp;

   IData.Clear;
   IData.Free;
end;


end.
