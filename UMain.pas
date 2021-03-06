﻿unit UMain;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, StdCtrls, Buttons, ComCtrls, PAExcel, Excel2000, Comobj,
  Clipbrd, CheckLst, FileCtrl, Grids, DBGrids, DB, ADODB, Gauges;

type
  TfmMain = class(TForm)
    Panel1: TPanel;
    Label1: TLabel;
    Label10: TLabel;
    Label11: TLabel;
    Label12: TLabel;
    Label14: TLabel;
    SpeedButton1: TSpeedButton;
    Bevel1: TBevel;
    ClassGroup: TRadioGroup;
    cmbExam: TComboBox;
    btnExit: TBitBtn;
    edtUser: TEdit;
    edtPW: TEdit;
    cmbIP: TComboBox;
    cmbDBName: TComboBox;
    PageControl1: TPageControl;
    TabSheet1: TTabSheet;
    TabSheet2: TTabSheet;
    TabSheet3: TTabSheet;
    TabSheet4: TTabSheet;
    Label2: TLabel;
    btnSchOut: TBitBtn;
    btnSchIn: TBitBtn;
    SaveDialog1: TSaveDialog;
    DBGrid1: TDBGrid;
    OpenDialog1: TOpenDialog;
    adods1: TADODataSet;
    Conn1: TADOConnection;
    Label3: TLabel;
    btnClsOut: TBitBtn;
    btnClsIn: TBitBtn;
    DBGrid2: TDBGrid;
    DBGrid3: TDBGrid;
    CheckBox1: TCheckBox;
    btnSet: TBitBtn;
    CheckListBox1: TCheckListBox;
    CheckBox2: TCheckBox;
    CheckListBox2: TCheckListBox;
    BitBtn1: TBitBtn;
    CheckBox3: TCheckBox;
    edtTNo: TEdit;
    BitBtn2: TBitBtn;
    Memo1: TMemo;
    TabSheet5: TTabSheet;
    CheckListBox3: TCheckListBox;
    Label4: TLabel;
    cbSub: TComboBox;
    Memo2: TMemo;
    btnAbsent: TButton;
    CheckBox4: TCheckBox;
    btnSearch: TButton;
    btnOut: TBitBtn;
    ProgressBar1: TProgressBar;
    cbSubT: TComboBox;
    Label5: TLabel;
    Edit1: TEdit;
    Edit2: TEdit;
    TabSheet6: TTabSheet;
    CheckListBox4: TCheckListBox;
    CheckBox5: TCheckBox;
    btnExcel: TButton;
    mmoCheck: TMemo;
    lbl1: TLabel;
    g1: TGauge;
    gTotal: TGauge;
    lblTotal: TLabel;
    lblNow: TLabel;
    lbl2: TLabel;
    btnExcelAll: TButton;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormDestroy(Sender: TObject);
    procedure btnExitClick(Sender: TObject);
    procedure cmbIPChange(Sender: TObject);
    procedure SpeedButton1Click(Sender: TObject);
    procedure cmbDBNameClick(Sender: TObject);
    procedure btnSchOutClick(Sender: TObject);
    procedure btnSchInClick(Sender: TObject);
    procedure btnClsOutClick(Sender: TObject);
    procedure btnClsInClick(Sender: TObject);
    procedure CheckBox1Click(Sender: TObject);
    procedure PageControl1Change(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure btnSetClick(Sender: TObject);
    procedure CheckBox2Click(Sender: TObject);
    procedure BitBtn1Click(Sender: TObject);
    procedure cbSubChange(Sender: TObject);
    procedure CheckBox4Click(Sender: TObject);
    procedure btnAbsentClick(Sender: TObject);
    procedure btnSearchClick(Sender: TObject);
    procedure btnOutClick(Sender: TObject);
    procedure BitBtn2Click(Sender: TObject);
    procedure cmbExamChange(Sender: TObject);
    procedure CheckListBox4ClickCheck(Sender: TObject);
    procedure CheckBox5Click(Sender: TObject);
    procedure btnExcelClick(Sender: TObject);
    procedure btnExcelAllClick(Sender: TObject);
  private
    { Private declarations }
  public
    SQLStr: string;
    { Public declarations }
    procedure TransStd(bRange: Boolean);  //學生資料轉檔
  end;

var
  fmMain: TfmMain;

implementation

{$R *.dfm}

uses
  DataM, StrUtil;

procedure TfmMain.btnExcelAllClick(Sender: TObject);
var
  str, str2, strTitle, strTitle2, strtemp,title: string;
  SQLtemp, SListA: TStringList;
  i, x, y, z, SubCount, QueCounut, RowsCount, ColCount: integer;
  ExcelApp: Variant;
  strFormat: OleVariant;
begin
  try
    ExcelApp := CreateOleObject('Excel.Application'); //Create Excel 物件
    ExcelApp.WorkBooks.Add; //新增工作簿(預設為三個工作表)
  except
    ShowMessage('尚未安裝任何 Excel 版本.');
    Exit;
  end;

  SQLtemp := TStringList.Create;
  SListA := TStringList.Create;

  {$region '插入標題'}
  strTitle := '學校代碼' + #9 + '學校名稱' + #9 + '分校/分班註記' + #9 + '測驗科目代碼' + #9 + '年級' + #9 + '班級代碼' + #9 + '班級名稱' + #9 + '座號' + #9 + '學生姓名' + #9 + '性別代碼' + #9 + '導師姓名' + #9 + '身分註記' + #9 + #9 + #9 + #9 + #9 + #9 + #9 + '學生原始作答反應' + #9;
  strTitle2 := #9 + #9 + #9 + #9 + #9 + #9 + #9 + #9 + #9 + #9 + #9 + '資賦優異' + #9 + '資源班' + #9 + '原住民子女' + #9'新住民子女' + #9 + '才藝班學生' + #9 + '體育班學生' + #9 + '非學校型態實驗教育者（在家教育）' + #9;
  {$endregion}

  SaveDialog1.Filter := '.xlsx|*.xlsx';
  SaveDialog1.FileName := '預設名稱';
  if SaveDialog1.Execute then
  begin
    try
      try
        {$region '儲存考試名稱'}
        SQLStr := 'SELECT Exam_Name FROM Exam ';
        OpenSQL(DM.qryTemp, SQLStr);
        title := DM.qryTemp.FieldByName('Exam_Name').AsString;
        {$endregion}

        {$region '判斷科目總數'}
        SQLStr := 'SELECT es.Sub_No,Sub_Name FROM Exam_Sub es '
                + 'left join Ex_Subject ej on es.Sub_No = ej.Sub_No '
                + 'where Exam_No =' + QuotedStr(cmbExam.text);
        OpenSQL(DM.qryTemp, SQLStr);
        SubCount := DM.qryTemp.RecordCount;
        {$endregion}

        gTotal.MaxValue := SubCount;

        //分類-科目
        for x := 1 to SubCount do
        begin
          g1.Progress := 0;

          {$region 'SQL科目語法'}
          SQLStr := ' SELECT left(st.[Address],6) AS 學校代碼, '
          //                  + '       s.Sch_Memo as 學校代碼, '
                  + '       s.Sch_Name as 學校名稱, '
                  + '       case len(s.Sch_Memo) when  6 then 0  when 7 then substring(s.Sch_Memo,7,7) else 0 end as 分校註記,'
                  + '       case Sub_No when 01 then ''C'' when 02 then ''M''  when 03 then ''E'' end AS 測驗科目代碼, '
                  + '       sc.grade as 年級, '
                  + '       sc.Sch_Code as 班級代碼, '//                  + '       B_Year as 班級代碼, '
                  + '       (sc.grade +''年''+rtrim(sc.Class_Name)+''班'') as 班級名稱, '
                  //                  + '       (sc.grade +''年''+rtrim(B_Year)+''班'') as 班級名稱, '
                  + '       st.Seat_No as 座號, '
                  + '       Student_Name as 學生姓名, '
                  + '       Sex as 性別代碼, '
                  + '       Remark1 as 導師姓名,'
                  + '       Remark as 身分註記,'
                  //                  + '       Remark1 as 身分註記,'
                  + '       ans as 原始作答反映 '
                  + ' FROM Sub_Score ss ' + ' left join School  s on left([Student_No],3) = s.Sch_Code '
                  + ' left join Student st on ss.Student_No = st.Student_No '
                  + ' left join Sch_Class sc on s.Sch_Code = sc.Sch_Code and sc.Class_No = st.Class_No '
                  + ' where Sub_No = ' + DM.qryTemp.FieldByName('Sub_No').AsString + ' and  Exam_No = ' + QuotedStr(cmbExam.Text)
//                  + ' where Sub_No = '+DM.qryTemp.FieldByName('Sub_No').AsString+' and left(ss.[Student_No],3) = '+ Copy(mmoCheck.Lines[i],0,3) + ' and  Exam_No = ' +QuotedStr(cmbExam.Text)
                  + ' order by s.Sch_Code,Sub_No,sc.grade,st.Seat_No';
          OpenSQL(DM.qrySearch, SQLStr);

          {$endregion}

          g1.MaxValue := DM.qrySearch.RecordCount + 7;
          g1.Progress := g1.Progress + 1;

          {$region 'SQL總題數語法'}
          SQLStr := 'Select Count(*) as iCount From Sub_Ans' + ' Where Exam_No=' + #39 + Trim(cmbExam.Text) + #39 + ' And Sub_No=' + #39 + DM.qryTemp.FieldByName('Sub_No').AsString + #39;
          OpenSQL(DM.qrySClass, SQLStr);
          QueCounut := DM.qrySClass.FieldByName('iCount').AsInteger;
          {$endregion}

          g1.Progress := g1.Progress + 1;

          {$region '儲存Excel Title'}
          strTitle2 := '';
          strTitle2 := #9 + #9 + #9 + #9 + #9 + #9 + #9 + #9 + #9 + #9 + #9 + '資賦優異' + #9 + '資源班' + #9 + '原住民子女' + #9'新住民子女' + #9 + '才藝班學生' + #9 + '體育班學生' + #9 + '非學校型態實驗教育者（在家教育）' + #9;
          for y := 1 to QueCounut do
          begin
            strTitle2 := strTitle2 + '第' + IntToStr(y) + '題' + #9;
          end;

          SListA.Append(strTitle);
          SListA.Append(strTitle2);
          {$endregion}

          g1.Progress := g1.Progress + 1;
          lblNow.Caption := '設定標題';

          {$region 'Excel初始化設定'}
          g1.Progress := g1.Progress + 1;
          ExcelApp.DisplayAlerts := False;  //關閉提示視窗
          ExcelApp.Visible := FALSE; //不顯示Excel 視窗
          if ExcelApp.WorkSheets.Count < 3 then
            ExcelApp.WorkBooks[1].WorkSheets.Add(null, null, 2, xlWorksheet);
          ExcelApp.WorkSheets[x].Activate;
          ExcelApp.WorkSheets[x].Name := '科目名稱_' + DM.qryTemp.FieldByName('Sub_Name').AsString; //工作表更名
          strFormat := '@'; //@: 儲存格格式改為文字
          ExcelApp.WorkSheets[x].Cells.NumberFormatLocal := strFormat; //設定儲格格式(一定要宣告OleVariant，直接等於'@'無
          {$endregion}

          g1.Progress := g1.Progress + 1;
          lblNow.Caption := '初始化Excel設定';

          //分類-考生
          for y := 0 to DM.qrySearch.RecordCount - 1 do
          begin
            g1.Progress := g1.Progress + 1;
            lblNow.Caption := '讀取考生資料';
            {$region '切割身分註記'}
            SQLtemp.Clear;
            str := '';
            SQLtemp.Add(DM.qrySearch.FieldByName('身分註記').AsString);
            SQLtemp.Delimiter := '-';
            SQLtemp.DelimitedText := SQLtemp.Text;
            for z := 0 to SQLtemp.Count - 1 do
            begin
              str := str + SQLtemp[z] + #9;
            end;
            {$endregion}

            {$region '切割原始作答'}
            str2 := '';
            SQLtemp.Clear;
            SQLtemp.Add(DM.qrySearch.FieldByName('原始作答反映').AsString);
            SQLtemp.Delimiter := ',';
            SQLtemp.DelimitedText := SQLtemp.Text;
            for z := 0 to SQLtemp.Count - 1 do
            begin
              if Length(SQLtemp[z]) > 1 then
                str2 := str2 + '?' + #9
              else
                str2 := str2 + SQLtemp[z] + #9;
            end;
            {$endregion}

            {$region '儲存資料至剪貼簿'}
            SListA.Add(DM.qrySearch.FieldByName('學校代碼').AsString + #9 + DM.qrySearch.FieldByName('學校名稱').AsString + #9 + DM.qrySearch.FieldByName('分校註記').AsString + #9 + DM.qrySearch.FieldByName('測驗科目代碼').AsString + #9 + DM.qrySearch.FieldByName('年級').AsString + #9 + Trim(DM.qrySearch.FieldByName('班級代碼').AsString) + #9 + Trim(DM.qrySearch.FieldByName('班級名稱').AsString) + #9 + DM.qrySearch.FieldByName('座號').AsString + #9 + DM.qrySearch.FieldByName('學生姓名').AsString + #9 + DM.qrySearch.FieldByName('性別代碼').AsString + #9 + DM.qrySearch.FieldByName('導師姓名').AsString + #9 + str + #9    //身分註記
            + str2);   //原始作答
            {$endregion}

            DM.qrySearch.Next;
          end;
          Clipboard.Clear; //先清空剪貼簿
          Clipboard.AsText := SListA.Text; //複製資料到剪貼簿

          {$region '設定Excel 格式'}
          ExcelApp.Range['A1'].Select;
          ExcelApp.Range['A1'].PasteSpecial; //在A1貼上
          RowsCount :=  ExcelApp.ActiveSheet.UsedRange.Rows.Count;
          ColCount :=   ExcelApp.ActiveSheet.UsedRange.Columns.Count;
          g1.Progress := g1.Progress + 1;
          lblNow.Caption := '設定Excel格式';

          //設定背景色
          ExcelApp.Selection.FormatConditions.Add(xlExpression,Formula1:='=MOD(ROW(),2)=0');
          ExcelApp.Selection.FormatConditions[1].Interior.Color := RGB(235, 241, 222);
          ExcelApp.Selection.FormatConditions.Add(xlExpression,Formula1:='=MOD(ROW(),2)=1');
          ExcelApp.Selection.FormatConditions[2].Interior.Color := RGB(242, 220, 219);
          ExcelApp.Rows['1:2'].FormatConditions.Delete;
          ExcelApp.Rows[1].Interior.Color := RGB(153, 204, 255);
          ExcelApp.Range['S1'].Interior.Color := RGB(255, 255, 0);
          ExcelApp.Range['L2:R2'].Interior.Color := RGB(197, 217, 241);
          ExcelApp.Range[ExcelApp.cells[2, 19], ExcelApp.cells[2, QueCounut + 18]].Interior.Color := RGB(255, 255, 204);
          //合併儲存格
          ExcelApp.ActiveSheet.Range[ExcelApp.cells[1, 19], ExcelApp.cells[1, QueCounut + 18]].Merge;
          ExcelApp.ActiveSheet.Range['L1:R1'].Merge;
          for y := 1 to 11 do
            ExcelApp.ActiveSheet.Range[ExcelApp.cells[1, y], ExcelApp.cells[2, y]].Merge;
          //框線
//          ExcelApp.rows['3:'+inttostr(RowsCount)].Borders[2].LineStyle := 1;
//          ExcelApp.rows['3:'+inttostr(RowsCount)].Borders[2].Weight := 2;
          ExcelApp.rows['3:'+inttostr(RowsCount)].Borders[4].LineStyle := 1;
          ExcelApp.rows['3:'+inttostr(RowsCount)].Borders[4].Weight := 2;
          //置中
          ExcelApp.Cells.HorizontalAlignment := -4108;
          //凍結欄位
          ExcelApp.ActiveWindow.SplitColumn := 0;
          ExcelApp.ActiveWindow.SplitRow  := 2;
          ExcelApp.ActiveWindow.FreezePanes := true;
          //自適欄寬
          ExcelApp.Cells.EntireColumn.AutoFit;
          ExcelApp.Cells.Font.Name := '微軟正黑體';
          {$endregion}

          SListA.clear;
          g1.Progress := g1.MaxValue;
          lblNow.Caption := '完成';
          gTotal.Progress := gTotal.Progress + 1;
          DM.qryTemp.next; //切換下一筆科目
        end;
        gTotal.Progress := gTotal.MaxValue;
        ExcelApp.ActiveWorkBook.Saved := True; //設定不存檔，若不設定關閉時會出現"是否存檔的對話框"
        ExcelApp.WorkBooks[1].SaveAs(ExtractFilePath(SaveDialog1.FileName)+title + '_作答反應.xlsx'); //存檔
        ExcelApp.WorkBooks.close; //關閉Excel
        ShowMessage('匯出完成');
      except
        ShowMessage('匯出錯誤 - '+lblNow.Caption);
      end;
    finally
      SListA.Free;
      SQLtemp.Free;
      ExcelApp.Quit; //離開Excel
      ExcelApp := Unassigned; //釋放ExcelApp;
    end;
  end
  else
    ShowMessage('請選擇儲存位置');
end;

procedure TfmMain.btnExcelClick(Sender: TObject);
var
  str, str2, strTitle, strTitle2, strtemp,schStr: string;
  SQLtemp, SListA: TStringList;
  i, x, y, z, SubCount, QueCounut, RowsCount, ColCount: integer;
  sch :integer ;
  ExcelApp: Variant;
  strFormat: OleVariant;
begin
  try
    ExcelApp := CreateOleObject('Excel.Application'); //Create Excel 物件
    ExcelApp.WorkBooks.Add; //新增工作簿(預設為三個工作表)
  except
    ShowMessage('尚未安裝任何 Excel 版本.');
    Exit;
  end;
  gTotal.MaxValue := mmoCheck.Lines.Count;
  SQLtemp := TStringList.Create;
  SListA := TStringList.Create;

  {$region '插入標題'}
  strTitle := '學校代碼' + #9 + '學校名稱' + #9 + '分校/分班註記' + #9 + '測驗科目代碼' + #9 + '年級' + #9 + '班級代碼' + #9 + '班級名稱' + #9 + '座號' + #9 + '學生姓名' + #9 + '性別代碼' + #9 + '導師姓名' + #9 + '身分註記' + #9 + #9 + #9 + #9 + #9 + #9 + #9 + '學生原始作答反應' + #9;
  strTitle2 := #9 + #9 + #9 + #9 + #9 + #9 + #9 + #9 + #9 + #9 + #9 + '資賦優異' + #9 + '資源班' + #9 + '原住民子女' + #9'新住民子女' + #9 + '才藝班學生' + #9 + '體育班學生' + #9 + '非學校型態實驗教育者（在家教育）' + #9;
  {$endregion}

  SaveDialog1.Filter := '.xlsx|*.xlsx';
  SaveDialog1.FileName := '預設名稱';
  if SaveDialog1.Execute then
  begin
    try
      try
        for sch := 0 to mmoCheck.Lines.Count - 1 do
        begin
          {$region '判斷科目總數'}
          SQLStr := 'SELECT es.Sub_No,Sub_Name FROM Exam_Sub es '
                  + 'left join Ex_Subject ej on es.Sub_No = ej.Sub_No '
                  + 'where Exam_No =' + QuotedStr(cmbExam.text);
          OpenSQL(DM.qryTemp, SQLStr);
          SubCount := DM.qryTemp.RecordCount;
          {$endregion}

          {$region '儲存學校名稱'}
          schStr := '' ;
          schStr := Copy(mmoCheck.Lines[sch],pos('-',mmoCheck.Lines[sch])+1,Length(mmoCheck.Lines[sch]));
          {$endregion}

          //分類-科目
          for x := 1 to SubCount do
          begin
            g1.Progress := 0;

            {$region 'SQL科目語法'}
            SQLStr := ' SELECT left(st.[Address],6) AS 學校代碼, '
            //                  + '       s.Sch_Memo as 學校代碼, '
                    + '       s.Sch_Name as 學校名稱, '
                    + '       case len(s.Sch_Memo) when  6 then 0  when 7 then substring(s.Sch_Memo,7,7) else 0 end as 分校註記,'
                    + '       case Sub_No when 01 then ''C'' when 02 then ''M''  when 03 then ''E'' end AS 測驗科目代碼, '
                    + '       sc.grade as 年級, '
                    + '       sc.Sch_Code as 班級代碼, '//                  + '       B_Year as 班級代碼, '
                    + '       (sc.grade +''年''+rtrim(sc.Class_Name)+''班'') as 班級名稱, '
                    //                  + '       (sc.grade +''年''+rtrim(B_Year)+''班'') as 班級名稱, '
                    + '       st.Seat_No as 座號, '
                    + '       Student_Name as 學生姓名, '
                    + '       Sex as 性別代碼, '
                    + '       Remark1 as 導師姓名,'
                    + '       Remark as 身分註記,'
                    //                  + '       Remark1 as 身分註記,'
                    + '       ans as 原始作答反映 '
                    + ' FROM Sub_Score ss ' + ' left join School  s on left([Student_No],3) = s.Sch_Code '
                    + ' left join Student st on ss.Student_No = st.Student_No '
                    + ' left join Sch_Class sc on s.Sch_Code = sc.Sch_Code and sc.Class_No = st.Class_No '
                    + ' where Sub_No = ' + DM.qryTemp.FieldByName('Sub_No').AsString + ' and  Exam_No = ' + QuotedStr(cmbExam.Text) + ' and s.Sch_Name = ' + QuotedStr(schStr)
  //                  + ' where Sub_No = '+DM.qryTemp.FieldByName('Sub_No').AsString+' and left(ss.[Student_No],3) = '+ Copy(mmoCheck.Lines[i],0,3) + ' and  Exam_No = ' +QuotedStr(cmbExam.Text)
                    + ' order by s.Sch_Code,Sub_No,sc.grade,st.Seat_No';
            OpenSQL(DM.qrySearch, SQLStr);

            {$endregion}

            g1.MaxValue := DM.qrySearch.RecordCount + 7;
            g1.Progress := g1.Progress + 1;

            {$region 'SQL總題數語法'}
            SQLStr := 'Select Count(*) as iCount From Sub_Ans' + ' Where Exam_No=' + #39 + Trim(cmbExam.Text) + #39 + ' And Sub_No=' + #39 + DM.qryTemp.FieldByName('Sub_No').AsString + #39;
            OpenSQL(DM.qrySClass, SQLStr);
            QueCounut := DM.qrySClass.FieldByName('iCount').AsInteger;
            {$endregion}

            g1.Progress := g1.Progress + 1;

            {$region '儲存Excel Title'}
            strTitle2 := '';
            strTitle2 := #9 + #9 + #9 + #9 + #9 + #9 + #9 + #9 + #9 + #9 + #9 + '資賦優異' + #9 + '資源班' + #9 + '原住民子女' + #9'新住民子女' + #9 + '才藝班學生' + #9 + '體育班學生' + #9 + '非學校型態實驗教育者（在家教育）' + #9;
            for y := 1 to QueCounut do
            begin
              strTitle2 := strTitle2 + '第' + IntToStr(y) + '題' + #9;
            end;

            SListA.Append(strTitle);
            SListA.Append(strTitle2);
            {$endregion}

            g1.Progress := g1.Progress + 1;
            lblNow.Caption := '設定標題';

            {$region 'Excel初始化設定'}
            g1.Progress := g1.Progress + 1;
            if ( x = 1)  then      //   and (sch <> 0)
            begin
              ExcelApp.WorkBooks.Add; //新增工作簿(預設為三個工作表)
              ExcelApp.DisplayAlerts := False;  //關閉提示視窗
              ExcelApp.Visible := False; //不顯示Excel 視窗
            end;
            if ExcelApp.WorkSheets.Count < 3 then
              ExcelApp.WorkBooks[1].WorkSheets.Add(null, null, 2, xlWorksheet);
            ExcelApp.WorkSheets[x].Activate;
            ExcelApp.WorkSheets[x].Name := '科目名稱_' + DM.qryTemp.FieldByName('Sub_Name').AsString; //工作表更名
            strFormat := '@'; //@: 儲存格格式改為文字
            ExcelApp.WorkSheets[x].Cells.NumberFormatLocal := strFormat; //設定儲格格式(一定要宣告OleVariant，直接等於'@'無
            {$endregion}

            g1.Progress := g1.Progress + 1;
            lblNow.Caption := '初始化Excel設定';

            //分類-考生
            for y := 0 to DM.qrySearch.RecordCount - 1 do
            begin
              g1.Progress := g1.Progress + 1;
              lblNow.Caption := '讀取考生資料';
              {$region '切割身分註記'}
              SQLtemp.Clear;
              str := '';
              SQLtemp.Add(DM.qrySearch.FieldByName('身分註記').AsString);
              SQLtemp.Delimiter := '-';
              SQLtemp.DelimitedText := SQLtemp.Text;
              for z := 0 to SQLtemp.Count - 1 do
              begin
                str := str + SQLtemp[z] + #9;
              end;
              {$endregion}

              {$region '切割原始作答'}
              str2 := '';
              SQLtemp.Clear;
              SQLtemp.Add(DM.qrySearch.FieldByName('原始作答反映').AsString);
              SQLtemp.Delimiter := ',';
              SQLtemp.DelimitedText := SQLtemp.Text;
              for z := 0 to SQLtemp.Count - 1 do
              begin
                if Length(SQLtemp[z]) > 1 then
                  str2 := str2 + '?' + #9
                else
                  str2 := str2 + SQLtemp[z] + #9;
              end;
              {$endregion}

              {$region '儲存資料至剪貼簿'}
              SListA.Add(DM.qrySearch.FieldByName('學校代碼').AsString + #9 + DM.qrySearch.FieldByName('學校名稱').AsString + #9 + DM.qrySearch.FieldByName('分校註記').AsString + #9 + DM.qrySearch.FieldByName('測驗科目代碼').AsString + #9 + DM.qrySearch.FieldByName('年級').AsString + #9 + Trim(DM.qrySearch.FieldByName('班級代碼').AsString) + #9 + Trim(DM.qrySearch.FieldByName('班級名稱').AsString) + #9 + DM.qrySearch.FieldByName('座號').AsString + #9 + DM.qrySearch.FieldByName('學生姓名').AsString + #9 + DM.qrySearch.FieldByName('性別代碼').AsString + #9 + DM.qrySearch.FieldByName('導師姓名').AsString + #9 + str + #9    //身分註記
              + str2);   //原始作答
              {$endregion}

              DM.qrySearch.Next;
            end;
            Clipboard.Clear; //先清空剪貼簿
            Clipboard.AsText := SListA.Text; //複製資料到剪貼簿

            {$region '設定Excel 格式'}
            ExcelApp.Range['A1'].Select;
            ExcelApp.Range['A1'].PasteSpecial; //在A1貼上
            RowsCount :=  ExcelApp.ActiveSheet.UsedRange.Rows.Count;
            ColCount :=   ExcelApp.ActiveSheet.UsedRange.Columns.Count;
            g1.Progress := g1.Progress + 1;
            lblNow.Caption := '設定Excel格式';

            //設定背景色
            ExcelApp.Selection.FormatConditions.Add(xlExpression,Formula1:='=MOD(ROW(),2)=0');
            ExcelApp.Selection.FormatConditions[1].Interior.Color := RGB(235, 241, 222);
            ExcelApp.Selection.FormatConditions.Add(xlExpression,Formula1:='=MOD(ROW(),2)=1');
            ExcelApp.Selection.FormatConditions[2].Interior.Color := RGB(242, 220, 219);
            ExcelApp.Rows['1:2'].FormatConditions.Delete;
            ExcelApp.Rows[1].Interior.Color := RGB(153, 204, 255);
            ExcelApp.Range['S1'].Interior.Color := RGB(255, 255, 0);
            ExcelApp.Range['L2:R2'].Interior.Color := RGB(197, 217, 241);
            ExcelApp.Range[ExcelApp.cells[2, 19], ExcelApp.cells[2, QueCounut + 18]].Interior.Color := RGB(255, 255, 204);
            //合併儲存格
            ExcelApp.ActiveSheet.Range[ExcelApp.cells[1, 19], ExcelApp.cells[1, QueCounut + 18]].Merge;
            ExcelApp.ActiveSheet.Range['L1:R1'].Merge;
            for y := 1 to 11 do
              ExcelApp.ActiveSheet.Range[ExcelApp.cells[1, y], ExcelApp.cells[2, y]].Merge;
            //框線
  //          ExcelApp.rows['3:'+inttostr(RowsCount)].Borders[2].LineStyle := 1;
  //          ExcelApp.rows['3:'+inttostr(RowsCount)].Borders[2].Weight := 2;
            ExcelApp.rows['3:'+inttostr(RowsCount)].Borders[4].LineStyle := 1;
            ExcelApp.rows['3:'+inttostr(RowsCount)].Borders[4].Weight := 2;
            //置中
            ExcelApp.Cells.HorizontalAlignment := -4108;
            //凍結欄位
            ExcelApp.ActiveWindow.SplitColumn := 0;
            ExcelApp.ActiveWindow.SplitRow  := 2;
            ExcelApp.ActiveWindow.FreezePanes := true;
            //自適欄寬
            ExcelApp.Cells.EntireColumn.AutoFit;
            ExcelApp.Cells.Font.Name := '微軟正黑體';
            {$endregion}

            SListA.clear;
            g1.Progress := g1.MaxValue;
            lblNow.Caption := '完成';
            DM.qryTemp.next; //切換下一筆科目
          end;
          gTotal.Progress := gTotal.Progress + 1;
          ExcelApp.ActiveWorkBook.Saved := True; //設定不存檔，若不設定關閉時會出現"是否存檔的對話框"
          ExcelApp.ActiveWorkBook.SaveAs(ExtractFilePath(SaveDialog1.FileName)+schStr + '_作答反應.xlsx'); //存檔
          ExcelApp.WorkBooks.close; //關閉Excel
        end;
        gTotal.Progress := gTotal.MaxValue;
        ShowMessage('匯出完成');
      except
        ShowMessage('匯出錯誤 - '+lblNow.Caption);
      end;
    finally
      SListA.Free;
      SQLtemp.Free;
      ExcelApp.Quit; //離開Excel
      ExcelApp := Unassigned; //釋放ExcelApp;
    end;
  end
  else
    ShowMessage('請選擇儲存位置');
end;

procedure TfmMain.btnExitClick(Sender: TObject);
begin
  Close;
end;

procedure TfmMain.TransStd(bRange: Boolean);  //學生資料轉檔
var
  SListA, SListB: TStringList;
  Str1, sTemp, sTemp1, sClass: string;
  ii, iSub, iCCount, iMod: Integer;
  ExcelApp: Variant;
  StrFormat: OleVariant;
begin
  SListA := TStringList.Create;
  SListB := TStringList.Create;

  iSub := POS('-', cbSubT.Text);

  SQLStr := 'Select a.Sub_No, b.Sub_Name' + '  From Exam_Sub a' + ' Inner Join Ex_Subject b on a.Sub_No=b.Sub_No' + ' Where a.Exam_No=' + #39 + Trim(cmbExam.Text) + #39;
  if bRange then   //部分
  begin
    SQLStr := SQLStr + '   And a.Sub_No=' + #39 + Copy(cbSubT.Text, 1, 2) + #39;
  end;
  SQLStr := SQLStr + ' Order by a.Sub_No;';
  OpenSQL(DM.qryTemp, SQLStr);

  SListA.Clear;

  while not DM.qryTemp.Eof do
  begin
    {
    SQLStr := 'Select REPLICATE(''0'', (5-LEN(CONVERT(Varchar(5),ROW_NUMBER() OVER(Order by Student_No)))))+ CONVERT(Varchar(5),ROW_NUMBER() OVER(Order by Student_No)), '
            +       #39+Trim(DM.qryTemp.FieldByName('Sub_Name').AsString)+#39+', X_No, Sch_Code, Grade, Class_No, Seat_No, Student_Name, Student_No, Group_No, '
            +       #39+DM.qryTemp.FieldByName('Sub_No').AsString+#39+', Student_No+Group_No+'#39+DM.qryTemp.FieldByName('Sub_No').AsString+#39+', '
            +       ' Remark1, Address, Remark, B_Year, ID '
            + '  From student '
            + ' Order by student_No;';
     }
    Str1 := '';
    sTemp := '';
    sTemp1 := '';
    for ii := 0 to CheckListBox2.Items.Count - 1 do
    begin
      if CheckListBox2.Checked[ii] then
      begin
        case ii of
          0:
            begin
              SListB.Append('REPLICATE(''0'', (5-LEN(CONVERT(Varchar(5),ROW_NUMBER() OVER(Order by a.Student_No)))))+ CONVERT(Varchar(5),ROW_NUMBER() OVER(Order by a.Student_No)) as iRow, ');
              sTemp := sTemp + '序號' + #9;
            end;
          1:
            begin
              SListB.Append(#39 + Trim(DM.qryTemp.FieldByName('Sub_Name').AsString) + #39 + ' as SubName,');
              sTemp := sTemp + '科目名稱' + #9;
            end;
          2:
            begin
              SListB.Append('a.X_No,');
              sTemp := sTemp + '准考證號' + #9;
            end;
          3:
            begin
              SListB.Append('a.Sch_Code,');
              sTemp := sTemp + '校代碼' + #9;
            end;
          4:
            begin
              SListB.Append('a.Grade,');
              sTemp := sTemp + '年級' + #9;
            end;
          5:
            begin
              SListB.Append('a.Class_No,');
              sTemp := sTemp + '班級' + #9;
            end;
          6:
            begin
              SListB.Append('a.Seat_No,');
              sTemp := sTemp + '座號' + #9;
            end;
          7:
            begin
              SListB.Append('a.Student_Name,');
              sTemp := sTemp + '姓名' + #9;
            end;
          8:
            begin
              SListB.Append('a.Student_No,');
              sTemp := sTemp + '電閱編號' + #9;
            end;
          9:
            begin
              SListB.Append('a.Group_No,');
              sTemp := sTemp + '類組' + #9;
            end;
          10:
            begin
              SListB.Append(#39 + DM.qryTemp.FieldByName('Sub_No').AsString + #39 + ' as Sub_No,');
              sTemp := sTemp + '科目代碼' + #9;
            end;
          11:
            begin
              SListB.Append('a.Student_No+a.Group_No+'#39 + DM.qryTemp.FieldByName('Sub_No').AsString + #39 + ' as BarCode, ');
              sTemp := sTemp + '條碼' + #9;
            end;
          12:
            begin
              SListB.Append('b.Sch_Name,');
              sTemp := sTemp + '備註1' + #9;
            end;
          13:
            begin
              SListB.Append('a.Address,');
              sTemp := sTemp + '住址/行政區' + #9;
            end;
          14:
            begin
              SListB.Append('a.Remark,');
              sTemp := sTemp + '備註' + #9;
            end;
          15:
            begin
              SListB.Append('c.Class_Name,');
              sTemp := sTemp + '出生年/原班級' + #9;
            end;
          16:
            begin
              SListB.Append('a.ID');
              sTemp := sTemp + '身分證號/原座號' + #9;
            end;
        end;
      end;
    end;

    Str1 := '';
    for ii := 0 to SListB.Count - 1 do
    begin
      Str1 := Str1 + SListB[ii];
    end;
    SListB.Clear;

    SQLStr := 'Select ' + Str1 + '  From Student a' + ' Inner Join School b on a.Sch_Code=b.Sch_Code';
    if CheckListBox2.Checked[15] then
    begin
      SQLStr := SQLStr + ' Inner Join Sch_Class c on a.Sch_Code=c.Sch_Code And a.Class_No=c.Class_No';
    end;

    if bRange then   //部分
    begin
      SQLStr := SQLStr + ' Where a.Student_No Between ' + #39 + Trim(Edit1.Text) + #39 + ' And ' + #39 + Trim(Edit2.Text) + #39;
    end;

    SQLStr := SQLStr + ' Order by a.Student_No;';
    OpenSQL(DM.qrySearch, SQLStr);
    SListA.Add(sTemp);

    sClass := '';
    iCCount := 0;
    while not DM.qrySearch.Eof do
    begin
      sClass := DM.qrySearch.FieldByName('Class_No').AsString;
      for ii := 0 to CheckListBox2.Items.Count - 1 do
      begin
        if CheckListBox2.Checked[ii] then
        begin
          case ii of
            0:
              begin
                sTemp1 := sTemp1 + DM.qrySearch.FieldByName('iRow').AsString + #9;
              end;
            1:
              begin
                sTemp1 := sTemp1 + DM.qrySearch.FieldByName('SubName').AsString + #9;
              end;
            2:
              begin
                sTemp1 := sTemp1 + DM.qrySearch.FieldByName('X_No').AsString + #9;
              end;
            3:
              begin
                sTemp1 := sTemp1 + DM.qrySearch.FieldByName('Sch_Code').AsString + #9;
              end;
            4:
              begin
                sTemp1 := sTemp1 + DM.qrySearch.FieldByName('Grade').AsString + #9;
              end;
            5:
              begin
                sTemp1 := sTemp1 + DM.qrySearch.FieldByName('Class_No').AsString + #9;
              end;
            6:
              begin
                sTemp1 := sTemp1 + DM.qrySearch.FieldByName('Seat_No').AsString + #9;
              end;
            7:
              begin
                sTemp1 := sTemp1 + DM.qrySearch.FieldByName('Student_Name').AsString + #9;
              end;
            8:
              begin
                sTemp1 := sTemp1 + DM.qrySearch.FieldByName('Student_No').AsString + #9;
              end;
            9:
              begin
                sTemp1 := sTemp1 + DM.qrySearch.FieldByName('Group_No').AsString + #9;
              end;
            10:
              begin
                sTemp1 := sTemp1 + DM.qrySearch.FieldByName('Sub_No').AsString + #9;
              end;
            11:
              begin
                sTemp1 := sTemp1 + DM.qrySearch.FieldByName('BarCode').AsString + #9;
              end;
            12:
              begin
                sTemp1 := sTemp1 + DM.qrySearch.FieldByName('Sch_Name').AsString + #9;
              end;
            13:
              begin
                sTemp1 := sTemp1 + DM.qrySearch.FieldByName('Address').AsString + #9;
              end;
            14:
              begin
                sTemp1 := sTemp1 + DM.qrySearch.FieldByName('Remark').AsString + #9;
              end;
            15:
              begin
                sTemp1 := sTemp1 + Trim(DM.qrySearch.FieldByName('Class_Name').AsString) + #9;
              end;
            16:
              begin
                sTemp1 := sTemp1 + DM.qrySearch.FieldByName('ID').AsString + #9;
              end;
          end;
        end;
      end;

      SListA.Add(sTemp1);
      sTemp1 := '';
      Inc(iCCount);

      DM.qrySearch.Next;

      if CheckBox3.Checked then
      begin
        if sClass <> DM.qrySearch.FieldByName('Class_No').AsString then
        begin
          if iCCount < StrToInt(Trim(edtTNo.Text)) then
          begin
            for ii := 1 to StrToInt(Trim(edtTNo.Text)) - iCCount do
              SListA.Add('　' + #9);
          end
          else if iCCount > StrToInt(Trim(edtTNo.Text)) then
          begin
            iMod := iCCount mod StrToInt(Trim(edtTNo.Text));
            for ii := 1 to StrToInt(Trim(edtTNo.Text)) - iMod do
              SListA.Add('　' + #9);
          end;
          sClass := DM.qrySearch.FieldByName('Class_No').AsString;
          iCCount := 0;
        end;
      end;
    end;
    DM.qryTemp.Next;

  end;

  if bRange then
  begin
    SaveDialog1.FileName := '學生答案卡檔案_' + FormatDateTime('YYYYMMDD', Now()) + '.xls';
  end
  else
  begin
    SaveDialog1.FileName := '學生答案卡檔案.xls';
  end;

  try
    if SaveDialog1.Execute then
    begin
      try
        ExcelApp := CreateOleObject('Excel.Application'); //Create Excel 物件
      except
        ShowMessage('尚未安裝任何 Excel 版本.');
        SListA.Free;
        Exit;
      end;

      if FileExists(SaveDialog1.FileName) then
        DeleteFile(SaveDialog1.FileName);

      ExcelApp.WorkBooks.Add; //新增工作簿(預設為三個工作表)
      ExcelApp.Visible := False; //不顯示Excel 視窗
      ExcelApp.WorkSheets[1].Activate;
      ExcelApp.WorkSheets[1].Name := 'Sheet1'; //工作表更名
      StrFormat := '@'; //@: 儲存格格式改為文字
      ExcelApp.WorkSheets[1].Cells.NumberFormatLocal := StrFormat; //設定儲格格式(一定要宣告OleVariant，直接等於'@'無
      Clipboard.Clear; //先清空剪貼簿
      Clipboard.AsText := SListA.Text; //複製資料到剪貼簿
      ExcelApp.Range['A1'].Select;
      ExcelApp.Range['A1'].PasteSpecial; //在A1貼上
      Clipboard.Clear; //用完清空剪貼簿
       //ExcelApp.Range['A1'].Select;
      ExcelApp.ActiveWorkBook.Saved := True; //設定不存檔，若不設定關閉時會出現"是否存檔的對話框"
      ExcelApp.WorkBooks[1].SaveAs(SaveDialog1.FileName); //存檔
      ExcelApp.WorkBooks.close; //關閉Excel
      ExcelApp.Quit; //離開Excel
      ExcelApp := Unassigned; //釋放ExcelApp;
    end;

    showmessage('轉檔完畢--' + SaveDialog1.FileName);
  except
    showmessage('轉檔失敗!');
  end;
   //Address --- 行政區
   //Remark  --- 群組
   //B_Year  --- 班級
   //ID      --- 座號

  SListA.Clear;
  SListA.Free;
  SListB.Clear;
  SListB.Free;

end;

procedure TfmMain.BitBtn1Click(Sender: TObject);
begin
  TransStd(False);   //全部

  //序,科目,准考證號,校代碼,年級,班級1,座號1,姓名,條碼文字,條碼文字1(類組),條碼文字2(科目),條碼,學校名稱,行政區,群組,班級,座號,卷別,
  //序,電閱號碼,級別,組別,准考證號,校代碼,學校簡稱,科目,條碼,條碼文字,


end;

procedure TfmMain.BitBtn2Click(Sender: TObject);
begin
  TransStd(True);   //部份
end;

procedure TfmMain.btnAbsentClick(Sender: TObject);
var
  ii, iSLen: Integer;
  sStd: string;
begin
  case ClassGroup.ItemIndex of
    0:
      iSLen := 12;
    1:
      iSLen := 9;
    2:
      iSLen := 12;
    3:
      iSLen := 9;
  end;

  for ii := 1 to Memo2.Lines.Count - 1 do
  begin
    sStd := Copy(Memo2.Lines.Strings[ii], 1, iSLen); //電閱號碼
    if (sStd <> '') then
    begin
      case classGroup.ItemIndex of
        0:
          SQLStr := 'Update [z' + Trim(cmbExam.Text) + Copy(sStd, 1, 6) + ']';
        1..3:
          SQLStr := 'Update [' + Trim(cmbExam.Text) + Copy(sStd, 1, 3) + ']';
      end;
      SQLStr := SQLStr + '  Set Absent=''1'', Absent1=''1'', Absent2=''1'' ';
      case classGroup.ItemIndex of
        0:
          SQLStr := SQLStr + ' Where Stud_No=' + #39 + sStd + #39;
        1..3:
          SQLStr := SQLStr + ' Where Student_No=' + #39 + sStd + #39;
      end;
      SQLStr := SQLStr + '  And Sub_No=' + #39 + Trim(cbSub.Text) + #39;
      SQLExec(DM.qryTemp, SQLStr);
    end;
  end;
  showmessage('全部已設定為缺考！');
  btnAbsent.Enabled := False;
end;

procedure TfmMain.btnClsInClick(Sender: TObject);
var
  Str1, sTemp: string;
  xlsFileName: string;
  ii, i: integer;
  EmptyKey: TStringList;
begin
   //匯入
  if TableExists('Sch_Class') then
    Str1 := 'Sch_Class'
  else
    Str1 := '';

  SQLStr := 'Select * From Sch_Class';
  OpenSQL(DM.qryTemp, SQLStr);

  if not DM.qryTemp.IsEmpty then
  begin
    sTemp := '資料庫裡頭有些學校班級資料, 是否先全部清除?';
    if application.Messagebox(PChar(sTemp), 'Message', MB_YesNo) = IDYES then
    begin
      SQLStr := 'Delete From Sch_Class';
      SQLExec(Dm.qryTemp, SQLStr);
    end;
  end;

  if Str1 <> '' then
  begin
     //1. 先產生年級
    SQLStr := 'Select count(name) as iCount From SysColumns ' + ' Where id=(Select id From Sysobjects ' + ' Where name=''Sch_Class'' ' + '   And name=''Grade'')';
    OpenSQL(DM.qryTemp, SQLStr);

    if DM.qryTemp.IsEmpty then  // 沒有這個欄位
    begin
      SQLStr := 'Alter table Sch_Class add [Grade] nchar(1) Not NULL DEFAULT ''3'';';
      SQLExec(DM.qryTemp, SQLStr);

      SQLStr := 'Update Sch_Class Set [Grade]=''3'';';
      SQLExec(DM.qryTemp, SQLStr);

      SQLStr := 'ALTER TABLE Sch_Class DROP CONSTRAINT PK_Sch_Class;';
      SQLExec(DM.qryTemp, SQLStr);

      SQLStr := 'ALTER TABLE Sch_Class ADD CONSTRAINT PK_Sch_Class PRIMARY KEY (Sch_Code, Grade, Class_No); ';
      SQLExec(DM.qryTemp, SQLStr);
    end;
     //2. 進行匯入
    EmptyKey := TStringList.Create;
    OpenDialog1.DefaultExt := '*.xls';
    OpenDialog1.InitialDir := ExtractFilePath(Application.ExeName);
    OpenDialog1.Filter := 'Excel(*.xls)|*.xls;Text(*.csv)|*.csv;Text(*.txt)|*.txt';
    OpenDialog1.FileName := '';
    OpenDialog1.Title := '選擇檔案!';
    if OpenDialog1.Execute then
    begin
      xlsFileName := OpenDialog1.FileName;
      if Conn1.Connected then
        Conn1.Close;
      Conn1.ConnectionString := 'Provider=Microsoft.ACE.OLEDB.12.0;Data Source=' + xlsFileName + ';' + 'Extended Properties="Excel 12.0;HDR=No;IMEX=1";'; //Excel 8.0  //Text
      Conn1.Open;
      adods1.Close;
      adods1.CommandText := '';
      adods1.CommandText := 'Select * from [Sheet1$] '; //Info.Exam_No+
      adods1.Open;
      adods1.Last;
      adods1.First;
      xlsFileName := Unassigned;
      while not adods1.Eof do
      begin
        if (adods1.FieldByName('F1').AsString <> '學校代碼') or ((adods1.FieldByName('F1').AsString >= '000') and (adods1.FieldByName('F1').AsString <= '999')) then
        begin
          if not (adods1.FieldByName('F1').IsNull) then //
          begin
            SQLStr := ' Insert Into Sch_Class (Sch_Code, Grade, Class_No, Class_Name)' + ' Values(' + #39 + Trim(adods1.FieldByName('F1').AsString) + #39 + ',' + #39 + Trim(adods1.FieldByName('F2').AsString) + #39 + ',' + #39 + Trim(adods1.FieldByName('F3').AsString) + #39 + ',' + #39 + Trim(adods1.FieldByName('F4').AsString) + #39 + ')';
            try
              SQLExec(Dm.qryTemp, SQLStr);
            except
              ShowMessage(SQLStr);
            end;
          end
          else
            EmptyKey.Add(IntToStr(adods1.RecNo) + 'th Row:' + #9 + adods1.FieldByName('F1').AsString + #9#44 + adods1.FieldByName('F3').AsString + #9#44 + adods1.FieldByName('F4').AsString + #9 + ' is Null');
        end;
        adods1.Next;
      end;

      Showmessage('匯入完成!!');
      if EmptyKey.Count > 0 then
      begin
        ShowMessage(EmptyKey.Text);
        EmptyKey.SaveToFile(ExtractFilePath(Application.ExeName) + 'EmptyKey.txt');
      end;
    end;
    EmptyKey.Clear;
    EmptyKey.Free;

    SQLStr := 'Select * From Sch_Class' + ' Order by Sch_Code, Class_No;';
    OpenSQL(DM.qrySClass, SQLStr);
    DBGrid2.DataSource := DM.dsSClass;
    DBGrid2.Columns[0].FieldName := 'Sch_Code';
    DBGrid2.Columns[1].FieldName := 'Grade';
    DBGrid2.Columns[2].FieldName := 'Class_No';
    DBGrid2.Columns[3].FieldName := 'Class_Name';
    PageControl1.ActivePageIndex := 1;
    PageControl1.OnChange(Self);

  end
  else
  begin
    showmessage('找不到學校班級資料表!');
  end;
end;

procedure TfmMain.btnClsOutClick(Sender: TObject);
var
  SListA: TStringList;
  Str1, Str2: string;
  ExcelApp: Variant;
  StrFormat: OleVariant;
begin
   //匯出
  if TableExists('Sch_Base') then
    Str1 := 'Sch_Base'
  else if TableExists('School') then
    Str1 := 'School'
  else
    Str1 := '';

  if Str1 <> '' then
  begin
     //1. 先產生年級
    SQLStr := 'Select count(name) as iCount From SysColumns ' + ' Where id=(Select id From Sysobjects ' + ' Where name=''Sch_Class'' ' + '   And name=''Grade'')';
    OpenSQL(DM.qryTemp, SQLStr);
    if not DM.qryTemp.IsEmpty then  // 沒有這個欄位
    begin
      SQLStr := 'Alter table Sch_Class add [Grade] nchar(1) Not NULL DEFAULT ''3'';';
      SQLExec(DM.qryTemp, SQLStr);

      SQLStr := 'Update Sch_Class Set [Grade]=''3'';';
      SQLExec(DM.qryTemp, SQLStr);

      SQLStr := 'ALTER TABLE Sch_Class DROP CONSTRAINT PK_Sch_Class;';
      SQLExec(DM.qryTemp, SQLStr);

      SQLStr := 'ALTER TABLE Sch_Class ADD CONSTRAINT PK_Sch_Class PRIMARY KEY (Sch_Code, Grade, Class_No); ';
      SQLExec(DM.qryTemp, SQLStr);
    end;

    SListA := TStringList.Create;

    SQLStr := 'Select * ' + '  From Sch_Class ' + ' Order by Sch_Code, Class_No;';
    OpenSQL(DM.qryTemp, SQLStr);

    SListA.Append('學校代碼' + #9 + '年級' + #9 + '班級代碼' + #9 + '班級名稱' + #9);
    while not DM.qryTemp.Eof do
    begin
      Str2 := DM.qryTemp.FieldByName('Sch_Code').AsString + #9 + DM.qryTemp.FieldByName('Grade').AsString + #9 + DM.qryTemp.FieldByName('Class_No').AsString + #9 + DM.qryTemp.FieldByName('Class_Name').AsString + #9;

      SListA.Append(Str2);
      Str2 := '';
      DM.qryTemp.Next;
    end;

    SaveDialog1.FileName := '學校班級資料檔.xls';
    if SaveDialog1.Execute then
    begin
      try
        ExcelApp := CreateOleObject('Excel.Application'); //Create Excel 物件
      except
        ShowMessage('尚未安裝任何 Excel 版本.');
        SListA.Free;
        Exit;
      end;

      if FileExists(SaveDialog1.FileName) then
        DeleteFile(SaveDialog1.FileName);

      ExcelApp.WorkBooks.Add; //新增工作簿(預設為三個工作表)
      ExcelApp.Visible := False; //不顯示Excel 視窗
      ExcelApp.WorkSheets[1].Activate;
      ExcelApp.WorkSheets[1].Name := 'Sheet1'; //工作表更名
      StrFormat := '@'; //@: 儲存格格式改為文字
      ExcelApp.WorkSheets[1].Cells.NumberFormatLocal := StrFormat; //設定儲格格式(一定要宣告OleVariant，直接等於'@'無
      Clipboard.Clear; //先清空剪貼簿
      Clipboard.AsText := SListA.Text; //複製資料到剪貼簿
      ExcelApp.Range['A1'].Select;
      ExcelApp.Range['A1'].PasteSpecial; //在A1貼上
      Clipboard.Clear; //用完清空剪貼簿
        //ExcelApp.Range['A1'].Select;
      ExcelApp.ActiveWorkBook.Saved := True; //設定不存檔，若不設定關閉時會出現"是否存檔的對話框"
      ExcelApp.WorkBooks[1].SaveAs(SaveDialog1.FileName); //存檔
      ExcelApp.WorkBooks.close; //關閉Excel
      ExcelApp.Quit; //離開Excel
      ExcelApp := Unassigned; //釋放ExcelApp;
    end;
    SListA.Clear;
    SListA.Free;
    ShowMessage('轉出學校班級基本資料檔完成!');
  end
  else
  begin
    showmessage('此資料庫沒有學校班級基本資料檔!');
  end;

end;

procedure TfmMain.btnSchInClick(Sender: TObject);
var
  xlsFileName, Str1, sTemp: string;
  ii, i: integer;
  EmptyKey: TStringList;
begin
  if TableExists('Sch_Base') then
    Str1 := 'Sch_Base'
  else if TableExists('School') then
    Str1 := 'School'
  else
    Str1 := '';

  SQLStr := 'Select * From ' + Str1;
  OpenSQL(DM.qryTemp, SQLStr);

  if not DM.qryTemp.IsEmpty then
  begin
    sTemp := '資料庫裡頭有些學校基本資料, 是否先全部清除?';
    if application.Messagebox(PChar(sTemp), 'Message', MB_YesNo) = IDYES then
    begin
      SQLStr := 'Delete From ' + Str1;
      SQLExec(Dm.qryTemp, SQLStr);
    end;
  end;

  if Str1 <> '' then
  begin
    EmptyKey := TStringList.Create;
    OpenDialog1.DefaultExt := '*.xls';
    OpenDialog1.InitialDir := ExtractFilePath(Application.ExeName);
    OpenDialog1.Filter := 'Excel(*.xls)|*.xls;Text(*.csv)|*.csv;Text(*.txt)|*.txt';
    OpenDialog1.FileName := '';
    OpenDialog1.Title := '選擇檔案!';
    if OpenDialog1.Execute then
    begin
      xlsFileName := OpenDialog1.FileName;
      if Conn1.Connected then
        Conn1.Close;
      Conn1.ConnectionString := 'Provider=Microsoft.ACE.OLEDB.12.0;Data Source=' + xlsFileName + ';' + 'Extended Properties="Excel 12.0;HDR=No;IMEX=1";'; //Excel 8.0  //Text
      Conn1.Open;
      adods1.Close;
      adods1.CommandText := '';
      adods1.CommandText := 'Select * from [Sheet1$] '; //Info.Exam_No+
      adods1.Open;
      adods1.Last;
      adods1.First;
      xlsFileName := Unassigned;
      while not adods1.Eof do
      begin
        if (adods1.FieldByName('F1').AsString <> '學校代碼') or ((adods1.FieldByName('F1').AsString >= '000') and (adods1.FieldByName('F1').AsString <= '999')) then
        begin
          if not (adods1.FieldByName('F1').IsNull) then //
          begin
            SQLStr := ' Insert Into ' + Str1 + ' (Sch_Code, Sch_Name, Sch_Add,';
            if Str1 = 'Sch_Base' then
              SQLStr := SQLStr + 'Sch_Area, Sch_Memo) '
            else
              SQLStr := SQLStr + 'Sch_ACode, Sch_Memo) ';
            SQLStr := SQLStr + ' Values(' + #39 + Trim(adods1.FieldByName('F1').AsString) + #39 + ',' + #39 + Trim(adods1.FieldByName('F2').AsString) + #39 + ',' + #39 + Trim(adods1.FieldByName('F3').AsString) + #39 + ',' + #39 + Trim(adods1.FieldByName('F4').AsString) + #39 + ',' + #39 + Trim(adods1.FieldByName('F5').AsString) + #39 + ');';
            try
              SQLExec(Dm.qryTemp, SQLStr);
            except
              ShowMessage(SQLStr);
            end;
          end
          else
            EmptyKey.Add(IntToStr(adods1.RecNo) + 'th Row:' + #9 + adods1.FieldByName('F1').AsString + #9#44 + adods1.FieldByName('F3').AsString + #9#44 + adods1.FieldByName('F4').AsString + #9 + ' is Null');
        end;
        adods1.Next;
      end;

      Showmessage('匯入完成!!');
      if EmptyKey.Count > 0 then
      begin
        ShowMessage(EmptyKey.Text);
        EmptyKey.SaveToFile(ExtractFilePath(Application.ExeName) + 'EmptyKey.txt');
      end;
    end;
    EmptyKey.Clear;
    EmptyKey.Free;

    SQLStr := 'Select * From ' + Str1 + ' Order by Sch_Code;';
    OpenSQL(DM.qrySch, SQLStr);
    DBGrid1.DataSource := DM.dsSch;
    DBGrid1.Columns[0].FieldName := 'Sch_Code';
    DBGrid1.Columns[1].FieldName := 'Sch_Name';
    DBGrid1.Columns[2].FieldName := 'Sch_Add';
    if Str1 = 'Sch_Base' then
      DBGrid1.Columns[3].FieldName := 'Sch_Area'
    else
      DBGrid1.Columns[3].FieldName := 'Sch_ACode';
    DBGrid1.Columns[4].FieldName := 'Sch_Memo';

    PageControl1.ActivePageIndex := 0;
    PageControl1.OnChange(Self);

  end
  else
  begin
    showmessage('資料庫中找不到學校基本資料表!');
  end;
end;

procedure TfmMain.btnSchOutClick(Sender: TObject);
var
  SListA: TStringList;
  Str1, Str2: string;
  ExcelApp: Variant;
  StrFormat: OleVariant;
begin
   //匯出
  if TableExists('Sch_Base') then
    Str1 := 'Sch_Base'
  else if TableExists('School') then
    Str1 := 'School'
  else
    Str1 := '';

  if Str1 <> '' then
  begin
    SListA := TStringList.Create;

    SQLStr := 'Select * ' + '  From ' + Str1 + ' Order by Sch_Code;';
    OpenSQL(DM.qryTemp, SQLStr);

    SListA.Append('學校代碼' + #9 + '學校簡稱' + #9 + '區域' + #9 + '路線代碼' + #9 + '校代碼(6)' + #9);
    while not DM.qryTemp.Eof do
    begin
      Str2 := DM.qryTemp.FieldByName('Sch_Code').AsString + #9 + DM.qryTemp.FieldByName('Sch_Name').AsString + #9 + DM.qryTemp.FieldByName('Sch_Add').AsString + #9;
      if Str1 = 'Sch_Base' then
        Str2 := Str2 + DM.qryTemp.FieldByName('Sch_Area').AsString + #9
      else
        Str2 := Str2 + DM.qryTemp.FieldByName('Sch_ACode').AsString + #9;
      Str2 := Str2 + DM.qryTemp.FieldByName('Sch_Memo').AsString + #9;
      SListA.Append(Str2);
      Str2 := '';
      DM.qryTemp.Next;
    end;

    SaveDialog1.FileName := '學校基本資料檔.xls';
    if SaveDialog1.Execute then
    begin
      try
        ExcelApp := CreateOleObject('Excel.Application'); //Create Excel 物件
      except
        ShowMessage('尚未安裝任何 Excel 版本.');
        SListA.Free;
        Exit;
      end;

      if FileExists(SaveDialog1.FileName) then
        DeleteFile(SaveDialog1.FileName);

      ExcelApp.WorkBooks.Add; //新增工作簿(預設為三個工作表)
      ExcelApp.Visible := False; //不顯示Excel 視窗
      ExcelApp.WorkSheets[1].Activate;
      ExcelApp.WorkSheets[1].Name := 'Sheet1'; //工作表更名
      StrFormat := '@'; //@: 儲存格格式改為文字
      ExcelApp.WorkSheets[1].Cells.NumberFormatLocal := StrFormat; //設定儲格格式(一定要宣告OleVariant，直接等於'@'無
      Clipboard.Clear; //先清空剪貼簿
      Clipboard.AsText := SListA.Text; //複製資料到剪貼簿
      ExcelApp.Range['A1'].Select;
      ExcelApp.Range['A1'].PasteSpecial; //在A1貼上
      Clipboard.Clear; //用完清空剪貼簿
          //ExcelApp.Range['A1'].Select;
      ExcelApp.ActiveWorkBook.Saved := True; //設定不存檔，若不設定關閉時會出現"是否存檔的對話框"
      ExcelApp.WorkBooks[1].SaveAs(SaveDialog1.FileName); //存檔
      ExcelApp.WorkBooks.close; //關閉Excel
      ExcelApp.Quit; //離開Excel
      ExcelApp := Unassigned; //釋放ExcelApp;
    end;
    SListA.Clear;
    SListA.Free;
    ShowMessage('轉出學校基本資料檔完成!');
  end
  else
  begin
    showmessage('此資料庫沒有學校基本資料檔!');
  end;
end;

procedure TfmMain.btnSearchClick(Sender: TObject);
var
  Str1, sAnsTemp: string;
  iQCount, ii, jj: Integer;
begin
  iQCount := 0;
  ProgressBar1.Min := 0;

  SQLStr := 'Select Count(*) as iCount From Sub_Ans' + ' Where Exam_No=' + #39 + Trim(cmbExam.Text) + #39 + '   And Sub_No=' + #39 + Trim(cbSub.Text) + #39;
  OpenSQL(DM.qryTemp, SQLStr);
  iQCount := DM.qryTemp.FieldByName('iCount').AsInteger;   //題數
  sAnsTemp := '';
  for ii := 0 to iQCount - 1 do
    sAnsTemp := sAnsTemp + ',';   //空白答案的資料　例如： 英聽20題　",,,,,,,,,,,,,,,,,,,,"

  Memo2.Clear;
  if Trim(cbSub.Text) <> '' then
  begin
    Memo2.Lines.Add('電閱號碼、作答情形');
    for ii := 0 to CheckListBox3.Items.Count - 1 do
    begin
      if CheckListBox3.Checked[ii] then
      begin
        if Length(Trim(cbSub.Text)) = 1 then
        begin
          SQLStr := 'Select * From [z' + Trim(cmbExam.Text) + Copy(CheckListBox3.Items.Strings[ii], 1, 6) + ']';
        end
        else
        begin
          SQLStr := 'Select * From [' + Trim(cmbExam.Text) + Copy(CheckListBox3.Items.Strings[ii], 1, 3) + ']';
        end;
        SQLStr := SQLStr + ' Where Sub_No=' + #39 + Trim(cbSub.Text) + #39 + '   And Left(Ans,' + IntToStr(iQCount) + ')=' + #39 + sAnsTemp + #39 + '   And Absent=''0'' ' + ' Order by Student_No;';
        OpenSQL(DM.qrySearch, SQLStr);
        while not DM.qrySearch.Eof do
        begin
          if Length(Trim(cbSub.Text)) = 1 then
          begin
            Memo2.Lines.Add(DM.qrySearch.FieldByName('Stud_No').AsString + '、' + DM.qrySearch.FieldByName('Ans').AsString);
          end
          else
          begin
            Memo2.Lines.Add(DM.qrySearch.FieldByName('Student_No').AsString + '、' + DM.qrySearch.FieldByName('Ans').AsString);
          end;
          DM.qrySearch.Next;
        end;
        ProgressBar1.StepIt;
      end;
    end;
    if Memo2.Lines.Count >= 2 then
    begin
      showmessage('檢查完畢！如確定均為缺考，請按［設為缺考］');
      btnAbsent.Enabled := True;
    end
    else
    begin
      showmessage('檢查完畢! 目前沒有需要設定的學生資料!');
    end;

  end
  else
  begin
    showmessage('請選擇一個科目！');
    cbSub.SetFocus;
  end;

end;

procedure TfmMain.btnSetClick(Sender: TObject);
var
  Str1: string;
  ii, jj: Integer;
begin
  if TableExists('Sch_Exam') then
    Str1 := 'Sch_Exam'
  else if TableExists('School_Exam') then
    Str1 := 'School_Exam'
  else
    Str1 := '';

  SQLStr := 'Delete From ' + Str1 + ' Where Exam_No=' + #39 + Trim(cmbExam.Text) + #39;
  SQLExec(DM.qryTemp, SQLStr);

  for ii := 0 to CheckListBox1.Items.Count - 1 do
  begin
    if CheckListBox1.Checked[ii] then
    begin
      jj := POS(':', Trim(CheckListBox1.Items.Strings[ii]));
      SQLStr := 'Insert Into ' + Str1 + '(Exam_No, Sch_Code, C_Flag)' + ' Values(' + #39 + Trim(cmbExam.Text) + #39 + ',' + #39 + Copy(CheckListBox1.Items.Strings[ii], 1, jj - 1) + #39 + ',' + #39 + '1' + #39 + ')';
      SQLExec(DM.qryTemp, SQLStr);
    end;

  end;
end;

procedure TfmMain.btnOutClick(Sender: TObject);
var
  SListA: TStringList;
  ii, iSCount: Integer;
  Str1, sStd, Str2: string;
  ExcelApp: Variant;
  StrFormat: OleVariant;
begin
  //單張缺考
  //序	校代碼	學校名稱	班級座號	姓名	科目
  SListA := TStringList.Create;
  SListA.Add('校代碼' + #9 + '學校名稱' + #9 + '年級' + #9 + '班級' + #9 + '座號' + #9 + '姓名' + #9 + '科目' + #9);
  ProgressBar1.Min := 1;
  if TableExists('Exam_SubQ') then
  begin
    Str1 := 'Exam_SubQ';
    sStd := 'Stud_No';
  end
  else if TableExists('Exam_Sub') then
  begin
    Str1 := 'Exam_Sub';
    sStd := 'Student_No';
  end
  else
  begin
    Str1 := '';
    sStd := '';
  end;

  SQLStr := 'Select Count(*) as iCount From ' + Str1 + ' Where Exam_No=' + #39 + Trim(cmbExam.Text) + #39;
  if Str1 = 'Exam_SubQ' then
    SQLStr := SQLStr + '   And Card_Cnt<>0 '
  else if Str1 = 'Exam_Sub' then
    SQLStr := SQLStr + '   And Card_Count<>0 ';
  OpenSQL(DM.qryTemp, SQLStr);
  iSCount := DM.qryTemp.FieldByName('iCount').AsInteger;

  for ii := 0 to CheckListBox3.Items.Count - 1 do
  begin
    if CheckListBox3.Checked[ii] then
    begin
      if Str1 = 'Exam_SubQ' then
        SQLStr := ' Select ' + sStd + ' From [z' + Trim(cmbExam.Text) + Copy(CheckListBox3.Items.Strings[ii], 1, 6) + '] '
      else if Str1 = 'Exam_Sub' then
        SQLStr := ' Select ' + sStd + ' From [' + Trim(cmbExam.Text) + Copy(CheckListBox3.Items.Strings[ii], 1, 3) + '] ';
      SQLStr := SQLStr + ' Where ErrCode like ''%2%'' And Ans is NULL ' + ' Group by ' + sStd + ' Having COUNT(*)=1 ';
      OpenSQL(DM.qryTemp, SQLStr);
      while not DM.qryTemp.Eof do
      begin
        SQLStr := 'Select Left(a.' + sStd + ', 6) as Sch_Code, b.Sch_Name,c.Grade, c.Stud_Name, c.Class_No, c.Seat_No, d.Sub_Name' + '  From [z' + Trim(cmbExam.Text) + Copy(CheckListBox3.Items.Strings[ii], 1, 6) + '] a ' + ' Inner Join Sch_Base b on Left(a.Stud_No,6)=b.Sch_Code ' + ' Inner Join Stud_Base c on a.Stud_No=c.Stud_No ' + ' Inner Join Ex_Subject d on a.Sub_No=d.Sub_No ' + ' Where a.ErrCode like ''%2%'' ' + '   And a.Stud_No=' + #39 + Trim(DM.qryTemp.FieldByName('Stud_No').AsString) + #39;
        OpenSQL(DM.qrySearch, SQLStr);

        Str2 := DM.qrySearch.FieldByName('Sch_Code').AsString + #9 + DM.qrySearch.FieldByName('Sch_Name').AsString + #9 + DM.qrySearch.FieldByName('Grade').AsString + #9 + DM.qrySearch.FieldByName('Class_No').AsString + #9 + DM.qrySearch.FieldByName('Seat_No').AsString + #9 + Trim(DM.qrySearch.FieldByName('Stud_Name').AsString) + #9 + Trim(DM.qrySearch.FieldByName('Sub_Name').AsString) + #9;
        SListA.Append(Str2);
        Str2 := '';
        DM.qryTemp.Next;
      end;
      ProgressBar1.StepIt;

    end;
  end;

  SaveDialog1.FileName := '單科缺卡' + Trim(cmbExam.Text) + '.xls';
  if SaveDialog1.Execute then
  begin
    try
      ExcelApp := CreateOleObject('Excel.Application'); //Create Excel 物件
    except
      ShowMessage('尚未安裝任何 Excel 版本.');
      SListA.Free;
      Exit;
    end;

    if FileExists(SaveDialog1.FileName) then
      DeleteFile(SaveDialog1.FileName);

    ExcelApp.WorkBooks.Add; //新增工作簿(預設為三個工作表)
    ExcelApp.Visible := False; //不顯示Excel 視窗
    ExcelApp.WorkSheets[1].Activate;
    ExcelApp.WorkSheets[1].Name := 'Sheet1'; //工作表更名
    StrFormat := '@'; //@: 儲存格格式改為文字
    ExcelApp.WorkSheets[1].Cells.NumberFormatLocal := StrFormat; //設定儲格格式(一定要宣告OleVariant，直接等於'@'無
    Clipboard.Clear; //先清空剪貼簿
    Clipboard.AsText := SListA.Text; //複製資料到剪貼簿
    ExcelApp.Range['A1'].Select;
    ExcelApp.Range['A1'].PasteSpecial; //在A1貼上
    Clipboard.Clear; //用完清空剪貼簿
      //ExcelApp.Range['A1'].Select;
    ExcelApp.ActiveWorkBook.Saved := True; //設定不存檔，若不設定關閉時會出現"是否存檔的對話框"
    ExcelApp.WorkBooks[1].SaveAs(SaveDialog1.FileName); //存檔
    ExcelApp.WorkBooks.close; //關閉Excel
    ExcelApp.Quit; //離開Excel
    ExcelApp := Unassigned; //釋放ExcelApp;
  end;

  showmessage('轉檔完畢!');

  SListA.Clear;
  SListA.Free;

end;

procedure TfmMain.cbSubChange(Sender: TObject);
begin
  if Trim(cbSub.Text) = '' then
  begin
    showmessage('請選擇一個科目代碼!');
    cbSub.SetFocus;
  end;
end;

procedure TfmMain.CheckBox1Click(Sender: TObject);
var
  ii: Integer;
begin
  if CheckBox1.Checked then
  begin
    for ii := 0 to CheckListBox1.Count - 1 do
      CheckListBox1.Checked[ii] := True;
  end
  else
  begin
    for ii := 0 to CheckListBox1.Count - 1 do
      CheckListBox1.Checked[ii] := False;
  end;
end;

procedure TfmMain.CheckBox2Click(Sender: TObject);
var
  ii: Integer;
begin
  if CheckBox2.Checked then
  begin
    for ii := 0 to CheckListBox2.Items.Count - 1 do
      CheckListBox2.Checked[ii] := True;
  end
  else
  begin
    for ii := 0 to CheckListBox2.Items.Count - 1 do
      CheckListBox2.Checked[ii] := False;
  end;
end;

procedure TfmMain.CheckBox4Click(Sender: TObject);
var
  ii: Integer;
begin
  if CheckBox4.Checked then
  begin
    for ii := 0 to CheckListBox3.Count - 1 do
      CheckListBox3.Checked[ii] := True;
  end
  else
  begin
    for ii := 0 to CheckListBox3.Count - 1 do
      CheckListBox3.Checked[ii] := False;
  end;

end;

procedure TfmMain.CheckBox5Click(Sender: TObject);
var
  ii: Integer;
begin
  if CheckBox5.Checked then
  begin
    for ii := 0 to CheckListBox4.Count - 1 do
      CheckListBox4.Checked[ii] := True;
  end
  else
  begin
    for ii := 0 to CheckListBox4.Count - 1 do
      CheckListBox4.Checked[ii] := False;
  end;
  CheckListBox4ClickCheck(nil);
end;

procedure TfmMain.CheckListBox4ClickCheck(Sender: TObject);
var
  ii: integer;
begin
  mmoCheck.Clear;
  for ii := 0 to CheckListBox4.Count - 1 do
    if CheckListBox4.Checked[ii] then
      mmoCheck.Lines.Add(CheckListBox4.Items[ii]);
end;

procedure TfmMain.cmbDBNameClick(Sender: TObject);
begin
  if Trim(cmbDBName.Text) <> '' then
  begin
    DM.ADOConnection1.Close;
//    DM.ADOConnection1.ConnectionString := 'Provider=SQLOLEDB.1'
//                                        +';User ID='+trim(edtUser.Text)+''
//                                        +';Password='+trim(edtPW.Text)+''
//                                        +';Data Source='+trim(cmbIP.Text)+';'
//                                        +'Initial Catalog='+trim(cmbDBName.Text)+'';
    DM.ADOConnection1.ConnectionString := 'Provider=SQLOLEDB.1' + ';Integrated Security=SSPI' + '' + ';Persist Security Info=False' + '' + ';Data Source= .' + ';' + 'Initial Catalog=' + trim(cmbDBName.Text) + '';
    DM.ADOConnection1.Open;
    DM.ADOQuery1.Close;
    DM.ADOQuery1.SQL.Clear;
    if ClassGroup.ItemIndex = 0 then
    begin
      if TableExists('Ex_Base') then
        DM.ADOQuery1.SQL.Add('SELECT Exam_No, Exam_Name FROM Ex_Base ORDER BY Exam_No')
      else
        DM.ADOQuery1.SQL.Add('SELECT Exam_No, Exam_Name FROM Exam ORDER BY Exam_No');
    end
    else
      DM.ADOQuery1.SQL.Add('SELECT Exam_No, Exam_Name FROM Exam ORDER BY Exam_No');
    DM.ADOQuery1.Open;

    cmbExam.Clear;
    while DM.ADOQuery1.eof <> True do
    begin
      cmbExam.Items.Add(DM.ADOQuery1.fieldByName('Exam_No').AsString);
      DM.ADOQuery1.Next;
    end;
    cmbExam.Text := '請選擇考試';
  end
  else
  begin
    showmessage('請選擇資料庫!');
  end;
end;

procedure TfmMain.cmbExamChange(Sender: TObject);
begin
  if Trim(cmbExam.Text) <> '' then
    PageControl1.ActivePageIndex := 0;
  PageControl1.OnChange(Self);
end;

procedure TfmMain.cmbIPChange(Sender: TObject);
begin
  if Trim(cmbIP.Text) <> '' then //有輸入IP
  begin
    edtUser.Text := 'sa';
    edtPW.Text := 'seat';
    if (Trim(cmbIP.Text) = '172.16.100.70') or (Trim(cmbIP.Text) = '172.16.44.208') or (Trim(cmbIP.Text) = '172.16.100.63') or (Trim(cmbIP.Text) = '172.16.100.72') then //國中
    begin
      ClassGroup.ItemIndex := 0;
    end
    else if (Trim(cmbIP.Text) = '172.16.100.60') or (Trim(cmbIP.Text) = '172.16.100.61') then
    begin   //高中 or 其他零星考試
      ClassGroup.ItemIndex := 1;
    end
    else if (Trim(cmbIP.Text) = '172.16.100.62') then //高職
    begin
      ClassGroup.ItemIndex := 2;
    end
    else
      ClassGroup.ItemIndex := 3;
  end
  else
  begin
    showmessage('確定不選IP嗎?');
  end;
end;

procedure TfmMain.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  Action := caFree;
end;

procedure TfmMain.FormDestroy(Sender: TObject);
begin
  fmMain := Nil;
end;

procedure TfmMain.FormShow(Sender: TObject);
begin
  PageControl1.ActivePageIndex := 0;
end;

procedure TfmMain.PageControl1Change(Sender: TObject);
var
  Str1, Str2: string;
begin
  if TableExists('Sch_Base') then
    Str1 := 'Sch_Base'
  else if TableExists('School') then
    Str1 := 'School'
  else
    Str1 := '';

  if TableExists('Sch_Exam') then
    Str2 := 'Sch_Exam'
  else if TableExists('School_Exam') then
    Str2 := 'School_Exam'
  else
    Str2 := '';

  case PageControl1.ActivePageIndex of
    0:
      begin   //學校基本資料
        if Str1 <> '' then
          SQLStr := 'Select * From ' + Str1 + ' Order by Sch_Code;';
        OpenSQL(DM.qrySch, SQLStr);

        if not DM.qrySch.IsEmpty then
        begin
          DBGrid1.DataSource := DM.dsSch;
          DBGrid1.Columns[0].FieldName := 'Sch_Code';
          DBGrid1.Columns[1].FieldName := 'Sch_Name';
          DBGrid1.Columns[2].FieldName := 'Sch_Add';
          if Str1 = 'Sch_Base' then
            DBGrid1.Columns[3].FieldName := 'Sch_Area'
          else
            DBGrid1.Columns[3].FieldName := 'Sch_ACode';
        end;
      end;
    1:
      begin   //學校班級名稱資料
        SQLStr := 'Select * From Sch_Class ' + ' Order by Sch_Code, Class_No;';
        OpenSQL(DM.qrySClass, SQLStr);
        if not DM.qrySClass.IsEmpty then
        begin
          DBGrid2.DataSource := DM.dsSClass;
          DBGrid2.Columns[0].FieldName := 'Sch_Code';
          DBGrid2.Columns[1].FieldName := 'Grade';
          DBGrid2.Columns[2].FieldName := 'Class_No';
          DBGrid2.Columns[3].FieldName := 'Class_Name';
        end;
      end;
    2:
      begin   //參加考試學校
        SQLStr := 'Select Sch_Code, Sch_Name' + '  From ' + Str1 + ' Order by Sch_Code;';
        OpenSQL(DM.qryTemp, SQLStr);
        CheckListBox1.Clear;
        while not DM.qryTemp.Eof do
        begin
          CheckListBox1.Items.Add(Trim(DM.qryTemp.FieldByName('Sch_Code').AsString) + ':' + Trim(DM.qryTemp.FieldByName('Sch_Name').AsString));
          DM.qryTemp.Next;
        end;
        if Str1 = 'Sch_Base' then
          SQLStr := 'Select a.Exam_No, a.Sch_Code, b.Sch_Name From Sch_Exam a'
        else if Str1 = 'School' then
          SQLStr := 'Select a.Exam_No, a.Sch_Code, b.Sch_Name From School_Exam a';
        SQLStr := SQLStr + ' Inner Join ' + Str1 + ' b on a.Sch_Code=b.Sch_Code ' + ' Where a.Exam_No=' + #39 + Trim(cmbExam.Text) + #39 + ' Order by a.Sch_Code;';
        OpenSQL(DM.qryExam, SQLStr);
        DBGrid3.DataSource := DM.dsExam;
        DBGrid3.Columns[0].FieldName := 'Exam_No';
        DBGrid3.Columns[1].FieldName := 'Sch_Code';
        DBGrid3.Columns[2].FieldName := 'Sch_Name';
      end;
    3:
      begin  //答案卡資料
        cbSubT.Clear;
        SQLStr := ' Select a.Sub_No, b.Sub_Name ';

        if Str1 = 'Sch_Base' then
          SQLStr := SQLStr + '   From Exam_SubQ a '
        else
          SQLStr := SQLStr + '   From Exam_Sub a ';

        SQLStr := SQLStr + '  Inner Join Ex_Subject b on a.Sub_No=b.Sub_No' + '  Where a.Exam_No=' + #39 + Trim(cmbExam.Text) + #39 + '  Order by a.Sub_No;';
        OpenSQL(DM.qryTemp, SQLStr);
        while not DM.qryTemp.Eof do
        begin
          cbSubT.Items.Add(DM.qryTemp.FieldByName('Sub_No').AsString + '-' + Trim(DM.qryTemp.FieldByName('Sub_Name').AsString));
          DM.qryTemp.Next;
        end;
      end;
    4:
      begin   //資料檢查
        cbSub.Clear;
        SQLStr := ' Select Sub_No ';

        if Str1 = 'Sch_Base' then
          SQLStr := SQLStr + '   From Exam_SubQ '
        else
          SQLStr := SQLStr + '   From Exam_Sub ';

        SQLStr := SQLStr + '  Where Exam_No=' + #39 + Trim(cmbExam.Text) + #39 + '  Order by Sub_No;';
        OpenSQL(DM.qryTemp, SQLStr);

        while not DM.qryTemp.Eof do
        begin
          cbSub.Items.Add(DM.qryTemp.FieldByName('Sub_No').AsString);
          DM.qryTemp.Next;
        end;

        SQLStr := ' Select a.Sch_Code, b.Sch_Name' + '   From ' + Str2 + ' a ' + '  Inner Join ' + Str1 + ' b on a.Sch_Code=b.Sch_Code' + '  Where a.Exam_No=' + #39 + Trim(cmbExam.Text) + #39 + '    And a.C_Flag>=''5'' ' + '  Order by a.Sch_Code;';
        OpenSQL(DM.qryTemp, SQLStr);
        CheckListBox3.Clear;
        while not DM.qryTemp.Eof do
        begin
          CheckListBox3.Items.Add(DM.qryTemp.FieldByName('Sch_Code').AsString + '-' + Trim(DM.qryTemp.FieldByName('Sch_Name').AsString));

          DM.qryTemp.Next;
        end;
      end;
    5:
      begin   //Excel匯出
        SQLStr := ' Select a.Sch_Code, b.Sch_Name' + '   From ' + Str2 + ' a ' + '  Inner Join ' + Str1 + ' b on a.Sch_Code=b.Sch_Code' + '  Where a.Exam_No=' + #39 + Trim(cmbExam.Text) + #39 + '    And a.C_Flag>=''5'' ' + '  Order by a.Sch_Code;';
        OpenSQL(DM.qryTemp, SQLStr);
        CheckListBox3.Clear;
        while not DM.qryTemp.Eof do
        begin
          CheckListBox4.Items.Add(DM.qryTemp.FieldByName('Sch_Code').AsString + '-' + Trim(DM.qryTemp.FieldByName('Sch_Name').AsString));

          DM.qryTemp.Next;
        end;
      end;
  end;

end;

procedure TfmMain.SpeedButton1Click(Sender: TObject);
begin
//  DM.ADOConnection1.Close;
//  DM.ADOConnection1.ConnectionString:='Provider=SQLOLEDB.1'
//                                     +';User ID='+Trim(edtUser.Text)+''
//                                      +';Password='+Trim(edtPW.Text)+''
//                                      +';Data Source='+trim(cmbIP.Text)+''
//                                      +';Initial Catalog=master';
  DM.ADOConnection1.ConnectionString := 'Provider=SQLOLEDB.1' + ';Integrated Security=SSPI' + '' + ';Persist Security Info=False' + '' + ';Data Source= .' + ';';
//                                        +'Initial Catalog='+trim(cmbDBName.Text)+'';
  try
    DM.ADOConnection1.Open;
    cmbDBName.Clear;
    DM.ADOQuery1.Close;
    DM.ADOQuery1.SQL.Clear;
    DM.ADOQuery1.SQL.Add('select name,dbid from sysdatabases order by dbid');
    DM.ADOQuery1.Open;
    while DM.ADOQuery1.eof <> True do
    begin
      cmbDBName.Items.Add(DM.ADOQuery1.fieldByName('name').AsString);
      DM.ADOQuery1.Next;
    end;
    cmbDBName.Text := '請選擇';
  except
    ShowMessage('『連結失敗』');
  end;
  cmbDBName.SetFocus;

end;

end.

