unit DataM;

interface

uses
  SysUtils, Classes, DB, ADODB;

type
  TDM = class(TDataModule)
    qrySch: TADOQuery;
    ADOConnection1: TADOConnection;
    ADOQuery1: TADOQuery;
    DataSource1: TDataSource;
    dsSch: TDataSource;
    qryTemp: TADOQuery;
    qrySearch: TADOQuery;
    qrySClass: TADOQuery;
    dsSClass: TDataSource;
    qryExam: TADOQuery;
    dsExam: TDataSource;
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  DM: TDM;

implementation

{$R *.dfm}

end.
