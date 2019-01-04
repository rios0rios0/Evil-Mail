program EM;

uses
  SysUtils,
  Forms,
  UEMP in 'UEMP.pas' {Form2},
  UEM in 'UEM.pas' {Form1};

{$R *.res}

begin
  Application.Initialize;
  Application.Title := 'Evil Mail';
  Application.CreateForm(TForm1, Form1);
  Application.CreateForm(TForm2, Form2);
  Application.Run;
end.
