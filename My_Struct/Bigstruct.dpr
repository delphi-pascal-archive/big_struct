program Bigstruct;

uses
  Forms,
  Bs in 'Bs.pas' {Form1},
  open in 'open.pas' {Form2};

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TForm1, Form1);
  Application.CreateForm(TForm2, Form2);
  Application.Run;
end.
