program CalculoTributario;

uses
  Forms,
  uFrmCalculoTributario in 'uFrmCalculoTributario.pas' {F_CalculoTributario};

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TF_CalculoTributario, F_CalculoTributario);
  Application.Run;
end.
