unit uFrmCalculoTributario;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, DB, DBClient, Grids, DBGrids, EDBGrid, ENumEd, Math,
  scExcelExport, ExtCtrls, Buttons;

type
  TF_CalculoTributario = class(TForm)
    Panel1: TPanel;
    GroupBox1: TGroupBox;
    Label15: TLabel;
    Label16: TLabel;
    Label17: TLabel;
    Label18: TLabel;
    Label20: TLabel;
    Label21: TLabel;
    Label22: TLabel;
    MTotalProduto2: TEvNumEdit;
    MtotalIPI2: TEvNumEdit;
    MTotalBaseICMS: TEvNumEdit;
    MTotalICMS: TEvNumEdit;
    MTotalBaseICMSST: TEvNumEdit;
    MTotal: TEvNumEdit;
    MTotalValorICMSST: TEvNumEdit;
    trb_itens: TClientDataSet;
    trb_itensCodProduto: TStringField;
    trb_itensItem: TIntegerField;
    trb_itensQtde: TFloatField;
    trb_itensPrecoUnitario: TFloatField;
    trb_itensTotalProduto: TFloatField;
    trb_itensAliqIPI: TFloatField;
    trb_itensTotalIPI: TFloatField;
    trb_itensAliqICMS: TFloatField;
    trb_itensBaseICMS: TFloatField;
    trb_itensValorICMS: TFloatField;
    trb_itensAliqICMSST: TFloatField;
    trb_itensBaseICMSST: TFloatField;
    trb_itensValorICMSST: TFloatField;
    trb_itensAliqIVA: TFloatField;
    DSItens: TDataSource;
    Excel: TscExcelExport;
    GroupBox2: TGroupBox;
    Label5: TLabel;
    Label6: TLabel;
    Label8: TLabel;
    Label9: TLabel;
    Label10: TLabel;
    Label11: TLabel;
    Label12: TLabel;
    Label13: TLabel;
    Label7: TLabel;
    MAliqIPI: TEvNumEdit;
    MValorIPI: TEvNumEdit;
    MAliqICMS: TEvNumEdit;
    MBaseICMS: TEvNumEdit;
    MValorICMS: TEvNumEdit;
    MAliqICMSST: TEvNumEdit;
    MBaseICMSST: TEvNumEdit;
    MValorICMSST: TEvNumEdit;
    MAliqIVA: TEvNumEdit;
    GroupBox3: TGroupBox;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Label14: TLabel;
    MQtde: TEvNumEdit;
    MPrecoUnitario: TEvNumEdit;
    MTotalProduto: TEvNumEdit;
    MCodProduto: TEdit;
    MItem: TEvNumEdit;
    MSalvar: TBitBtn;
    MCancelar: TBitBtn;
    MGeraExcel: TBitBtn;
    Panel2: TPanel;
    MGrade: TEvDBGrid3D;
    Panel3: TPanel;
    lbl_QtdeItens: TLabel;
    MCalculaICMS: TCheckBox;
    trb_itensCalculaICMS: TStringField;
    procedure MPrecoUnitarioExit(Sender: TObject);
    procedure MAliqIPIExit(Sender: TObject);
    procedure MAliqICMSExit(Sender: TObject);
    procedure MAliqICMSSTExit(Sender: TObject);
    procedure MAliqIVAExit(Sender: TObject);
    procedure MSalvarClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure inicializa;
    procedure FormShow(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure BitBtn1Click(Sender: TObject);
    procedure MCancelarClick(Sender: TObject);
    procedure MGeraExcelClick(Sender: TObject);
    procedure DSItensDataChange(Sender: TObject; Field: TField);
    procedure MGradeDblClick(Sender: TObject);
    procedure CalculaTotal;
    procedure FormKeyPress(Sender: TObject; var Key: Char);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  F_CalculoTributario: TF_CalculoTributario;

implementation

{$R *.dfm}

procedure TF_CalculoTributario.MSalvarClick(Sender: TObject);
begin

   if MCodProduto.Text <> EmptyStr then
   begin
      trb_itens.Append;
      trb_itensCodProduto.AsString   := MCodProduto.Text;
      trb_itensItem.Text             := MItem.Text;
      trb_itensQtde.AsFloat          := MQtde.Value;
      trb_itensPrecoUnitario.AsFloat := MPrecoUnitario.Value;
      trb_itensTotalProduto.AsFloat  := MTotalProduto.Value;
      trb_itensAliqIPI.AsFloat       := MAliqIPI.Value;
      trb_itensTotalIPI.AsFloat      := MValorIPI.Value;
      trb_itensAliqICMS.AsFloat      := MAliqICMS.Value;
      trb_itensBaseICMS.AsFloat      := MBaseICMS.Value;
      trb_itensValorICMS.AsFloat     := MValorICMS.Value;
      trb_itensAliqIVA.AsFloat       := MAliqIVA.Value;
      trb_itensBaseICMSST.AsFloat    := MBaseICMSST.Value;
      trb_itensAliqICMSST.AsFloat    := MAliqICMSST.Value;
      trb_itensValorICMSST.Value     := MValorICMSST.Value;
      trb_itens.Post;

      MTotalProduto2.Value := MTotalProduto2.Value + trb_itensTotalProduto.AsFloat;
      MtotalIPI2.Value     := MtotalIPI2.Value + trb_itensTotalIPI.AsFloat;
      MTotalBaseICMS.Value := MTotalBaseICMS.Value + trb_itensBaseICMS.AsFloat;
      MTotalICMS.Value     := MTotalICMS.Value + trb_itensValorICMS.AsFloat;
      MTotalBaseICMSST.Value := MTotalBaseICMSST.Value + trb_itensBaseICMSST.AsFloat;
      MTotalValorICMSST.Value := MTotalValorICMSST.Value + trb_itensValorICMSST.AsFloat;
      MTotal.Value := MTotalProduto.Value + MValorIPI.Value + MTotalValorICMSST.Value;
   end;
   
   inicializa;

end;

procedure TF_CalculoTributario.BitBtn1Click(Sender: TObject);
begin
   if (MCodProduto.Text <> EmptyStr) and (MItem.Value > 0) then
   begin

      if trb_itens.Locate('CodProduto; Item', VarArrayOf([MCodProduto.Text, MItem.Text]), []) then
      begin
         if Application.MessageBox('Produto Com Tributo Já Calculado!' + #13 + 'Deseja Alterar?','Atenção', MB_YESNO + MB_ICONQUESTION + MB_DEFBUTTON2) = IDNo then
         begin
            MCodProduto.SetFocus;
            Exit;
         end
         else
            trb_itens.Edit;
      end
      else
         trb_itens.Append;

      trb_itensCodProduto.AsString   := MCodProduto.Text;
      trb_itensItem.Text             := MItem.Text;
      trb_itensQtde.AsFloat          := MQtde.Value;
      trb_itensPrecoUnitario.AsFloat := MPrecoUnitario.Value;
      trb_itensTotalProduto.AsFloat  := MTotalProduto.Value;
      trb_itensAliqIPI.AsFloat       := MAliqIPI.Value;
      trb_itensTotalIPI.AsFloat      := MValorIPI.Value;
      trb_itensAliqICMS.AsFloat      := MAliqICMS.Value;
      trb_itensBaseICMS.AsFloat      := MBaseICMS.Value;
      trb_itensValorICMS.AsFloat     := MValorICMS.Value;
      trb_itensAliqIVA.AsFloat       := MAliqIVA.Value;
      trb_itensBaseICMSST.AsFloat    := MBaseICMSST.Value;
      trb_itensAliqICMSST.AsFloat    := MAliqICMSST.Value;
      trb_itensValorICMSST.Value     := MValorICMSST.Value;

      if MCalculaICMS.Checked then
         trb_itensCalculaICMS.AsString  := 'S'
      else
         trb_itensCalculaICMS.AsString  := 'N';

      trb_itens.Post;

      CalculaTotal;
   end;

   inicializa;
end;

procedure TF_CalculoTributario.MCancelarClick(Sender: TObject);
begin
   inicializa;
end;

procedure TF_CalculoTributario.MGeraExcelClick(Sender: TObject);
begin
   if trb_itens.RecordCount > 0 then
   begin
      Excel.Dataset:= trb_itens;
      Excel.Connect;
      Excel.ExportDataset;
      Excel.Disconnect;
   end;

end;

procedure TF_CalculoTributario.Button1Click(Sender: TObject);
begin
   if trb_itens.RecordCount > 0 then
   begin
      Excel.Dataset:= trb_itens;
      Excel.Connect;
      Excel.ExportDataset;
      Excel.Disconnect;
   end;

end;

procedure TF_CalculoTributario.CalculaTotal;
begin
   MTotalProduto2.Value    := 0;
   MtotalIPI2.Value        := 0;
   MTotalBaseICMS.Value    := 0;
   MTotalICMS.Value        := 0;
   MTotalBaseICMSST.Value  := 0;
   MTotalValorICMSST.Value := 0;
   MTotal.Value            := 0;

   if trb_itens.RecordCount > 0 then
   begin
      trb_itens.ControlsDisabled;
   
      trb_itens.First;
      while not trb_itens.Eof do
      begin
         MTotalProduto2.Value    := MTotalProduto2.Value + trb_itensTotalProduto.AsFloat;
         MtotalIPI2.Value        := MtotalIPI2.Value + trb_itensTotalIPI.AsFloat;
         MTotalBaseICMS.Value    := MTotalBaseICMS.Value + trb_itensBaseICMS.AsFloat;
         MTotalICMS.Value        := MTotalICMS.Value + trb_itensValorICMS.AsFloat;
         MTotalBaseICMSST.Value  := MTotalBaseICMSST.Value + trb_itensBaseICMSST.AsFloat;
         MTotalValorICMSST.Value := MTotalValorICMSST.Value + trb_itensValorICMSST.AsFloat;
         MTotal.Value            := MTotalProduto.Value + MValorIPI.Value + MTotalValorICMSST.Value;        
         
         trb_itens.Next;
      end;

      trb_itens.First;
      trb_itens.EnableControls;
   end;
end;

procedure TF_CalculoTributario.DSItensDataChange(Sender: TObject;
  Field: TField);
begin
   lbl_QtdeItens.Caption := IntToStr(trb_itens.RecNo) + ' de ' + IntToStr(trb_itens.RecordCount);
end;

procedure TF_CalculoTributario.FormCreate(Sender: TObject);
begin
   trb_itens.Close;
   trb_itens.CreateDataSet;
   trb_itens.Open;
end;

procedure TF_CalculoTributario.FormKeyPress(Sender: TObject; var Key: Char);
begin
    case key of
      #13: Perform(WM_NEXTDLGCTL,0,0);
    end;
end;

procedure TF_CalculoTributario.FormShow(Sender: TObject);
begin
   inicializa;
end;

procedure TF_CalculoTributario.inicializa;
begin
   MCodProduto.Text     := EmptyStr;
   MItem.Value          := 0;
   MQtde.Value          := 0;
   MPrecoUnitario.Value := 0;
   MTotalProduto.Value  := 0;
   MAliqIPI.Value       := 0;
   MValorIPI.Value      := 0;
   MAliqICMS.Value      := 0;
   MBaseICMS.Value      := 0;
   MValorICMS.Value     := 0;
   MAliqIVA.Value       := 0;
   MBaseICMSST.Value    := 0;
   MAliqICMSST.Value    := 0;
   MValorICMSST.Value   := 0;
   MCalculaICMS.Checked := False; 

   MCodProduto.SetFocus;
end;

procedure TF_CalculoTributario.MAliqICMSExit(Sender: TObject);
Var
   BaseICMS : real;
begin
   if MAliqICMS.Value > 0 then
   begin
      if MTotalProduto.Value > 0 then
      begin
         if MCalculaICMS.Checked then
         begin
            MBaseICMS.Value := MTotalProduto.Value;
            MValorICMS.Value := RoundTo((MAliqICMS.Value * MBaseICMS.Value) / 100, -2);
         end;
      end;
   end
   else
   begin
      MValorICMS.Value := 0;
      MBaseICMS.Value  := 0;
   end;

   MAliqIVA.SetFocus;
end;

procedure TF_CalculoTributario.MAliqICMSSTExit(Sender: TObject);
var
   BaseICMS, ValorICMS : Real;
begin
   ValorICMS := 0;
   BaseICMS  := 0;

   if MAliqICMSST.Value > 0 then
   begin
      if MBaseICMSST.Value > 0 then
      begin
         if MAliqICMS.Value = 0 then
         begin
            ShowMessage('Por Favor Informe Aliquota de ICMS para que seja possivel Calcular o valor ICMS ST!');
            MAliqICMS.SetFocus;
            Exit;
         end;

         BaseICMS  := MTotalProduto.Value;
         ValorICMS := RoundTo((BaseICMS * MAliqICMS.Value) / 100, -2);

         MValorICMSST.Value := RoundTo(((MBaseICMSST.Value * MAliqICMSST.Value) / 100) - ValorICMS, -2);
      end;
   end;

   MSalvar.SetFocus;
end;

procedure TF_CalculoTributario.MAliqIPIExit(Sender: TObject);
begin
   if MAliqIPI.Value > 0 then
   begin
      if MTotalProduto.Value > 0 then
      begin
         MValorIPI.Value := RoundTo((MAliqIPI.Value * MTotalProduto.Value) / 100, -2);
      end;
   end
   else
      MValorIPI.Value := 0;

   MAliqICMS.SetFocus;

end;

procedure TF_CalculoTributario.MAliqIVAExit(Sender: TObject);
var
   BaseST : Real;
begin
   if MAliqIVA.Value > 0 then
   begin
      if MTotalProduto.Value > 0 then
      begin
         BaseST := MTotalProduto.Value + MValorIPI.Value;
         MBaseICMSST.Value := RoundTo(BaseST * (1 + (MAliqIVA.Value / 100)), -2);
      end;
   end
   else
   begin
      MValorICMSST.Value := 0;
      MBaseICMSST.Value  := 0;
      MValorICMSST.Value := 0;
      MAliqICMSST.Value  := 0;
   end;

   MAliqICMSST.SetFocus;
end;

procedure TF_CalculoTributario.MGradeDblClick(Sender: TObject);
begin
   if trb_itens.RecordCount > 0 then
   begin
      MCodProduto.Text     := trb_itensCodProduto.AsString;
      MItem.Value          := trb_itensItem.AsFloat;
      MQtde.Value          := trb_itensQtde.AsFloat;
      MPrecoUnitario.Value := trb_itensPrecoUnitario.AsFloat;
      MTotalProduto.Value  := trb_itensTotalProduto.AsFloat;
      MAliqIPI.Value       := trb_itensAliqIPI.AsFloat;
      MValorIPI.Value      := trb_itensTotalIPI.AsFloat;
      MAliqICMS.Value      := trb_itensAliqICMS.AsFloat;
      MBaseICMS.Value      := trb_itensBaseICMS.AsFloat;
      MValorICMS.Value     := trb_itensValorICMS.AsFloat;
      MAliqIVA.Value       := trb_itensAliqIVA.AsFloat;
      MBaseICMSST.Value    := trb_itensBaseICMSST.AsFloat;
      MAliqICMSST.Value    := trb_itensAliqICMSST.AsFloat;
      MValorICMSST.Value   := trb_itensValorICMSST.AsFloat;

      if trb_itensCalculaICMS.AsString = 'S' then
         MCalculaICMS.Checked := True
      else
         MCalculaICMS.Checked := False;

   end;

end;

procedure TF_CalculoTributario.MPrecoUnitarioExit(Sender: TObject);
begin
   if MQtde.Value > 0 then
   begin
      MTotalProduto.Value := RoundTo(MQtde.Value * MPrecoUnitario.Value, -2);
   end;

end;

end.
