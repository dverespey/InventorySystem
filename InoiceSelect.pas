unit InoiceSelect;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, DB, Grids, DBGrids, StdCtrls, ExtCtrls;

type
  TInvoiceSelect_Form = class(TForm)
    CarTruckShip_Label: TLabel;
    Button_Panel: TPanel;
    CreateASN_Button: TButton;
    Close_Button: TButton;
    GroupBox1: TGroupBox;
    DBGrid1: TDBGrid;
    DataSource1: TDataSource;
  private
    { Private declarations }
  public
    { Public declarations }
    procedure Execute;
  end;

var
  InvoiceSelect_Form: TInvoiceSelect_Form;

implementation

{$R *.dfm}

procedure TInvoiceSelect_Form.Execute;
begin
end;

end.
