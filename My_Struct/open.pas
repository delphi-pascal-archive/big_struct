unit open;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls;

type
  TForm2 = class(TForm)
    ListBox1: TListBox;
    Button1: TButton;
    Button2: TButton;
    Button3: TButton;
    procedure ListBox1Click(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure ListBox1DblClick(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form2: TForm2;

implementation

uses Bs;

{$R *.dfm}

procedure TForm2.ListBox1Click(Sender: TObject);
begin
 if ListBox1.ItemIndex=-1
 then Form2.Caption:='Открыть/Удалить'+' - пусто'
 else Form2.Caption:='Открыть/Удалить'+' - '+ListBox1.Items[ListBox1.ItemIndex];
end;

procedure TForm2.ListBox1DblClick(Sender: TObject);
begin
 if ListBox1.ItemIndex=-1
 then Form2.Caption:='Открыть/Удалить'+' - пусто'
 else Form2.Caption:='Открыть/Удалить'+' - '+ListBox1.Items[ListBox1.ItemIndex];
 ModalResult:=MrOk;
end;

procedure TForm2.Button1Click(Sender: TObject);
begin
 ModalResult:=MrOk;
end;

procedure TForm2.Button2Click(Sender: TObject);
begin
 Close;
end;

procedure TForm2.Button3Click(Sender: TObject);
var
 i,j: integer;
begin
 j:=-1;
 if MessageDlg('Вы действительно хотите удалить текущую запись?', mtConfirmation,
    [mbYes, mbNo],0)=mrNo
 then
 else
  begin
   for i:=0 to ListBox1.Items.Count-1 do
    begin
     if ListBox1.Selected[i]=true
     then
      begin
       j:=i;
       Break;
      end;
    end;
   if j<>-1
   then
    begin
     Form1.Table1.Open;
     Form1.Table1.First;
     //
     while not (Form1.Table1.Eof) do
      begin
       if Form1.Table1.Fields[0].AsString=ListBox1.Items.Strings[j]
       then
        begin
         Form1.Table1.Delete;
         Form1.Table1.Close;
         Form1.Table1.Open;
         ShowMessage('Запись успешно удалена!');
         Button2.Click;
         Close;
         Exit;
        end;
       Form1.Table1.Next;
      end;
    end
   else ShowMessage('Запись для удаления не выбрана!');
  end;
end;

end.
