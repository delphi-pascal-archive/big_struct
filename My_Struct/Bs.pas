unit Bs;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Menus, ExtCtrls, Word2000, OleServer, DB, DBTables,
  Mask, ComCtrls, Gauges, ActnMan, ActnColorMaps, WordXP;

type
  TForm1 = class(TForm)
    Edit1: TEdit;
    Edit2: TEdit;
    Edit3: TEdit;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    Label6: TLabel;
    Memo1: TMemo;
    Label7: TLabel;
    Memo2: TMemo;
    Edit4: TEdit;
    ComboBox1: TComboBox;
    Label8: TLabel;
    Label9: TLabel;
    Label10: TLabel;
    Edit5: TEdit;
    Edit6: TEdit;
    MainMenu1: TMainMenu;
    N1: TMenuItem;
    N2: TMenuItem;
    N4: TMenuItem;
    N5: TMenuItem;
    N6: TMenuItem;
    N7: TMenuItem;
    Bevel1: TBevel;
    Button1: TButton;
    DataSource1: TDataSource;
    Table1: TTable;
    Table2: TTable;
    SaveDialog1: TSaveDialog;
    WordApplication1: TWordApplication;
    WordDocument1: TWordDocument;
    Label11: TLabel;
    Edit7: TEdit;
    ME1: TMaskEdit;
    Gauge1: TGauge;
    procedure N7Click(Sender: TObject);
    procedure Edit5KeyPress(Sender: TObject; var Key: Char);
    procedure FormCreate(Sender: TObject);
    procedure N2Click(Sender: TObject);
    procedure N4Click(Sender: TObject);
    procedure Edit7KeyPress(Sender: TObject; var Key: Char);
    procedure Button1Click(Sender: TObject);
    procedure ComboBox1Select(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure N6Click(Sender: TObject);
    procedure Edit1KeyPress(Sender: TObject; var Key: Char);
  private
    { Private declarations }
  public
    procedure ClearAllFields(Sender: TObject);
    procedure LoadDocument(Const Name: String);
    procedure SaveDocumentAs(Name: String);
    procedure SaveDocument;
    procedure ReqOpen(Sender: TObject);
    procedure ReqLoad(Sender: TObject);
    procedure ReqSave(Sender: TObject);
    procedure LoadNames(Sender: TObject);
    { Public declarations }
  end;

var
  Form1: TForm1;
  DocName: string;
  SaveTo: String;

implementation

uses open;

{$R *.dfm}

function GetProgramPath : String;
begin
 GetProgramPath:=ExtractFilePath(ParamStr(0));
end;

function GetPath:String;
begin
 GetPath:=GetProgramPath+'shablon.doc';
end;

procedure TForm1.N2Click(Sender: TObject);
begin
 ClearAllFields(Sender); // очистка всех полей ввода
end;

procedure TForm1.N4Click(Sender: TObject);
begin
 if DocName=''
 then DocName:='Новый документ';
 SaveDocumentAs(DocName);
 if DocName='Новый документ'
 then DocName:='';
end;

procedure TForm1.N6Click(Sender: TObject);
begin
 ReqSave(Sender);
 DocName:='';
 LoadNames(Sender);
 case Form2.ShowModal of
  mrOk: LoadDocument(Form2.ListBox1.Items[Form2.ListBox1.ItemIndex]);
 end;
end;

procedure TForm1.Edit5KeyPress(Sender: TObject; var Key: Char);
begin
 if not (Key in [#0..#31,'0'..'9']) then Key:=#0;
end;

procedure TForm1.Edit1KeyPress(Sender: TObject; var Key: Char);
begin
 if not (Key in [#0..#31,'0'..'9']) then Key:=#0;
end;

procedure TForm1.FormCreate(Sender: TObject);
begin
 if not (DirectoryExists(GetProgramPath+'data')) then
  begin                      // если директория не существует, то
   ForceDirectories(GetProgramPath+'\data'); // создать директорию
  end;
 ReqOpen(Sender);
 ReqLoad(Sender);
end;

procedure TForm1.ReqOpen(Sender: TObject);
begin
 Table2.Active:=false;        // если таблицы не существует, то
 Table2.TableName:=GetProgramPath+'data\'+'Req.db';
 if not (FileExists(GetProgramPath+'data\'+'Req.db')) then
 if not Table2.Exists then
  begin                       // создаем новую базу
   with Table2.FieldDefs do   // название поля
    begin
     Clear;
     with AddFieldDef do      // добавление нового поля
      begin
       Name:='Name';          // имя
       DataType:=ftString;    // тип
       Size:=150;             // размер
      end;
     with AddFieldDef do
      begin
       Name:='Address';
       DataType:=ftString;
       Size:=150;
      end;
     with AddFieldDef do
      begin
       Name:='Telephone';
       DataType:=ftString;
       Size:=30;
      end;
     with AddFieldDef do
      begin
       Name:='Director';
       DataType:=ftString;
       Size:=150;
      end;
     Table2.CreateTable;      // создание таблицы
    end;
  end;
 Table2.Open;
end;

procedure TForm1.ReqLoad(Sender: TObject);
begin
 if not (Table2.Active) then
  begin
   ShowMessage('База с реквизитами не открыта');
   Exit;
  end;
 ComboBox1.Items.BeginUpdate;
 ComboBox1.Items.Clear;
 Table2.First;             // на начало таблицы
 while not Table2.Eof do   // до конца файла
  begin                    // добавляем в список поле с именем
   ComboBox1.Items.Add(Table2.Fields[0].AsString);
   Table2.Next;
  end;
  ComboBox1.Items.EndUpdate;
 Table2.Close;
end;

procedure TForm1.ReqSave(Sender: TObject);
begin
 if Trim(ComboBox1.Text)='' then Exit;
 ReqOpen(Sender);
 Table2.First;
 if not(Table2.Locate('Name',ComboBox1.Text,[LoCaseInsensitive])) then
  begin
   Table2.Last;
   Table2.Insert;
   Table2.FieldValues['Name']:=ComboBox1.Text;
   Table2.FieldValues['Address']:=Edit3.Text;
   Table2.FieldValues['Telephone']:=Edit4.Text;
   Table2.FieldValues['Director']:=Edit7.Text;
   Table2.Post;
   ReqLoad(Sender);
  end;
 Table2.Close;
end;

procedure TForm1.ComboBox1Select(Sender: TObject);
var
 i: longint;
 temp: TComponent;
begin                         // цикл по всем компонентам формы
 for i:=Form1.ComponentCount-1 downto 0 do
  begin
   temp:=Components[i];
   if (Temp is TEdit)
   then TEdit(Components[i]).Text:='';
   if (Temp is TComboBox)
   then TComboBox(Components[0]).Text:='';
   if Temp is TMemo
   then TMemo(Components[i]).Lines.Clear;
   if Temp is TMaskEdit
   then TMaskEdit(Components[i]).Text:='';
  end;
 SaveTo:='';
 if Trim(ComboBox1.Text)='' then Exit;
 ReqOpen(Sender);
 Table2.First;
 if Table2.Locate('Name',ComboBox1.Text,[]) then
  begin
   Edit3.Text:=Table2.FieldValues['Address'];
   Edit4.Text:=Table2.FieldValues['Telephone'];
   Edit7.Text:=Table2.FieldValues['Director'];
   Exit;
  end;
end;

procedure TForm1.SaveDocumentAs(Name: string);
 function FieldExists(name: string):Boolean;
 var
  i: integer;
 begin
  FieldExists:=false;
  for i:=0 to Form2.ListBox1.Items.Count-1 do
   begin
    if Name=UpperCase(Form2.ListBox1.Items[i]) then
     begin
      FieldExists:=true;
      Exit;
     end;
   end;
 end;
Label 1;
var
 st: string;
 i: Integer;
 Temp: TComponent;
 Buf: array [0..150] of Char;
begin
 Table1.Active:=false;
 Table1.TableName:=GetProgramPath+'data\'+'AllReq.db';
 if not (FileExists(GetProgramPath+'data\'+'AllReq.db')) then
 begin // создание новой базы
  Table1.Active:=false;
  Table1.TableName:=GetProgramPath+'data\'+'AllReq.db';
  if not Table1.Exists then
   begin
    with Table1.FieldDefs do
     begin
      Clear;
      with AddFieldDef do
       begin
        Name:='DocName';
        DataType:=ftString;
        Size:=240;
       end;
      for i:=Form1.ComponentCount-1 downto 0 do
       begin
        Temp:=Components[i];
        if (Temp is TEdit) or (Temp is TComboBox) then
         with AddFieldDef do
          begin
           Name:=TControl(Temp).Name;
           DataType:=ftString;
           Size:=200;
          end;
        if (Temp is TMaskEdit) then
         with AddFieldDef do
          begin
           Name:=TControl(Temp).Name;
           DataType:=ftString;
           Size:=20;
          end;
        if (Temp Is TMemo) then
         with AddFieldDef do
          begin
           Name:=TControl(Temp).Name;
           DataType:=ftMemo;
           Size:=500;
          end;
        end;
       Table1.CreateTable;
     end;
    Table1.Open;
  end;
 end;

 Table1.Active:=true;
 Table1.Open;
 Table1.First;
 Form2.ListBox1.Items.BeginUpdate;
 Form2.ListBox1.Items.Clear;
  while not (Table1.Eof) do
   begin
    Form2.ListBox1.Items.Add(Table1.Fields[0].AsString);
    Table1.Next;
   end;
 Form2.ListBox1.Items.EndUpdate;
 st:=Name;
 1:
  st:=InputBox('Сохранение','Введите название:','');
   begin
    if FieldExists(Uppercase(st)) then
     begin
      StrPCopy(Buf,'Заменить существующий '+st);
      i:=Application.MessageBox(Buf,'',MB_YESNO);
      if i=IDYES then
       begin
        DocName:=st;
       end
      else
       goto 1;
     end
    else DocName:=st;
    SaveDocument;
   end;
end;

procedure TForm1.SaveDocument;
Label 2;
var
 i: longint;
 Temp: TComponent;
 NameExists: Boolean;
begin
 ReqSave(Self);
 if DocName='' then
  begin
   ShowMessage('Ошибка при записи документа '+''''''+DocName+'''''');
   Exit;
  end;
 NameExists:=false;
 Table1.Open;
 Table1.First;
 while not(Table1.Eof) do
  begin
   if (UpperCase(DocName)=UpperCase(Table1.Fields[0].AsString)) then
    begin
     NameExists:=True;
     goto 2;
    end;
   Table1.Next;
  end;
 2:
   if NameExists then Table1.Edit else Table1.Append;
   Table1.FieldValues['DocName']:=DocName;
   for i:=Form1.ComponentCount-1 downto 0 do
    begin
     Temp:=Components[i];
      if (Temp is TEdit) or (Temp is TComboBox) then
       begin
        Table1.FieldValues[TControl(Temp).Name]:=TCustomEdit(Temp).Text;
       end;
      if (Temp is TMemo) then
       begin
        Table1.FieldValues[TControl(Temp).Name]:=TMemo(Temp).Text;
       end;
    end;
 Table1.Post;
end;

procedure TForm1.LoadNames(Sender: TObject);
begin
 Table1.Active:=False;
 Table1.TableName:=GetProgramPath+'data\'+'AllReq.db';
 Table1.Open;
 if Table1.Exists then
  begin
   Form2.ListBox1.Items.Clear;
   Form2.ListBox1.Items.BeginUpdate;
   Table1.First;
   while not(Table1.Eof) do
    begin
     Form2.ListBox1.Items.Add(Table1.Fields[0].AsString);
     Table1.Next;
    end;
   Form2.ListBox1.Items.EndUpdate;
  end;
 Table1.Close;
end;

procedure TForm1.LoadDocument(Const name: string);
var
 i: integer;
 temp: TComponent;
begin
 Table1.Open;
 Table1.First;
 while not (Table1.Eof) do
  begin
   if UpperCase(name)=UpperCase(Table1.Fields[0].AsString) then
    begin
     for i:=Form1.ComponentCount-1 downto 0 do
      begin
       Temp:=Components[i];
       if (Temp is TEdit) or (Temp is TComboBox)
       then TCustomEdit(Temp).Text:=Table1.Fieldbyname(TControl(Temp).Name).AsString;
       if (Temp is TMaskEdit)
       then TMaskEdit(Temp).Text:=Table1.Fieldbyname(TControl(Temp).Name).AsString;
       if (Temp Is TMemo)
       then TMemo(Temp).Text:=Table1.Fieldbyname(TControl(Temp).Name).AsString;
      end;
     DocName:=Name;
     Table1.Close;
     Exit; // выход
    end;
   Table1.Next;
  end;
 ShowMessage('Неверное имя документа или база данных повреждена!');
 if Table1.Active then Table1.Close;
end;

procedure TForm1.ClearAllFields(Sender: TObject);
var
 i: Longint;
 temp: TComponent;
begin                        // цикл по всем компонентам формы
 for i:=Form1.ComponentCount-1 downto 0 do
  begin
   temp:=Components[i];
   if (Temp is TEdit)
   then TEdit(Components[i]).Text:='';
   if (Temp is TComboBox)
   then TComboBox(Components[i]).Text:='';
   if Temp is TMemo
   then TMemo(Components[i]).Lines.Clear;
   if Temp is TMaskEdit
   then TMaskEdit(Components[i]).Text:='';
  end;
 SaveTo:='';
end;

procedure TForm1.Edit7KeyPress(Sender: TObject; var Key: Char);
begin
 if not (Key in [#0..#31,'0'..'9']) then Key:=#0;
end;

procedure TForm1.Button1Click(Sender: TObject);
var
 DocName, FN, gp, Psw, PswTmp, Revert, WritePsw, WritePswTmp, Fmt, Index: OleVariant;
 _Prod, Prod,ConfConv,ReadOnly,AddToRecFiles,replace, FalseParam: OleVariant;
 TrueParam: OleVariant;
 Wrap     : OleVariant;
 i1       : OleVariant; // Temporaly Var <-|
 SP       : OleVariant; //  Save Prompt  <-|
 EP       : OleVariant; //  Empty param  <-|
 i        : integer;
 EmptyParam: OleVariant; // Пустой параметр
 S,S1     : string;
begin
 DocName:=GetProgramPath+'shablon.doc';
 ConfConv:=true;
 ReadOnly:=False;
 AddToRecFiles:=False;
 Psw:='';
 PswTmp:='';
 Revert:=False;
 WritePsw:='False';
 WritePswTmp:='False';
 Fmt:=WdOpenFormatAuto;

 FalseParam:=False;
 TrueParam:=True;
 Wrap:=WdFindContinue;

 WordApplication1.Connect;
 WordApplication1.Documents.OpenOld(DocName, ConfConv, ReadOnly, AddToRecFiles, Psw, PswTmp, Revert, WritePsw, WritePswTmp, Fmt);
 WordDocument1.ConnectTo(WordApplication1.ActiveDocument);

 Replace:=WdReplaceAll;
 Index:=0;

 FN:=InputBox('Сохранение','Введите имя файла:','xXx.doc');

 for i:=1 to WordApplication1.Documents.Count do
  begin
   i1:=i;
   S:=FN;
   if Pos('.',S)<>0
   then S:=Copy(S,1,Pos('.',S));
   S:=S+'doc';
   S1:=WordApplication1.Documents.Item(I1).Name;
   S:=ExtractFileName(S);
   S1:=ExtractFileName(S1);  // усли имена файлов одинаковы
   if UpperCase(S1)=UpperCase(S) then // то файл закрывается
    begin
     WordApplication1.Visible:=True;
     SP:=wdPromptToSaveChanges;
     EP:=wdOriginalDocumentFormat;
     WordApplication1.Documents.Item(i1).Close(SP,EP,EP);
    end;
  end;
 gp:=GetProgramPath+FN; // сохраняем шаблон под именем new_1.doc
 WordDocument1.SaveAs(gp);

////////////////////////// Заявка /////////////////////////////////

 Gauge1.Progress:=Gauge1.Progress+9;
 _Prod:='Edit1';
 Prod:=Edit1.Text;
 WordDocument1.Range.Find.ExecuteOld(_Prod, EmptyParam, EmptyParam, EmptyParam,
     EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam, Prod, replace);

 Gauge1.Progress:=Gauge1.Progress+9;
 _Prod:='Edit2';
 Prod:=Edit2.Text;
 WordDocument1.Range.Find.ExecuteOld(_Prod, EmptyParam, EmptyParam, EmptyParam,
     EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam, Prod, replace);

 Gauge1.Progress:=Gauge1.Progress+9;
 _Prod:='ComboBox1';
 Prod:=ComboBox1.Text;
 WordDocument1.Range.Find.ExecuteOld(_Prod, EmptyParam, EmptyParam, EmptyParam,
     EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam, Prod, replace);

 Gauge1.Progress:=Gauge1.Progress+9;
 _Prod:='Edit3';
 Prod:=Edit3.Text;
 WordDocument1.Range.Find.ExecuteOld(_Prod, EmptyParam, EmptyParam, EmptyParam,
     EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam, Prod, replace);

 Gauge1.Progress:=Gauge1.Progress+9;
 _Prod:='Edit4';
 Prod:=Edit4.Text;
 WordDocument1.Range.Find.ExecuteOld(_Prod, EmptyParam, EmptyParam, EmptyParam,
     EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam, Prod, replace);

 Gauge1.Progress:=Gauge1.Progress+9;
 _Prod:='ME1';
 Prod:=ME1.Text;
 WordDocument1.Range.Find.ExecuteOld(_Prod, EmptyParam, EmptyParam, EmptyParam,
     EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam, Prod, replace);

 Gauge1.Progress:=Gauge1.Progress+9;
 _Prod:='Memo1';
 Prod:='';
 for i:=0 to Memo1.Lines.Count-1 do
  begin
   Prod:=Prod+Memo1.Lines.Strings[i];
  end;
 WordDocument1.Range.Find.ExecuteOld(_Prod, EmptyParam, EmptyParam, EmptyParam,
     EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam, Prod, replace);

 Gauge1.Progress:=Gauge1.Progress+9;
 _Prod:='Memo2';
 Prod:='';
 for i:=0 to Memo2.Lines.Count-1 do
  begin
   Prod:=Prod+Memo2.Lines.Strings[i];
  end;
 WordDocument1.Range.Find.ExecuteOld(_Prod, EmptyParam, EmptyParam, EmptyParam,
     EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam, Prod, replace);

 Gauge1.Progress:=Gauge1.Progress+9;
 _Prod:='Edit5';
 Prod:=Edit5.Text;
 WordDocument1.Range.Find.ExecuteOld(_Prod, EmptyParam, EmptyParam, EmptyParam,
     EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam, Prod, replace);

 Gauge1.Progress:=Gauge1.Progress+9;
 _Prod:='Edit6';
 Prod:=Edit6.Text;
 WordDocument1.Range.Find.ExecuteOld(_Prod, EmptyParam, EmptyParam, EmptyParam,
     EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam, Prod, replace);

 Gauge1.Progress:=Gauge1.Progress+9;
 _Prod:='Edit7';
 Prod:=Edit7.Text;
 WordDocument1.Range.Find.ExecuteOld(_Prod, EmptyParam, EmptyParam, EmptyParam,
     EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam, Prod, replace);

 WordApplication1.Application.Visible:=True;
 WordApplication1.Disconnect;

 Gauge1.Progress:=0;
end;

procedure TForm1.FormClose(Sender: TObject; var Action: TCloseAction);
begin
 ReqSave(Sender);
end;

procedure TForm1.N7Click(Sender: TObject);
begin
 Close;
end;

end.
