unit AutoRenamerMain;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, ExtCtrls, FileCtrl, Registry, ShellAPI;

type { You MUST define/augment class TEdit BEFORE TForm1 because TForm1
 uses it. }
  TEdit = class(StdCtrls.TEdit)
    procedure WMDropFiles(var Message: TWMDropFiles); message WM_DROPFILES;
  end;

type
  TForm1 = class(TForm)
    Button1: TButton;
    Button2: TButton;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Edit1: TEdit;
    Edit2: TEdit;
    Edit3: TEdit;
    Edit4: TEdit;
    CheckBox1: TCheckBox;
    CheckBox2: TCheckBox;
    Timer1: TTimer;
    IconData : TNotifyIconData;
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure Timer1Timer(Sender: TObject);
    function FixPath(TheFilePath, TheNullPath: string): Boolean;
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure AppOnMinimize(Sender: TObject);
    procedure AppOnRestore(Sender: TObject);
    procedure ThisFormMinimise(const DoMinimize: Boolean);
    private
      { Private declarations }
      procedure WndProc(var Msg : TMessage); override;
    public
      { Public declarations }
    end;

var
  Form1: TForm1;
  SearchRec: TSearchRec;
  f: file;
  MyRegistry: TRegistry;
  s, t, u, FileDirectory, NullDirectory: string;
  n, o, p: Integer;
  q: Boolean; 

implementation

{$R *.DFM}

{ -- TEdit class augmentation -- }
procedure TEdit.WMDropFiles(var Message: TWMDropFiles);
  var
    c: integer;
    fn: array[0..MAX_PATH-1] of char;
  begin
    c := DragQueryFile(Message.Drop, $FFFFFFFF, fn, MAX_PATH);
    if c <> 1 then
    begin
      MessageBox(Handle, 'Too many files.', 'Drag and drop error', MB_ICONERROR);
      Exit;
    end;
    if DragQueryFile(Message.Drop, 0, fn, MAX_PATH) = 0 then Exit;
    Text := fn; { Text is a variable within class TEdit }
  end;
{ -- TEdit class augmentation ends-- }

procedure TForm1.WndProc(var Msg : TMessage);
begin
  case Msg.Msg of WM_USER + 1:
    case Msg.lParam of WM_LBUTTONDOWN:
    begin
      { The order of the following function calls is important. }
      ShowWindow(Application.Handle, SW_SHOW);
      Application.Restore; { restores from minimized }
    end
    end; { end case }
  end; { end case }
  inherited;
end;

procedure TForm1.FormCreate(Sender: TObject);
begin
  DragAcceptFiles(Edit3.Handle, True);
  DragAcceptFiles(Edit4.Handle, True);
  MyRegistry := TRegistry.Create; { Class constructor must be explicitly
   called to create object! }
  if MyRegistry.OpenKey('Software\DelphiAutoRenamer\', False) then
  begin
    FileDirectory := MyRegistry.ReadString('FileDirectory');
    NullDirectory := MyRegistry.ReadString('NullDirectory');
    q := MyRegistry.ReadBool('DelZeroSize');
    MyRegistry.CloseKey;
  end;
  if Length(FileDirectory) > 0 then Edit3.Text := FileDirectory;
  if Length(NullDirectory) > 0 then Edit4.Text := NullDirectory;
  CheckBox2.Checked := q;
  Application.OnMinimize := AppOnMinimize;
  Application.OnRestore := AppOnRestore;
end;

procedure TForm1.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  Shell_NotifyIcon(NIM_DELETE, @IconData);
  Application.ProcessMessages;
end;

procedure TForm1.AppOnMinimize(Sender: TObject);
begin
  if Timer1.Enabled = True then ThisFormMinimise(False);
end;

procedure TForm1.AppOnRestore(Sender: TObject);
begin
  if Timer1.Enabled = True then Shell_NotifyIcon(NIM_DELETE, @IconData);
end;

procedure TForm1.Button1Click(Sender: TObject);
begin
  if (n = 0) then
  begin
    MessageDlg('Please Set Starting No.', mtError, [mbOK], 0);
    Button2.SetFocus;
    Exit;
  end;
  u := Trim(Edit2.Text);
  FileDirectory := Trim(Edit3.Text);
  NullDirectory := Trim(Edit4.Text);
  if CheckBox1.Checked = True then Timer1.Interval := 500
  else Timer1.Interval := 2000;
  if Timer1.Enabled = False then
  begin
    if FixPath(FileDirectory, NullDirectory) = True then
      Exit; { True means bad path }
    Timer1.Enabled := True;
    Button1.Caption := 'Press To Stop';
    Button2.Enabled := False;
    CheckBox1.Enabled := False;
    ThisFormMinimise(True);
  end
  else if Timer1.Enabled = True then
  begin
    Timer1.Enabled := False;
    Button1.Caption := 'Press To Start';
    Button2.Enabled := True;
    Button2.SetFocus;
    CheckBox1.Enabled := True;
    //Edit1.Text := IntToStr(n); { not really helpful }
    n := 0;
  end;
end;

procedure TForm1.Button2Click(Sender: TObject);
begin
  try
    n := StrToInt(Edit1.Text);
    if n > 0 then
    begin
      MessageDlg('Will start at: ' + IntToStr(n), mtInformation, [mbOK], 0);
      Button1.SetFocus;
    end
    else
      Abort; { silent exception }
  except
    MessageDlg('Please Enter a Valid Number', mtError, [mbOK], 0);
    n := 0;
  end;
end;

procedure TForm1.Timer1Timer(Sender: TObject);
label BoRL;
begin
  o := FindFirst('.\*' + u, faAnyFile, SearchRec);
  if (o = 0) then
  begin
    try
      Timer1.Enabled := False;
      t := IntToStr(n);
      o := Length(t);
      if (o = 1) then t := '000' + t;
      if (o = 2) then t := '00' + t;
      if (o = 3) then t := '0' + t;
      t := t + '_1000' + u;
      s := SearchRec.Name;
      if SearchRec.Size > 0 then
      begin
        AssignFile(f, s);
        Rename(f, t);
        if RenameFile('.\' + t, FileDirectory + t) = False then
        begin
          Timer1.Enabled := False;
          MessageDlg('Could not move file! Check Path or Duplicates!', mtError, [mbOK], 0);
          FindClose(SearchRec);
          Application.Terminate;
        end;
        n := n + 1;
      end
      else
      begin
        AssignFile(f, s);
        o := 0;
        BoRL: { Beggining of Rename Loop }
        p := p + 1; o := o + 1;
        t := IntToStr(p) + '_bad' + u;
        Rename(f, t);
        if CheckBox2.Checked = True then
        begin
          q := DeleteFile(t);
          p := p - 1; { will stop incrementation if user presses checkbox while running }
          if q = False then
          begin
            MessageDlg('Error deleting file!', mtError, [mbOK], 0);
            FindClose(SearchRec);
            Application.Terminate;
          end;
        end
        else
        if RenameFile('.\' + t, NullDirectory + t) = False then
          if o < 250 then goto BoRL
          else
          begin
            Timer1.Enabled := False;
            MessageDlg('Error moving file!', mtError, [mbOK], 0);
            FindClose(SearchRec);
            Application.Terminate;
          end;
      end;
    finally
      FindClose(SearchRec);
      Timer1.Enabled := True;
    end;
  end;
end;

function TForm1.FixPath(TheFilePath, TheNullPath: string): Boolean;
begin
  FileDirectory := TheFilePath; { FileDirectory and TheFilePath might be
 different }
  NullDirectory := TheNullPath; { Ditto }
  if FileDirectory[Length(FileDirectory)] <> '\' then
    FileDirectory := FileDirectory + '\';
  if NullDirectory[Length(NullDirectory)] <> '\' then
    if DirectoryExists(NullDirectory) then
      NullDirectory := NullDirectory + '\';
  if (DirectoryExists(FileDirectory) = False) or
    ((DirectoryExists(NullDirectory) = False) and
    (CheckBox2.Checked = False))
  then
  begin
    MessageDlg('Destination Path is Invalid!', mtError, [mbOK], 0);
    Timer1.Enabled := False;
    Button1.Caption := 'Press To Start';
    Button2.Enabled := True;
    Button1.SetFocus;
    CheckBox1.Enabled := True;
    FixPath := True;
  end
  else
  begin
    Edit3.Text := FileDirectory;
    if CheckBox2.Checked = False then Edit4.Text := NullDirectory;
    if MyRegistry.OpenKey('Software\DelphiAutoRenamer\', True) then
    begin
      MyRegistry.WriteString('FileDirectory', FileDirectory);
      MyRegistry.WriteString('NullDirectory', NullDirectory);
      MyRegistry.WriteBool('DelZeroSize', CheckBox2.Checked);
      MyRegistry.CloseKey;
    end;
    FixPath := False;
  end;
end;

procedure TForm1.ThisFormMinimise(const DoMinimize: Boolean);
begin
  IconData.cbSize := sizeof(IconData);
  IconData.Wnd := Handle;
  IconData.uID := 100;
  IconData.uFlags := NIF_MESSAGE + NIF_ICON + NIF_TIP;
  IconData.uCallbackMessage := WM_USER + 1;
  IconData.hIcon := Application.Icon.Handle;
  StrPCopy(IconData.szTip, Application.Title);
  Shell_NotifyIcon(NIM_ADD, @IconData);
  { The order of the following function calls is important. }
  if DoMinimize then ShowWindow(Application.Handle, SW_MINIMIZE);
  ShowWindow(Application.Handle, SW_HIDE); { hides the bottom tile }
end;

end.
