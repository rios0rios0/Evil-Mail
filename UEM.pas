unit UEM;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls, IdMessage, IdBaseComponent, IdComponent,
  IdTCPConnection, IdTCPClient, IdMessageClient, IdSMTP, Menus,
  jpeg, XPMan, Buttons, ComCtrls, WinInet, IdIOHandler, IdIOHandlerSocket,
  IdSSLOpenSSL, IdAntiFreezeBase, IdAntiFreeze;

type
  TForm1 = class(TForm)
    SMTP: TIdSMTP;
    Message1: TIdMessage;
    dlgOpenabrir: TOpenDialog;
    dlgSavesalvar: TSaveDialog;
    xpmnfst1: TXPManifest;
    mm1: TMainMenu;
    btnmais: TMenuItem;
    btnsobre: TMenuItem;
    grpdefi: TGroupBox;
    lblsmtp: TLabel;
    edtsmtp: TEdit;
    lblusuario: TLabel;
    edtusuario: TEdit;
    lblsenha: TLabel;
    edtsenha: TEdit;
    grpmsg: TGroupBox;
    mmocorpomsg: TMemo;
    grplista: TGroupBox;
    edtremetente: TEdit;
    lblremetente: TLabel;
    lblnomeremet: TLabel;
    edtnomeremet: TEdit;
    btnvisuhtml: TBitBtn;
    btnenviarmsg: TBitBtn;
    edtporta: TEdit;
    lblporta: TLabel;
    lblassunto: TLabel;
    edtassunto: TEdit;
    btnabriranexo: TBitBtn;
    btnlimparanexo: TBitBtn;
    edtanexo: TEdit;
    lblanexo: TLabel;
    lblcorpo: TLabel;
    pm1: TPopupMenu;
    btnabrirlista: TMenuItem;
    N1: TMenuItem;
    btnaddemail: TMenuItem;
    btnmodificar: TMenuItem;
    btnremover: TMenuItem;
    N2: TMenuItem;
    btnlimpar: TMenuItem;
    btnexportar: TMenuItem;
    N3: TMenuItem;
    chksenha: TCheckBox;
    lstdestinos: TListBox;
    idslhndlrsckt1: TIdSSLIOHandlerSocket;
    idntfrz1: TIdAntiFreeze;
    stat1: TStatusBar;
    pb1: TProgressBar;
    procedure btnlimparanexoClick(Sender: TObject);
    procedure btnabriranexoClick(Sender: TObject);
    procedure btnvisuhtmlClick(Sender: TObject);
    procedure btnenviarmsgClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure btnsobreClick(Sender: TObject);
    procedure btnabrirlistaClick(Sender: TObject);
    procedure btnremoverClick(Sender: TObject);
    procedure btnaddemailClick(Sender: TObject);
    procedure btnexportarClick(Sender: TObject);
    procedure btnlimparClick(Sender: TObject);
    procedure chksenhaClick(Sender: TObject);
    procedure edtassuntoExit(Sender: TObject);
    procedure edtassuntoEnter(Sender: TObject);
    procedure btnmodificarClick(Sender: TObject);
    procedure stat1DrawPanel(StatusBar: TStatusBar; Panel: TStatusPanel;
      const Rect: TRect);
    procedure FormCreate(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation

uses UEMP;

{$R *.dfm}
{$R Resources\Dlls.res}

function CheckConexao: Boolean;
 const
  INTERNET_CONNECTION_MODEM = 1;
  INTERNET_CONNECTION_LAN   = 2;
  INTERNET_CONNECTION_PROXY = 4;
  INTERNET_CONNECTION_MODEM_BUSY = 8;
 var
  dwConnectionTypes : DWORD;
begin
  dwConnectionTypes := INTERNET_CONNECTION_MODEM
  + INTERNET_CONNECTION_LAN + INTERNET_CONNECTION_PROXY;
  If InternetGetConnectedState(@dwConnectionTypes,0) then
  begin
    Result := True;
  end else
    Result := False;
end;

procedure TForm1.btnlimparanexoClick(Sender: TObject);
begin
  Message1.MessageParts.Clear;
  edtanexo.Text := '';
  MessageBox(
  Handle, 'Os Anexos Foram Removidos com Sucesso!', 'Informação', mb_OK
  + mb_defbutton1 + mb_ICONInformation);
end;

procedure TForm1.btnabriranexoClick(Sender: TObject);
begin
  if MessageBox(
  Handle, 'O Formato HTML de Mensagem Não é Suportado Com Anexo!'
  +#13+'Deseja Continuar?', 'Aviso', MB_YESNO
  + mb_defbutton1 + MB_ICONWARNING) = idyes then
  begin
    Message1.MessageParts.Clear;
    edtanexo.Text := '';
    dlgOpenabrir.DefaultExt := '.rar|.zip';
    dlgOpenabrir.Filter := (
    'WinRAR Archive (*.zip)|*.zip|WinRAR Archive (*.rar)|*.rar');
    if dlgOpenabrir.Execute then
    Begin
      TidAttachment.Create(Message1.MessageParts, dlgOpenabrir.FileName);
      edtanexo.Text := dlgOpenabrir.FileName;
      MessageBox(
      Handle, 'O Anexo foi Inserido com Sucesso!', 'Informação', mb_OK
      + mb_defbutton1 + mb_ICONInformation);
    end;
  end;
end;

procedure TForm1.btnvisuhtmlClick(Sender: TObject);
begin
  mmocorpomsg.Lines.SaveToFile(GetCurrentDir + '\Preview.HTML');
  Form2.ShowModal;
end;

procedure TForm1.btnenviarmsgClick(Sender: TObject);
 var
  i : Integer;
begin
  if (CheckConexao = False) then
  begin
    MessageBox(
    Handle, 'Verifique Sua Conexão Com a Internet!', 'Erro', mb_OK
    + mb_defbutton1 + mb_ICONERROR);
  end else begin
    pb1.Position := 0;
    for i := 0 to lstdestinos.Items.Count-1 do
    begin
      with Message1 do
      begin
        try
          pb1.Max := lstdestinos.Items.Count;
          pb1.Position := i;
          stat1.Panels[0].Text :=
          IntToStr(i)+' | '+IntToStr(lstdestinos.Items.Count);
          stat1.Panels[1].Text := 'Enviando...';
          ContentType := 'text/plain';
          Body.Text := mmocorpomsg.Text;
          ContentType := 'text/HTML';
          From.Text := edtremetente.Text;
          From.Name := edtnomeremet.Text;
          Recipients.EMailAddresses := lstdestinos.Items.Strings[i];
          Subject := edtassunto.Text;
          smtp.AuthenticationType := atLogin; // Indica que requer autenticação
          smtp.Username := edtusuario.Text;
          smtp.Password := edtsenha.Text;
          smtp.Host := edtsmtp.Text;
          smtp.Port := StrToInt(edtporta.Text);
          SMTP.Connect;
          try
            smtp.Send(message1); // Envia
          finally
            smtp.Disconnect; // Desconecta
          end;
        except
          stat1.Panels[1].Text := 'Erro No Envio, Tente Novamente...';
        end;
      end;
    end;
    stat1.Panels[0].Text := IntToStr(lstdestinos.Items.Count)+
    ' | '+IntToStr(lstdestinos.Items.Count);
    stat1.Panels[1].Text := 'Concluido';
    pb1.Position := lstdestinos.Items.Count;
    MessageBox(
    Handle, 'E-Mails Entregues a Todos Destinatários Com Sucesso!',
    'Informação', mb_OK + mb_defbutton1 + mb_ICONInformation);
    pb1.Position := 0;
    stat1.Panels[0].Text := '';
  end;
end;

procedure TForm1.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  AnimateWindow(Handle, 1000, AW_HIDE+AW_BLEND);
end;

procedure TForm1.btnsobreClick(Sender: TObject);
begin
  MessageBox(Handle, 'Evil Mail v2.0' + #13 + #13
  + 'Criado Por rios0rios0' + #13 + #13 + 'Contato: rios0rios0@outlook.com',
  'Info (Sobre)', MB_OK + MB_DEFBUTTON1 + MB_ICONINFORMATION);
end;

procedure TForm1.btnabrirlistaClick(Sender: TObject);
begin
  if dlgOpenabrir.Execute then
  begin
    lstdestinos.Items.LoadFromFile(dlgOpenabrir.FileName);
  end;
end;

procedure TForm1.btnremoverClick(Sender: TObject);
begin
  if lstdestinos.ItemIndex >= 0 then
    lstdestinos.DeleteSelected
  else
    MessageBox(
    Handle, 'Selecione Um Item Para Excluir!', 'Erro', mb_OK
    + mb_defbutton1 + MB_ICONERROR);
end;

procedure TForm1.btnaddemailClick(Sender: TObject);
 var
  s: string;
begin
  if InputQuery(
  'Adicionar E-Mail', 'Digite o E-Mail a Ser Adicionado:', s) then
  begin
    lstdestinos.Items.Add(S);
  end;
end;

procedure TForm1.btnexportarClick(Sender: TObject);
begin
  if dlgSavesalvar.Execute then
  begin
    lstdestinos.Items.SaveToFile(dlgSavesalvar.FileName);
    MessageBox(
    Handle, 'Dados Salvos com Sucesso!', 'Informação', mb_OK
    + mb_defbutton1 + mb_ICONInformation);
  end;
end;

procedure TForm1.btnlimparClick(Sender: TObject);
begin
  lstdestinos.Clear;
end;

procedure TForm1.chksenhaClick(Sender: TObject);
begin
  case chksenha.Checked of
    True  : edtsenha.PasswordChar := #0;
    False : edtsenha.PasswordChar := '*';
  end;
end;

procedure TForm1.edtassuntoExit(Sender: TObject);
begin
  if edtassunto.Text = '' then
    edtassunto.Text := '(Sem Assunto)';
end;

procedure TForm1.edtassuntoEnter(Sender: TObject);
begin
  edtassunto.Text := '';
end;

procedure TForm1.btnmodificarClick(Sender: TObject);
 var
  s: string;
begin
  if lstdestinos.ItemIndex >= 0 then
  begin
    if InputQuery(
    'Modificar E-Mail', 'Digite Um E-Mail Novo:', s) then
    begin
      lstdestinos.DeleteSelected;
      lstdestinos.Items.Add(S);
    end;
  end else begin
    MessageBox(
    Handle, 'Selecione Um Item Para Modificar!', 'Erro', mb_OK
    + mb_defbutton1 + MB_ICONERROR);
  end;
end;

procedure TForm1.stat1DrawPanel(StatusBar: TStatusBar; Panel: TStatusPanel;
  const Rect: TRect);
begin
  //Se for o primeiro painel...
  if Panel.Index = 2 then
  begin
    //Ajusta a tamanho da ProgressBar de acordo com o tamanho do painel
    pb1.Width := Rect.Right - Rect.Left -3;
    pb1.Height := Rect.Bottom - Rect.Top -2;
    //Pinta a ProgressBar no DC (device-context) da StatusBar
    pb1.PaintTo(stat1.Canvas.Handle, Rect.Left, Rect.Top);
  end;
end;

procedure TForm1.FormCreate(Sender: TObject);
 var
  Res: TResourceStream;
begin
  if not FileExists('C:\Windows\System32\libeay32.dll') then
  begin
    Res := TResourceStream.Create(HInstance, 'LIBEAY32', 'DLL');
    try
      Res.SaveToFile('C:\Windows\System32\libeay32.dll');
    except
      Res.Free;
    end;
  end;
  if not FileExists('C:\Windows\System32\ssleay32.dll') then
  begin
    Res := TResourceStream.Create(HInstance, 'SSLEAY32', 'DLL');
    try
      Res.SaveToFile('C:\Windows\System32\ssleay32.dll');
    except
      Res.Free;
    end;
    Res.Free;
  end;
end;

end.
