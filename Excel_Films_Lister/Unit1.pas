unit Unit1;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls,  FileCtrl, ComCtrls, ShellCtrls, ComObj, registry,
  jpeg, ExtCtrls;

type
  TForm1 = class(TForm)
    Button1: TButton;
    ListBox1: TListBox;
    ShellTreeView1: TShellTreeView;
    Image1: TImage;
    Image2: TImage;
    Button2: TButton;
    procedure ShellTreeView1Change(Sender: TObject; Node: TTreeNode);
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
  private
    { D�clarations priv�es }
  public
    { D�clarations publiques }
    procedure WMNCHitTest(var msg: TWMNCHitTest); message WM_NCHITTEST;
  end;

var
  //index de fin du fichier excel (pour rajouter des fichiers � la fin du fichier excel)
  indfin : Integer;
  //index parcourant le nombre de film total dans la listbox pour les mettre
  //ensuite dans le fichier excel
  chiffre : Integer;
  Form1: TForm1;
  //Enregistrement pour recherche des fichiers
  F: TSearchRec;
  // entier � 0 si il reste des fichiers � parcourir
  n: Integer;
  //chaine de caract�re contenant l'extension du fichier courant
  ext: String;
  //var relatives � excel
  vXLWorkbook, vXLWorkbooks : variant;
  vWorksheet : variant;
  vCell : variant;
  aFileName : AnsiString;
  vMSExcel : variant;

implementation

{$R *.dfm}
//fonction retournant les differents dossiers sp�ciaux de windows
// suivant ce qu'on met en param�tre (par exemple avec Personal
//on obtient le chemin d'acces de Mes Documents)
//trouv� sur DelphiFR
//besoin car le fichier excel se trouve dans 'Mes Documents'
function GetSpecialFolder(folder:string) :string;
var
  Reg : TRegistry;
  res : string;

begin

try
  Reg := TRegistry.Create;
  Reg.RootKey := HKEY_CURRENT_USER;
  if Reg.OpenKey('\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders', False)
  then res := Reg.ReadString(folder)
  else res := '';
  finally
  Reg.CloseKey;
  Reg.Free;
  end;
  result := res;
end;

//Fonction permettant de lister les fichiers d'un chemin sp�cifi� en param�tre
// elle retourne aussi un entier d�nombrant les fichiers je pensais en avoir besoin pour plus tard
// ... en fait non :)
Function ListeFichiers(Chemin:String):Integer;
Var
  S:TSearchRec;
  fold : string;
Begin
  fold:='';
  //rajouter le '/' � la fin du chemin
  Chemin:=IncludeTrailingPathDelimiter(Chemin);
  Result:=0;
  // Recherche de la premi�re entr�e du r�pertoire
  If FindFirst(Chemin+'*.*',faAnyFile,S)=0
  Then Begin
    Repeat
      // Il faut absolument dans le cas d'une proc�dure r�cursive ignorer
      // les . et .. qui sont toujours plac�s en d�but de r�pertoire
      // sinon la proc�dure va boucler sur elle-m�me.
      If (S.Name<>'.')And(s.Name<>'..')
      Then Begin
        If (S.Attr And faDirectory)<>0
          // Dans le cas d'un sous-repertoire on appelle la m�me proc�dure
          Then Result:=Result+ListeFichiers(Chemin+S.FindData.cFileName)
          // Sinon on compte simplement le fichier
          Else
          begin
            //v�rification de l'extension
            ext:=ExtractFileExt(S.Name);
            //si l'extension correspond � un fichier video
            //  /!\ penser � chercher une fa�on plus simple de v�rifier �a /!\
            If ((ext = '.avi') or (ext = '.mkv') or (ext = '.flv') or (ext = '.mp4') or (ext = '.divx')) then
            begin
              fold:='';
	            //on l'ajoute � la liste
              // si on est dans un sous-dossier
              if (Chemin <> IncludeTrailingPathDelimiter(Form1.ShellTreeView1.Path))then
                //rajouter le nom de ce sous-dossier avant le fichier video dans la Listbox
                fold := ExtractFileName( ExtractFileDir(Chemin+S.FindData.cFileName)) + ' | ';
	            Form1.ListBox1.Items.Add(fold+ChangeFileExt(S.Name, ''));
            end;
          end;
      End;
    // Recherche du suivant
    Until FindNext(S)<>0;
    FindClose(S);
  End;
End;

// pour plus tard : penser � faire une proc�dure lan�ant le fichier video lors d'un double clic dans la listbox

//proc�dure execut�e � chaque changement dans l'arborescence (ex : clic sur nouveau dossier)
procedure TForm1.ShellTreeView1Change(Sender: TObject; Node: TTreeNode);
begin
  //effacement de la listbox
  ListBox1.Clear;
  //listing des films contenus dans le repertoire point� pas le ShellTreeView
  ListeFichiers(ShellTreeView1.Path);
end;

procedure TForm1.Button1Click(Sender: TObject);
begin
  //ouvrir instance d'excel
  vMSExcel := CreateOleObject('Excel.Application');
  vMSExcel.Visible := true;
  //ouvrir le fichier Listedefilms.xls se trouvant dans mes documents
  aFileName := 'Listedefilms.xls';
  vXLWorkbooks := vMSExcel.Workbooks;
  // si le fichier n'existe pas, le creer
  if not (FileExists(aFileName)) then
  begin
    vXLWorkbook := vXLWorkbooks.Add;
    vXLWorkbook.SaveAs(aFileName);
    vXLWorkbook.Close(true, aFileName);
  end;
  //dans tous les cas, l'ouvrir
  vXLWorkbook := vXLWorkbooks.Open(aFileName);
  //acceder � la Feuil1
  vWorksheet := vXLWorkbook.Worksheets['����1'];
  vCell := vWorksheet.Range['A1'];
  //formatage excel
  //titre sur la premi�re cellule
  vCell.Value:='Nom des films : ';
  vWorksheet.Range['A1'].Borders.LineStyle:=true;
  indfin :=1;
  //recherche de la premi�re cellule de vide
  while vWorksheet.Range['A'+IntToStr(indfin)].Value<>'' do
  begin
    indfin:=indfin+1;
  end;
  //insertion des films
  for chiffre := 0 to (ListBox1.Items.Count -1) do
  begin
    vWorksheet.Range['A'+IntToStr(chiffre+indfin)].Value:=ListBox1.Items[chiffre];
  end;
  //redimensionne la colone en fonction de la taille du plus grand nom de fichier
  vWorksheet.Range['A1','A'+IntToStr(chiffre+indfin)].Columns.AutoFit;
  //sauvegarde du fichier
  vXLWorkbook.Save;
  //fermeture d'excel
  vMSExcel.Quit;
  vMSExcel := unassigned;
end;

// proc�dure permettant de bouger la fenetre en cliquant sur le fond de la fenetre n'importe ou
//trouv� sur DelphiFR !
procedure TForm1.WMNCHitTest(var msg: TWMNCHittest);
var
  pt: TPoint;
begin
  inherited;
  pt:= ScreenToClient(Point(msg.XPos, msg.YPos));
  if PtInRect(Rect(0, 0, ClientWidth, ClientHeight), pt) then
    msg.Result:= HTCAPTION;
end;

// Lors du clic sur le bouton fermer.... on ferme la f�n�tre !

procedure TForm1.Button2Click(Sender: TObject);
begin
  Close;
end;

end.
