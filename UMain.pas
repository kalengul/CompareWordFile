unit UMain;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, ComObj, Vcl.StdCtrls, Vcl.ExtCtrls;

type
  TFMain = class(TForm)
    Panel1: TPanel;
    RadioGroup1: TRadioGroup;
    OpenDialog: TOpenDialog;
    BtLoadSuos: TButton;
    Panel2: TPanel;
    MeProt: TMemo;
    Panel3: TPanel;
    Panel4: TPanel;
    Panel5: TPanel;
    Panel6: TPanel;
    Panel7: TPanel;
    MeFileShab: TMemo;
    MeFileSYOS: TMemo;
    Button1: TButton;
    LaNameFile: TLabel;
    LaNameShablon: TLabel;
    LaNameFGOS: TLabel;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    LaColOsh: TLabel;
    Label5: TLabel;
    LaGlava: TLabel;
    Label6: TLabel;
    Label7: TLabel;
    Label8: TLabel;
    LaGlavaSYOS: TLabel;
    LaGlavaFGOS: TLabel;
    BtSprFile: TButton;
    procedure FormActivate(Sender: TObject);
    procedure BtLoadSuosClick(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure BtSprFileClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

Procedure GoFoundWord(Word,WordShablon:variant; var aShablon,bShablon,a,b:Longword; TypeChange:byte);

var
  FMain: TFMain;
  Word,WordShablon,WordFGOS: Variant;   //Файлы Ворд СУОС, Шаблон СУОС, ФГОС
  FileNameOb,FileNameFGOS:string;       //Файлы Шаблона СУОС и ФГОС
  Excel:Variant;                        //Файл Эксель для хранения информации
  NomRowExcel,NomColExcel:Longword;     //Параметры файла Ексель
  CurrentDir:string;                    //Текущая папка
  Glava,GlavaShablon,GlavaFGOS:string;  //Текущая глава
  EndText,EndTextShablon:Longword;      //Размер фалов ВОрд
  St,StShablon:string;                  //Строки для сравнения
  Napravlenie:string;                   //Направление подготовки
  KolOsh:longword;                      //Счетчик ошибок в тексте
  NotFound,FGOSOpen:Boolean;            //Флаги

implementation

const
wdYellow=7;
wdRed = 6;
wdGreen = 11;

{$R *.dfm}

Function DeleteSpace(st:string):string;
begin
while Pos(Chr(7),st)<>0 do
  Delete(st,Pos(Chr(7),st),1);
while Pos(Chr(9),st)<>0 do
  Delete(st,Pos(Chr(9),st),1);
while Pos(Chr(11),st)<>0 do
  Delete(st,Pos(Chr(11),st),1);
while Pos(Chr(13),st)<>0 do
  Delete(st,Pos(Chr(13),st),1);
while Pos(Chr(12),st)<>0 do
  Delete(st,Pos(Chr(12),st),1);
while Pos(Chr(32),st)<>0 do
  Delete(st,Pos(Chr(32),st),1);
if (Pos('_Toc',st)<>0) or (Pos('PAGEREF',st)<>0)  or (Pos('HYPERLINK',st)<>0) or (Pos('.ru',st)<>0) then st:='';
result:=st;
end;

Function EndWord(st:string):boolean;
begin
result:= not((Length(st)>0) and(
             ((st[1]>=Chr(33)))));
//result:= (st=Chr(7)) or (st=Chr(9)) or (st=Chr(11)) or (st=Chr(13)) or (st=Chr(32));
end;

Function EndGlava(st:string; TypeGlava:byte):boolean;
begin
if (length(st)>3) and (st[1]>'0') and (St[1]<'9') and (st[2]='.') and (st[3]>'0') and (St[3]<'9')  then
  begin
  Case TypeGlava of
    0: Glava:=St;
    1: GlavaShablon:=St;
    2: GlavaFGOS:=st;
  End;
  result:=true;
  end
else
  result:=false;
end;

//Пройтись по сноскам
Procedure GoFotones;
var
  NomFotonesText,NomFotonesShablon:longword;
begin
//Проходим по всем сноскам документов (Footnotes object)
If (WordShablon.ActiveDocument.Footnotes.Count >0) and (Word.ActiveDocument.Footnotes.Count >0) Then
  begin

  end;
end;

Procedure GoSnoskaFGOS(var AFGOS,BFGOS:Longword);
var
  stFGOS:string;
  NomSnoskaFGOS:Longword;
begin
With FMain do
begin
stFGOS:=WordFGOS.ActiveDocument.Range(AFGOS,BFGOS).Text;
NomSnoskaFGOS:=0;
if stFGOS='--------------------------------' then
  begin
  //Это сноска
  MeProt.Lines.Add(GlavaFGOS+' Snoska FGOS '+IntToStr(AFGOS)+','+IntToStr(BFGOS));

  //Ищем номер сноски
  while EndWord(WordFGOS.ActiveDocument.Range(BFGOS,BFGOS+1).Text) do
    inc(BFGOS);
  AFGOS:=BFGOS;
  while not EndWord(WordFGOS.ActiveDocument.Range(BFGOS,BFGOS+1).Text) do
    inc(BFGOS);
  MeProt.Lines.Add(GlavaFGOS+' Nom Snoska FGOS !'+WordFGOS.ActiveDocument.Range(AFGOS,BFGOS).Text+'!'+IntToStr(AFGOS)+','+IntToStr(BFGOS));
  while EndWord(WordFGOS.ActiveDocument.Range(BFGOS,BFGOS+1).Text) do
    inc(BFGOS);
  while not EndWord(WordFGOS.ActiveDocument.Range(BFGOS,BFGOS+1).Text) do
    inc(BFGOS);
  AFGOS:=BFGOS;
  //Пробегаем по всей сноске до Enter
  while WordFGOS.ActiveDocument.Range(BFGOS,BFGOS+1).Text<>Chr(13) do
    inc(BFGOS);
  MeProt.Lines.Add(GlavaFGOS+' Text Snoska FGOS !'+WordFGOS.ActiveDocument.Range(AFGOS,BFGOS).Text+'!'+IntToStr(AFGOS)+','+IntToStr(BFGOS));
  while EndWord(WordFGOS.ActiveDocument.Range(BFGOS,BFGOS+1).Text) do
    inc(BFGOS);
  AFGOS:=BFGOS;
  while not EndWord(WordFGOS.ActiveDocument.Range(BFGOS,BFGOS+1).Text) do
    inc(BFGOS);
  MeProt.Lines.Add(GlavaFGOS+' Next Snoska FGOS !'+WordFGOS.ActiveDocument.Range(AFGOS,BFGOS).Text+'!'+IntToStr(AFGOS)+','+IntToStr(BFGOS));

  end;
end;
end;

Procedure SearchFgos(aShablon,bShablon,a,b:Longword; var AFGOS,BFGOS:Longword);
var
  aShDo,bShDo,bNext:Longword;
  stSh1,stSh2:string;
  stFgos1,stFgos2:string;
  i:byte;
begin
With FMain do
begin
//В ФГОС найти главу соответствующую главе из Шаблона
MeProt.Lines.Add(GlavaFGOS+' FGOS поиск главы '+GlavaShablon);
while (GlavaFGOS<>GlavaShablon) do
  begin
  //Ищем следующее нежелтое слово ФГОС
  AFGOS:=BFGOS;
  while not EndWord(WordFGOS.ActiveDocument.Range(BFGOS,BFGOS+1).Text) do
    inc(BFGOS);
  GoSnoskaFGOS(AFGOS,BFGOS);
  //Проверяем его на главу
  if EndGlava(WordFGOS.ActiveDocument.Range(AFGOS,BFGOS).Text,2) then
   MeProt.Lines.Add(GlavaFGOS+' FGOS !'+WordFGOS.ActiveDocument.Range(AFGOS,BFGOS).Text+'!');
  while EndWord(WordFGOS.ActiveDocument.Range(BFGOS,BFGOS+1).Text) do
    inc(BFGOS);
  end;
//Найти 2 слова до в Шаблоне
aShDo:=aShablon;
while EndWord(WordShablon.ActiveDocument.Range(aShDo-1,aShDo).Text) do
  dec(aShDo);
bShDo:=aShDo;
while not EndWord(WordShablon.ActiveDocument.Range(aShDo-1,aShDo).Text) do
  dec(aShDo);
stSh1:=WordShablon.ActiveDocument.Range(aShDo,bShDo).Text;
while EndWord(WordShablon.ActiveDocument.Range(aShDo-1,aShDo).Text) do
  dec(aShDo);
bShDo:=aShDo;
while not EndWord(WordShablon.ActiveDocument.Range(aShDo-1,aShDo).Text) do
  dec(aShDo);
stSh2:=WordShablon.ActiveDocument.Range(aShDo,bShDo).Text;
//Найти 2 слова до в ФГОС
AFGOS:=BFGOS-1;
stFgos1:=GlavaFGOS;
while not (EndGlava(WordFGOS.ActiveDocument.Range(AFGOS,BFGOS).Text,2))and
      not ((stSh1=stFgos1) and (stSh2=stFgos2)) do
  begin
  stFgos2:=stFgos1;
  AFGOS:=BFGOS;
  while not EndWord(WordFGOS.ActiveDocument.Range(BFGOS,BFGOS+1).Text) do
    inc(BFGOS);
  GoSnoskaFGOS(AFGOS,BFGOS);
  stFgos1:=WordFGOS.ActiveDocument.Range(AFGOS,BFGOS).Text;
  while EndWord(WordFGOS.ActiveDocument.Range(BFGOS,BFGOS+1).Text) do
    inc(BFGOS);
  MeProt.Lines.Add(GlavaFGOS+' FGOS !'+stFgos1+'!'+stFgos2+'! Shablon !'+stSh1+'!'+stSh2+'!');
  end;

AFGOS:=BFGOS;
while not EndWord(WordFGOS.ActiveDocument.Range(BFGOS,BFGOS+1).Text) do
  inc(BFGOS);
//Сравнить то, что во ФГОС и сделанный СУОС
NotFound:=true;
bNext:=b;
while (Word.ActiveDocument.Range(bNext,bNext+1).HighlightColorIndex=wdYellow) or (Word.ActiveDocument.Range(bNext,bNext+1).HighlightColorIndex=wdRed) do
  begin
  GoSnoskaFGOS(AFGOS,BFGOS);
  MeProt.Lines.Add(GlavaFGOS+' FGOS !AFGOS='+IntToStr(AFGOS)+'!'+IntToStr(BFGOS)+'! a=!'+IntToStr(a)+'!'+IntToStr(b)+'!');
  Word.ActiveDocument.Range(a,b).HighlightColorIndex:=wdGreen;
  GoFoundWord(Word,WordFGOS,AFGOS,BFGOS,a,b,1);

  bNext:=b;
  while EndWord(Word.ActiveDocument.Range(bNext,bNext+1).Text) do
    inc(bNext);
  if not NotFound  then
    repeat
    AFGOS:=BFGOS+1;
    inc(BFGOS);
    while not(EndWord(WordFGOS.ActiveDocument.Range(BFGOS,BFGOS+1).Text)) do
      inc(BFGOS);
    until not ((WordFGOS.ActiveDocument.Range(AFGOS,BFGOS).Text='') or (EndWord(WordFGOS.ActiveDocument.Range(AFGOS,BFGOS).Text)))
  else
    begin
//    for i := 1 to 2 do
    begin
    while EndWord(Word.ActiveDocument.Range(bNext,bNext+1).Text) do
      inc(bNext);
    while not EndWord(Word.ActiveDocument.Range(bNext,bNext+1).Text) do
      inc(bNext);
    while EndWord(Word.ActiveDocument.Range(bNext,bNext+1).Text) do
      inc(bNext);
    end;
  end;
  MeProt.Lines.Add(GlavaFGOS+' FGOS1!AFGOS='+IntToStr(AFGOS)+'!'+IntToStr(BFGOS)+'! a=!'+IntToStr(a)+'!'+IntToStr(b)+'!'+IntToStr(bNext)+'!'+IntToStr(Word.ActiveDocument.Range(bNext,bNext+1).HighlightColorIndex));
  end;
GoSnoskaFGOS(AFGOS,BFGOS);
MeProt.Lines.Add(GlavaFGOS+' FGOS2!AFGOS='+IntToStr(AFGOS)+'!'+IntToStr(BFGOS)+'! a=!'+IntToStr(a)+'!'+IntToStr(b)+'!');
Word.ActiveDocument.Range(a,b).HighlightColorIndex:=wdGreen;
GoFoundWord(Word,WordFGOS,AFGOS,BFGOS,a,b,1);
end;
end;

//Пройтись по желтому тексту
Procedure GoYellow(var aShablon,bShablon,a,b,AFGOS,BFGOS:Longword);
var
  YellowStShablon,YellowSt:string;  //Что храниться в желтом тексте
  an,b2n,bn,az,bz,aShablonYe,bShablonYe:Longword;             //Дополнительные указатели
  SR: TSearchRec;   // поисковая переменная
  FindRes: Integer; // переменная для записи результата поиска
  stOld,StShablonOld,stExcel:string;
begin
With FMain do
begin
//Проходим по всему желтому
inc(bShablon);
while WordShablon.ActiveDocument.Range(bShablon,bShablon+1).HighlightColorIndex=wdYellow do
  inc(bShablon);
while EndWord(WordShablon.ActiveDocument.Range(bShablon,bShablon+1).Text) do
  inc(bShablon);
//сохраняем в эксель данные из отмеченных желтым словам
YellowStShablon:=WordShablon.ActiveDocument.Range(aShablon,bShablon).Text;
Excel.Cells[NomRowExcel,1]:=YellowStShablon;
MeFileShab.Lines.Add(YellowStShablon);
aShablonYe:=aShablon; bShablonYe:=bShablon;

//Ищем следующее нежелтое слово в ШАблоне
aShablon:=bShablon;
while not EndWord(WordShablon.ActiveDocument.Range(bShablon,bShablon+1).Text) do
  inc(bShablon);
StShablon:=DeleteSpace(WordShablon.ActiveDocument.Range(aShablon,bShablon).Text);
{while EndWord(WordShablon.ActiveDocument.Range(bShablon,bShablon+1).Text) do
  inc(bShablon);}
//Ищем такое слово в тексте
MeFileShab.Lines.Add(GlavaShablon+' Sh---!'+StShablon+'!');
St:=DeleteSpace(Word.ActiveDocument.Range(a,b).Text);
MeFileSYOS.Lines.Add(Glava+' Sh1---!'+St+'!'+IntToStr(a)+','+IntToStr(b));
while EndWord(Word.ActiveDocument.Range(b,b+1).Text) do
  inc(b);
az:=a;
bz:=b;
//    repeat
b2n:=b;
a:=az; b:=bz;
an:=a; bn:=b;
while (b<EndText) and (St<>StShablon) and (not EndGlava(st,0)) do
  begin
  bn:=b;
  a:=b;
  while EndWord(Word.ActiveDocument.Range(b,b+1).Text) do
    inc(b);
  while not(EndWord(Word.ActiveDocument.Range(b,b+1).Text)) do
    inc(b);
  St:=DeleteSpace(Word.ActiveDocument.Range(a,b).Text);
  MeFileSYOS.Lines.Add(Glava+' Sh---!'+St+'!'+IntToStr(a)+','+IntToStr(b));
{  while EndWord(Word.ActiveDocument.Range(b,b+1).Text) do
    inc(b); }
  end;

if (bn<EndText) then
  begin
  MeFileSYOS.Lines.Add(Word.ActiveDocument.Range(an,bn).Text);
  Word.ActiveDocument.Range(an,bn).HighlightColorIndex:=wdYellow;
  end;
MeProt.Lines.Add(Glava+' End= !'+St+'!'+IntToStr(a)+','+IntToStr(b)+'-!'+StShablon+'!');
  //Ищем начало в ФГОС
if (FGOSOpen) and (Pos('ФГОС ВО',YellowStShablon)<>0) then
  begin
  MeProt.Lines.Add('End= !a='+IntToStr(a)+','+IntToStr(b)+'-!aShablon='+IntToStr(aShablon)+','+IntToStr(bShablon)+'!');
  stOld:=St; StShablonOld:=StShablon;
  SearchFgos(aShablonYe,bShablonYe,an,b2n-1,AFGOS,BFGOS);
  St:=stOld; StShablon:=StShablonOld;
  MeProt.Lines.Add('End= !'+St+'!'+IntToStr(a)+','+IntToStr(b)+'-!'+StShablon+'!');
  MeProt.Lines.Add('End= !a='+IntToStr(a)+','+IntToStr(b)+'-!aShablon='+IntToStr(aShablon)+','+IntToStr(bShablon)+'!');
//  dec(bShablon);
  end;

YellowSt:=Word.ActiveDocument.Range(an,bn).Text;
Excel.Cells[NomRowExcel,NomColExcel]:=YellowSt;
inc(NomRowExcel);

if Napravlenie='' then
  begin
  Napravlenie:=YellowSt; //Запомнить направление
  MeProt.Lines.Add('Napravlenie:'+Napravlenie);
  //Найти столбец в EXCEL файле
  NomColExcel:=2;
  stExcel:=Excel.Cells[1,NomColExcel];
  while (stExcel<>'') and (stExcel<>Napravlenie) do
    begin
    inc(NomColExcel);
    stExcel:=Excel.Cells[1,NomColExcel];
    MeProt.Lines.Add('NomColExcel:'+IntTostr(NomColExcel)+'!'+stExcel);
    end;
  MeProt.Lines.Add('NomColExcel:'+IntTostr(NomColExcel));
  //Открыть соответствующий файл ФГОС
  MeProt.Lines.Add(CurrentDir+'\Документы\'+Copy(Napravlenie,1,2)+'*'+Copy(Napravlenie,4,2)+'*'+Copy(Napravlenie,7,2)+'*.doc*');
  FindRes := FindFirst(CurrentDir+'\Документы\'+Copy(Napravlenie,1,2)+'*'+Copy(Napravlenie,4,2)+'*'+Copy(Napravlenie,7,2)+'*.doc*', faAnyFile, SR);
  if FindRes = 0 then // Если нашли файл
    begin
    FileNameFGOS:=CurrentDir+'\Документы\'+SR.Name;
    FGOSOpen:=true;
    WordFGOS.Documents.Open(FileNameFGOS);
    MeProt.Lines.Add('Загружен ФГОС из файла:'+FileNameFGOS);
    LaNameFGOS.Caption:=FileNameFGOS;
    end;
  end;
NotFound:=true;
{    if not ((b<EndText) and (St<>StShablon)) then
      begin
      StShablonProv:=StShablon;
      //Ищем следующее нежелтое слово
      aShablon:=bShablon;
      while not EndWord(WordShablon.ActiveDocument.Range(bShablon,bShablon+1).Text) do
        inc(bShablon);
      StShablon:=WordShablon.ActiveDocument.Range(aShablon,bShablon).Text;
      MeFileShab.Lines.Add('Sh---!'+StShablon+'!');
      while EndWord(WordShablon.ActiveDocument.Range(bShablon,bShablon+1).Text) do
        inc(bShablon);
      dec(bShablon);
      end;
    until not ((b<EndText) and (St<>StShablon)); }

end;
end;

Procedure GoNotFoundWord(Word,WordShablon:variant; var aShablon,bShablon,a,b:Longword; TypeChange:byte);
var
aShSearch,bShSearch,aSearch,bSearch:Longword;
MaxLenSearch,LenSh,LenShShablon:Longword;
aRed,bRed:Longword;
CommentsRange:Variant;
stvivod:string;
begin
With FMain do
begin
NotFound:=true;
//Если не нашли, т.е. следующее слово другое
 Case TypeChange of
   0: stvivod:=GlavaShablon+' Shab';
   1: stvivod:=GlavaFGOS+' FGOS';
 End;
MeFileShab.Lines.Add(stvivod+' -!('+IntToStr(aShablon)+')'+WordShablon.ActiveDocument.Range(aShablon,bShablon).Text+'!');
MeFileSYOS.Lines.Add(Glava+' -!('+IntToStr(a)+')'+Word.ActiveDocument.Range(a,b).Text+'!');
inc(KolOsh); LaColOsh.Caption:=IntToStr(KolOsh);
//Поиск места в котором слова будут совпадать

//Установить максимальное расстояние между словами
MaxLenSearch:=0;
//Пока не нашли совпадение
while (St<>StShablon) and (not ((EndGlava(St,0)) and (EndGlava(StShablon,1)))) do
  begin
  aShSearch:=aShablon; bShSearch:=bShablon;
  aSearch:=a; bSearch:=b;
  inc(MaxLenSearch); //Увеличиваем максимальное расстояние между словами
  MeProt.Lines.Add(IntToStr(MaxLenSearch));
  St:=DeleteSpace(Word.ActiveDocument.Range(aSearch,bSearch).Text);
  LenSh:=1;  //Ищем слово в тексте на максимальном расстоянии
  while (LenSh<MaxLenSearch) and (bSearch<EndText) and (not EndGlava(St,1)) do
    begin
    while (EndWord(Word.ActiveDocument.Range(bSearch,bSearch+1).Text)) do
      inc(bSearch);
    aSearch:=bSearch;
    inc(bSearch);
    while not(EndWord(Word.ActiveDocument.Range(bSearch,bSearch+1).Text)) do
      inc(bSearch);
    St:=DeleteSpace(Word.ActiveDocument.Range(aSearch,bSearch).Text);
    inc(LenSh);
    end;
  St:=DeleteSpace(Word.ActiveDocument.Range(aSearch,bSearch).Text);
  MeProt.Lines.Add(Glava+' '+IntToStr(MaxLenSearch)+'!'+St+'!'+IntToStr(aSearch)+','+IntToStr(bSearch));
  aShSearch:=aShablon; bShSearch:=bShablon;
  StShablon:=DeleteSpace(WordShablon.ActiveDocument.Range(aShSearch,bShSearch).Text);
  LenSh:=1;
  while (St<>StShablon) and (LenSh<=MaxLenSearch) and (bShSearch<EndTextShablon) and (not EndGlava(StShablon,1)) do //Сравниваем с текущим словом шаблона
    begin
    //Увеличиваем расстояние текущего слово шаблона, пока не достигнем максимума
    while (EndWord(WordShablon.ActiveDocument.Range(bShSearch,bShSearch+1).Text)) do
      inc(bShSearch);
    aShSearch:=bShSearch;
    inc(bShSearch);
    while not (EndWord(WordShablon.ActiveDocument.Range(bShSearch,bShSearch+1).Text)) do
      inc(bShSearch);
    If WordShablon.ActiveDocument.Range(aShSearch,bShSearch).HighlightColorIndex<>wdYellow then
      begin
      StShablon:=DeleteSpace(WordShablon.ActiveDocument.Range(aShSearch,bShSearch).Text);
      if St<>StShablon then
        MeProt.Lines.Add('-'+IntToStr(LenSh)+'!'+StShablon+'!'+IntToStr(aShSearch)+','+IntToStr(bShSearch)+' - '+IntToStr(MaxLenSearch)+'!'+St+'!'+IntToStr(aSearch)+','+IntToStr(bSearch))
      else
        MeProt.Lines.Add('+'+IntToStr(LenSh)+'!'+StShablon+'!'+IntToStr(aShSearch)+','+IntToStr(bShSearch)+' - '+IntToStr(MaxLenSearch)+'!'+St+'!'+IntToStr(aSearch)+','+IntToStr(bSearch));
      inc(LenSh);
      end;
    end;
  //Аналогично но текст и шаблон меняем местами.
  if (St<>StShablon) then
    begin
    MeProt.Lines.Add('NextSr!');
    aShSearch:=aShablon; bShSearch:=bShablon;
    while (WordShablon.ActiveDocument.Range(aShSearch,bShSearch).HighlightColorIndex=wdYellow) or ((EndWord(WordShablon.ActiveDocument.Range(bShSearch,bShSearch+1).Text))) do
      inc(bShSearch);
    StShablon:=DeleteSpace(WordShablon.ActiveDocument.Range(aShSearch,bShSearch).Text);
    LenSh:=1;  //Ищем слово в тексте на максимальном расстоянии
    while (LenSh<MaxLenSearch) and (bShSearch<EndTextShablon) and (not EndGlava(StShablon,1)) do
      begin
      while  ((EndWord(WordShablon.ActiveDocument.Range(bShSearch,bShSearch+1).Text))) do
        inc(bShSearch);
      aShSearch:=bShSearch;
      inc(bShSearch);
      while  (not (EndWord(WordShablon.ActiveDocument.Range(bShSearch,bShSearch+1).Text))) do
        inc(bShSearch);
      StShablon:=DeleteSpace(WordShablon.ActiveDocument.Range(aShSearch,bShSearch).Text);
      inc(LenSh);
      end;
    StShablon:=DeleteSpace(WordShablon.ActiveDocument.Range(aShSearch,bShSearch).Text);
    MeProt.Lines.Add(IntToStr(MaxLenSearch)+'!'+StShablon+'!'+IntToStr(aShSearch)+','+IntToStr(bShSearch));
    aSearch:=a; bSearch:=b;
    St:=DeleteSpace(Word.ActiveDocument.Range(aSearch,bSearch).Text);
    LenSh:=1;
    while (St<>StShablon) and (LenSh<=MaxLenSearch) and (bSearch<EndText) and (not EndGlava(St,0)) do //Сравниваем с текущим словом
      begin
      //Увеличиваем расстояние текущего слово, пока не достигнем максимума
      while (EndWord(Word.ActiveDocument.Range(bSearch,bSearch+1).Text)) do
        inc(bSearch);
      aSearch:=bSearch;
      inc(bSearch);
      while not(EndWord(Word.ActiveDocument.Range(bSearch,bSearch+1).Text)) do
        inc(bSearch);
      St:=DeleteSpace(Word.ActiveDocument.Range(aSearch,bSearch).Text);
      if St<>StShablon then
        MeProt.Lines.Add('-'+IntToStr(MaxLenSearch)+'!'+StShablon+'!'+IntToStr(aShSearch)+','+IntToStr(bShSearch)+' - '+IntToStr(LenSh)+'!'+st+'!'+IntToStr(aSearch)+','+IntToStr(bSearch))
      else
        MeProt.Lines.Add('+'+IntToStr(MaxLenSearch)+'!'+StShablon+'!'+IntToStr(aShSearch)+','+IntToStr(bShSearch)+' - '+IntToStr(LenSh)+'!'+st+'!'+IntToStr(aSearch)+','+IntToStr(bSearch));
      inc(LenSh);
      end;
    end;
  If EndGlava(St,0) then  MeProt.Lines.Add('EndSt!'+St);
  If EndGlava(StShablon,1) then MeProt.Lines.Add('EndStShablon!'+StShablon);
  If (St<>StShablon) then  MeProt.Lines.Add('St<>StShablon!'+st+'!'+StShablon);
  MeProt.Lines.Add('Circle');
  end;
MeProt.Lines.Add('Next');
//Тут вывести, несовпадающие элементы
MeFileShab.Lines.Add('?'+WordShablon.ActiveDocument.Range(aShablon,aShSearch).Text+'('+IntToStr(aShablon)+','+IntToStr(aShSearch)+')');
MeFileSYOS.Lines.Add('?'+Word.ActiveDocument.Range(a,aSearch).Text+'('+IntToStr(a)+','+IntToStr(aSearch)+')');
bRed:=a;
while (EndWord(Word.ActiveDocument.Range(bRed,bRed+1).Text)) do
  inc(bRed);
aRed:=bRed;
while not(EndWord(Word.ActiveDocument.Range(bRed,bRed+1).Text)) do
  inc(bRed);
MeFileSYOS.Lines.Add('?RED '+Word.ActiveDocument.Range(aRed,bRed).Text+'('+IntToStr(aRed)+','+IntToStr(bRed)+')');
Word.ActiveDocument.Range(aRed,bRed).HighlightColorIndex:=wdRed;

//Добавить примечание
if aShablon+1<aShSearch then
  begin
  CommentsRange:=Word.ActiveDocument.Range(aRed,bSearch);
  Word.ActiveDocument.Comments.Add(CommentsRange,WordShablon.ActiveDocument.Range(aShablon,aShSearch).Text);
  end;

//Перейти на место совпадения
aShablon:=aShSearch; bShablon:=bShSearch;
a:=aSearch; b:=bSearch;
end;
end;

Procedure GoFoundWord(Word,WordShablon:variant; var aShablon,bShablon,a,b:Longword; TypeChange:byte);
var
stvivod:string;
begin
With FMain do
begin
St:=DeleteSpace(Word.ActiveDocument.Range(a,b).Text);
StShablon:=DeleteSpace(WordShablon.ActiveDocument.Range(aShablon,bShablon).Text);
EndGlava(St,0); EndGlava(StShablon,1);
LaGlava.Caption:=Glava; LaGlavaSYOS.Caption:=GlavaShablon; LaGlavaFGOS.Caption:=GlavaFGOS;
{if (Pos('_Toc',st)<>0) or (Pos('PAGEREF',st)<>0)  or (Pos('HYPERLINK',st)<>0) or (Pos('.ru',st)<>0) then st:='';
if (Pos('_Toc',StShablon)<>0) or (Pos('PAGEREF',StShablon)<>0) or (Pos('HYPERLINK',StShablon)<>0) or (Pos('.ru',StShablon)<>0) then StShablon:='';
}
MeProt.Lines.Add('Sr= !'+St+'!'+IntToStr(a)+','+IntToStr(b)+'-!'+StShablon+'!'+IntToStr(aShablon)+','+IntToStr(bShablon));
if St=StShablon then
  begin
  //Если нашли
  Case TypeChange of
   0: stvivod:=GlavaShablon+' Shab';
   1: stvivod:=GlavaFGOS+' FGOS';
  End;
  MeFileShab.Lines.Add(stvivod+' +!('+IntToStr(aShablon)+')'+WordShablon.ActiveDocument.Range(aShablon,bShablon).Text+'!');
  MeFileSYOS.Lines.Add(Glava+' +!('+IntToStr(a)+')'+Word.ActiveDocument.Range(a,b).Text+'!');
    repeat
    while (EndWord(Word.ActiveDocument.Range(b,b+1).Text)) do
      inc(b);
    a:=b;
    while not(EndWord(Word.ActiveDocument.Range(b,b+1).Text)) do
      inc(b);
    until not ((Word.ActiveDocument.Range(a,b).Text='') or (EndWord(Word.ActiveDocument.Range(a,b).Text)));
  NotFound:=false;
  end
else
  begin
  GoNotFoundWord(Word,WordShablon,aShablon,bShablon,a,b,TypeChange);
  end;
end;
end;

Procedure GoFoundIfFootnotes;
var
ColFotonesText,ColFotonesShab:LOngword;
NomFotones:Longword;
aFt,bFt:Longword;
aFtShablon,bFtShablon:Longword;
FotonesText,FotonesShab:Variant;
begin
ColFotonesText:=Word.ActiveDocument.Footnotes.Count;
ColFotonesShab:=WordShablon.ActiveDocument.Footnotes.Count;
if (ColFotonesText>0) and (ColFotonesShab>0) and (ColFotonesText=ColFotonesShab) then
  for NomFotones := 1 to ColFotonesShab do
    begin
    aFt:=0; bFt:=1;
    aFtShablon:=0; bFtShablon:=1;
    FotonesText:=Word.ActiveDocument.Footnotes(NomFotones);
    FotonesShab:=WordShablon.ActiveDocument.Footnotes(NomFotones);
    while (bFt<Length(FotonesText.Range)) and (bFtShablon<Length(FotonesShab.Range)) do
      begin

      end;
    end;
end;

Procedure VivodFotones;
var
ColFotonesText,NomFotones:Longword;
begin
ColFotonesText:=Word.ActiveDocument.Footnotes.Count;
for NomFotones := 1 to ColFotonesText do
  FMain.MeProt.Lines.Add(Word.ActiveDocument.Footnotes.Range.Text);
end;

procedure TFMain.BtLoadSuosClick(Sender: TObject);
var
i:longword;
aShablon,bShablon:Longword;
a,b:Longword;
aShablonFGOS,bShablonFGOS,AFgos,BFgos:Longword;

RangeShablon,RangeText:Variant;
StringShablon,StringText:string;
begin
KolOsh:=0; LaColOsh.Caption:=IntToStr(KolOsh);
NomRowExcel:=1; NomColExcel:=1;
Glava:='1.0.';GlavaShablon:='1.0.';GlavaFGOS:='1.0.';
FGOSOpen:=false;
Napravlenie:='';
case RadioGroup1.ItemIndex of
  0:FileNameOb:=CurrentDir+'\Документы\МАКЕТ СУОС ВО 3   БАКАЛАВР';
  1:FileNameOb:=CurrentDir+'\Документы\МАКЕТ СУОС ВО 3   МАГИСТР';
  2:FileNameOb:=CurrentDir+'\Документы\МАКЕТ СУОС ВО 3  СПЕЦИАЛИСТ';
end;
if FileExists(FileNameOb+'.docx') then
begin
Word := CreateOleObject('Word.Application'); Word.Visible := false;
WordShablon := CreateOleObject('Word.Application'); WordShablon.Visible := false;
WordFGOS := CreateOleObject('Word.Application'); WordFGOS.Visible := false;
Excel := CreateOleObject('Excel.Application'); Excel.Visible := false;

WordShablon.Documents.Open(FileNameOb+'.docx');
MeProt.Lines.Add('Загружен макет СУОС из файла:'+FileNameOb+'.docx');
Excel.Workbooks.Open(FileNameOb+'.xlsx');
MeProt.Lines.Add('Для работы открыт файл Excel:'+FileNameOb+'.xlsx');
NomColExcel:=Excel.Cells[NomRowExcel,1];
inc(NomColExcel);
Excel.Cells[NomRowExcel,1]:=NomColExcel;
inc(NomRowExcel);
LaNameShablon.Caption:=FileNameOb;
if OpenDialog.Execute then
begin
Word.Documents.Open(OpenDialog.FileName);
LaNameFile.Caption:=OpenDialog.FileName;
//Эхопечать
for i:=1 to WordShablon.Documents.Count do
    MeFileShab.Lines.Add(WordShablon.Documents.Item(i).Name);
for i:=1 to Word.Documents.Count do
    MeFileSYOS.Lines.Add(Word.Documents.Item(i).Name);

//Определение начальных параметров
StringShablon:=WordShablon.ActiveDocument.Content.Text;
EndTextShablon:=length(StringShablon)*2;// ComputeStatistics(5);
MeFileShab.Lines.Add(IntToStr(EndTextShablon));
aShablon:=0; bShablon:=1;
StringText:=Word.ActiveDocument.Content.Text;
Word.ActiveDocument.Range.HighlightColorIndex:=0;
{if Word.ActiveDocument.Comments.Count>0 then
  begin
  Word.ActiveDocument.Comments.NextComment;
  Word.ActiveDocument.Comments.DeleteAllCommentsInDoc;
  end;                  }
EndText:=length(StringText)*2; //.ComputeStatistics(5);
MeFileSYOS.Lines.Add(IntToStr(EndText));
a:=0; b:=1;
AFGOS:=0; BFGOS:=1;
NotFound:=false;
//Идем по всему тексту шаблона, по словам
while (bShablon<EndTextShablon) and (b<EndText) do
  begin
  try
  //Выбираем слово из шаблона
  RangeShablon:=WordShablon.ActiveDocument.Range(aShablon,bShablon);
  if RangeShablon.HighlightColorIndex=wdYellow then    //Если слово помечено желтым,
    begin
    GoYellow(aShablon,bShablon,a,b,AFGOS,BFGOS);
    end
  else
    begin
    //Иначе ищем слово в тексте
    GoFoundWord(Word,WordShablon,aShablon,bShablon,a,b,0);
    end;

  if not NotFound  then
  repeat
  while (EndWord(WordShablon.ActiveDocument.Range(bShablon,bShablon+1).Text)) do
    inc(bShablon);
  aShablon:=bShablon;
  while not(EndWord(WordShablon.ActiveDocument.Range(bShablon,bShablon+1).Text)) do
    inc(bShablon);
  until not ((WordShablon.ActiveDocument.Range(aShablon,bShablon).Text='') or (EndWord(WordShablon.ActiveDocument.Range(aShablon,bShablon).Text)));
  except
    bShablon:=EndTextShablon+100;
  end;
  end;
MeFileShab.Lines.Add(IntToStr(bShablon));
MeFileShab.Lines.Add(IntToStr(b));
//GoFoundIfFootnotes;

Word.Documents.Close; Word.Quit; Word := Unassigned;
WordShablon.Documents.Close; WordShablon.Quit; WordShablon := Unassigned;
if FGOSOpen then WordFGOS.Documents.Close; WordFGOS.Quit; WordFGOS := Unassigned;
Excel.Workbooks[1].save; Excel.Workbooks.Close; Excel.Quit; Excel := Unassigned;
end;
end
else
  MeProt.Lines.Add('Не найден файл макета СУОС:'+FileNameOb);
end;

procedure TFMain.Button1Click(Sender: TObject);
var
st:string;
i:Longword;
begin
WordShablon := CreateOleObject('Word.Application');
WordShablon.Visible := false;
WordShablon.Documents.Open(CurrentDir+'\Документы\1.docx');
St:=WordShablon.ActiveDocument.Range(1,700).text;
i:=1;
while i<length(st) do
  begin
  MeProt.Lines.Add(st[i]+' - '+IntToStr(ord(St[i])));
  inc(i);
  end;
WordShablon.Documents.Close;
WordShablon.Quit;
WordShablon := Unassigned;
end;

procedure TFMain.BtSprFileClick(Sender: TObject);
var
b,Max:Longword;
St,st1:String;
begin
WordShablon := CreateOleObject('Word.Application');
WordShablon.Visible := false;
if OpenDialog.Execute then
begin
WordShablon.Documents.Open(OpenDialog.FileName);
St:=WordShablon.ActiveDocument.Content.Text;
Max:=length(St)*2;
b:=2;
while (b<Max) do
  begin
  try
  st:=WordShablon.ActiveDocument.Range(b-1,b).Text;
  st1:=st1+st;
  if (Length(st)=1) then
  begin
  If St=Chr(8211) then
   WordShablon.ActiveDocument.Range(b-1,b).Text:=Chr(45);    //Дефис - тире
 { If St=Chr(171) then
   WordShablon.ActiveDocument.Range(b-1,b).Text:=Chr(34);     //Ковычки
  If St=Chr(187) then
   WordShablon.ActiveDocument.Range(b-1,b).Text:=Chr(34);     }
 { if St='N' then
   WordShablon.ActiveDocument.Range(b-1,b).Text:='№';//Номер   }
  end;
  inc(b);
  if b mod 256 =0 then
    begin
    MeProt.Lines.Add(st1);
    st1:='';
    end;
  except
    b:=Max+100;
  end;
  end;
MeProt.Lines.Add('Исправлен Файл:'+OpenDialog.FileName);
end;
WordShablon.Documents.Close;
WordShablon.Quit;
WordShablon := Unassigned;

end;

procedure TFMain.FormActivate(Sender: TObject);
begin
CurrentDir := GetCurrentDir;
end;

end.
