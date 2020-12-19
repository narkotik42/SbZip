unit uSbZip;
{-----------------------------------------------------------------------------
 Unit Name: uSbZip
 Author:    Salih BAÐCI
 Date:      19-Ara-2020
-----------------------------------------------------------------------------}
interface

  uses SysUtils, Classes, Controls, Zip, IOUtils ,Types, Masks, StrUtils;

  type
    TFilePathStr = record
    Drive  : String;
    Folder : String;
    Name   : String;
  end;

  type
  TSbZip = class(TComponent)
  strict private
    FCompressionType: TZipCompression;
    FSourceFileList: TStrings;
    FSourceFolderList: TStrings;
    FDestFileList: TStrings;
    FDestFileFoundDelete: Boolean;
    FSourceSubFolders: Boolean;
    FDestSubFolders: Boolean;
    FMaskExt: String;
    FMaskNotExt: String;
    FMaskNotFileName: String;
    procedure setSourceFileList(const Value: TStrings);
    procedure setSourceFolderList(const Value: TStrings);
    procedure setDestFileList(const Value: TStrings);
  private
    function GetFileNameDest(const AFilePath:String):TFilePathStr;
    function GetSourceFileList(const AFolder:String):TStringDynArray;overload;
    function GetSourceFileList:TStringDynArray;overload;
    function GetDestFileList(const ASourceFileList:TStringDynArray):TStringDynArray;
    function GetSourceFileNotExits(const ASourceFileList:TStringDynArray):String;
    function GetDestFileExits:String;
  public
    constructor Create(AOwner:TComponent); override;
    destructor Destroy;override;
    procedure RunCompress;virtual;
  published
    property CompressionType:TZipCompression read FCompressionType write FCompressionType; // Sýkýþtýrma türü
    property SourceFileList: TStrings read FSourceFileList write setSourceFileList; // Zip iþlemi yapýlacak dosyalar listesi
    property SourceFolderList: TStrings read FSourceFolderList write setSourceFolderList; // Zip iþlemi yapýlacak klasörler listesi
    property DestFileList: TStrings read FDestFileList write setDestFileList; // Zip iþlemi çýkartýlacak dosyalar listesi, birden çok yere çýkartýlabilir
    property DestFileFoundDelete: Boolean read FDestFileFoundDelete write FDestFileFoundDelete default True; // Hedef zip dosyasý varsa sil
    property SourceSubFolders: Boolean read FSourceSubFolders write FSourceSubFolders default False; // Alt klasörleri tara
    property DestSubFolders: Boolean read FDestSubFolders write FDestSubFolders default False; // Yeni zip dosyasýnda klasörleri ile oluþtur
    property MaskExt: String read FMaskExt write FMaskExt; // Geçerli uzantýlar Örn: *.* veya *.json;*.xml
    property MaskNotExt: String read FMaskNotExt write FMaskNotExt; // Hariç tutulacak uzantýlar Örn: *.* veya *.json;*.xml
    property MaskNotFileName: String read FMaskNotFileName write FMaskNotFileName; // Hariç tutulacak dosya isimleri Örn: abc.dll;xyz.txt
  end;

implementation

{ TSbZip }

constructor TSbZip.Create(AOwner:TComponent);
begin
  inherited Create(AOwner);
  if FSourceFileList = nil then
    FSourceFileList := TStringList.Create;
  if FSourceFolderList = nil then
    FSourceFolderList := TStringList.Create;
  if FDestFileList = nil then
    FDestFileList := TStringList.Create;
  FCompressionType := zcDeflate;
  FDestFileFoundDelete := True;
  FMaskExt := '*.*';
end;

destructor TSbZip.Destroy;
begin
  if Assigned(FDestFileList) then
    FDestFileList.Free;
  if Assigned(FSourceFolderList) then
    FSourceFolderList.Free;
  if Assigned(FSourceFileList) then
    FSourceFileList.Free;
  inherited;
end;

function TSbZip.GetDestFileExits: String;
var
  Ind : Integer;
begin
  Result := '';
  for Ind := 0 to Pred(DestFileList.Count) do
  begin
    if FileExists(DestFileList[Ind]) then
    begin
      if DestFileFoundDelete then
        DeleteFile(DestFileList[Ind])
      else
        Exit(DestFileList[Ind]);
    end;
  end;
end;

function TSbZip.GetDestFileList(const ASourceFileList: TStringDynArray): TStringDynArray;
var
  Ind      : Integer;
  xFileStr : TFilePathStr;
  xName    : String;
begin
  SetLength(Result,Length(ASourceFileList));
  for Ind := Low(ASourceFileList) to High(ASourceFileList) do
  begin
    xName    := '';
    xFileStr := GetFileNameDest(ASourceFileList[Ind]);
    if DestSubFolders then
      xName := xFileStr.Folder + '\';
    xName := Concat(xName,xFileStr.Name);
    Result[Ind] := xName;
  end;
end;

function TSbZip.GetFileNameDest(const AFilePath: String): TFilePathStr;
begin
  // C:\FOLDER\FOLDER2\a.txt
  Result.Drive  := ExtractFileDrive(AFilePath); // C:
  Result.Folder := TPath.GetDirectoryName(AFilePath);
  Result.Folder := StringReplace(Result.Folder,Result.Drive + '\','',[rfReplaceAll]); // FOLDER\FOLDER2
  Result.Name   := TPath.GetFileName(AFilePath); // a.txt
end;

function TSbZip.GetSourceFileList: TStringDynArray;
var
  Ind      : Integer;
  xFileArr : TStringDynArray;
begin
  SetLength(xFileArr,SourceFileList.Count);
  for Ind := 0 to Pred(SourceFileList.Count) do
    xFileArr[Ind] := SourceFileList[Ind];
  if Length(xFileArr) > 0 then
    Result := Concat(Result,xFileArr);

  for Ind := 0 to Pred(SourceFolderList.Count) do
    Result := Concat(Result,GetSourceFileList(SourceFolderList[Ind]));
end;

function TSbZip.GetSourceFileNotExits(const ASourceFileList: TStringDynArray): String;
var
  Ind : Integer;
begin
  Result := '';
  for Ind := Low(ASourceFileList) to High(ASourceFileList) do
    if not FileExists(ASourceFileList[Ind]) then
      Exit(ASourceFileList[Ind]);
end;

procedure TSbZip.RunCompress;
var
  Ind        : Integer;
  Ind2       : Integer;
  xZipper    : TZipFile;
  xSFileList : TStringDynArray;
  xDFileList : TStringDynArray;
  xGcc       : String;
  xFPathStr  : TFilePathStr;
begin
  xGcc := GetDestFileExits;
  if xGcc <> '' then
    raise Exception.Create('Destination file found: ' + xGcc);

  xSFileList := GetSourceFileList;
  xGcc := GetSourceFileNotExits(xSFileList);
  if xGcc <> '' then
    raise Exception.Create('Source file not found: ' + xGcc);

  if Length(xSFileList) > 0 then
  begin
    xDFileList := GetDestFileList(xSFileList);
    for Ind := 0 to Pred(DestFileList.Count) do
    begin
      xFPathStr := GetFileNameDest(DestFileList[Ind]);
      if xFPathStr.Folder <> '' then
        ForceDirectories(TPath.GetDirectoryName(DestFileList[Ind]));

      xZipper := TZipFile.Create;
      try
        xZipper.Open(DestFileList[Ind],zmWrite);
        for Ind2 := Low(xSFileList) to High(xSFileList) do
          xZipper.Add(xSFileList[Ind2],xDFileList[Ind2],CompressionType);
      finally
        xZipper.Free;
      end;
    end;
  end;
end;

function TSbZip.GetSourceFileList(const AFolder:String):TStringDynArray;
var
  xPredicate      : TDirectory.TFilterPredicate;
  xMaskArr        : TStringDynArray;
  xMaskNotArr     : TStringDynArray;
  xMaskNotFileArr : TStringDynArray;
begin
  xMaskArr        := SplitString(MaskExt,';');
  xMaskNotArr     := SplitString(MaskNotExt,';');
  xMaskNotFileArr := SplitString(MaskNotFileName,';');

  xPredicate :=
    function(const Path: string; const SearchRec: TSearchRec): Boolean
    var
      xMask        : String;
      xMaskNot     : String;
      xMaskNotFile : String;
    begin
      for xMask in xMaskArr do
      begin
        if MatchesMask(SearchRec.Name, xMask) then
        begin
          for xMaskNot in xMaskNotArr do
          begin
            if MatchesMask(SearchRec.Name, xMaskNot) then
              Exit(False);
          end;
          for xMaskNotFile in xMaskNotFileArr do
          begin
            if SearchRec.Name = xMaskNotFile then
              Exit(False);
          end;
          Exit(True);
        end;
      end;
      Exit(False);
    end;

    if not DirectoryExists(AFolder) then
      raise Exception.Create('Source folder not found: ' + AFolder);

    if SourceSubFolders then
      Result := TDirectory.GetFiles(AFolder,TSearchOption.soAllDirectories,xPredicate)
    else
      Result := TDirectory.GetFiles(AFolder,xPredicate);
end;

procedure TSbZip.setDestFileList(const Value: TStrings);
begin
  if Assigned(FDestFileList) then
    FDestFileList.Assign(Value)
  else
    FDestFileList := Value;
end;

procedure TSbZip.setSourceFileList(const Value: TStrings);
begin
  if Assigned(FSourceFileList) then
    FSourceFileList.Assign(Value)
  else
    FSourceFileList := Value;
end;

procedure TSbZip.setSourceFolderList(const Value: TStrings);
begin
  if Assigned(FSourceFolderList) then
    FSourceFolderList.Assign(Value)
  else
    FSourceFolderList := Value;
end;

end.
