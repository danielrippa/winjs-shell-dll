unit Win32FileDialogs;


{$mode delphi}

interface

  function FileOpenDialog(aFileName, aDefaultExt, aFilter, aInitialDir, aTitle: UnicodeString; aFilterIndex: Integer): UnicodeString;
  function FileSaveDialog(aFileName, aDefaultExt, aFilter, aInitialDir, aTitle: UnicodeString; aFilterIndex: Integer): UnicodeString;

implementation

  uses
    CommDlg, SysUtils, Windows;

  function DialogOptions(aFileName, aDefaultExt, aFilter, aInitialDir, aTitle: UnicodeString; aFilterIndex: Integer): TOpenFilenameW;
    function Filter: UnicodeString;
    var
      I: Word;
    begin
      Result := aFilter + #0;
      for I := 1 to Length(aFilter) do begin
        if Result[I] = '|' then
          Result[I] := #0;
      end;
    end;
  var
    TempFilter, TempFileName, TempExt: UnicodeString;
    Succeeded: Boolean;
  begin
    TempFilter := Filter + #0;
    SetLength(TempFileName, MAX_PATH + 2);
    lstrcpyW(PWideChar(TempFileName), PWideCHar(aFileName));
    TempExt := aDefaultExt;
    if TempExt = '' then begin
      TempExt := ExtractFileExt(aFileName);
      Delete(TempExt, 1, 1);
    end;

    FillChar(Result, SizeOf(Result), 0);
    With Result do begin
      lStructSize := SizeOf(OpenFileNameW);
      lpstrFilter := PWideChar(TempFilter);
      nFilterIndex := aFilterIndex;
      nMaxFile := MAX_PATH;
      lpstrFile := PWideChar(TempFileName);
      if aInitialDir = '' then
        lpstrInitialDir := '.'
      else
        lpstrInitialDir := PWideChar(aInitialDir);
      lpstrTitle := PWideChar(aTitle);
      if TempExt <> '' then
        lpstrDefExt := PWideChar(TempExt);
    end;
  end;

  function FileOpenDialog(aFileName, aDefaultExt, aFilter, aInitialDir, aTitle: UnicodeString; aFilterIndex: Integer): UnicodeString;
  var
    Options: TOpenFilenameW;
  begin
    Options := DialogOptions(aFileName, aDefaultExt, aFilter, aInitialDir, aTitle, aFilterIndex);
    if GetOpenFileNameW(@Options) then
      Result := Options.lpstrFile;
  end;

  function FileSaveDialog(aFileName, aDefaultExt, aFilter, aInitialDir, aTitle: UnicodeString; aFilterIndex: Integer): UnicodeString;
  var
    Options: TOpenFilenameW;
  begin
    Options := DialogOptions(aFileName, aDefaultExt, aFilter, aInitialDir, aTitle, aFilterIndex);
    if GetSaveFileNameW(@Options) then
      Result := Options.lpstrFile;
  end;

end.