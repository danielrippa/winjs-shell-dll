unit Win32ShellExecute;

{$mode delphi}

interface

  function ShellExecuteVerb(aVerb, aFileName, aParameters, aClassName, aWorkingDirectory: String; aWindowState, aFlags: Integer; aMonitor, aWindowHandle: THandle): Boolean;

implementation

  uses ShellApi;

  function ShellExecuteVerb;
  var
    ExecuteInfo: TShellExecuteInfoA;
    Size: Integer;
  begin

    Size := SizeOf(ExecuteInfo);
    FillChar(ExecuteInfo, Size, 0);

    with ExecuteInfo do begin
      cbSize := Size;
      fMask := aFlags;
      Wnd := aWindowHandle;
      lpVerb := PChar(aVerb);
      lpFile := PChar(aFileName);
      lpParameters := PChar(aParameters);
      lpDirectory := PChar(aWorkingDirectory);
      nShow := aWindowState;
      lpClass := PChar(aClassName);
      DUMMYUNIONNAME.hMonitor := aMonitor;

      // TODO: Implement more of SHELLEXECUTEINFOA
      // https://learn.microsoft.com/en-us/windows/win32/api/shellapi/ns-shellapi-shellexecuteinfoa

    end;

    Result := ShellExecuteExA(@ExecuteInfo);

  end;

end.
