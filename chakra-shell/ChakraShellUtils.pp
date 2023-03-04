unit ChakraShellUtils;

{$mode delphi}

interface

  uses Variants, ChakraTypes;

  function GetWScriptShell: OleVariant;

  function ExpandEnvVar(EnvVarName: WideString): WideString;

  function ExecuteProcess(Executable, Params, WorkingFolder: WideString): TJsValue;

  function WScriptShellRun(Command: WideString; WaitOnReturn: Boolean): TJsValue;

implementation

  uses Chakra, ComObj, ActiveX, SysUtils, Types, StrUtils, Process;

  function GetWScriptShell;
  begin
    Result := CreateOleObject('WScript.Shell');
  end;

  function ExpandEnvVar;
  var
    Shell: OleVariant;
    Query: String;
    Expanded: String;
  begin
    Shell := GetWScriptShell();
    Query := '%' + EnvVarName + '%';

    Expanded := Shell.ExpandEnvironmentStrings(Query);

    if CompareText(EnvVarName, Query) = 0 then begin
      Result := '';
    end else begin
      Result := Expanded;
    end;
  end;

  function WScriptShellRun;
  var
    Shell: OleVariant;
  begin
    Shell := GetWScriptShell();
    Result := IntAsJsNumber(Shell.Run(Command, 0, WaitOnReturn));
  end;

  type TUnicodeStringArray = array of UnicodeString;

  function UnicodeSplit(Value: UnicodeString): TUnicodeStringArray;
  var
    Params: TStringDynArray;
    I: Integer;
  begin
    Params := SplitString(Value, ' ');
    SetLength(Result, Length(Params));
    for I := 0 to Length(Params) do begin
      Result[I] := Params[I];
    end;
  end;

  type

    TProcessOutcome = record
      Stdout: String;
      StdErr: String;
      Succeeded: Boolean;
      ExitStatus: Integer;
    end;

  function ExecProcess(aExecutable, aParams: WideString; aWorkingFolder: WideString = ''): TProcessOutcome;
  var
    Process: TProcess;
    Params: TUnicodeStringArray;
    I: Integer;
  begin

    Result.Succeeded := False;

    try

      Process := TProcess.Create(Nil);

      Params := UnicodeSplit(aParams);

      with Process do begin
        Executable := aExecutable;
        CurrentDirectory := aWorkingFolder;
        for I := 0 to Length(Params) - 1 do begin
          Parameters.Add(Params[I]);
        end;
      end;

      with Result do begin
        Succeeded := Process.RunCommandLoop(StdOut, StdErr, ExitStatus) = 0;
      end;

    finally
      Process.Free;
    end;
  end;

  function ExecuteProcess;
  begin
    Result := CreateObject;

    with ExecProcess(Executable, Params, WorkingFolder) do begin
      SetProperty(Result, 'stdout', StringAsJsString(StdOut));
      SetProperty(Result, 'stderr', StringAsJsString(StdErr));
      SetProperty(Result, 'succeeded', BooleanAsJsBoolean(Succeeded));
      SetProperty(Result, 'exitStatus', IntAsJsNumber(ExitStatus));
    end;
  end;

initialization

  CoInitialize(Nil);

finalization

  CoUninitialize;

end.