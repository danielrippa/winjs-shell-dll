unit ChakraShellUtils;

{$mode delphi}

interface

  uses Variants, ChakraTypes;

  type
    WideStringArray = array of WideString;

  function GetWScriptShell: OleVariant;

  function ExpandEnvVar(EnvVarName: WideString): WideString;
  function ExpandEnvVarString(EnvVarString: WideString): WideString;

  function ExecuteProcess(Executable: WideString; Params: WideStringArray; WorkingFolder: WideString; Priority, BufferSize: Integer; OutputProc: TJsValue): TJsValue;

  function StartProcess(Executable: WideString, Params: WideStringArray; WorkingFolder: WideString; Priority: Integer; WaitOnReturn: Boolean): Integer;

  function WScriptShellRun(Command: WideString; WaitOnReturn: Boolean): TJsValue;

implementation

  uses Chakra, ComObj, ActiveX, SysUtils, Types, StrUtils, Process, Classes, ChakraError;

  function GetWScriptShell;
  begin
    Result := CreateOleObject('WScript.Shell');
  end;

  function ExpandEnvVarString;
  var
    Shell: OleVariant;
  begin
    Shell := GetWScriptShell();
    Result := Shell.ExpandEnvironmentStrings(EnvVarString);
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

  procedure InvokeOutputCallback(aCallback: TJsValue; Output: String);
  var
    Args: Array of TJsValue;
    ArgCount: Word;
  begin
    ArgCount := 2;
    Args := [ Undefined, StringAsJsString(Output) ];

    CallFunction(aCallback, @Args[0], ArgCount);
  end;

  procedure EmitOutput(aStream: TStringStream; aBuffer: Pointer; aBytesRead: DWord; aOutputProc: TJsValue);
  begin
    with aStream do begin
      Size := 0;
      Write(aBuffer^, aBytesRead);

      InvokeOutputCallback(aOutputProc, DataString);
    end;
  end;

  procedure EmitProcessOutput(aProcess: TProcess; aBufferSize: Integer; aOutputProc: TJsValue);
  var
    Stream: TStringStream;
    Buffer: Pointer;
    BytesRead: DWord;
  begin

    Stream := TStringStream.Create(EmptyStr);
    GetMem(Buffer, aBufferSize);

    try

      aProcess.Execute;

      while aProcess.Running and aProcess.Active do begin

        BytesRead := aProcess.Output.Read(Buffer^, aBufferSize);
        if BytesRead = 0 then begin
          Sleep(10);
        end else begin
          EmitOutput(Stream, Buffer, BytesRead, aOutputProc);
        end;

      end;

      repeat

        BytesRead := aProcess.Output.Read(Buffer^, aBufferSize);
        if BytesRead <> 0 then begin
          EmitOutput(Stream, Buffer, BytesRead, aOutputProc);
        end;

      until BytesRead = 0;

    finally
      FreeMem(Buffer, aBufferSize);
      Stream.Free;
    end;

  end;

  type

    TProcessOutcome = record
      Stdout: String;
      StdErr: String;
      ExitStatus: Integer;
    end;

  procedure AddParameters(aParams: WideStringArray; Parameters: TProcessStrings);
  var
    i: Integer;
  begin
    for i := 0 to Length(aParams) - 1 do begin
      Parameters.Add(aParams[i]);
    end;
  end;

  function ExecProcess(aExecutable: WideString; aParams: WideStringArray; aWorkingFolder: WideString; aPriority, aBufferSize: Integer; aOutputProc: TJsValue): TProcessOutcome;
  var
    Process: TProcess;
    I: Integer;
    ProcessPriority: TProcessPriority;
  begin

    try
      ProcessPriority := TProcessPriority(aPriority);
    except
      on E: Exception do
        ThrowError('Invalid Priority Value %d', [aPriority]);
    end;

    try

      Process := TProcess.Create(Nil);

      with Process do begin

        Executable := aExecutable;
        AddParameters(aParams, Parameters);
        CurrentDirectory := aWorkingFolder;
        Priority := ProcessPriority;

        if aBufferSize > 0 then begin
          Options := [poUsePipes, poStdErrToOutput];
        end;
      end;

      with Result do begin

        if aBufferSize > 0 then begin
          EmitProcessOutput(Process, aBufferSize, aOutputProc);
          ExitStatus := Process.ExitStatus;
        end else begin
          Process.RunCommandLoop(StdOut, StdErr, ExitStatus);
        end;

      end;

    finally
      Process.Free;
    end;
  end;

  function ExecuteProcess;
  begin
    Result := CreateObject;

    with ExecProcess(Executable, Params, WorkingFolder, Priority, BufferSize, OutputProc) do begin
      SetProperty(Result, 'stdout', StringAsJsString(StdOut));
      SetProperty(Result, 'stderr', StringAsJsString(StdErr));
      SetProperty(Result, 'exitStatus', IntAsJsNumber(ExitStatus));
    end;
  end;

  function WScriptShellRun;
  var
    Shell: OleVariant;
  begin
    Shell := GetWScriptShell();
    Result := IntAsJsNumber(Shell.Run(Command, 0, WaitOnReturn));
  end;

  function StartProcess;
  begin

    Result := -1;

    try
      ProcessPriority := TProcessPriority(aPriority);
    except
      on E: Exception do
        ThrowError('Invalid Priority Value %d', [aPriority]);
    end;

    try

      Process := TProcess.Create(Nil);

      with Process do begin

        Executable := aExecutable;
        AddParameters(aParams, Parameters);
        CurrentDirectory := aWorkingFolder;
        Priority := ProcessPriority;

        Result := ProcessID;

        if AWaitOnReturn then begin
          Options := [poWaitOnExit];
        end;

        Execute;

      end;

    finally
      Process.Free;
    end;

  end;

initialization

  CoInitialize(Nil);

finalization

  CoUninitialize;

end.