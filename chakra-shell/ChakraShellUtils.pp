unit ChakraShellUtils;

{$mode delphi}

interface

  uses Variants, ChakraTypes;

  function GetWScriptShell: OleVariant;

  function ExpandEnvVar(EnvVarName: WideString): WideString;

  function ExecuteProcess(Executable, Params, WorkingFolder: WideString; Priority, BufferSize: Integer; OutputProc: TJsValue): TJsValue;

implementation

  uses Chakra, ChakraErr, ComObj, ActiveX, SysUtils, Types, StrUtils, Process, Classes;

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

  function ExecProcess(aExecutable, aParams, aWorkingFolder: WideString; aPriority, aBufferSize: Integer; aOutputProc: TJsValue): TProcessOutcome;
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
        CurrentDirectory := aWorkingFolder;
        Parameters.Add(aParams);
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

initialization

  CoInitialize(Nil);

finalization

  CoUninitialize;

end.