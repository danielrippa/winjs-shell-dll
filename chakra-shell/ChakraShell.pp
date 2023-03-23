unit ChakraShell;

{$mode delphi}

interface

  uses ChakraTypes;

  function GetJsValue: TJsValue;

implementation

  uses Chakra, ChakraUtils, ChakraErr, Variants, SysUtils, Win32ShellExecute, Win32Taskbar, ChakraShellUtils;

  function ShellExpandEnvVar(Args: PJsValue; ArgCount: Word): TJsValue;
  var
    EnvVarName: WideString;
  begin
    CheckParams('expandEnvVar', Args, ArgCount, [jsString], 1);
    EnvVarName := JsStringAsString(Args^);
    Result := StringAsJsString(ExpandEnvVar(EnvVarName));
  end;

  function ShellExec(Args: PJsValue; ArgCount: Word): TJsValue;
  var
    Executable, Params, WorkingFolder: WideString;
    Priority, BufferSize: Integer;
    OutputProc: TJsValue;
  begin
   CheckParams('exec', Args, ArgCount, [jsString, jsString, jsString, jsNumber, jsNumber, jsFunction], 6);

    Executable := JsStringAsString(Args^); Inc(Args);
    Params := JsStringAsString(Args^); Inc(Args);
    WorkingFolder := JsStringAsString(Args^); Inc(Args);

    Priority := JsNumberAsInt(Args^); Inc(Args);
    BufferSize := JsNumberAsInt(Args^); Inc(Args);

    OutputProc := Args^;

    Result := ExecuteProcess(Executable, Params, WorkingFolder, Priority, BufferSize, OutputProc);
  end;

  function ShellRun(Args: PJsValue; ArgCount: Word): TJsValue;
  var
    Command: WideString;
    WaitOnReturn: Boolean;
  begin
    WaitOnReturn := False;
    CheckParams('exec', Args, ArgCount, [jsString, jsBoolean], 1);

    Command := JsStringAsString(Args^);

    if ArgCount > 1 then begin
      Inc(Args);
      WaitOnReturn := JsBooleanAsBoolean(Args^);
    end;

    Result := WScriptShellRun(Command, WaitOnReturn);
  end;

  function ShellExecVerb(Args: PJsValue; ArgCount: Word): TJsValue;
  var
    Verb: String;
    FileName: String;
    Parameters: String;
    ClassName: String;
    WorkingDirectory: String;
    WindowState: Integer;
    Flags: Integer;
    Monitor: THandle;
    WindowHandle: THandle;
  begin
    CheckParams('execVerb', Args, ArgCount, [jsString, jsString, jsString, jsString, jsNumber, jsNumber, jsNumber, jsNumber], 8);

    Verb := JsStringAsString(Args^); Inc(Args);
    FileName := JsStringAsString(Args^); Inc(Args);
    Parameters := JsStringAsString(Args^); Inc(Args);
    WorkingDirectory := JsStringAsString(Args^); Inc(Args);
    WindowState := JsNumberAsInt(Args^); Inc(Args);
    Flags := JsNumberAsInt(Args^); Inc(Args);
    Monitor := THandle(JsNumberAsInt(Args^)); Inc(Args);
    WindowHandle := THandle(JsNumberAsInt(Args^));

    Result := BooleanAsJsBoolean(ShellExecuteVerb(Verb, FileName, Parameters, ClassName, WorkingDirectory, WindowState, Flags, Monitor, WindowHandle));
  end;

  function TaskbarIconSetProgressPercentage(Args: PJsValue; ArgCount: Word): TJsValue;
  var
    Percentage: TTaskbarProgressPercentage;
  begin
    Result := Undefined;
    CheckParams('setProgressPercentage', Args, ArgCount, [jsNumber], 1);

    try
      Percentage := JsNumberAsInt(Args^);
    except
      on E: Exception do begin
        ThrowError(E.Message, []);
      end;
    end;

    SetTaskbarProgressPercentage(Percentage);
  end;

  function TaskbarIconSetProgressState(Args: PJsValue; ArgCount: Word): TJsValue;
  var
    State: TTaskbarProgressState;
  begin
    Result := Undefined;
    CheckParams('setProgressState', Args, ArgCount, [jsNumber], 1);

    try
      State := TTaskbarProgressState(JsNumberAsInt(Args^));
    except
      on E: Exception do begin
        ThrowError(E.Message, []);
      end;
    end;

    SetTaskbarProgressState(State);
  end;

  function GetTaskbarIcon: TJsValue;
  begin
    Result := CreateObject;

    SetFunction(Result, 'setProgressPercentage', TaskbarIconSetProgressPercentage);
    SetFunction(Result, 'setProgressState', TaskbarIconSetProgressState);
  end;

  function GetJsValue;
  begin
    Result := CreateObject;
    SetFunction(Result, 'expandEnvVar', ShellExpandEnvVar);
    SetFunction(Result, 'exec', ShellExec);
    SetFunction(Result, 'execVerb', ShellExecVerb);
    SetFunction(Result, 'run', ShellRun);

    SetProperty(Result, 'taskbarIcon', GetTaskbarIcon);
  end;

end.