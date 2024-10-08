unit ChakraShell;

{$mode delphi}

interface

  uses ChakraTypes;

  function GetJsValue: TJsValue;

implementation

  uses Chakra, Variants, SysUtils, Win32ShellExecute, Win32Taskbar, ChakraShellUtils, Win32FileDialogs, ChakraError;

  function ShellExpandEnvVarString(Args: PJsValue; ArgCount: Word): TJsValue;
  var
    EnvVarString: WideString;
  begin
    CheckParams('expandEnvVarString', Args, ArgCount, [jsString], 1);
    EnvVarString := JsStringAsString(Args^);
    Result := StringAsJsString(ExpandEnvVarString(EnvVarString));
  end;

  function ShellExpandEnvVar(Args: PJsValue; ArgCount: Word): TJsValue;
  var
    EnvVarName: WideString;
  begin
    CheckParams('expandEnvVar', Args, ArgCount, [jsString], 1);
    EnvVarName := JsStringAsString(Args^);
    Result := StringAsJsString(ExpandEnvVar(EnvVarName));
  end;

  function JsArrayAsWideStringArray(aArray: TJsValue): WideStringArray;
  var
    ArrayLength: Integer;
    jsArrayLength: TJsValue;
    i: Integer;
    Item: TJsValue;
  begin
    Result := [];

    ArrayLength := GetIntProperty(aArray, 'length');

    if ArrayLength > 0 then begin

      SetLength(Result, ArrayLength);

      for i := 0 to ArrayLength - 1 do begin
        Item := GetArrayItem(aArray, i);
        Result[i] := JsValueAsString(Item);
      end;

    end;

  end;

  function ShellExec(Args: PJsValue; ArgCount: Word): TJsValue;
  var
    Executable, WorkingFolder: WideString;
    Params: array of WideString;
    Priority, BufferSize: Integer;
    OutputProc: TJsValue;
  begin
    CheckParams('exec', Args, ArgCount, [jsString, jsArray, jsString, jsNumber, jsNumber, jsFunction], 6);

    Executable := JsStringAsString(Args^); Inc(Args);
    Params := JsArrayAsWideStringArray(Args^); Inc(Args);
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
    CheckParams('run', Args, ArgCount, [jsString, jsBoolean], 1);

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

  function ShellFileOpenDialog(Args: PJsValue; ArgCount: Word): TJsValue;
  var
    aFileName, aDefaultExtension, aFilter, aInitialFolder, aTitle: WideString;
    aFilterIndex: Integer;
  begin
    CheckParams('fileOpenDialog', Args, ArgCount, [jsString, jsString, jsString, jsString, jsString, jsNumber], 6);

    aFileName := JsStringAsString(Args^); Inc(Args);
    aDefaultExtension := JsStringAsString(Args^); Inc(Args);
    aFilter := JsStringAsString(Args^); Inc(Args);
    aInitialFolder := JsStringAsString(Args^); Inc(Args);
    aTitle := JsStringAsString(Args^); Inc(Args);

    aFilterIndex := JsNumberAsInt(Args^);

    Result := StringAsJsString(FileOpenDialog(aFileName, aDefaultExtension, aFilter, aInitialFolder, aTitle, aFilterIndex));
  end;

  function ShellFileSaveDialog(Args: PJsValue; ArgCount: Word): TJsValue;
  var
    aFileName, aDefaultExtension, aFilter, aInitialFolder, aTitle: WideString;
    aFilterIndex: Integer;
  begin
    CheckParams('fileSaveDialog', Args, ArgCount, [jsString, jsString, jsString, jsString, jsString, jsNumber], 6);

    aFileName := JsStringAsString(Args^); Inc(Args);
    aDefaultExtension := JsStringAsString(Args^); Inc(Args);
    aFilter := JsStringAsString(Args^); Inc(Args);
    aInitialFolder := JsStringAsString(Args^); Inc(Args);
    aTitle := JsStringAsString(Args^); Inc(Args);

    aFilterIndex := JsNumberAsInt(Args^);

    Result := StringAsJsString(FileSaveDialog(aFileName, aDefaultExtension, aFilter, aInitialFolder, aTitle, aFilterIndex));
  end;

  function ShellStart(Args: PJsValue; ArgCount: Word): TJsValue;
  var
    Executable, WorkingFolder: WideString;
    Params: array of WideString;
    Priority: Integer;
    WaitOnReturn: Boolean;
  begin
    CheckParams('exec', Args, ArgCount, [jsString, jsArray, jsString, jsNumber, jsBoolean], 5);

    Executable := JsStringAsString(Args^); Inc(Args);
    Params := JsArrayAsWideStringArray(Args^); Inc(Args);
    WorkingFolder := JsStringAsString(Args^); Inc(Args);

    Priority := JsNumberAsInt(Args^); Inc(Args);
    WaitOnReturn := JsBooleanAsBoolean(Args^);

    Result := StartProcess(Executable, Params, WorkingFolder, Priority, WaitOnReturn);
  end;

  function GetJsValue;
  begin
    Result := CreateObject;
    SetFunction(Result, 'expandEnvVar', ShellExpandEnvVar);
    SetFunction(Result, 'expandEnvVarString', ShellExpandEnvVarString);

    SetFunction(Result, 'exec', ShellExec);
    SetFunction(Result, 'start', ShellStart);
    SetFunction(Result, 'execVerb', ShellExecVerb);
    SetFunction(Result, 'run', ShellRun);

    SetProperty(Result, 'taskbarIcon', GetTaskbarIcon);

    SetFunction(Result, 'fileOpenDialog', ShellFileOpenDialog);
    SetFunction(Result, 'fileSaveDialog', ShellFileSaveDialog);
  end;

end.
