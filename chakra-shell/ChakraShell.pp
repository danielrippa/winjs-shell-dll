unit ChakraShell;

{$mode delphi}

interface

  uses ChakraTypes;

  function GetJsValue: TJsValue;

implementation

  uses Chakra, ChakraUtils, ChakraShellUtils, Variants;

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
  begin
   CheckParams('exec', Args, ArgCount, [jsString, jsString, jsString], 2);

    Executable := JsStringAsString(Args^);
    Inc(Args);

    Params := JsStringAsString(Args^);

    WorkingFolder := '';

    if ArgCount > 2 then begin
      Inc(Args);
      WorkingFolder := JsStringAsString(Args^);
    end;

    Result := ExecuteProcess(Executable, Params, WorkingFolder);
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

  function GetJsValue;
  begin
    Result := CreateObject;
    SetFunction(Result, 'expandEnvVar', ShellExpandEnvVar);
    SetFunction(Result, 'exec', ShellExec);
    SetFunction(Result, 'run', ShellRun);
  end;

end.