unit Win32Taskbar;

  {$mode delphi}

interface

  type

    TTaskBarProgressState = (tbpsNone, tbpsIndeterminate, tbpsNormal, tbpsError, tbpsPaused);
    TTaskbarProgressPercentage = 0..100;

  procedure SetTaskbarProgressState(aState: TTaskBarProgressState);
  procedure SetTaskbarProgressPercentage(aPercentage: TTaskbarProgressPercentage);

implementation

  uses
    Windows, ComObj, SysUtils, ChakraErr;

  type

    // https://github.com/nightcap79/cheat-engine/blob/89bab360fdb2a029dbe8ff0a27ebc3a301529d36/Cheat%20Engine/windows7taskbar.pas

    ITaskBarList3 = interface(IUnknown)
    ['{EA1AFB91-9E28-4B86-90E9-9E9F8A5EEFAF}']
      procedure HrInit(); stdcall;
      procedure AddTab(hwnd: THandle); stdcall;
      procedure DeleteTab(hwnd: THandle); stdcall;
      procedure ActivateTab(hwnd: THandle); stdcall;
      procedure SetActiveAlt(hwnd: THandle); stdcall;

      procedure MarkFullscreenWindow(hwnd: THandle; fFullscreen: Boolean); stdcall;

      procedure SetProgressValue(hwnd: THandle; ullCompleted: UInt64; ullTotal: UInt64); stdcall;
      procedure SetProgressState(hwnd: THandle; tbpFlags: Cardinal); stdcall;

      procedure RegisterTab(hwnd: THandle; hwndMDI: THandle); stdcall;
      procedure UnregisterTab(hwndTab: THandle); stdcall;
      procedure SetTabOrder(hwndTab: THandle; hwndInsertBefore: THandle); stdcall;
      procedure SetTabActive(hwndTab: THandle; hwndMDI: THandle; tbatFlags: Cardinal); stdcall;
      procedure ThumbBarAddButtons(hwnd: THandle; cButtons: Cardinal; pButtons: Pointer); stdcall;
      procedure ThumbBarUpdateButtons(hwnd: THandle; cButtons: Cardinal; pButtons: Pointer); stdcall;
      procedure ThumbBarSetImageList(hwnd: THandle; himl: THandle); stdcall;
      procedure SetOverlayIcon(hwnd: THandle; hIcon: THandle; pszDescription: PChar); stdcall;
      procedure SetThumbnailTooltip(hwnd: THandle; pszDescription: PChar); stdcall;
      procedure SetThumbnailClip(hwnd: THandle; var prcClip: TRect); stdcall;
    end;

  var
    TaskbarList: ITaskbarList3;

  function GetTaskbarList: ITaskbarList3;
  const
    CLSID_TaskbarList: TGUID = '{56FDF344-FD6D-11d0-958A-006097C9A090}';
  var
    aInterface: IInterface;
  begin
    if TaskbarList = Nil then begin

      CoInitializeEx(Nil, 0);

      TaskbarList := CreateComObject(CLSID_TaskbarList) as ITaskbarList3;
      TaskbarList.HrInit;

    end;

    Result := TaskbarList;

  end;

  procedure SetTaskbarProgressState;
  const
    ProgressState: Array[TTaskbarProgressState] of Cardinal = (0, 1, 2, 4, 8);
  begin
    try
      GetTaskbarList.SetProgressState(GetConsoleWindow, ProgressState[aState]);
    except
      on E: Exception do
        ThrowError(E.Message, []);
    end;
  end;

  procedure SetTaskbarProgressPercentage;
  var
    CurrentValue, MaximumValue: UInt64;
  begin
    CurrentValue := aPercentage;
    MaximumValue := 100;

    try
      GetTaskbarList.SetProgressValue(GetConsoleWindow, CurrentValue, MaximumValue);
    except
      on E: Exception do
        ThrowError(E.Message, []);
    end;
  end;

  finalization

    if TaskbarList <> nil then begin
      TaskBarList := Nil;
    end;

end.
