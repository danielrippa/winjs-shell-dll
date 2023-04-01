# WinjShell.dll

```
get-js-value: ->

  expand-env-var: (var-name: string) -> string
  exec: (executable: string, parameters: string, working-folder: string, priority: number, buffer-size: number, output-callback: function) ->
  
    stdout: string
    stderr: string
    exit-status: number
    
  exec-verb: (verb: string, filename: string, working-folder: string, window-state: number, flags: number, monitor: number, window-handle: number) ->
  
  taskbar-icon:
    
    set-progress-percentage: (percentage: number) -> void
    set-progress-state: (state: number) -> void
  
  file-open-dialog: (filename: string, default-extension: string, filter: string, initial-folder: string, title: string, filter-index: number): string
  
  file-save-dialog: (filename: string, default-extension: string, filter: string, initial-folder: string, title: string, filter-index: number): string
  
```
