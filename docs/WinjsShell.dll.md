# WinjsWeb.dll

```
get-js-value: ->

  get-connection-state: ->

    connection-name: string
    is-connected: boolean
    is-configured: boolean
    is-offline: boolean
    is-proxy: boolean

  get-content: (url: string) ->

    status-code: number
    response-string: string

  get-file: (url: string, filename: string) -> number
  
  file-open-dialog: (filename: string, default-extension: string, filter: string, initial-folder: string, title: string, filter-index: number): string
  
  file-save-dialog: (filename: string, default-extension: string, filter: string, initial-folder: string, title: string, filter-index: number): string
  
```
