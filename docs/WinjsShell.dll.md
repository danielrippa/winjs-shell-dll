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
  
```
