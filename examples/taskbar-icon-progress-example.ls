
  { taskbar-icon: { set-progress-state, set-progress-percentage } } = winjs.load-library 'WinjsShell.dll'

  taskbar-icon-state =

    none: 0
    indeterminate: 1
    normal: 2
    error: 4
    paused: 8

  set-progress-state taskbar-icon-state.error

  for percentage til 100

    process.sleep 500

    set-progress-percentage percentage

  set-progress-state taskbar-icon-state.none
