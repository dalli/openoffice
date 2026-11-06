/** Use the studio's existing history; never allocate snapshots outside its resource budget. */
export function createWorkspaceUndoHistory(host) {
  let pending = null
  const command = 'ext:ieum-workspace-history'
  try {
    host.automation.registerCommand({
      id: command,
      label: 'Check workspace undo history',
      execute(services) {
        pending?.(services)
      },
    })
  } catch {
    // A host without the verified studio hooks cannot offer whole-run undo.
  }
  const access = (fn) => {
    let result = null
    pending = (services) => {
      const history = services.getInputHandler()?.history
      const capacity = services.wasm.snapshotCapacity()
      if (Array.isArray(history?.undoStack) && Number.isSafeInteger(capacity)) {
        result = fn(history, capacity, host.automation.getContext())
      }
    }
    try {
      host.automation.execute(command)
      return result
    } finally {
      pending = null
    }
  }
  return {
    capture: () => access((history) => history),
    undo(checkpoint, steps) {
      return (
        access((history, capacity, context) => {
          const suffix = history.undoStack.slice(-steps)
          if (
            history !== checkpoint ||
            context.isEditable !== true ||
            context.isFormMode === true ||
            !Number.isSafeInteger(steps) ||
            steps < 1 ||
            // Undo first saves an after-snapshot. Keep the same two-slot headroom
            // as CommandHistory, plus that first save, so no pending step is evicted.
            steps > capacity - 3 ||
            suffix.length !== steps ||
            suffix.some(
              (entry) =>
                !entry.type?.startsWith('snapshot:plugin:ieum-agent:') ||
                entry.snapshotResourceCount?.() !== 1,
            )
          )
            return false
          for (let index = 0; index < steps; index++) {
            const expected = suffix[steps - index - 1]
            if (history.peekUndoTop() !== expected) return false
            const result = host.automation.execute('edit:undo')
            if (result?.ok !== true || history.peekRedoTop() !== expected) return false
          }
          return true
        }) === true
      )
    },
  }
}
