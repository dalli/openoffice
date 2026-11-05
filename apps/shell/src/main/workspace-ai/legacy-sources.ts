import type { WebContents } from 'electron'
import type { LegacyChatEvent, LegacyChatRef } from '../../../../docs/src/main/legacy-chat-events'
import type { WorkspacePanelSource } from '../../shared/workspace-panel-api'

/** Keeps only live host references. Chat bodies and pending messages stay in ProjectStore. */
export class WorkspaceLegacySources {
  private readonly hosts = new Map<WebContents, { ref: LegacyChatRef; off(): void }>()

  receive(event: LegacyChatEvent): void {
    if (event.kind !== 'resolved' || event.sender.isDestroyed()) return
    const { sender } = event
    this.hosts.get(sender)?.off()
    const remove = () => {
      this.hosts.get(sender)?.off()
      this.hosts.delete(sender)
    }
    const navigate = (_event: unknown, _url: string, inPlace: boolean, mainFrame: boolean) => {
      if (mainFrame && !inPlace) remove()
    }
    sender.on('destroyed', remove)
    sender.on('render-process-gone', remove)
    sender.on('did-start-navigation', navigate)
    this.hosts.set(sender, {
      ref: { ...event.ref },
      off() {
        sender.removeListener('destroyed', remove)
        sender.removeListener('render-process-gone', remove)
        sender.removeListener('did-start-navigation', navigate)
      },
    })
  }

  sources(titleFor: (sender: WebContents) => string | undefined): WorkspacePanelSource[] {
    return [...this.hosts].flatMap(([sender, { ref }]) => {
      const title = !sender.isDestroyed() ? titleFor(sender) : undefined
      return title === undefined ? [] : [{ origin: { kind: 'legacy-document', ...ref }, title }]
    })
  }
}
