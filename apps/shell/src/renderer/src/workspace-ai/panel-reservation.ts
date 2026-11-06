import type { WorkspaceLayout, WorkspacePanelLayout } from '../../../shared/workspace-layout'
import type { WorkspacePanelClient } from './panel-client'

/** Serializes reservations and ignores acknowledgements from replaced or closed hosts. */
export class WorkspacePanelReservation {
  private live = true
  private epoch = 0
  private pending = false
  private pushed: WorkspaceLayout | null = null
  private tail: Promise<void> = Promise.resolve()
  private readonly off: () => void

  constructor(
    private readonly client: Pick<WorkspacePanelClient, 'layout' | 'subscribe'>,
    private readonly change: (layout: WorkspaceLayout | null) => void,
    private readonly failed: () => void,
  ) {
    this.off = client.subscribe((event) => {
      if (!this.live) return
      if (event.kind === 'closed') {
        this.dispose()
        this.change(null)
      } else if (event.kind === 'layout') {
        if (this.pending) this.pushed = event.layout
        else this.change(event.layout)
      }
    })
  }

  set(panel: WorkspacePanelLayout | null): void {
    const epoch = ++this.epoch
    this.pending = true
    this.tail = this.tail.then(async () => {
      if (!this.live || epoch !== this.epoch) return
      this.pushed = null
      try {
        const layout = await this.client.layout(panel)
        if (this.live && epoch === this.epoch) {
          this.pending = false
          this.change(this.pushed ?? layout)
        }
      } catch {
        if (this.live && epoch === this.epoch) {
          this.dispose()
          this.change(null)
          this.failed()
        }
      }
    })
  }

  dispose(): void {
    this.live = false
    this.off()
  }
}
