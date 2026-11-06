import type { SetStateAction } from 'react'
import type { CommentInfo, NoteInfo } from '@genoffice/docx-engine'

export interface DocsReviewState {
  comments: CommentInfo[]
  footnotes: NoteInfo[]
  endnotes: NoteInfo[]
  commentsDirty: boolean
  notesDirty: boolean
}

/** App-owned review state. Like React state, its arrays and entries are immutable. */
export class DocsReviewStore {
  private state: DocsReviewState = {
    comments: [],
    footnotes: [],
    endnotes: [],
    commentsDirty: false,
    notesDirty: false,
  }
  private listeners = new Set<() => void>()
  private contentRevision = 0

  get revision(): number {
    return this.contentRevision
  }

  getSnapshot = (): DocsReviewState => this.state

  subscribe = (listener: () => void): (() => void) => {
    this.listeners.add(listener)
    return () => {
      this.listeners.delete(listener)
    }
  }

  private update(patch: Partial<DocsReviewState>): void {
    const previous = this.state
    const next = { ...previous, ...patch }
    if (
      Object.keys(patch).every((key) =>
        Object.is(previous[key as keyof DocsReviewState], next[key as keyof DocsReviewState]),
      )
    )
      return
    if (
      previous.comments !== next.comments ||
      previous.footnotes !== next.footnotes ||
      previous.endnotes !== next.endnotes
    )
      this.contentRevision++
    this.state = next
    for (const listener of this.listeners) listener()
  }

  private set<K extends keyof DocsReviewState>(
    key: K,
    action: SetStateAction<DocsReviewState[K]>,
  ): void {
    const value = typeof action === 'function' ? action(this.state[key]) : action
    this.update({ [key]: value })
  }

  setComments = (action: SetStateAction<CommentInfo[]>): void => this.set('comments', action)
  setFootnotes = (action: SetStateAction<NoteInfo[]>): void => this.set('footnotes', action)
  setEndnotes = (action: SetStateAction<NoteInfo[]>): void => this.set('endnotes', action)
  setCommentsDirty = (action: SetStateAction<boolean>): void => this.set('commentsDirty', action)
  setNotesDirty = (action: SetStateAction<boolean>): void => this.set('notesDirty', action)

  /** Only retained in this renderer; never serialized into a conversation or IPC reply. */
  capture(): DocsReviewState {
    return { ...this.state }
  }

  restore(snapshot: DocsReviewState): void {
    const current = this.state
    this.update({
      ...snapshot,
      // Restoring after a save must write the restored lists on the next save.
      commentsDirty:
        current.commentsDirty || snapshot.commentsDirty || current.comments !== snapshot.comments,
      notesDirty:
        current.notesDirty ||
        snapshot.notesDirty ||
        current.footnotes !== snapshot.footnotes ||
        current.endnotes !== snapshot.endnotes,
    })
  }
}
