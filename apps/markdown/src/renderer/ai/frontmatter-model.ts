/**
 * Frontmatter stays raw YAML outside the block editor's redaction tree. A
 * persisted withheld span cannot be safely reconstructed there, so refuse the
 * read rather than returning the text beneath that marker.
 */
export function frontmatterHasPrivateMarkup(inner: string): boolean {
  return /data-redaction/i.test(inner)
}
