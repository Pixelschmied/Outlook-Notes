import type { Selection } from './types'

/**
 * Derive a *stable* key for the selected mail. We prefer `internetMessageId`
 * (the RFC 5322 Message-ID header): it is globally unique, does not change when
 * a mail is moved between folders, and survives restarts — exactly what we need
 * so a mail always maps back to the same note. We fall back to the item id and
 * finally the conversation id when the message id is not exposed.
 */
export function readSelection(): Selection {
  const item = Office.context.mailbox?.item
  if (!item) return { status: 'none' }

  // Compose items and unusual item types have no stable message id we can trust.
  const messageId = (item as Office.MessageRead).internetMessageId
  const key = messageId || item.itemId || item.conversationId
  if (!key) return { status: 'unsupported' }

  const subject = item.subject ?? ''
  const from = (item as Office.MessageRead).from
  const sender = from ? from.displayName || from.emailAddress || '' : ''

  return { status: 'ok', mailKey: key, subject, sender }
}
