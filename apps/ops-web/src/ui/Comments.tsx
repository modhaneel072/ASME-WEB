import { useState } from 'react';
import { Send } from 'lucide-react';
import { Avatar } from './Avatar';
import { Button } from './Button';
import { ConfirmDialog } from './Dialog';
import type { Comment, UserRef } from '@/contracts/types';
import { fmtRelative } from '@/lib/format';

export function CommentThread({
  comments,
  currentUser,
  canComment,
  onSubmit,
  onEdit,
  onDelete,
  submitting,
}: {
  comments: Comment[];
  currentUser: UserRef;
  canComment: boolean;
  onSubmit: (body: string) => Promise<unknown>;
  onEdit?: (comment: Comment, body: string) => Promise<unknown>;
  onDelete?: (comment: Comment) => Promise<unknown>;
  submitting?: boolean;
}) {
  const [editing, setEditing] = useState<Comment | null>(null);
  const [draft, setDraft] = useState('');
  const [confirm, setConfirm] = useState<Comment | null>(null);
  return (
    <div>
      {comments.length === 0 && <div className="text-muted text-label">No comments yet.</div>}
      {comments.map((c) => (
        <div key={c.id} className="comment" data-testid="comment">
          <Avatar user={c.author} size="sm" />
          <div className="comment-body-wrap">
            <div className="comment-head">
              <span className="comment-author">{c.author.name}</span>
              <span className="comment-time">
                {fmtRelative(c.created_at)}
                {c.edited_at && ' · edited'}
              </span>
            </div>
            {editing?.id === c.id ? (
              <div className="composer-box" style={{ marginTop: 6 }}>
                <textarea
                  className="field-control"
                  value={draft}
                  onChange={(e) => setDraft(e.target.value)}
                  rows={3}
                  aria-label="Edit comment"
                />
                <div className="row">
                  <Button
                    size="sm"
                    variant="primary"
                    onClick={async () => {
                      await onEdit?.(c, draft.trim());
                      setEditing(null);
                    }}
                    disabled={!draft.trim()}
                  >
                    Save
                  </Button>
                  <Button size="sm" variant="ghost" onClick={() => setEditing(null)}>
                    Cancel
                  </Button>
                </div>
              </div>
            ) : c.deleted ? (
              <div className="comment-text comment-deleted">This comment was removed.</div>
            ) : (
              <div className="comment-text">{c.body}</div>
            )}
            {!c.deleted && c.author.id === currentUser.id && editing?.id !== c.id && (onEdit || onDelete) && (
              <div className="comment-actions">
                {onEdit && (
                  <button
                    type="button"
                    className="link-button text-caption"
                    onClick={() => {
                      setEditing(c);
                      setDraft(c.body);
                    }}
                  >
                    Edit
                  </button>
                )}
                {onDelete && (
                  <button
                    type="button"
                    className="link-button text-caption"
                    style={{ color: 'var(--color-danger)' }}
                    onClick={() => setConfirm(c)}
                  >
                    Delete
                  </button>
                )}
              </div>
            )}
          </div>
        </div>
      ))}
      {canComment ? (
        <CommentComposer user={currentUser} onSubmit={onSubmit} submitting={submitting} />
      ) : (
        <div className="text-muted text-caption" style={{ paddingTop: 8 }}>
          Your role can view this discussion but not post in it.
        </div>
      )}
      <ConfirmDialog
        open={!!confirm}
        onClose={() => setConfirm(null)}
        title="Delete comment?"
        body="The comment will be marked as removed for everyone. This is recorded in the history."
        confirmLabel="Delete"
        danger
        onConfirm={async () => {
          if (confirm) await onDelete?.(confirm);
          setConfirm(null);
        }}
      />
    </div>
  );
}

export function CommentComposer({
  user,
  onSubmit,
  submitting,
}: {
  user: UserRef;
  onSubmit: (body: string) => Promise<unknown>;
  submitting?: boolean;
}) {
  const [body, setBody] = useState('');
  const send = async () => {
    const text = body.trim();
    if (!text) return;
    await onSubmit(text);
    setBody('');
  };
  return (
    <div className="composer">
      <Avatar user={user} size="sm" />
      <div className="composer-box">
        <textarea
          className="field-control"
          placeholder="Write a comment… (Ctrl+Enter to send)"
          rows={2}
          value={body}
          onChange={(e) => setBody(e.target.value)}
          onKeyDown={(e) => {
            if ((e.ctrlKey || e.metaKey) && e.key === 'Enter') void send();
          }}
          aria-label="New comment"
          data-testid="comment-input"
        />
        <div className="composer-foot">
          <span className="text-caption text-muted">Visible to everyone who can see this item.</span>
          <Button
            size="sm"
            variant="primary"
            icon={<Send />}
            onClick={send}
            disabled={!body.trim()}
            loading={submitting}
            data-testid="comment-submit"
          >
            Comment
          </Button>
        </div>
      </div>
    </div>
  );
}
