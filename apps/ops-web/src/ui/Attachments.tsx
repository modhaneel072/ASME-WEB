import { useRef, useState } from 'react';
import { Download, FileText, Paperclip, Trash2, UploadCloud } from 'lucide-react';
import { api, ApiError } from '@/api/client';
import type { Attachment, UserRef } from '@/contracts/types';
import { fmtBytes, fmtRelative } from '@/lib/format';
import { Button, IconButton } from './Button';
import { ConfirmDialog } from './Dialog';
import { useToast } from './Toast';
import { cn } from '@/lib/cn';

export function AttachmentList({
  files,
  currentUser,
  canDelete,
  onChanged,
}: {
  files: Attachment[];
  currentUser: UserRef;
  canDelete?: (file: Attachment) => boolean;
  onChanged: () => void;
}) {
  const toast = useToast();
  const [confirm, setConfirm] = useState<Attachment | null>(null);
  const [deleting, setDeleting] = useState(false);
  if (files.length === 0) return <div className="text-muted text-label">No files yet.</div>;
  return (
    <div className="file-list">
      {files.map((f) => (
        <div key={f.id} className="file-row" data-testid="file-row">
          {f.kind === 'image' ? (
            <img className="file-thumb" src={f.download_url} alt="" loading="lazy" />
          ) : (
            <span className="file-thumb">
              <FileText />
            </span>
          )}
          <div style={{ flex: 1, minWidth: 0 }}>
            <div className="file-name truncate">{f.original_name}</div>
            <div className="file-meta">
              {fmtBytes(f.size_bytes)} · {f.uploaded_by?.name || 'Unknown'} · {fmtRelative(f.created_at)}
              {f.scan_status &&
                f.scan_status !== 'clean' &&
                f.scan_status !== 'skipped' &&
                ` · scan: ${f.scan_status}`}
            </div>
          </div>
          <div className="inline-actions">
            <a
              className="btn btn-ghost btn-icon btn-sm"
              href={f.download_url}
              target="_blank"
              rel="noreferrer"
              aria-label={`Download ${f.original_name}`}
              title="Download"
            >
              <Download />
            </a>
            {(canDelete ? canDelete(f) : f.uploaded_by?.id === currentUser.id) && (
              <IconButton label={`Delete ${f.original_name}`} size="sm" onClick={() => setConfirm(f)}>
                <Trash2 />
              </IconButton>
            )}
          </div>
        </div>
      ))}
      <ConfirmDialog
        open={!!confirm}
        onClose={() => setConfirm(null)}
        title="Delete file?"
        body={`"${confirm?.original_name}" will be removed. The upload stays in the audit history.`}
        confirmLabel="Delete"
        danger
        loading={deleting}
        onConfirm={async () => {
          if (!confirm) return;
          setDeleting(true);
          try {
            await api(`/files/${confirm.id}`, { method: 'DELETE' });
            toast.success('File deleted');
            onChanged();
          } catch (err) {
            toast.error('Could not delete file', err instanceof ApiError ? err.message : undefined);
          } finally {
            setDeleting(false);
            setConfirm(null);
          }
        }}
      />
    </div>
  );
}

export function AttachmentUploader({
  entityType,
  entityId,
  onUploaded,
  compact = false,
}: {
  entityType: 'work_order' | 'project' | 'asset';
  entityId: string;
  onUploaded: () => void;
  compact?: boolean;
}) {
  const toast = useToast();
  const inputRef = useRef<HTMLInputElement>(null);
  const [dragging, setDragging] = useState(false);
  const [busy, setBusy] = useState(false);

  const upload = async (list: FileList | File[]) => {
    const files = Array.from(list);
    if (files.length === 0) return;
    setBusy(true);
    let ok = 0;
    for (const file of files) {
      const form = new FormData();
      form.append('entity_type', entityType);
      form.append('entity_id', entityId);
      form.append('file', file);
      try {
        await api('/files', { method: 'POST', form });
        ok += 1;
      } catch (err) {
        toast.error(`Could not upload ${file.name}`, err instanceof ApiError ? err.message : undefined);
      }
    }
    setBusy(false);
    if (ok) {
      toast.success(ok === 1 ? 'File uploaded' : `${ok} files uploaded`);
      onUploaded();
    }
    if (inputRef.current) inputRef.current.value = '';
  };

  return (
    <div
      className={cn('uploader', dragging && 'dragging')}
      onDragOver={(e) => {
        e.preventDefault();
        setDragging(true);
      }}
      onDragLeave={() => setDragging(false)}
      onDrop={(e) => {
        e.preventDefault();
        setDragging(false);
        void upload(e.dataTransfer.files);
      }}
      data-testid="uploader"
    >
      <input
        ref={inputRef}
        type="file"
        multiple
        onChange={(e) => e.target.files && upload(e.target.files)}
        aria-label="Choose files"
        data-testid="file-input"
      />
      <div className={compact ? 'row' : 'stack'} style={{ alignItems: 'center', justifyContent: 'center' }}>
        {!compact && <UploadCloud size={22} />}
        <span>Drag files here, or</span>
        <Button size="sm" icon={<Paperclip />} loading={busy} onClick={() => inputRef.current?.click()}>
          Choose files
        </Button>
      </div>
      <div className="text-caption text-muted" style={{ marginTop: 6 }}>
        Images, PDFs, CAD exports and office documents. Links expire and refresh on reload.
      </div>
    </div>
  );
}
