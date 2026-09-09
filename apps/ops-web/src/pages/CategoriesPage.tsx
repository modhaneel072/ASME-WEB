import { useState } from 'react';
import { useForm } from 'react-hook-form';
import { zodResolver } from '@hookform/resolvers/zod';
import { Archive, Plus, Tags } from 'lucide-react';
import { api, ApiError } from '@/api/client';
import { keys, useCategories, useInvalidate } from '@/api/hooks';
import { useCurrentSession } from '@/auth/SessionProvider';
import { categorySchema, type CategoryInput } from '@/contracts/schemas';
import type { Category } from '@/contracts/types';
import {
  Button,
  ConfirmDialog,
  EmptyState,
  ErrorState,
  PageHeader,
  SideSheet,
  Skeleton,
  TextArea,
  TextField,
  useToast,
} from '@/ui';

const ICONS = [
  'tag',
  'wrench',
  'zap',
  'code',
  'cpu',
  'hammer',
  'search-check',
  'shield-alert',
  'calendar-check',
  'alert-triangle',
  'flag',
  'calendar',
  'shopping-cart',
  'file-text',
  'truck',
  'printer',
];

export function CategoriesPage() {
  const { can } = useCurrentSession();
  const canManage = can('category.manage');
  const categories = useCategories();
  const [editing, setEditing] = useState<Category | 'new' | null>(null);
  const [archiving, setArchiving] = useState<Category | null>(null);
  const invalidate = useInvalidate();
  const toast = useToast();
  const rows = categories.data || [];
  return (
    <div className="page">
      <PageHeader
        title="Categories"
        subtitle="Labels that route and report on work: Mechanical, Electrical, Safety…"
        testId="categories-page"
        actions={
          canManage && (
            <Button
              variant="primary"
              icon={<Plus />}
              onClick={() => setEditing('new')}
              data-testid="new-category"
            >
              New category
            </Button>
          )
        }
      />
      <div className="page-body">
        {categories.isPending ? (
          <Skeleton lines={5} />
        ) : categories.error ? (
          <ErrorState error={categories.error} onRetry={() => categories.refetch()} />
        ) : rows.length === 0 ? (
          <div className="card">
            <EmptyState
              icon={<Tags />}
              title="No categories yet"
              body="Categories tag work orders so reports can group by discipline."
              actions={
                canManage ? (
                  <Button variant="primary" icon={<Plus />} onClick={() => setEditing('new')}>
                    New category
                  </Button>
                ) : undefined
              }
            />
          </div>
        ) : (
          <div className="table-wrap">
            <table className="table" data-testid="categories-table">
              <thead>
                <tr>
                  <th>Category</th>
                  <th>Icon</th>
                  <th className="num">Work orders</th>
                  <th />
                </tr>
              </thead>
              <tbody>
                {rows.map((c) => (
                  <tr
                    key={c.id}
                    className={canManage ? 'clickable' : undefined}
                    onClick={() => canManage && setEditing(c)}
                    data-testid="category-row"
                  >
                    <td>
                      <span className="row">
                        <span className="chip-dot" style={{ background: c.color, width: 12, height: 12 }} />
                        <span style={{ fontWeight: 600 }}>{c.name}</span>
                      </span>
                      {c.description && <div className="text-caption text-muted">{c.description}</div>}
                    </td>
                    <td className="mono">{c.icon}</td>
                    <td className="num">{c.work_order_count}</td>
                    <td onClick={(e) => e.stopPropagation()}>
                      {canManage && (
                        <Button
                          size="sm"
                          variant="ghost"
                          icon={<Archive />}
                          onClick={() => setArchiving(c)}
                          aria-label={`Archive ${c.name}`}
                        />
                      )}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </div>
      {editing && (
        <CategoryForm category={editing === 'new' ? undefined : editing} onClose={() => setEditing(null)} />
      )}
      <ConfirmDialog
        open={!!archiving}
        onClose={() => setArchiving(null)}
        title={`Archive ${archiving?.name}?`}
        body={`${archiving?.work_order_count || 0} work order(s) keep this tag; it just stops being offered for new ones.`}
        confirmLabel="Archive"
        danger
        onConfirm={async () => {
          if (!archiving) return;
          try {
            await api(`/categories/${archiving.id}`, { method: 'DELETE' });
            invalidate(keys.categories, keys.lookups, keys.setup);
            toast.success('Category archived');
          } catch (err) {
            toast.error('Could not archive', err instanceof ApiError ? err.message : undefined);
          }
          setArchiving(null);
        }}
      />
    </div>
  );
}

function CategoryForm({ category, onClose }: { category?: Category; onClose: () => void }) {
  const invalidate = useInvalidate();
  const toast = useToast();
  const [banner, setBanner] = useState<string | null>(null);
  const form = useForm<CategoryInput>({
    resolver: zodResolver(categorySchema),
    defaultValues: {
      name: category?.name || '',
      color: category?.color || '#0878d1',
      icon: category?.icon || 'tag',
      description: category?.description || '',
    },
  });
  const icon = form.watch('icon');
  const submit = form.handleSubmit(async (values) => {
    setBanner(null);
    const data = categorySchema.parse(values);
    try {
      if (category) {
        const body: Record<string, unknown> = {};
        const dirty = form.formState.dirtyFields as Record<string, unknown>;
        for (const key of Object.keys(data) as (keyof CategoryInput)[])
          if (dirty[key]) body[key] = data[key] ?? null;
        if (Object.keys(body).length) await api(`/categories/${category.id}`, { method: 'PATCH', body });
        toast.success('Category updated');
      } else {
        await api('/categories', { method: 'POST', body: data });
        toast.success('Category created');
      }
      invalidate(keys.categories, keys.lookups, keys.setup);
      onClose();
    } catch (err) {
      if (err instanceof ApiError) {
        const fields = err.fieldErrors();
        Object.entries(fields).forEach(([f, m]) => form.setError(f as keyof CategoryInput, { message: m }));
        if (!Object.keys(fields).length) setBanner(err.message);
      }
    }
  });
  const e = form.formState.errors;
  return (
    <SideSheet
      open
      onClose={onClose}
      title={category ? `Edit ${category.name}` : 'New category'}
      testId="category-form"
      footer={
        <>
          <span className="grow" />
          <Button variant="ghost" onClick={onClose}>
            Cancel
          </Button>
          <Button
            variant="primary"
            onClick={submit}
            loading={form.formState.isSubmitting}
            data-testid="category-submit"
          >
            {category ? 'Save changes' : 'Create category'}
          </Button>
        </>
      }
    >
      <form className="stack" style={{ gap: 18 }} onSubmit={submit} noValidate>
        {banner && (
          <div className="form-banner form-banner-error" role="alert">
            {banner}
          </div>
        )}
        <TextField
          label="Name"
          required
          error={e.name?.message}
          {...form.register('name')}
          data-autofocus
          data-testid="category-name"
        />
        <div className="field-row">
          <TextField
            label="Colour"
            type="color"
            error={e.color?.message}
            {...form.register('color')}
            style={{ padding: 4, width: 80 }}
          />
          <div className="field">
            <label className="field-label" htmlFor="category-icon">
              Icon
            </label>
            <select id="category-icon" className="field-control" {...form.register('icon')}>
              {[...new Set([icon, ...ICONS])].filter(Boolean).map((i) => (
                <option key={i} value={i}>
                  {i}
                </option>
              ))}
            </select>
          </div>
        </div>
        <TextArea label="Description" rows={2} {...form.register('description')} />
      </form>
    </SideSheet>
  );
}
