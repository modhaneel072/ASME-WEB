import { useMemo, useState } from 'react';
import { Controller, useForm } from 'react-hook-form';
import { zodResolver } from '@hookform/resolvers/zod';
import { Archive, MapPin, Plus, Star } from 'lucide-react';
import { api, ApiError } from '@/api/client';
import { keys, useInvalidate, useLocations } from '@/api/hooks';
import { useCurrentSession } from '@/auth/SessionProvider';
import { locationSchema, type LocationInput } from '@/contracts/schemas';
import type { Location } from '@/contracts/types';
import {
  Badge,
  Button,
  ConfirmDialog,
  EmptyState,
  ErrorState,
  PageHeader,
  Picker,
  SideSheet,
  Skeleton,
  TextArea,
  TextField,
  useToast,
} from '@/ui';

interface Node {
  location: Location;
  depth: number;
}

function flatten(locations: Location[]): Node[] {
  const byParent = new Map<string | null, Location[]>();
  for (const l of locations) {
    const key =
      l.parent_location_id && locations.some((x) => x.id === l.parent_location_id)
        ? l.parent_location_id
        : null;
    if (!byParent.has(key)) byParent.set(key, []);
    byParent.get(key)!.push(l);
  }
  const out: Node[] = [];
  const walk = (parent: string | null, depth: number) => {
    for (const l of byParent.get(parent) || []) {
      out.push({ location: l, depth });
      walk(l.id, depth + 1);
    }
  };
  walk(null, 0);
  return out;
}

export function LocationsPage() {
  const { can } = useCurrentSession();
  const canManage = can('location.manage');
  const locations = useLocations();
  const [editing, setEditing] = useState<Location | 'new' | null>(null);
  const [archiving, setArchiving] = useState<Location | null>(null);
  const invalidate = useInvalidate();
  const toast = useToast();
  const nodes = useMemo(() => flatten(locations.data || []), [locations.data]);

  const setDefault = async (l: Location) => {
    try {
      await api(`/locations/${l.id}`, { method: 'PATCH', body: { is_default: true } });
      invalidate(keys.locations, keys.lookups);
      toast.success(`${l.name} is now the default location`);
    } catch (err) {
      toast.error('Could not update', err instanceof ApiError ? err.message : undefined);
    }
  };

  return (
    <div className="page">
      <PageHeader
        title="Locations"
        subtitle="Where assets live and where work happens. Nest rooms under buildings."
        testId="locations-page"
        actions={
          canManage && (
            <Button
              variant="primary"
              icon={<Plus />}
              onClick={() => setEditing('new')}
              data-testid="new-location"
            >
              New location
            </Button>
          )
        }
      />
      <div className="page-body">
        {locations.isPending ? (
          <Skeleton lines={5} />
        ) : locations.error ? (
          <ErrorState error={locations.error} onRetry={() => locations.refetch()} />
        ) : nodes.length === 0 ? (
          <div className="card">
            <EmptyState
              icon={<MapPin />}
              title="No locations yet"
              body="Add the shop, lab and storage rooms so assets and work orders have a home."
              actions={
                canManage ? (
                  <Button variant="primary" icon={<Plus />} onClick={() => setEditing('new')}>
                    New location
                  </Button>
                ) : undefined
              }
            />
          </div>
        ) : (
          <div className="table-wrap">
            <table className="table" data-testid="locations-table">
              <thead>
                <tr>
                  <th>Location</th>
                  <th>Building / room</th>
                  <th className="num">Assets</th>
                  <th />
                </tr>
              </thead>
              <tbody>
                {nodes.map(({ location: l, depth }) => (
                  <tr
                    key={l.id}
                    className={canManage ? 'clickable' : undefined}
                    onClick={() => canManage && setEditing(l)}
                    data-testid="location-row"
                  >
                    <td>
                      <span className="tree-row" style={{ paddingLeft: depth * 18 }}>
                        <MapPin size={14} className="text-muted" />
                        <span style={{ fontWeight: 600 }}>{l.name}</span>
                        {l.is_default && <Badge tone="info">Default</Badge>}
                      </span>
                      {l.description && (
                        <div className="text-caption text-muted" style={{ paddingLeft: depth * 18 + 22 }}>
                          {l.description}
                        </div>
                      )}
                    </td>
                    <td>{[l.building, l.room].filter(Boolean).join(' · ') || '—'}</td>
                    <td className="num">{l.asset_count}</td>
                    <td onClick={(e) => e.stopPropagation()}>
                      {canManage && (
                        <span className="inline-actions">
                          {!l.is_default && (
                            <Button
                              size="sm"
                              variant="ghost"
                              icon={<Star />}
                              onClick={() => setDefault(l)}
                              aria-label={`Make ${l.name} the default`}
                              title="Make default"
                            />
                          )}
                          {!l.is_default && (
                            <Button
                              size="sm"
                              variant="ghost"
                              icon={<Archive />}
                              onClick={() => setArchiving(l)}
                              aria-label={`Archive ${l.name}`}
                              title="Archive"
                            />
                          )}
                        </span>
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
        <LocationForm
          location={editing === 'new' ? undefined : editing}
          all={locations.data || []}
          onClose={() => setEditing(null)}
        />
      )}
      <ConfirmDialog
        open={!!archiving}
        onClose={() => setArchiving(null)}
        title={`Archive ${archiving?.name}?`}
        body="Archived locations are hidden from selectors. Assets keep their history."
        confirmLabel="Archive"
        danger
        onConfirm={async () => {
          if (!archiving) return;
          try {
            await api(`/locations/${archiving.id}`, { method: 'DELETE' });
            invalidate(keys.locations, keys.lookups, keys.setup);
            toast.success('Location archived');
          } catch (err) {
            toast.error('Could not archive', err instanceof ApiError ? err.message : undefined);
          }
          setArchiving(null);
        }}
      />
    </div>
  );
}

function LocationForm({
  location,
  all,
  onClose,
}: {
  location?: Location;
  all: Location[];
  onClose: () => void;
}) {
  const invalidate = useInvalidate();
  const toast = useToast();
  const [banner, setBanner] = useState<string | null>(null);
  const form = useForm<LocationInput>({
    resolver: zodResolver(locationSchema),
    defaultValues: {
      name: location?.name || '',
      description: location?.description || '',
      parent_location_id: location?.parent_location_id || '',
      building: location?.building || '',
      room: location?.room || '',
    },
  });
  const submit = form.handleSubmit(async (values) => {
    setBanner(null);
    const data = locationSchema.parse(values);
    try {
      if (location) {
        const body: Record<string, unknown> = {};
        const dirty = form.formState.dirtyFields as Record<string, unknown>;
        for (const key of Object.keys(data) as (keyof LocationInput)[])
          if (dirty[key]) body[key] = data[key] ?? null;
        if (Object.keys(body).length) await api(`/locations/${location.id}`, { method: 'PATCH', body });
        toast.success('Location updated');
      } else {
        await api('/locations', { method: 'POST', body: data });
        toast.success('Location created');
      }
      invalidate(keys.locations, keys.lookups, keys.setup);
      onClose();
    } catch (err) {
      if (err instanceof ApiError) {
        const fields = err.fieldErrors();
        Object.entries(fields).forEach(([f, m]) => form.setError(f as keyof LocationInput, { message: m }));
        if (!Object.keys(fields).length) setBanner(err.message);
      }
    }
  });
  const e = form.formState.errors;
  return (
    <SideSheet
      open
      onClose={onClose}
      title={location ? `Edit ${location.name}` : 'New location'}
      testId="location-form"
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
            data-testid="location-submit"
          >
            {location ? 'Save changes' : 'Create location'}
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
          data-testid="location-name"
        />
        <Controller
          control={form.control}
          name="parent_location_id"
          render={({ field }) => (
            <Picker
              label="Inside"
              multiple={false}
              options={all
                .filter((l) => l.id !== location?.id)
                .map((l) => ({
                  value: l.id,
                  label: l.name,
                  hint: [l.building, l.room].filter(Boolean).join(' · ') || undefined,
                }))}
              value={field.value ? [field.value] : []}
              onChange={(v) => field.onChange(v[0] || '')}
              placeholder="Top level"
            />
          )}
        />
        <div className="field-row">
          <TextField label="Building" {...form.register('building')} />
          <TextField label="Room" {...form.register('room')} />
        </div>
        <TextArea label="Description" rows={2} {...form.register('description')} />
      </form>
    </SideSheet>
  );
}
