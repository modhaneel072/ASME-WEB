import { useState } from 'react';
import { Controller, useForm } from 'react-hook-form';
import { zodResolver } from '@hookform/resolvers/zod';
import { api, ApiError } from '@/api/client';
import { keys, useInvalidate, useLookups } from '@/api/hooks';
import { projectSchema, type ProjectInput } from '@/contracts/schemas';
import type { ProjectDetail } from '@/contracts/types';
import { userOptions } from '@/lib/options';
import { Button, Picker, SelectField, SideSheet, TextArea, TextField } from '@/ui';
import { useToast } from '@/ui';

const EDITABLE: (keyof ProjectInput)[] = [
  'name',
  'code',
  'description',
  'status',
  'risk_level',
  'lead_user_id',
  'faculty_advisor_user_id',
  'competition',
  'academic_year',
  'start_date',
  'target_date',
  'budget_amount',
  'budget_code',
  'repository_url',
  'cad_url',
  'requirements_url',
  'visibility',
];

function initial(project?: ProjectDetail): ProjectInput {
  return {
    name: project?.name || '',
    code: project?.code || '',
    description: project?.description || '',
    status: project?.status || 'active',
    risk_level: project?.risk_level || 'medium',
    lead_user_id: project?.lead?.id ?? null,
    faculty_advisor_user_id: project?.faculty_advisor?.id ?? null,
    competition: project?.competition || '',
    academic_year: project?.academic_year || '',
    start_date: project?.start_date || '',
    target_date: project?.target_date || '',
    budget_amount: project?.budget_amount ?? undefined,
    budget_code: project?.budget_code || '',
    repository_url: project?.repository_url || '',
    cad_url: project?.cad_url || '',
    requirements_url: project?.requirements_url || '',
    visibility: (project?.visibility as ProjectInput['visibility']) || 'members',
  };
}

export function ProjectForm({
  project,
  onClose,
  onSaved,
}: {
  project?: ProjectDetail;
  onClose: () => void;
  onSaved: (project: ProjectDetail) => void;
}) {
  const toast = useToast();
  const lookups = useLookups();
  const invalidate = useInvalidate();
  const [banner, setBanner] = useState<string | null>(null);
  const form = useForm<ProjectInput>({
    resolver: zodResolver(projectSchema),
    defaultValues: initial(project),
  });
  const e = form.formState.errors;

  const submit = form.handleSubmit(async (values) => {
    setBanner(null);
    const data = projectSchema.parse(values);
    try {
      let saved: ProjectDetail;
      if (project) {
        const body: Record<string, unknown> = {};
        const dirty = form.formState.dirtyFields as Record<string, unknown>;
        for (const key of EDITABLE)
          if (dirty[key]) body[key] = (data as Record<string, unknown>)[key] ?? null;
        if (Object.keys(body).length === 0) return onClose();
        saved = await api<ProjectDetail>(`/projects/${project.id}`, { method: 'PATCH', body });
        invalidate(['projects'], keys.project(project.id), keys.lookups);
        toast.success('Project updated');
      } else {
        saved = await api<ProjectDetail>('/projects', { method: 'POST', body: data });
        invalidate(['projects'], keys.lookups, keys.setup);
        toast.success('Project created', saved.name);
      }
      onSaved(saved);
    } catch (err) {
      if (err instanceof ApiError) {
        const fields = err.fieldErrors();
        Object.entries(fields).forEach(([f, m]) => form.setError(f as keyof ProjectInput, { message: m }));
        if (!Object.keys(fields).length) setBanner(err.message);
      } else setBanner('Could not save. Try again.');
    }
  });

  return (
    <SideSheet
      open
      onClose={onClose}
      title={project ? `Edit ${project.name}` : 'New project'}
      subtitle={project ? undefined : 'A project groups teams, assets, milestones and work orders.'}
      testId="project-form"
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
            data-testid="project-submit"
          >
            {project ? 'Save changes' : 'Create project'}
          </Button>
        </>
      }
    >
      <form className="stack" style={{ gap: 20 }} onSubmit={submit} noValidate>
        {banner && (
          <div className="form-banner form-banner-error" role="alert">
            {banner}
          </div>
        )}
        <div className="field-row">
          <TextField
            label="Name"
            required
            error={e.name?.message}
            {...form.register('name')}
            data-autofocus
            data-testid="project-name"
            wrapClassName="grid-span"
          />
        </div>
        <div className="field-row">
          <TextField
            label="Code"
            placeholder="Auto from name"
            hint="Short code used in work orders, e.g. CCR"
            error={e.code?.message}
            {...form.register('code')}
          />
          <SelectField
            label="Status"
            options={[
              { value: 'planning', label: 'Planning' },
              { value: 'active', label: 'Active' },
              { value: 'on_hold', label: 'On hold' },
              { value: 'completed', label: 'Completed' },
              { value: 'archived', label: 'Archived' },
            ]}
            {...form.register('status')}
          />
        </div>
        <TextArea
          label="Description"
          rows={3}
          error={e.description?.message}
          {...form.register('description')}
        />
        <div className="form-section">
          <div className="form-section-title">People</div>
          <Controller
            control={form.control}
            name="lead_user_id"
            render={({ field, fieldState }) => (
              <Picker
                label="Project lead"
                multiple={false}
                options={userOptions(lookups.data)}
                value={field.value ? [field.value] : []}
                onChange={(v) => field.onChange(v[0] ?? null)}
                placeholder="Choose a lead"
                error={fieldState.error?.message}
                id="project-lead"
              />
            )}
          />
          <Controller
            control={form.control}
            name="faculty_advisor_user_id"
            render={({ field }) => (
              <Picker
                label="Faculty advisor"
                multiple={false}
                options={userOptions(lookups.data)}
                value={field.value ? [field.value] : []}
                onChange={(v) => field.onChange(v[0] ?? null)}
                placeholder="Optional"
              />
            )}
          />
        </div>
        <div className="form-section">
          <div className="form-section-title">Season</div>
          <div className="field-row">
            <TextField
              label="Competition"
              placeholder="e.g. Lunar Crater Challenge 2027"
              {...form.register('competition')}
            />
            <TextField
              label="Academic year"
              placeholder="2026-27"
              error={e.academic_year?.message}
              {...form.register('academic_year')}
            />
          </div>
          <div className="field-row">
            <TextField label="Start date" type="date" {...form.register('start_date')} />
            <TextField label="Target date" type="date" {...form.register('target_date')} />
          </div>
          <SelectField
            label="Risk level"
            options={[
              { value: 'low', label: 'Low' },
              { value: 'medium', label: 'Medium' },
              { value: 'high', label: 'High' },
            ]}
            {...form.register('risk_level')}
          />
        </div>
        <div className="form-section">
          <div className="form-section-title">Budget &amp; links</div>
          <div className="field-row">
            <TextField
              label="Budget (USD)"
              type="number"
              min={0}
              step="0.01"
              error={e.budget_amount?.message}
              {...form.register('budget_amount', { valueAsNumber: true })}
            />
            <TextField label="Budget code" {...form.register('budget_code')} />
          </div>
          <TextField
            label="Repository URL"
            type="url"
            placeholder="https://github.com/…"
            error={e.repository_url?.message}
            {...form.register('repository_url')}
          />
          <TextField label="CAD URL" type="url" error={e.cad_url?.message} {...form.register('cad_url')} />
          <TextField
            label="Requirements URL"
            type="url"
            error={e.requirements_url?.message}
            {...form.register('requirements_url')}
          />
          <SelectField
            label="Visibility"
            options={[
              { value: 'private', label: 'Private: project members only' },
              { value: 'members', label: 'Members: everyone in the chapter' },
              { value: 'public', label: 'Public: linked from the public site' },
            ]}
            {...form.register('visibility')}
          />
        </div>
      </form>
    </SideSheet>
  );
}
