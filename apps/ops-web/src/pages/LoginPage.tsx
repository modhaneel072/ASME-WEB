import { useState } from 'react';
import { useForm } from 'react-hook-form';
import { zodResolver } from '@hookform/resolvers/zod';
import { Navigate, useLocation, useNavigate } from 'react-router-dom';
import { useQueryClient } from '@tanstack/react-query';
import { LogIn } from 'lucide-react';
import { api, ApiError } from '@/api/client';
import { keys, useSession } from '@/api/hooks';
import { loginSchema, type LoginInput } from '@/contracts/schemas';
import type { Session } from '@/contracts/types';
import { Button, TextField } from '@/ui';

export function LoginPage() {
  const navigate = useNavigate();
  const location = useLocation();
  const client = useQueryClient();
  const existing = useSession({ retry: false });
  const [banner, setBanner] = useState<string | null>(null);
  const next = new URLSearchParams(location.search).get('next') || '/work-orders';
  const reason = (location.state as { reason?: string } | null)?.reason;

  const form = useForm<LoginInput>({
    resolver: zodResolver(loginSchema),
    defaultValues: { identifier: '', password: '' },
  });

  if (existing.data) return <Navigate to={next} replace />;

  const onSubmit = form.handleSubmit(async (values) => {
    setBanner(null);
    try {
      const session = await api<Session>('/ops/session/login', { method: 'POST', body: values });
      client.setQueryData(keys.session, session);
      navigate(next, { replace: true });
    } catch (err) {
      if (err instanceof ApiError) {
        if (err.code === 'rate_limited') setBanner(err.message);
        else if (err.code === 'membership_inactive')
          setBanner('Your membership is not active. Contact a chapter officer.');
        else if (err.details.length)
          Object.entries(err.fieldErrors()).forEach(([field, message]) =>
            form.setError(field as keyof LoginInput, { message }),
          );
        else setBanner(err.message || 'Sign-in failed.');
      } else setBanner('Sign-in failed. Check your connection and try again.');
    }
  });

  return (
    <div className="login-page">
      <section className="login-hero" aria-hidden="true">
        <div className="row" style={{ gap: 12 }}>
          <span className="sidebar-brand-mark">UI</span>
          <span style={{ fontWeight: 700 }}>ASME Ops · University of Iowa</span>
        </div>
        <div className="stack" style={{ gap: 20 }}>
          <h1>
            Every build, inspection and part <span className="gold">in one place.</span>
          </h1>
          <p>
            Work orders, projects, assets and teams for the chapter. Built for the Crater Cruncher Rover
            season and everything after it.
          </p>
        </div>
        <p style={{ fontSize: 13 }}>Members sign in with the same account as the public site.</p>
      </section>
      <section className="login-panel">
        <form className="login-card" onSubmit={onSubmit} noValidate data-testid="login-form">
          <div>
            <h2>Sign in</h2>
            <p className="text-muted" style={{ marginTop: 4 }}>
              Use your chapter email or username.
            </p>
          </div>
          {reason === 'membership_inactive' && (
            <div className="form-banner form-banner-info">Your membership is not active in ASME Ops yet.</div>
          )}
          {banner && (
            <div className="form-banner form-banner-error" role="alert" data-testid="login-error">
              {banner}
            </div>
          )}
          <TextField
            label="Email or username"
            autoComplete="username"
            required
            error={form.formState.errors.identifier?.message}
            {...form.register('identifier')}
            data-testid="login-identifier"
            autoFocus
          />
          <TextField
            label="Password"
            type="password"
            autoComplete="current-password"
            required
            error={form.formState.errors.password?.message}
            {...form.register('password')}
            data-testid="login-password"
          />
          <Button
            type="submit"
            variant="primary"
            icon={<LogIn />}
            loading={form.formState.isSubmitting}
            block
            data-testid="login-submit"
          >
            Sign in
          </Button>
          <p className="text-caption text-muted">
            Forgot your password? Ask a chapter admin to reset it from Teams &amp; Users.{' '}
            <a href="/">Back to the public site</a>.
          </p>
        </form>
      </section>
    </div>
  );
}
