import { QueryClient, QueryClientProvider } from '@tanstack/react-query';
import { BrowserRouter, Navigate, Route, Routes, useParams } from 'react-router-dom';
import { ApiError } from '@/api/client';
import { RequireAuth, RequirePermission } from '@/auth/SessionProvider';
import { AppShell } from '@/shell/AppShell';
import { ToastProvider } from '@/ui';
import { LoginPage } from '@/pages/LoginPage';
import { SetupCenterPage } from '@/pages/SetupCenterPage';
import { DashboardPage } from '@/pages/DashboardPage';
import { WorkOrdersPage } from '@/pages/work-orders/WorkOrdersPage';
import { ProjectsPage } from '@/pages/projects/ProjectsPage';
import { ProjectDetailPage } from '@/pages/projects/ProjectDetailPage';
import { TeamsUsersPage } from '@/pages/TeamsUsersPage';
import { LocationsPage } from '@/pages/LocationsPage';
import { CategoriesPage } from '@/pages/CategoriesPage';
import { AssetsPage } from '@/pages/AssetsPage';
import { SettingsPage } from '@/pages/SettingsPage';
import { NotFoundPage } from '@/pages/NotFoundPage';

const queryClient = new QueryClient({
  defaultOptions: {
    queries: {
      retry: (count, error) =>
        !(error instanceof ApiError && error.status >= 400 && error.status < 500) && count < 2,
      staleTime: 10_000,
      refetchOnWindowFocus: false,
    },
  },
});

function WorkOrderRedirect() {
  const { id } = useParams();
  return <Navigate to={`/work-orders?wo=${id}`} replace />;
}

export function App() {
  return (
    <QueryClientProvider client={queryClient}>
      <ToastProvider>
        <BrowserRouter basename="/app">
          <Routes>
            <Route path="/login" element={<LoginPage />} />
            <Route
              element={
                <RequireAuth>
                  <AppShell />
                </RequireAuth>
              }
            >
              <Route index element={<Navigate to="/work-orders" replace />} />
              <Route
                path="/setup"
                element={
                  <RequirePermission permission="setup.manage">
                    <SetupCenterPage />
                  </RequirePermission>
                }
              />
              <Route
                path="/reporting"
                element={
                  <RequirePermission permission="report.view">
                    <DashboardPage />
                  </RequirePermission>
                }
              />
              <Route path="/work-orders" element={<WorkOrdersPage />} />
              <Route path="/work-orders/:id" element={<WorkOrderRedirect />} />
              <Route
                path="/projects"
                element={
                  <RequirePermission permission="project.read">
                    <ProjectsPage />
                  </RequirePermission>
                }
              />
              <Route
                path="/projects/:id/:tab?"
                element={
                  <RequirePermission permission="project.read">
                    <ProjectDetailPage />
                  </RequirePermission>
                }
              />
              <Route
                path="/assets"
                element={
                  <RequirePermission permission="asset.read">
                    <AssetsPage />
                  </RequirePermission>
                }
              />
              <Route path="/locations" element={<LocationsPage />} />
              <Route path="/categories" element={<CategoriesPage />} />
              <Route
                path="/teams-users/:tab?"
                element={
                  <RequirePermission permission="team.read">
                    <TeamsUsersPage />
                  </RequirePermission>
                }
              />
              <Route
                path="/settings/:tab?"
                element={
                  <RequirePermission permission={['org.read', 'org.manage', 'audit.read']}>
                    <SettingsPage />
                  </RequirePermission>
                }
              />
              <Route path="*" element={<NotFoundPage />} />
            </Route>
          </Routes>
        </BrowserRouter>
      </ToastProvider>
    </QueryClientProvider>
  );
}
