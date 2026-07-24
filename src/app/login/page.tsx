import { login } from "./actions";

export default async function LoginPage({
  searchParams,
}: {
  searchParams: Promise<{ error?: string }>;
}) {
  const { error } = await searchParams;

  return (
    <div className="relative flex min-h-screen items-center justify-center bg-slate-50 selection:bg-blue-600/20">
      {/* Decoración de fondo limpia y corporativa (patrón sutil) */}
      <div className="pointer-events-none absolute inset-0 z-0 bg-[linear-gradient(to_right,#80808012_1px,transparent_1px),linear-gradient(to_bottom,#80808012_1px,transparent_1px)] bg-[size:24px_24px]"></div>
      <div className="absolute left-0 right-0 top-0 -z-10 m-auto h-[400px] w-[400px] rounded-full bg-blue-500 opacity-[0.08] blur-[100px]"></div>

      <div className="relative z-10 w-full max-w-[420px] p-6">
        <form
          action={login}
          className="flex flex-col gap-6 rounded-2xl border border-slate-200 bg-white p-8 shadow-xl shadow-slate-200/50 transition-all"
        >
          <div className="mb-2 flex flex-col items-center text-center">
            <div className="mb-6 flex h-14 w-14 items-center justify-center rounded-xl bg-blue-600 shadow-md shadow-blue-600/20">
              {/* Ícono de empresa / banco */}
              <svg xmlns="http://www.w3.org/2000/svg" className="h-7 w-7 text-white" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round">
                <path d="M22 12h-4l-3 9L9 3l-3 9H2" />
              </svg>
            </div>
            <h1 className="text-2xl font-bold tracking-tight text-slate-900">
              Bienvenido
            </h1>
            <p className="mt-1 text-sm text-slate-500">
              Conciliación diaria de bancos
            </p>
          </div>

          {error && (
            <div className="animate-in fade-in zoom-in-95 rounded-lg border border-red-200 bg-red-50 p-4">
              <p className="text-center text-sm font-medium text-red-600">
                {error}
              </p>
            </div>
          )}

          <div className="space-y-5">
            <div className="group relative">
              <label className="mb-1.5 block text-xs font-semibold uppercase tracking-wider text-slate-500 transition-colors group-focus-within:text-blue-600">
                Correo Electrónico
              </label>
              <input
                name="email"
                type="email"
                required
                autoComplete="email"
                placeholder="tu@correo.com"
                className="w-full rounded-lg border border-slate-300 bg-white px-4 py-2.5 text-sm text-slate-900 placeholder:text-slate-400 outline-none transition-all hover:border-slate-400 focus:border-blue-600 focus:ring-4 focus:ring-blue-600/10"
              />
            </div>

            <div className="group relative">
              <label className="mb-1.5 block text-xs font-semibold uppercase tracking-wider text-slate-500 transition-colors group-focus-within:text-blue-600">
                Contraseña
              </label>
              <input
                name="password"
                type="password"
                required
                autoComplete="current-password"
                placeholder="••••••••"
                className="w-full rounded-lg border border-slate-300 bg-white px-4 py-2.5 text-sm text-slate-900 placeholder:text-slate-400 outline-none transition-all hover:border-slate-400 focus:border-blue-600 focus:ring-4 focus:ring-blue-600/10"
              />
            </div>
          </div>

          <button
            type="submit"
            className="group mt-4 flex h-11 w-full items-center justify-center gap-2 rounded-lg bg-blue-600 font-medium text-white transition-all hover:bg-blue-700 focus:outline-none focus:ring-4 focus:ring-blue-600/20 active:scale-[0.98]"
          >
            Entrar
            <svg className="h-4 w-4 transition-transform group-hover:translate-x-1" xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
              <line x1="5" y1="12" x2="19" y2="12"></line>
              <polyline points="12 5 19 12 12 19"></polyline>
            </svg>
          </button>
        </form>
        
        <p className="mt-8 text-center text-xs font-medium text-slate-500">
          © {new Date().getFullYear()} Firplak S.A. Todos los derechos reservados.
        </p>
      </div>
    </div>
  );
}
