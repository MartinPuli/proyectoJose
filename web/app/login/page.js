"use client";
import { signIn, useSession } from "next-auth/react";
import { Suspense, useEffect, useState } from "react";
import { useRouter, useSearchParams } from "next/navigation";

function LoginInner() {
  const { status } = useSession();
  const router = useRouter();
  const params = useSearchParams();
  const error = params.get("error");
  const [yendo, setYendo] = useState(false);

  useEffect(() => {
    if (status === "authenticated") router.replace("/");
  }, [status, router]);

  async function entrar() {
    setYendo(true);
    await signIn("google", { callbackUrl: "/" });
  }

  return (
    <div className="login-card">
      <p className="eyebrow">Finanzas Familia</p>
      <h1>Entrar</h1>
      <p className="login-sub">Solo los emails autorizados pueden ver y modificar la planilla.</p>
      <button type="button" className="btn-google" onClick={entrar} disabled={yendo || status === "loading"}>
        <svg width="18" height="18" viewBox="0 0 18 18" aria-hidden="true">
          <path d="M17.64 9.2c0-.64-.06-1.25-.16-1.84H9v3.48h4.84c-.21 1.13-.84 2.09-1.79 2.73v2.27h2.9c1.69-1.56 2.69-3.86 2.69-6.64z" fill="#4285F4"/>
          <path d="M9 18c2.43 0 4.46-.81 5.95-2.18l-2.9-2.27c-.81.54-1.83.86-3.05.86-2.34 0-4.33-1.58-5.04-3.71H.96v2.34A9 9 0 0 0 9 18z" fill="#34A853"/>
          <path d="M3.96 10.7A5.41 5.41 0 0 1 3.68 9c0-.59.1-1.16.28-1.7V4.96H.96A9 9 0 0 0 0 9c0 1.45.35 2.83.96 4.04l3-2.34z" fill="#FBBC05"/>
          <path d="M9 3.58c1.32 0 2.5.45 3.44 1.35l2.58-2.58A9 9 0 0 0 9 0 9 9 0 0 0 .96 4.96l3 2.34C4.67 5.16 6.66 3.58 9 3.58z" fill="#EA4335"/>
        </svg>
        {yendo ? "Conectando…" : "Continuar con Google"}
      </button>
      {error && (
        <p className="login-error">
          {error === "AccessDenied"
            ? "Ese email no está autorizado. Pedile a José que te agregue a la whitelist."
            : "No pudimos completar el login. Probá de nuevo."}
        </p>
      )}
    </div>
  );
}

export default function LoginPage() {
  return (
    <div className="login-wrap">
      <Suspense fallback={<div className="login-card"><p className="eyebrow">Finanzas Familia</p><h1>Entrar</h1></div>}>
        <LoginInner />
      </Suspense>
    </div>
  );
}
