import { NextResponse } from "next/server";
import { getToken } from "next-auth/jwt";

function clean(s) { return String(s || "").replace(/^﻿/, "").trim(); }

function authEnabled() {
  return clean(process.env.AUTH_ENABLED).toLowerCase() === "true"
    && clean(process.env.GOOGLE_CLIENT_ID)
    && clean(process.env.GOOGLE_CLIENT_SECRET)
    && clean(process.env.AUTHORIZED_EMAILS);
}

const PUBLICOS = ["/login", "/api/auth"];

export async function middleware(req) {
  if (!authEnabled()) return NextResponse.next();

  const { pathname } = req.nextUrl;
  if (PUBLICOS.some((p) => pathname === p || pathname.startsWith(p + "/"))) {
    return NextResponse.next();
  }

  const token = await getToken({ req, secret: clean(process.env.NEXTAUTH_SECRET) });
  if (token?.email) return NextResponse.next();

  if (pathname.startsWith("/api/")) {
    return NextResponse.json({ error: "No autenticado" }, { status: 401 });
  }
  const url = req.nextUrl.clone();
  url.pathname = "/login";
  url.search = "";
  return NextResponse.redirect(url);
}

export const config = {
  matcher: ["/((?!_next/static|_next/image|favicon.ico).*)"],
};
