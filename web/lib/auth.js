import GoogleProvider from "next-auth/providers/google";

function clean(s) { return String(s || "").replace(/^﻿/, "").trim(); }

// Whitelist en formato "email1:Nombre,email2:Nombre". Si está vacía o falta GOOGLE_CLIENT_ID,
// la auth queda deshabilitada (ver middleware.js).
export function parseWhitelist() {
  const raw = clean(process.env.AUTHORIZED_EMAILS);
  if (!raw) return [];
  return raw.split(",")
    .map((s) => s.trim()).filter(Boolean)
    .map((entry) => {
      const idx = entry.indexOf(":");
      const email = (idx === -1 ? entry : entry.slice(0, idx)).toLowerCase().trim();
      const nombre = idx === -1 ? "" : entry.slice(idx + 1).trim();
      return { email, nombre };
    })
    .filter((e) => e.email);
}

export function authEnabled() {
  return clean(process.env.AUTH_ENABLED).toLowerCase() === "true"
    && clean(process.env.GOOGLE_CLIENT_ID)
    && clean(process.env.GOOGLE_CLIENT_SECRET);
}

export const authOptions = {
  providers: [
    GoogleProvider({
      clientId: clean(process.env.GOOGLE_CLIENT_ID),
      clientSecret: clean(process.env.GOOGLE_CLIENT_SECRET),
    }),
  ],
  secret: clean(process.env.NEXTAUTH_SECRET),
  pages: { signIn: "/login" },
  callbacks: {
    async signIn({ user }) {
      const wl = parseWhitelist();
      if (!wl.length) return false;
      const email = String(user?.email || "").toLowerCase();
      return wl.some((a) => a.email === email);
    },
    async jwt({ token, user }) {
      if (user?.email) token.email = String(user.email).toLowerCase();
      return token;
    },
    async session({ session, token }) {
      const email = String(token?.email || session.user?.email || "").toLowerCase();
      const match = parseWhitelist().find((a) => a.email === email);
      if (session.user) {
        session.user.email = email;
        session.user.nombreFamilia = match?.nombre || null;
      }
      return session;
    },
  },
};
