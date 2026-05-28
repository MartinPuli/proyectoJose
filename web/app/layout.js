import "./globals.css";

export const metadata = {
  title: "Finanzas Familia · Libro de cuentas",
  description: "Carga por audio, texto o documento con IA",
};

export const viewport = {
  themeColor: "#f6f5f1",
  width: "device-width",
  initialScale: 1,
};

// Aplica el tema antes del primer paint para evitar el flash claro→oscuro.
const setearTema = `(function(){try{var t=localStorage.getItem('tema');if(!t){t=window.matchMedia('(prefers-color-scheme: dark)').matches?'dark':'light';}document.documentElement.setAttribute('data-theme',t);}catch(e){}})();`;

export default function RootLayout({ children }) {
  return (
    <html lang="es">
      <head>
        <script dangerouslySetInnerHTML={{ __html: setearTema }} />
        <link rel="preconnect" href="https://fonts.googleapis.com" />
        <link rel="preconnect" href="https://fonts.gstatic.com" crossOrigin="anonymous" />
        <link
          href="https://fonts.googleapis.com/css2?family=Schibsted+Grotesk:wght@400;500;600;700;800&family=Hanken+Grotesk:wght@400;500;600;700&display=swap"
          rel="stylesheet"
        />
      </head>
      <body>{children}</body>
    </html>
  );
}
